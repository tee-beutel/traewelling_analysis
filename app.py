import os
import re
import time
import sqlite3
import requests
import uuid
import threading
import json
import shutil
from datetime import datetime
from flask import Flask, request, session, redirect, url_for, render_template, send_from_directory, abort, jsonify, render_template_string
import traewelling_analysis as ta

import secrets
from dotenv import load_dotenv, set_key

app = Flask(__name__)

ENV_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env')
if not os.path.exists(ENV_PATH):
    with open(ENV_PATH, 'w', encoding='utf-8') as f:
        f.write(f"FLASK_SECRET_KEY={secrets.token_hex(32)}\n")
        f.write("CLIENT_ID=\n")
        f.write("CLIENT_SECRET=\n")
        f.write("REDIRECT_URI=\n")
        f.write("TRAEWELLING_API_TOKEN=\n")

load_dotenv(ENV_PATH)
flask_secret = os.getenv('FLASK_SECRET_KEY')
if not flask_secret:
    flask_secret = secrets.token_hex(32)
    set_key(ENV_PATH, 'FLASK_SECRET_KEY', flask_secret)
    os.environ['FLASK_SECRET_KEY'] = flask_secret

app.secret_key = flask_secret

def get_env_val(key, default=''):
    return os.getenv(key, default).strip()

def save_env_val(key, value):
    set_key(ENV_PATH, key, value)
    os.environ[key] = value

def get_oauth_config():
    return (
        get_env_val('CLIENT_ID'),
        get_env_val('CLIENT_SECRET'),
        get_env_val('REDIRECT_URI')
    )

AUTH_URL = 'https://traewelling.de/oauth/authorize'
TOKEN_URL = 'https://traewelling.de/oauth/token'

# ==========================================
# THREADING & BACKGROUND TASK MANAGER
# ==========================================
TASKS = {}
WAITING_QUEUE = []
queue_lock = threading.Lock()
generation_lock = threading.Lock()

class StopTaskException(Exception):
    """Spezifische Exception um einen laufenden Thread sicher abzubrechen, ohne die DB zu zerstören."""
    pass

def cleanup_old_files():
    """Löscht alle Exporte in ALLEN Ordnern, die älter als 24 Stunden sind."""
    base_dir = os.path.join(os.getcwd(), 'finished_exports')
    if not os.path.exists(base_dir): return
    now = time.time()
    for root, dirs, files in os.walk(base_dir):
        for file in files:
            filepath = os.path.join(root, file)
            try:
                if now - os.path.getmtime(filepath) > 86400:
                    os.remove(filepath)
            except OSError:
                pass

def move_to_user_dir(creator_username, filename, folder_type):
    """Verschiebt generierte Dateien in den privaten Ordner des Erstellers."""
    src = os.path.join(os.getcwd(), 'finished_exports', folder_type, filename)
    dest_dir = os.path.join(os.getcwd(), 'finished_exports', creator_username, folder_type)
    
    if os.path.exists(src):
        os.makedirs(dest_dir, exist_ok=True)
        dest = os.path.join(dest_dir, filename)
        if os.path.exists(dest):
            os.remove(dest) # Überschreiben
        shutil.move(src, dest)

# ==========================================
# HINTERGRUND-WORKER
# ==========================================
def run_generation_task(task_id, creator, api_token, selected_users, start_date, end_date, options):
    def check_stop():
        """Wirft eine Exception, wenn das Frontend einen Stopp angefordert hat."""
        if TASKS[task_id].get('stop_requested', False):
            raise StopTaskException("⚠️ Auswertung wurde vom Benutzer gestoppt.")

    def log(msg, replace=False, ignore_stop=False):
        if not ignore_stop:
            check_stop()
        if task_id not in TASKS: return
        if replace and len(TASKS[task_id]['logs']) > 0:
            TASKS[task_id]['logs'][-1] = msg
        else:
            TASKS[task_id]['logs'].append(msg)

    def add_live_file(filename, folder_type):
        if folder_type == 'user': group_name = 'Excel-Fahrtenbücher'
        elif folder_type == 'gang_linien': group_name = 'Tageszeiten (Ganglinien)'
        elif folder_type == 'verspaetungs_plots': group_name = 'Verspätungs-Verteilungen'
        elif folder_type == 'shared': group_name = 'Team-Auswertungen'
        elif folder_type == 'timelines':
            group_name = 'Timelines'
        elif folder_type == 'polylines':
            if 'stops_map' in filename:
                group_name = 'Haltestellenkarten'
            else:
                group_name = 'Streckenkarten'
        else:
            group_name = folder_type.capitalize()
            
        if task_id in TASKS:
            TASKS[task_id]['new_files'].append({
                "name": filename,
                "folder": folder_type,
                "group_name": group_name
            })

    with queue_lock:
        WAITING_QUEUE.append(task_id)

    while True:
        with queue_lock:
            if WAITING_QUEUE and WAITING_QUEUE[0] == task_id:
                break
            pos = WAITING_QUEUE.index(task_id) if task_id in WAITING_QUEUE else 0
            total_waiting = len(WAITING_QUEUE) - 1
        
        if TASKS[task_id].get('stop_requested', False):
            with queue_lock:
                if task_id in WAITING_QUEUE: WAITING_QUEUE.remove(task_id)
            TASKS[task_id]['status'] = 'stopped'
            TASKS[task_id]['logs'].append("⚠️ Auswertung in der Warteschlange gestoppt.")
            return

        log(f"⏳ Warteschlange: Platz {pos} von {total_waiting} wartenden Prozessen...", replace=True)
        time.sleep(1)

    with generation_lock:
        try:
            log('🚀 Freie Server-Kapazitäten erhalten! Starte Initialisierung...')
            conn = ta.init_db()

            ta.AUTH_USER = creator
            ta.TRUSTED_USERS = ta.fetch_paginated_users(api_token, 'https://traewelling.de/api/v1/user/self/trusted-by')
            ta.FOLLOWED_USERS = ta.fetch_paginated_users(api_token, 'https://traewelling.de/api/v1/user/self/followings')
            ta.FOLLOWERS_USERS = ta.fetch_paginated_users(api_token, 'https://traewelling.de/api/v1/user/self/followers')

            total = len(selected_users)

            needs_individual_loop = (
                options.get('update_db', True) or 
                options.get('gen_excel') or 
                options.get('gen_time') or 
                options.get('gen_delay') or 
                (len(selected_users) == 1 and (options.get('gen_map') or options.get('gen_stop_map') or options.get('gen_timeline')))
            )

            if needs_individual_loop:
                for idx, user in enumerate(selected_users):
                    log(f'\n--- 👤 Platz {idx+1} von {total}: Verarbeite Account {user} ---')
                    
                    if options.get('update_db', True):
                        log(f'Bereite Daten für {user} vor...')
                    else:
                        log(f'Lade bestehende Daten für {user} aus der Datenbank...')

                    ta.fetch_data_to_db(user, conn=conn, api_token=api_token, reload_all=False, log_callback=log, update_db=options.get('update_db', True))

                    if ta.check_for_user_journeys(conn, user, start_date=start_date, end_date=end_date):

                        if options['gen_excel']:
                            check_stop()
                            log(f'Erstelle Excel-Fahrtenbuch für {user}...')
                            excel_filename = ta.create_user_excel_file(conn, user, start_date=start_date, end_date=end_date, create_picture=options['gen_time'])
                            check_stop()
                            move_to_user_dir(creator, excel_filename, "user")
                            add_live_file(excel_filename, "user")
                            log(f'   📄 {excel_filename} steht jetzt zum Download bereit!')

                        if options['gen_time']:
                            check_stop()
                            log(f'Erstelle Tageszeiten-Grafiken für {user}...')
                            for grouping in ['time_of_day', 'weekday', 'month', 'month_year', 'date', 'year']:
                                img_filename = ta.create_ganglinie_plot(username=user, grouping_type=grouping, conn=conn, start_date=start_date, end_date=end_date)
                                if img_filename:
                                    check_stop()
                                    move_to_user_dir(creator, img_filename, "gang_linien")
                                    add_live_file(img_filename, "gang_linien")

                        if options['gen_map'] and len(selected_users) == 1:
                            check_stop()
                            log(f'Erstelle interaktive Streckenkarte für {user}...')
                            map_filename = ta.create_polyline(conn, api_token, usernames=[user], start_date=start_date, end_date=end_date, log_callback=log, lite_version=options['gen_map_lite'])
                            if map_filename:
                                check_stop()
                                move_to_user_dir(creator, map_filename, "polylines")
                                add_live_file(map_filename, "polylines")
                                log(f'   🗺️ {map_filename} steht jetzt zur Ansicht bereit!')
                                
                        if options['gen_stop_map'] and len(selected_users) == 1:
                            check_stop()
                            log(f'Erstelle interaktive Haltestellenkarte für {user}...')
                            stop_map_filename = ta.create_stops_map(conn, usernames=[user], start_date=start_date, end_date=end_date, log_callback=log, lite_version=options['gen_stop_map_lite'])
                            if stop_map_filename:
                                check_stop()
                                move_to_user_dir(creator, stop_map_filename, "polylines")
                                add_live_file(stop_map_filename, "polylines")
                                log(f'   📍 {stop_map_filename} steht jetzt zur Ansicht bereit!')

                        if options.get('gen_timeline') and len(selected_users) == 1:
                            check_stop()
                            log(f'Erstelle interaktive Timeline für {user}...')
                            timeline_filename = ta.create_timeline(conn, usernames=[user], start_date=start_date, end_date=end_date, log_callback=log)
                            if timeline_filename:
                                check_stop()
                                move_to_user_dir(creator, timeline_filename, "timelines")
                                add_live_file(timeline_filename, "timelines")
                                log(f'   ⏱️ {timeline_filename} steht jetzt zur Ansicht bereit!')

                        if options['gen_delay']:
                            check_stop()
                            log(f'Erstelle Verspätungs-Graphen für {user}...')
                            delay_filenames = ta.create_verspaetungs_verteilungen_plot(user, conn=conn, start_date=start_date, end_date=end_date)
                            for d_file in (delay_filenames or []):
                                check_stop()
                                move_to_user_dir(creator, d_file, "verspaetungs_plots")
                                add_live_file(d_file, "verspaetungs_plots")
                    else:
                        log(f'⚠️ Keine passenden Fahrten für {user} im angegebenen Zeitraum gefunden.')
            else:
                log(f'\n[INFO] Datenbank-Update übersprungen. Nutze direkt bestehende Daten für alle {total} User.')

            if len(selected_users) > 1:
                log('\n--- 👥 Starte Team-Auswertung ---')

                if options.get('gen_shared_excel'):
                    check_stop()
                    log('Erstelle gemeinsames Excel-Dashboard...')
                    shared_excel = ta.create_shared_excel_file(conn=conn, usernames=selected_users, start_date=start_date, end_date=end_date)
                    if shared_excel:
                        check_stop()
                        move_to_user_dir(creator, shared_excel, "shared")
                        add_live_file(shared_excel, "shared")
                        log(f'   📊 {shared_excel} steht jetzt zum Download bereit!')

                if options['gen_map']:
                    check_stop()
                    log('Erstelle gemeinsame interaktive Streckenkarte...')
                    shared_map = ta.create_polyline(conn, api_token, usernames=selected_users, start_date=start_date, end_date=end_date, log_callback=log, lite_version=options['gen_map_lite'])
                    if shared_map:
                        check_stop()
                        move_to_user_dir(creator, shared_map, "polylines")
                        add_live_file(shared_map, "polylines")
                        log(f'   🗺️ {shared_map} steht jetzt zur Ansicht bereit!')
                        
                if options['gen_stop_map']:
                    check_stop()
                    log('Erstelle gemeinsame interaktive Haltestellenkarte...')
                    shared_stop_map = ta.create_stops_map(conn, usernames=selected_users, start_date=start_date, end_date=end_date, log_callback=log, lite_version=options['gen_stop_map_lite'])
                    if shared_stop_map:
                        check_stop()
                        move_to_user_dir(creator, shared_stop_map, "polylines")
                        add_live_file(shared_stop_map, "polylines")
                        log(f'   📍 {shared_stop_map} steht jetzt zur Ansicht bereit!')

                if options.get('gen_timeline'):
                    check_stop()
                    log('Erstelle gemeinsame interaktive Timeline...')
                    timeline_filename = ta.create_timeline(conn, usernames=selected_users, start_date=start_date, end_date=end_date, log_callback=log)
                    if timeline_filename:
                        check_stop()
                        move_to_user_dir(creator, timeline_filename, "timelines")
                        add_live_file(timeline_filename, "timelines")
                        log(f'   ⏱️ {timeline_filename} steht jetzt zur Ansicht bereit!')

            log('\n✅ Alle Auswertungen erfolgreich abgeschlossen!')
            TASKS[task_id]['status'] = 'completed'

        except StopTaskException as e:
            log(f'\n{str(e)}', ignore_stop=True)
            TASKS[task_id]['status'] = 'stopped'
        except Exception as e:
            log(f'\n❌ KRITISCHER FEHLER: {str(e)}', ignore_stop=True)
            TASKS[task_id]['status'] = 'error'
        finally:
            with queue_lock:
                if task_id in WAITING_QUEUE:
                    WAITING_QUEUE.remove(task_id)

def run_debug_task(task_id, username, fix_type):
    """Hintergrund-Task für automatische Reparaturen im Debug-Modus."""
    def log(msg):
        if task_id in TASKS:
            TASKS[task_id]['logs'].append(msg)

    conn = None
    try:
        conn = ta.init_db()
        cursor = conn.cursor()
        cursor.execute("SELECT access_token FROM user_login_data WHERE owner=?", (username,))
        row = cursor.fetchone()
        if not row:
            log("❌ Kein Benutzerkonto in user_login_data gefunden.")
            if task_id in TASKS:
                TASKS[task_id]['status'] = 'error'
            return

        api_token = row[0] if isinstance(row, tuple) else row['access_token']
        if not api_token:
            log("❌ Kein API-Token für Benutzer vorhanden.")
            if task_id in TASKS:
                TASKS[task_id]['status'] = 'error'
            return

        if fix_type == 'stops':
            log("🔍 Suche nach unvollständigen Haltestellen...")
            cursor.execute("""
                SELECT id, stopName FROM stops 
                WHERE latitude IS NULL OR longitude IS NULL OR countryCode = 'XX' OR countryName = 'Unbekannt'
            """)
            stops_to_fix = cursor.fetchall()
            total = len(stops_to_fix)
            log(f"📋 {total} unvollständige Haltestellen gefunden.")

            for idx, row_stop in enumerate(stops_to_fix, 1):
                if TASKS[task_id].get('stop_requested', False):
                    log("⚠️ Debug-Reparatur wurde vom Benutzer abgebrochen.")
                    TASKS[task_id]['status'] = 'stopped'
                    return
                stop_id = row_stop[0]
                stop_name = row_stop[1]
                if stop_id is not None:
                    cursor.execute("DELETE FROM stops WHERE id = ?", (stop_id,))
                    conn.commit()
                    log(f"[{idx}/{total}] Lade Haltestelle '{stop_name or stop_id}' (ID: {stop_id}) neu...")
                    ta.get_or_fetch_stop_info(conn, api_token, stop_id, stop_name)

            log("✅ Alle unvollständigen Haltestellen wurden nachgeladen.")

        elif fix_type == 'orphaned':
            log("🔍 Suche nach Fahrten mit fehlenden Haltestellen...")
            cursor.execute("""
                SELECT DISTINCT statusId FROM journeys 
                WHERE originId NOT IN (SELECT id FROM stops) 
                   OR destinationId NOT IN (SELECT id FROM stops)
            """)
            journeys_to_fix = [r[0] for r in cursor.fetchall()]
            total = len(journeys_to_fix)
            log(f"📋 {total} betroffene Fahrten gefunden.")

            for idx, status_id in enumerate(journeys_to_fix, 1):
                if TASKS[task_id].get('stop_requested', False):
                    log("⚠️ Debug-Reparatur wurde vom Benutzer abgebrochen.")
                    TASKS[task_id]['status'] = 'stopped'
                    return
                log(f"[{idx}/{total}] Lade Fahrt {status_id} neu...")
                ta.reload_single_status(conn, api_token, status_id)

            log("✅ Alle betroffenen Fahrten wurden neu geladen.")

        else:
            log(f"⚠️ Unbekannter fix_type: {fix_type}")

        if task_id in TASKS:
            TASKS[task_id]['status'] = 'completed'

    except Exception as e:
        log(f"❌ Fehler bei der Reparatur: {str(e)}")
        if task_id in TASKS:
            TASKS[task_id]['status'] = 'error'
    finally:
        if conn:
            try:
                conn.close()
            except Exception:
                pass


# ==========================================
# ROUTES
# ==========================================
@app.route('/')
def index():
    if 'username' in session: return redirect(url_for('dashboard'))
    cid, csecret, red_uri = get_oauth_config()
    has_oauth = bool(cid and csecret and red_uri)
    stored_token = get_env_val('TRAEWELLING_API_TOKEN')
    has_token = bool(stored_token)
    token_preview = ""
    if has_token:
        token_preview = stored_token[:12] + "..." + stored_token[-8:] if len(stored_token) > 24 else "***"

    return render_template('index.html',
                           has_oauth=has_oauth,
                           has_token=has_token,
                           client_id=cid,
                           redirect_uri=red_uri,
                           token_preview=token_preview,
                           error=request.args.get('error'))

@app.route('/login')
def login():
    cid, csecret, red_uri = get_oauth_config()
    if not (cid and csecret and red_uri):
        return redirect(url_for('index', error="Bitte hinterlege zuerst die OAuth-Zugangsdaten (Client-ID, Secret, Redirect-URI)."))
    scopes = "read-statuses+read-settings+read-settings-followers+read-search"
    return redirect(f"{AUTH_URL}?client_id={cid}&redirect_uri={red_uri}&response_type=code&scope={scopes}")

@app.route('/setup_oauth', methods=['POST'])
def setup_oauth():
    cid = request.form.get('client_id', '').strip()
    csecret = request.form.get('client_secret', '').strip()
    red_uri = request.form.get('redirect_uri', '').strip()
    if not (cid and csecret and red_uri):
        return redirect(url_for('index', error="Bitte alle OAuth-Felder vollständig ausfüllen."))

    save_env_val('CLIENT_ID', cid)
    save_env_val('CLIENT_SECRET', csecret)
    save_env_val('REDIRECT_URI', red_uri)
    return redirect(url_for('login'))

@app.route('/login_token', methods=['POST'])
def login_token():
    token = request.form.get('api_token', '').strip()
    if not token:
        token = get_env_val('TRAEWELLING_API_TOKEN')

    if not token:
        return redirect(url_for('index', error="Bitte gib einen API-Token ein."))

    try:
        user_resp = requests.get(
            'https://traewelling.de/api/v1/auth/user',
            headers={
                'Authorization': f'Bearer {token}',
                'User-Agent': 'TraewellingAnalysis_WebApp v9',
                'Accept': 'application/json'
            },
            timeout=15
        )
    except Exception as e:
        return redirect(url_for('index', error=f"Verbindungsfehler zur Träwelling-API: {e}"))

    if user_resp.status_code != 200:
        return redirect(url_for('index', error=f"Ungültiger API-Token (HTTP {user_resp.status_code})."))

    username = user_resp.json().get('data', {}).get('username')
    if not username:
        return redirect(url_for('index', error="Benutzername konnte nicht ermittelt werden."))

    # In .env speichern
    save_env_val('TRAEWELLING_API_TOKEN', token)

    trusted = ta.fetch_paginated_users(token, 'https://traewelling.de/api/v1/user/self/trusted-by')
    followed = ta.fetch_paginated_users(token, 'https://traewelling.de/api/v1/user/self/followings')
    followers = ta.fetch_paginated_users(token, 'https://traewelling.de/api/v1/user/self/followers')

    conn = ta.init_db()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO user_login_data (owner, access_token, followed_users, trusted_users, followers, last_updated) 
        VALUES (?, ?, ?, ?, ?, ?) 
        ON CONFLICT(owner) DO UPDATE SET 
            access_token = excluded.access_token,
            followed_users = excluded.followed_users,
            trusted_users = excluded.trusted_users,
            followers = excluded.followers,
            last_updated = excluded.last_updated
    """, (
        username, 
        token, 
        json.dumps(list(followed)), 
        json.dumps(list(trusted)), 
        json.dumps(list(followers)),
        datetime.now().isoformat()
    ))
    conn.commit()
    conn.close()

    session['username'] = username
    return redirect(url_for('dashboard'))

@app.route('/callback')
def callback():
    code = request.args.get('code')
    if not code: return "Fehler beim Login", 400

    cid, csecret, red_uri = get_oauth_config()
    token_data = {'grant_type': 'authorization_code', 'client_id': cid, 'client_secret': csecret, 'redirect_uri': red_uri, 'code': code}
    
    resp = requests.post(TOKEN_URL, data=token_data, headers={'User-Agent': 'TraewellingAnalysis_WebApp v9'})
    if resp.status_code != 200: return f"Token-Fehler: {resp.text}", 400

    token_json = resp.json()
    access_token = token_json.get('access_token')
    
    user_resp = requests.get(
        'https://traewelling.de/api/v1/auth/user', 
        headers={
            'Authorization': f'Bearer {access_token}',
            'User-Agent': 'TraewellingAnalysis_WebApp v9',
            'Accept': 'application/json'
        }
    )

    if user_resp.status_code != 200:
        return f"Fehler von der Träwelling-API: {user_resp.status_code} - {user_resp.text}", 400

    username = user_resp.json().get('data', {}).get('username')

    # In .env speichern
    save_env_val('TRAEWELLING_API_TOKEN', access_token)

    trusted = ta.fetch_paginated_users(access_token, 'https://traewelling.de/api/v1/user/self/trusted-by')
    followed = ta.fetch_paginated_users(access_token, 'https://traewelling.de/api/v1/user/self/followings')
    followers = ta.fetch_paginated_users(access_token, 'https://traewelling.de/api/v1/user/self/followers')

    conn = ta.init_db()
    cursor = conn.cursor()
    
    cursor.execute("""
        INSERT INTO user_login_data (owner, access_token, followed_users, trusted_users, followers, last_updated) 
        VALUES (?, ?, ?, ?, ?, ?) 
        ON CONFLICT(owner) DO UPDATE SET 
            access_token = excluded.access_token,
            followed_users = excluded.followed_users,
            trusted_users = excluded.trusted_users,
            followers = excluded.followers,
            last_updated = excluded.last_updated
    """, (
        username, 
        access_token, 
        json.dumps(list(followed)), 
        json.dumps(list(trusted)), 
        json.dumps(list(followers)),
        datetime.now().isoformat()
    ))
    conn.commit()
    conn.close()

    session['username'] = username
    return redirect(url_for('dashboard'))

@app.route('/logout')
def logout():
    session.pop('username', None)
    return redirect(url_for('index'))

@app.route('/dashboard')
def dashboard():
    if 'username' not in session: return redirect(url_for('login'))
    username = session['username']
    cleanup_old_files()
    
    conn = ta.init_db()
    cursor = conn.cursor()
    
    cursor.execute("SELECT followed_users, trusted_users, followers FROM user_login_data WHERE owner = ?", (username,))
    row = cursor.fetchone()
    
    available_users = set()
    if row:
        for item in row:
            if item:
                try:
                    parsed = json.loads(item)
                    if isinstance(parsed, list):
                        available_users.update(parsed)
                except Exception:
                    pass
    
    user_files = {}
    
    base_dir = os.path.join(os.getcwd(), 'finished_exports', username)
    if os.path.exists(base_dir):
        for root, dirs, files in os.walk(base_dir):
            for file in sorted(files):
                folder_type = os.path.basename(root)
                
                if folder_type == 'user': display_group = 'Excel-Fahrtenbücher'
                elif folder_type == 'gang_linien': display_group = 'Tageszeiten (Ganglinien)'
                elif folder_type == 'shared': display_group = 'Team-Auswertungen'
                elif folder_type == 'verspaetungs_plots': display_group = 'Verspätungs-Verteilungen'
                elif folder_type == 'timelines': display_group = 'Timelines'
                elif folder_type == 'polylines':
                    if 'stops_map' in file: display_group = 'Haltestellenkarten'
                    else: display_group = 'Streckenkarten'
                else:
                    display_group = folder_type.capitalize()

                if display_group not in user_files:
                    user_files[display_group] = []
                user_files[display_group].append({"name": file, "folder": folder_type})

    category_order = [
        'Excel-Fahrtenbücher',
        'Team-Auswertungen',
        'Timelines',
        'Streckenkarten',
        'Haltestellenkarten',
        'Tageszeiten (Ganglinien)',
        'Verspätungs-Verteilungen'
    ]

    sorted_user_files = {
        k: user_files[k]
        for k in sorted(user_files.keys(), key=lambda x: category_order.index(x) if x in category_order else 99)
    }

    return render_template(
        'dashboard.html', 
        username=username, 
        files=sorted_user_files, 
        available_users=sorted(list(available_users))
    )

@app.route('/start_generation', methods=['POST'])
def start_generation():
    if 'username' not in session: return jsonify({'error': 'Nicht eingeloggt'}), 401
    username = session['username']
    
    start_date = request.form.get('start_date', '').strip()
    end_date = request.form.get('end_date', '').strip()
    
    checked_users = request.form.getlist('users')
    manual_users_str = request.form.get('manual_users', '')
    
    selected_users = set(checked_users)
    for u in manual_users_str.split(','):
        u = u.strip()
        if u: selected_users.add(u)
    selected_users = list(selected_users)
    if not selected_users: selected_users = [username]

    options = {
        'update_db': request.form.get('update_db') == 'on',
        'gen_excel': request.form.get('gen_excel') == 'on',
        'gen_shared_excel': request.form.get('gen_shared_excel') == 'on',
        'gen_timeline': request.form.get('gen_timeline') == 'on',
        'gen_time': request.form.get('gen_time') == 'on',
        'gen_map': request.form.get('gen_map') == 'on',
        'gen_map_lite': request.form.get('gen_map_lite') == 'on',
        'gen_stop_map': request.form.get('gen_stop_map') == 'on',
        'gen_stop_map_lite': request.form.get('gen_stop_map_lite') == 'on',
        'gen_delay': request.form.get('gen_delay') == 'on'
    }

    conn = ta.init_db()
    cursor = conn.cursor()
    cursor.execute("SELECT access_token FROM user_login_data WHERE owner=?", (username,))
    row = cursor.fetchone()
    if not row or not row[0]: return jsonify({'error': 'Kein Token'}), 401
    
    task_id = str(uuid.uuid4())
    TASKS[task_id] = {
        'status': 'running', 
        'logs': ['Prozess in Warteschlange eingereiht...'],
        'new_files': [],
        'stop_requested': False,
        'creator': username,
        'type': 'Auswertung',
        'target_users': selected_users,
        'created_at': datetime.now().strftime('%H:%M:%S')
    }

    thread = threading.Thread(target=run_generation_task, args=(task_id, username, row[0], selected_users, start_date, end_date, options))
    thread.daemon = True
    thread.start()

    return jsonify({'task_id': task_id})

@app.route('/stop_task/<task_id>', methods=['POST'])
def stop_task(task_id):
    if 'username' not in session: return jsonify({'error': 'Nicht eingeloggt'}), 401
    if task_id in TASKS:
        TASKS[task_id]['stop_requested'] = True
        return jsonify({'success': True})
    return jsonify({'error': 'Task nicht gefunden'}), 404

@app.route('/delete_files', methods=['POST'])
def delete_files():
    if 'username' not in session: return jsonify({'error': 'Nicht eingeloggt'}), 401
    username = session['username']
    base_dir = os.path.join(os.getcwd(), 'finished_exports', username)
    if os.path.exists(base_dir):
        shutil.rmtree(base_dir)
    return jsonify({'success': True})

@app.route('/delete_category', methods=['POST'])
def delete_category():
    if 'username' not in session: return jsonify({'error': 'Nicht eingeloggt'}), 401
    username = session['username']
    category_to_delete = request.form.get('category')
    
    base_dir = os.path.join(os.getcwd(), 'finished_exports', username)
    if not os.path.exists(base_dir):
        return jsonify({'success': True})
        
    for root, dirs, files in os.walk(base_dir):
        folder_type = os.path.basename(root)
        for file in files:
            if folder_type == 'user': display_group = 'Excel-Fahrtenbücher'
            elif folder_type == 'gang_linien': display_group = 'Tageszeiten (Ganglinien)'
            elif folder_type == 'shared': display_group = 'Team-Auswertungen'
            elif folder_type == 'verspaetungs_plots': display_group = 'Verspätungs-Verteilungen'
            elif folder_type == 'timelines': display_group = 'Timelines'
            elif folder_type == 'polylines':
                if 'stops_map' in file: display_group = 'Haltestellenkarten'
                else: display_group = 'Streckenkarten'
            else:
                display_group = folder_type.capitalize()
                
            if display_group == category_to_delete:
                try:
                    os.remove(os.path.join(root, file))
                except OSError:
                    pass
                    
    # Leere Ordner aufräumen
    for root, dirs, files in os.walk(base_dir, topdown=False):
        if not os.listdir(root):
            try:
                os.rmdir(root)
            except OSError:
                pass
                
    return jsonify({'success': True})

@app.route('/task_status/<task_id>')
def task_status(task_id):
    return jsonify(TASKS.get(task_id, {'status': 'error', 'logs': ['Task nicht gefunden'], 'new_files': []}))

@app.route('/files/<action>/<folder>/<filename>')
def handle_file(action, folder, filename):
    if 'username' not in session: abort(401)
    username = session['username']
    
    safe_folder = os.path.basename(folder)
    safe_filename = os.path.basename(filename)
    
    directory = os.path.join(os.getcwd(), 'finished_exports', username, safe_folder)
    filepath = os.path.join(directory, safe_filename)
    
    if not os.path.exists(filepath): abort(404)
    return send_from_directory(directory, safe_filename, as_attachment=(action == 'download'))


# ==========================================
# ERWEITERTE DEBUG & DATENBANK-CLEANUP ROUTEN
# ==========================================



@app.route('/debug')
def debug_page():
    if 'username' not in session: return redirect(url_for('login'))

    conn = ta.init_db()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    # 1. Dummy-Haltestellen (Zählung & Liste ohne hartes Limit)
    cursor.execute("""
                   SELECT id, stopName, countryCode, latitude, longitude
                   FROM stops
                   WHERE latitude IS NULL
                      OR longitude IS NULL
                      OR countryCode = 'XX'
                      OR countryName = 'Unbekannt'
                   """)
    check1 = [dict(row) for row in cursor.fetchall()]

    # 2. Verwaiste Haltestellen-Referenzen (Zählung & Liste ohne hartes Limit)
    cursor.execute("""
                   SELECT DISTINCT j.statusId, j.username, j.originId, j.destinationId
                   FROM journeys j
                   WHERE j.originId NOT IN (SELECT id FROM stops)
                      OR j.destinationId NOT IN (SELECT id FROM stops)
                   """)
    check2 = [dict(row) for row in cursor.fetchall()]

    # 3. Fahrten mit 0 Distanz oder Zeit
    cursor.execute(
        "SELECT statusId, username, lineName, distance, durationReal FROM journeys WHERE distance = 0 OR durationReal <= 0")
    check3 = [dict(row) for row in cursor.fetchall()]

    # 4. Unplausible Geschwindigkeiten
    cursor.execute('''
                   SELECT statusId,
                          username,
                          lineName,
                          category,
                          distance,
                          durationReal,
                          ROUND((distance / 1000.0) / (durationReal / 60.0), 2) as speed_kmh
                   FROM journeys
                   WHERE durationReal > 0
                     AND (distance / 1000.0) / (durationReal / 60.0) > CASE
                                                                           WHEN category = 'Fernverkehr' THEN 350
                                                                           WHEN category = 'Regionalzug' THEN 250
                                                                           WHEN category = 'S-Bahn' THEN 160
                                                                           WHEN category = 'U-Bahn' THEN 120
                                                                           WHEN category = 'Tram' THEN 100
                                                                           WHEN category = 'Bus' THEN 130
                                                                           WHEN category = 'Fähre' THEN 100
                                                                           WHEN category = 'Taxi' THEN 200
                                                                           WHEN category = 'Igitt ein Flugzeug'
                                                                               THEN 1200
                                                                           ELSE 350
                       END
                   ''')
    check4 = [dict(row) for row in cursor.fetchall()]

    # 5. Zeitreisen (Ankunft vor Abfahrt)
    cursor.execute(
        "SELECT statusId, username, lineName, departureFinal, arrivalFinal FROM journeys WHERE departureFinal > arrivalFinal AND departureFinal IS NOT NULL AND arrivalFinal IS NOT NULL")
    check5 = [dict(row) for row in cursor.fetchall()]

    # 6. Extreme Verspätungen
    cursor.execute(
        "SELECT statusId, username, lineName, departureDelay, arrivalDelay FROM journeys WHERE arrivalDelay > 1200 OR arrivalDelay < -1200 OR departureDelay > 1200 OR departureDelay < -1200")
    check6 = [dict(row) for row in cursor.fetchall()]

    # 7. Ungültige Geo-Koordinaten
    cursor.execute(
        "SELECT id, stopName, latitude, longitude FROM stops WHERE CAST(latitude AS REAL) < -90 OR CAST(latitude AS REAL) > 90 OR CAST(longitude AS REAL) < -180 OR CAST(longitude AS REAL) > 180")
    check7 = [dict(row) for row in cursor.fetchall()]

    # 8. Inkonsistente User
    cursor.execute("SELECT DISTINCT userId, username FROM journeys WHERE userId NOT IN (SELECT id FROM users)")
    check8 = [dict(row) for row in cursor.fetchall()]

    # 9. Kaputte / Verwaiste Polylines
    cursor.execute("SELECT statusId FROM polylines WHERE geometry = '' OR geometry IS NULL")
    check9_empty = [dict(row) for row in cursor.fetchall()]
    cursor.execute("SELECT statusId FROM polylines WHERE statusId NOT IN (SELECT statusId FROM journeys)")
    check9_orphaned = [dict(row) for row in cursor.fetchall()]

    # Sammle alle eindeutigen betroffenen User über alle Checks hinweg für den Filter
    all_items = check1 + check2 + check3 + check4 + check5 + check6 + check7 + check8
    unique_users = sorted(list(set([r.get('username') for r in all_items if r.get('username')])))

    return render_template('debug.html',
                           check1=check1, check2=check2, check3=check3,
                           check4=check4, check5=check5, check6=check6,
                           check7=check7, check8=check8,
                           check9_empty=check9_empty, check9_orphaned=check9_orphaned,
                           unique_users=unique_users)


@app.route('/debug/queue')
def debug_queue():
    if 'username' not in session:
        return jsonify({'error': 'Nicht eingeloggt'}), 401

    with queue_lock:
        current_queue = list(WAITING_QUEUE)

    tasks_info = []
    # Alle Tasks anzeigen, neueste zuerst
    for tid, tinfo in reversed(list(TASKS.items())):
        is_waiting = tid in current_queue
        q_pos = (current_queue.index(tid) + 1) if is_waiting else None
        
        status = tinfo.get('status', 'running')
        if is_waiting and status == 'running':
            status = 'waiting'

        tasks_info.append({
            'task_id': tid,
            'creator': tinfo.get('creator', 'Unbekannt'),
            'type': tinfo.get('type', 'Auswertung'),
            'target_users': tinfo.get('target_users', []),
            'created_at': tinfo.get('created_at', ''),
            'queue_pos': q_pos,
            'status': status,
            'latest_log': tinfo['logs'][-1] if tinfo.get('logs') else ''
        })

    return jsonify({'tasks': tasks_info})


@app.route('/debug/run_fix', methods=['POST'])
def debug_run_fix():
    if 'username' not in session: return jsonify({'error': 'Nicht eingeloggt'}), 401
    username = session['username']
    fix_type = request.form.get('fix_type')

    task_id = str(uuid.uuid4())
    TASKS[task_id] = {
        'status': 'running',
        'logs': [f'Starte automatischen Reparatur-Task für: {fix_type}...'],
        'new_files': [],
        'stop_requested': False,
        'creator': username,
        'type': f'Debug ({fix_type})',
        'target_users': [username],
        'created_at': datetime.now().strftime('%H:%M:%S')
    }

    thread = threading.Thread(target=run_debug_task, args=(task_id, username, fix_type))
    thread.daemon = True
    thread.start()

    return jsonify({'task_id': task_id})


@app.route('/debug/api_check/<check_type>/<item_id>', methods=['POST'])
def debug_api_check(check_type, item_id):
    if 'username' not in session: return jsonify({'error': 'Nicht eingeloggt'}), 401
    username = session['username']

    conn = ta.init_db()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cursor.execute("SELECT access_token FROM user_login_data WHERE owner=?", (username,))
    row = cursor.fetchone()
    if not row or not row['access_token']: return jsonify({'error': 'Kein Token'}), 401
    api_token = row['access_token']

    differs = False
    diff_reasons = []

    train_type_to_readable_dict = {
        'regional': 'Regionalzug', 'bus': 'Bus', 'ferry': 'Fähre',
        'national': 'Regionalzug', 'regionalExp': 'Fernverkehr',
        'nationalExpress': 'Fernverkehr', 'suburban': 'S-Bahn', 'subway': 'U-Bahn',
        'taxi': 'Taxi', 'tram': 'Tram', 'plane': 'Igitt ein Flugzeug'
    }

    try:
        if check_type == 'journey':
            cursor.execute("SELECT * FROM journeys WHERE statusId = ?", (item_id,))
            db_row = cursor.fetchone()

            resp = requests.get(f'https://traewelling.de/api/v1/status/{item_id}', headers={
                'accept': 'application/json',
                'Authorization': f'Bearer {api_token}',
                'User-Agent': 'TraewellingAnalysis_WebApp v9'
            })

            if resp.status_code == 404:
                return jsonify({'success': True, 'api_data': None, 'differs': True, 'diff_reasons': [
                    'Fahrt existiert in der API nicht mehr (404 Not Found) - Löschen empfohlen']})
            elif resp.status_code == 200:
                api_data = resp.json().get('data', {})
                if db_row:
                    api_dist = api_data.get('checkin', {}).get('distance', 0)
                    if api_dist != db_row['distance']:
                        differs = True
                        diff_reasons.append(f"Distanz abweichend: API={api_dist}m, DB={db_row['distance']}m")

                    train_data = api_data.get('checkin', {})
                    api_line_raw = train_data.get('lineName', 'Unknown Line')
                    api_line_clean = re.sub(r'\s*\(\d+\)', '', api_line_raw).strip()

                    given_train_type = train_type_to_readable_dict.get(train_data.get('category', None),
                                                                       train_data.get('category', None))
                    if re.match(r"^(REX|RE|RB|Stoptrein|FEX|Sprinter)", api_line_clean) and given_train_type != 'Bus':
                        api_cat = 'Regionalzug'
                    elif re.match(r"^(RS)", api_line_clean):
                        api_cat = 'S-Bahn'
                    else:
                        api_cat = given_train_type

                    if api_cat != db_row['category']:
                        differs = True
                        diff_reasons.append(f"Kategorie abweichend: API={api_cat}, DB={db_row['category']}")

                    if api_line_clean != db_row['lineName']:
                        differs = True
                        diff_reasons.append(f"Linie abweichend: API={api_line_clean}, DB={db_row['lineName']}")

                if not differs:
                    diff_reasons.append(
                        "✅ Werte in der Datenbank stimmen exakt mit der API überein. (Ein Löschen ändert die Werte nicht)")

                return jsonify(
                    {'success': True, 'api_data': api_data, 'differs': differs, 'diff_reasons': diff_reasons})
            else:
                return jsonify({'success': False, 'message': f'API Fehler: {resp.status_code} - {resp.text}'})

        elif check_type == 'stop':
            cursor.execute("SELECT * FROM stops WHERE id = ?", (item_id,))
            db_row = cursor.fetchone()

            resp = requests.get(f'https://traewelling.de/api/v1/stations/{item_id}', headers={
                'accept': 'application/json',
                'Authorization': f'Bearer {api_token}',
                'User-Agent': 'TraewellingAnalysis_WebApp v9'
            })
            if resp.status_code == 404:
                return jsonify({'success': True, 'api_data': None, 'differs': True, 'diff_reasons': [
                    'Haltestelle existiert in der API nicht mehr (404 Not Found) - Löschen empfohlen']})
            elif resp.status_code == 200:
                api_data = resp.json().get('data', {})
                if db_row:
                    api_lat = str(api_data.get('latitude', ''))
                    db_lat = str(db_row['latitude']) if db_row['latitude'] else ''
                    if api_lat != db_lat:
                        differs = True
                        diff_reasons.append(f"Latitude abweichend: API={api_lat}, DB={db_lat}")

                    api_lon = str(api_data.get('longitude', ''))
                    db_lon = str(db_row['longitude']) if db_row['longitude'] else ''
                    if api_lon != db_lon:
                        differs = True
                        diff_reasons.append(f"Longitude abweichend: API={api_lon}, DB={db_lon}")

                if not differs:
                    diff_reasons.append(
                        "✅ Koordinaten in der Datenbank stimmen mit der API überein. (Ein Löschen ändert die Werte nicht)")

                return jsonify(
                    {'success': True, 'api_data': api_data, 'differs': differs, 'diff_reasons': diff_reasons})
            else:
                return jsonify({'success': False, 'message': f'API Fehler: {resp.status_code}'})

        return jsonify({'success': False, 'message': 'Unbekannter Check-Typ'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})


@app.route('/debug/fix_item/<check_type>/<item_id>', methods=['POST'])
def debug_fix_item(check_type, item_id):
    if 'username' not in session: return jsonify({'error': 'Nicht eingeloggt'}), 401

    conn = ta.init_db()
    cursor = conn.cursor()
    try:
        if check_type == 'journey':
            cursor.execute("DELETE FROM journeys WHERE statusId = ?", (item_id,))
            cursor.execute("DELETE FROM polylines WHERE statusId = ?", (item_id,))
            msg = "Fahrt und zugehörige Polyline gelöscht. Sie wird beim nächsten Sync neu geladen."
        elif check_type == 'stop':
            cursor.execute("DELETE FROM stops WHERE id = ?", (item_id,))
            msg = "Haltestelle gelöscht. Sie wird bei der nächsten Fahrtenberechnung neu geladen."
        elif check_type == 'polyline':
            cursor.execute("DELETE FROM polylines WHERE statusId = ?", (item_id,))
            msg = "Polyline gelöscht. Sie wird beim nächsten Sync neu geladen."
        else:
            return jsonify({'success': False, 'message': 'Unbekannter Typ'})

        conn.commit()
        return jsonify({'success': True, 'message': msg})
    except Exception as e:
        return jsonify({'success': False, 'message': f'Fehler: {str(e)}'})


@app.route('/debug/reload_status', methods=['POST'])
def debug_reload_status():
    if 'username' not in session: return jsonify({'error': 'Nicht eingeloggt'}), 401
    username = session['username']
    status_id = request.form.get('status_id')

    if not status_id:
        return jsonify({'success': False, 'message': 'Keine Statusnummer angegeben.'})

    try:
        status_id = int(status_id)
    except ValueError:
        return jsonify({'success': False, 'message': 'Ungültige Statusnummer.'})

    conn = ta.init_db()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT access_token FROM user_login_data WHERE owner=?", (username,))
    row = cursor.fetchone()
    if not row or not row['access_token']:
        return jsonify({'error': 'Kein Token'}), 401
    api_token = row['access_token']

    success, diffs = ta.reload_single_status(conn, api_token, status_id)
    if success:
        return jsonify({
            'success': True,
            'message': f'Status {status_id} erfolgreich korrigiert und neu berechnet (Polyline gelöscht).',
            'differences': diffs
        })
    else:
        return jsonify({'success': False, 'message': f'Fehler beim Neuladen des Status {status_id}.'})


if __name__ == '__main__':
    conn = ta.init_db()
    app.run(host='0.0.0.0', port=8080, debug=True, ssl_context= 'adhoc')