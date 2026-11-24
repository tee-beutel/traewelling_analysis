import os
import re
import time
import sqlite3
import requests
import uuid
import threading
import json
import math
import shutil
import hashlib
import urllib.parse
from pathlib import Path
from datetime import datetime, timedelta
from flask import Flask, request, session, redirect, url_for, render_template, send_from_directory, abort, jsonify, render_template_string, make_response
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
        f.write("GEMINI_API_KEY=\n")
        f.write("ADMIN_USER=tee_beutel,west_train\n")

load_dotenv(ENV_PATH, override=True)
flask_secret = os.getenv('FLASK_SECRET_KEY')
if not flask_secret:
    flask_secret = secrets.token_hex(32)
    set_key(ENV_PATH, 'FLASK_SECRET_KEY', flask_secret)
    os.environ['FLASK_SECRET_KEY'] = flask_secret

from werkzeug.middleware.proxy_fix import ProxyFix
from werkzeug.exceptions import HTTPException
import traceback

app.secret_key = flask_secret
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=12)
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_prefix=1)

@app.errorhandler(Exception)
def handle_unexpected_error(e):
    username = session.get('username') if 'username' in session else None
    if isinstance(e, HTTPException):
        if e.code >= 500:
            ta.log_error(
                username=username,
                error=e,
                context=f"HTTP {e.code} ({request.method} {request.path})"
            )
        return e

    tb_str = traceback.format_exc()
    ta.log_error(
        username=username,
        error=e,
        traceback_str=tb_str,
        context=f"Unhandled 500 ({request.method} {request.path})"
    )
    if request.path.startswith('/api/') or request.is_json:
        return jsonify({'error': 'Interner Serverfehler: ' + str(e)}), 500
    return render_template_string(
        "<h1>500 Interner Serverfehler</h1><p>{{ error }}</p><p><a href='/'>Zurück zur Startseite</a></p>",
        error=str(e)
    ), 500

@app.before_request
def make_session_permanent():
    session.permanent = True
    session.modified = True
    username = session.get('username')

    orig_ip = request.environ.get('orig_remote_addr') or request.environ.get('REMOTE_ADDR', '-')
    request.environ['orig_remote_addr'] = orig_ip
    request.environ['REMOTE_ADDR'] = f"@{username}" if username else f"Gast ({orig_ip})"
    request.environ['REMOTE_USER'] = username or '-'

    if username and not request.path.startswith('/static/'):
        now_ts = time.time()
        if now_ts - session.get('_last_online_sync', 0) > 60:
            session['_last_online_sync'] = now_ts
            try:
                conn = ta.init_db()
                cur = conn.cursor()
                cur.execute("UPDATE user_login_data SET last_online = ? WHERE owner = ?", (datetime.now().isoformat(), username))
                conn.commit()
                conn.close()
            except Exception:
                pass

@app.after_request
def update_log_username(response):
    username = session.get('username')
    orig_ip = request.environ.get('orig_remote_addr') or request.environ.get('REMOTE_ADDR', '-')
    request.environ['REMOTE_ADDR'] = f"@{username}" if username else f"Gast ({orig_ip})"
    request.environ['REMOTE_USER'] = username or '-'
    return response

def get_env_val(key, default=''):
    load_dotenv(ENV_PATH, override=True)
    return os.getenv(key, default).strip()

def save_env_val(key, value):
    set_key(ENV_PATH, key, value)
    os.environ[key] = value

def get_admin_users():
    raw = get_env_val('ADMIN_USER', 'tee_beutel,west_train')
    if not raw:
        return {'tee_beutel', 'west_train'}
    try:
        parsed = json.loads(raw)
        if isinstance(parsed, list):
            return {str(u).strip() for u in parsed if str(u).strip()}
    except Exception:
        pass
    return {u.strip() for u in raw.split(',') if u.strip()}

def is_admin_user(username):
    if not username:
        return False
    return username in get_admin_users()

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
USER_DOWNLOAD_LOCKS = {}
user_locks_mutex = threading.Lock()

def get_user_download_lock(username: str) -> threading.Lock:
    """Gibt ein nutzerspezifisches Lock zurück, damit Downloads und Syncs desselben Nutzers serialisiert werden,
    während andere Nutzer parallel arbeiten können."""
    with user_locks_mutex:
        if username not in USER_DOWNLOAD_LOCKS:
            USER_DOWNLOAD_LOCKS[username] = threading.Lock()
        return USER_DOWNLOAD_LOCKS[username]

StopTaskException = ta.StopTaskException

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

    user_lock = get_user_download_lock(creator)

    # Warten, falls für denselben Nutzer noch ein Download läuft
    while not user_lock.acquire(blocking=False):
        if TASKS[task_id].get('stop_requested', False):
            TASKS[task_id]['status'] = 'stopped'
            TASKS[task_id]['logs'].append("⚠️ Auswertung gestoppt.")
            return
        log(f"⏳ Warte auf Abschluss des vorherigen Downloads für @{creator}...", replace=True)
        time.sleep(0.5)

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
            user_lock.release()
            TASKS[task_id]['status'] = 'stopped'
            TASKS[task_id]['logs'].append("⚠️ Auswertung in der Warteschlange gestoppt.")
            return

        log(f"⏳ Warteschlange: Platz {pos} von {total_waiting} wartenden Prozessen...", replace=True)
        time.sleep(1)

    with generation_lock:
        try:
            log('🚀 Freie Server-Kapazitäten erhalten! Starte Initialisierung...')
            conn = ta.init_db()
            api_token = ta.get_valid_access_token(conn, creator, min_validity_sec=1800) or api_token

            trusted_users = ta.fetch_paginated_users(api_token, 'https://traewelling.de/api/v1/user/self/trusted-by')
            followed_users = ta.fetch_paginated_users(api_token, 'https://traewelling.de/api/v1/user/self/followings')
            followers_users = ta.fetch_paginated_users(api_token, 'https://traewelling.de/api/v1/user/self/followers')
            ta.set_auth_user(creator)
            ta.set_trusted_users(trusted_users)
            ta.set_followed_users(followed_users)
            ta.set_followers_users(followers_users)
            user_excel_settings = ta.get_user_excel_settings(conn, creator)

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

                    def on_user_progress(p):
                        check_stop()
                        if task_id not in TASKS:
                            return
                        s_curr = p.get('step_current', 0)
                        s_tot = p.get('step_total', 0)
                        s_idx = p.get('step_index', 0)
                        s_pct = p.get('step_percent')
                        if s_pct is None:
                            s_pct = round((s_curr / s_tot) * 100) if s_tot > 0 else (min(90, s_curr * 15) if s_idx == 1 else 100)
                        overall_pct = round((((idx) * 4 + (s_idx - 1) + (s_pct / 100.0)) / (total * 4)) * 100)
                        overall_pct = max(0, min(100, overall_pct))

                        rem_val = p.get('rate_limit_remaining', 500)
                        max_limit = p.get('rate_limit_max', max(500, int(rem_val)))
                        TASKS[task_id]['progress'] = {
                            'user_index': idx + 1,
                            'user_total': total,
                            'current_user': user,
                            'step_index': s_idx,
                            'step_name': p.get('step_name', ''),
                            'overall_percent': overall_pct,
                            'step_current': s_curr,
                            'step_total': s_tot,
                            'step_percent': s_pct,
                            'step_detail': p.get('step_detail', ''),
                            'rate_limit_remaining': rem_val,
                            'rate_limit_max': max_limit,
                            'rate_limit_waiting': p.get('rate_limit_waiting', False),
                            'rate_limit_sleep_remaining': p.get('rate_limit_sleep_remaining', 0),
                            'eta_seconds': p.get('eta_seconds'),
                            'eta_based_on_history': p.get('eta_based_on_history', False),
                            'eta_sample_count': p.get('eta_sample_count', 0)
                        }

                    def on_new_snippets(new_snippets):
                        if task_id not in TASKS:
                            return
                        existing_list = TASKS[task_id].setdefault('status_snippets', [])
                        existing_ids = {s.get('statusId') for s in existing_list}
                        for snip in new_snippets:
                            sid = snip.get('statusId')
                            if sid and sid not in existing_ids:
                                if len(existing_list) < 500:
                                    existing_ids.add(sid)
                                    existing_list.append(snip)
                                else:
                                    break

                    ta.fetch_data_to_db(user, conn=conn, api_token=api_token, reload_all=False,
                                        log_callback=log, update_db=options.get('update_db', True),
                                        progress_callback=on_user_progress,
                                        on_statuses_callback=on_new_snippets)

                    if ta.check_for_user_journeys(conn, user, start_date=start_date, end_date=end_date):

                        if options['gen_excel']:
                            check_stop()
                            log(f'Erstelle Excel-Fahrtenbuch für {user}...')
                            excel_filename = ta.create_user_excel_file(conn, user, start_date=start_date, end_date=end_date, create_picture=options['gen_time'], custom_settings=user_excel_settings, creator_username=creator)
                            check_stop()
                            add_live_file(excel_filename, "user")
                            log(f'   📄 {excel_filename} steht jetzt zum Download bereit!')

                        if options['gen_time']:
                            check_stop()
                            log(f'Erstelle Tageszeiten-Grafiken für {user}...')
                            for grouping in ['time_of_day', 'weekday', 'month', 'month_year', 'date', 'year']:
                                img_filename = ta.create_ganglinie_plot(username=user, grouping_type=grouping, conn=conn, start_date=start_date, end_date=end_date, creator_username=creator)
                                if img_filename:
                                    check_stop()
                                    add_live_file(img_filename, "gang_linien")

                        if options['gen_map'] and len(selected_users) == 1:
                            check_stop()
                            log(f'Erstelle interaktive Streckenkarte für {user}...')
                            map_filename = ta.create_polyline(conn, api_token, usernames=[user], start_date=start_date, end_date=end_date, log_callback=log, lite_version=options['gen_map_lite'], creator_username=creator)
                            if map_filename:
                                check_stop()
                                add_live_file(map_filename, "polylines")
                                log(f'   🗺️ {map_filename} steht jetzt zur Ansicht bereit!')
                                
                        if options['gen_stop_map'] and len(selected_users) == 1:
                            check_stop()
                            log(f'Erstelle interaktive Haltestellenkarte für {user}...')
                            stop_map_filename = ta.create_stops_map(conn, usernames=[user], start_date=start_date, end_date=end_date, log_callback=log, lite_version=options['gen_stop_map_lite'], creator_username=creator)
                            if stop_map_filename:
                                check_stop()
                                add_live_file(stop_map_filename, "polylines")
                                log(f'   📍 {stop_map_filename} steht jetzt zur Ansicht bereit!')

                        if options.get('gen_timeline') and len(selected_users) == 1:
                            check_stop()
                            log(f'Erstelle interaktive Timeline für {user}...')
                            timeline_filename = ta.create_timeline(conn, usernames=[user], start_date=start_date, end_date=end_date, log_callback=log, creator_username=creator)
                            if timeline_filename:
                                check_stop()
                                add_live_file(timeline_filename, "timelines")
                                log(f'   ⏱️ {timeline_filename} steht jetzt zur Ansicht bereit!')

                        if options['gen_delay']:
                            check_stop()
                            log(f'Erstelle Verspätungs-Graphen für {user}...')
                            delay_filenames = ta.create_verspaetungs_verteilungen_plot(user, conn=conn, start_date=start_date, end_date=end_date, creator_username=creator)
                            for d_file in (delay_filenames or []):
                                check_stop()
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
                    shared_excel = ta.create_shared_excel_file(conn=conn, usernames=selected_users, start_date=start_date, end_date=end_date, custom_settings=user_excel_settings, creator_username=creator)
                    if shared_excel:
                        check_stop()
                        add_live_file(shared_excel, "shared")
                        log(f'   📊 {shared_excel} steht jetzt zum Download bereit!')

                if options['gen_map']:
                    check_stop()
                    log('Erstelle gemeinsame interaktive Streckenkarte...')
                    shared_map = ta.create_polyline(conn, api_token, usernames=selected_users, start_date=start_date, end_date=end_date, log_callback=log, lite_version=options['gen_map_lite'], creator_username=creator)
                    if shared_map:
                        check_stop()
                        add_live_file(shared_map, "polylines")
                        log(f'   🗺️ {shared_map} steht jetzt zur Ansicht bereit!')
                        
                if options['gen_stop_map']:
                    check_stop()
                    log('Erstelle gemeinsame interaktive Haltestellenkarte...')
                    shared_stop_map = ta.create_stops_map(conn, usernames=selected_users, start_date=start_date, end_date=end_date, log_callback=log, lite_version=options['gen_stop_map_lite'], creator_username=creator)
                    if shared_stop_map:
                        check_stop()
                        add_live_file(shared_stop_map, "polylines")
                        log(f'   📍 {shared_stop_map} steht jetzt zur Ansicht bereit!')

                if options.get('gen_timeline'):
                    check_stop()
                    log('Erstelle gemeinsame interaktive Timeline...')
                    timeline_filename = ta.create_timeline(conn, usernames=selected_users, start_date=start_date, end_date=end_date, log_callback=log, creator_username=creator)
                    if timeline_filename:
                        check_stop()
                        add_live_file(timeline_filename, "timelines")
                        log(f'   ⏱️ {timeline_filename} steht jetzt zur Ansicht bereit!')

            log('\n✅ Alle Auswertungen erfolgreich abgeschlossen!')
            TASKS[task_id]['status'] = 'completed'
            if 'progress' in TASKS[task_id]:
                TASKS[task_id]['progress']['overall_percent'] = 100
                TASKS[task_id]['progress']['step_percent'] = 100
                TASKS[task_id]['progress']['step_name'] = 'Abgeschlossen'
                TASKS[task_id]['progress']['step_detail'] = 'Alle Schritte erfolgreich abgeschlossen.'
                TASKS[task_id]['progress']['eta_seconds'] = 0

        except StopTaskException as e:
            log(f'\n{str(e)}', ignore_stop=True)
            TASKS[task_id]['status'] = 'stopped'
        except Exception as e:
            tb_str = traceback.format_exc()
            print(tb_str)
            ta.log_error(
                username=creator,
                error=e,
                traceback_str=tb_str,
                context=f"run_generation_task (task_id: {task_id}, target_users: {selected_users})"
            )
            log(f'\n❌ KRITISCHER FEHLER: {str(e)}\n{tb_str.splitlines()[-3] if len(tb_str.splitlines()) >= 3 else ""}', ignore_stop=True)
            TASKS[task_id]['status'] = 'error'
        finally:
            ta.clear_user_context()
            with queue_lock:
                if task_id in WAITING_QUEUE:
                    WAITING_QUEUE.remove(task_id)
            user_lock.release()


def run_debug_task(task_id, username, fix_type):
    """Hintergrund-Task für automatische Reparaturen im Debug-Modus."""
    def log(msg):
        if task_id in TASKS:
            TASKS[task_id]['logs'].append(msg)

    conn = None
    try:
        conn = ta.init_db()
        cursor = conn.cursor()
        api_token = ta.get_valid_access_token(conn, username)
        if not api_token:
            log("❌ Kein Benutzerkonto oder gültiger API-Token in user_login_data gefunden.")
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
        ta.log_error(
            username=username,
            error=e,
            context=f"run_debug_task (fix_type: {fix_type}, task_id: {task_id})"
        )
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
    username = session.get('username')
    if username:
        conn = ta.init_db()
        is_admin = is_admin_user(username)
        sync_users = ta.get_users_sync_overview(conn, username, is_admin=is_admin)
        return render_template('index.html', logged_in=True, username=username, sync_users=sync_users, is_admin=is_admin)

    cid, csecret, red_uri = get_oauth_config()
    has_oauth = bool(cid and csecret and red_uri)
    stored_token = get_env_val('TRAEWELLING_API_TOKEN')
    has_token = bool(stored_token)
    token_preview = ""
    if has_token:
        token_preview = stored_token[:12] + "..." + stored_token[-8:] if len(stored_token) > 24 else "***"

    return render_template('index.html',
                           logged_in=False,
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
    cid, csecret, red_uri = get_oauth_config()
    if cid and csecret and red_uri:
        return redirect(url_for('index', error="OAuth-Zugangsdaten sind bereits in der .env hinterlegt."))

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
    cid, csecret, red_uri = get_oauth_config()
    has_oauth = bool(cid and csecret and red_uri)

    if not token and not has_oauth:
        token = get_env_val('TRAEWELLING_API_TOKEN')

    if not token:
        return redirect(url_for('index', error="Bitte gib einen API-Token ein."))

    try:
        user_resp = requests.get(
            'https://traewelling.de/api/v1/auth/user',
            headers={
                'Authorization': f'Bearer {token}',
                'User-Agent': 'TraewellingAnalysis_WebApp v10',
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

    # Nur in .env speichern, wenn KEIN OAuth hinterlegt ist
    if not has_oauth:
        save_env_val('TRAEWELLING_API_TOKEN', token)
        session['session_only'] = False
    else:
        session['session_only'] = True

    trusted = ta.fetch_paginated_users(token, 'https://traewelling.de/api/v1/user/self/trusted-by')
    followed = ta.fetch_paginated_users(token, 'https://traewelling.de/api/v1/user/self/followings')
    followers = ta.fetch_paginated_users(token, 'https://traewelling.de/api/v1/user/self/followers')

    now_iso = datetime.now().isoformat()
    conn = ta.init_db()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO user_login_data (owner, access_token, refresh_token, expires_at, followed_users, trusted_users, followers, last_updated, last_online) 
        VALUES (?, ?, NULL, NULL, ?, ?, ?, ?, ?) 
        ON CONFLICT(owner) DO UPDATE SET 
            access_token = excluded.access_token,
            refresh_token = NULL,
            expires_at = NULL,
            followed_users = excluded.followed_users,
            trusted_users = excluded.trusted_users,
            followers = excluded.followers,
            last_updated = excluded.last_updated,
            last_online = excluded.last_online
    """, (
        username, 
        token, 
        json.dumps(list(followed)), 
        json.dumps(list(trusted)), 
        json.dumps(list(followers)),
        now_iso,
        now_iso
    ))
    conn.commit()
    conn.close()

    session['username'] = username
    return redirect(url_for('index'))

@app.route('/callback')
def callback():
    code = request.args.get('code')
    if not code: return "Fehler beim Login", 400

    cid, csecret, red_uri = get_oauth_config()
    token_data = {'grant_type': 'authorization_code', 'client_id': cid, 'client_secret': csecret, 'redirect_uri': red_uri, 'code': code}
    
    resp = requests.post(TOKEN_URL, data=token_data, headers={'User-Agent': 'TraewellingAnalysis_WebApp v10'})
    if resp.status_code != 200: return f"Token-Fehler: {resp.text}", 400

    token_json = resp.json()
    access_token = token_json.get('access_token')
    refresh_token = token_json.get('refresh_token')
    expires_in = token_json.get('expires_in')
    expires_at = (time.time() + float(expires_in)) if expires_in else None
    
    user_resp = requests.get(
        'https://traewelling.de/api/v1/auth/user', 
        headers={
            'Authorization': f'Bearer {access_token}',
            'User-Agent': 'TraewellingAnalysis_WebApp v10',
            'Accept': 'application/json'
        }
    )

    if user_resp.status_code != 200:
        return f"Fehler von der Träwelling-API: {user_resp.status_code} - {user_resp.text}", 400

    username = user_resp.json().get('data', {}).get('username')

    trusted = ta.fetch_paginated_users(access_token, 'https://traewelling.de/api/v1/user/self/trusted-by')
    followed = ta.fetch_paginated_users(access_token, 'https://traewelling.de/api/v1/user/self/followings')
    followers = ta.fetch_paginated_users(access_token, 'https://traewelling.de/api/v1/user/self/followers')

    conn = ta.init_db()
    cursor = conn.cursor()
    
    now_iso = datetime.now().isoformat()
    cursor.execute("""
        INSERT INTO user_login_data (owner, access_token, refresh_token, expires_at, followed_users, trusted_users, followers, last_updated, last_online) 
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?) 
        ON CONFLICT(owner) DO UPDATE SET 
            access_token = excluded.access_token,
            refresh_token = excluded.refresh_token,
            expires_at = excluded.expires_at,
            followed_users = excluded.followed_users,
            trusted_users = excluded.trusted_users,
            followers = excluded.followers,
            last_updated = excluded.last_updated,
            last_online = excluded.last_online
    """, (
        username, 
        access_token, 
        refresh_token,
        expires_at,
        json.dumps(list(followed)), 
        json.dumps(list(trusted)), 
        json.dumps(list(followers)),
        now_iso,
        now_iso
    ))
    conn.commit()
    conn.close()

    session['username'] = username
    session['session_only'] = False
    return redirect(url_for('index'))

@app.route('/logout')
def logout():
    username = session.pop('username', None)
    is_session_only = session.pop('session_only', False)
    if is_session_only and username:
        try:
            conn = ta.init_db()
            cursor = conn.cursor()
            cursor.execute("UPDATE user_login_data SET access_token = '' WHERE owner = ?", (username,))
            conn.commit()
            conn.close()
        except Exception:
            pass
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
    api_token = ta.get_valid_access_token(conn, username, min_validity_sec=1800)
    conn.close()
    if not api_token:
        return jsonify({'error': 'Kein Token'}), 401
    
    task_id = str(uuid.uuid4())
    source = request.form.get('source', 'dashboard')
    TASKS[task_id] = {
        'status': 'running', 
        'logs': ['Prozess in Warteschlange eingereiht...'],
        'new_files': [],
        'status_snippets': [],
        'stop_requested': False,
        'creator': username,
        'type': 'Auswertung',
        'source': source,
        'target_users': selected_users,
        'created_at': datetime.now().strftime('%H:%M:%S'),
        'progress': {
            'user_index': 1,
            'user_total': len(selected_users),
            'current_user': selected_users[0],
            'step_index': 0,
            'step_name': 'Initialisierung...',
            'overall_percent': 0,
            'step_current': 0,
            'step_total': 0,
            'step_percent': 0,
            'step_detail': 'Warte auf Start...',
            'rate_limit_remaining': 500,
            'rate_limit_max': 500,
            'rate_limit_waiting': False,
            'rate_limit_sleep_remaining': 0,
            'eta_seconds': None,
            'eta_based_on_history': False,
            'eta_sample_count': 0
        }
    }

    thread = threading.Thread(target=run_generation_task, args=(task_id, username, api_token, selected_users, start_date, end_date, options))
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
    if task_id not in TASKS:
        return jsonify({'status': 'error', 'logs': ['Task nicht gefunden'], 'new_files': []})
    task_data = dict(TASKS[task_id])
    since = request.args.get('since', type=int)
    if since is not None and 'logs' in task_data:
        task_data['logs'] = task_data['logs'][since:]
        task_data['log_offset'] = since

    snippet_offset = request.args.get('snippet_offset', type=int)
    snippets = task_data.pop('status_snippets', [])
    if snippet_offset is not None:
        task_data['new_snippets'] = snippets[snippet_offset:]
    else:
        task_data['new_snippets'] = snippets
    task_data['snippet_total'] = len(snippets)

    with queue_lock:
        is_in_queue = task_id in WAITING_QUEUE
        is_waiting = is_in_queue and (len(WAITING_QUEUE) > 0 and WAITING_QUEUE[0] != task_id)
        queue_pos = WAITING_QUEUE.index(task_id) if is_in_queue else 0
        total_waiting = max(0, len(WAITING_QUEUE) - 1)

    task_data['is_waiting'] = is_waiting
    task_data['queue_pos'] = queue_pos
    task_data['total_waiting'] = total_waiting
    return jsonify(task_data)

@app.route('/api/active_task')
def get_active_task():
    if 'username' not in session:
        return jsonify({'active': False})
    username = session['username']

    with queue_lock:
        queue_snapshot = list(WAITING_QUEUE)

    for tid, tinfo in reversed(list(TASKS.items())):
        if tinfo.get('creator') == username and tinfo.get('status') in ('running', 'waiting'):
            is_in_queue = tid in queue_snapshot
            is_waiting = is_in_queue and (len(queue_snapshot) > 0 and queue_snapshot[0] != tid)
            queue_pos = queue_snapshot.index(tid) if is_in_queue else 0
            total_waiting = max(0, len(queue_snapshot) - 1)

            task_copy = dict(tinfo)
            task_copy.pop('status_snippets', None)
            task_copy['is_waiting'] = is_waiting
            task_copy['queue_pos'] = queue_pos
            task_copy['total_waiting'] = total_waiting

            return jsonify({
                'active': True,
                'task_id': tid,
                'task': task_copy,
                'is_waiting': is_waiting,
                'queue_pos': queue_pos,
                'total_waiting': total_waiting
            })

    return jsonify({'active': False})

@app.route('/api/sync_users')
def api_sync_users():
    if 'username' not in session:
        return jsonify({'error': 'Nicht eingeloggt'}), 401
    conn = ta.init_db()
    username = session['username']
    is_admin = is_admin_user(username)
    users_overview = ta.get_users_sync_overview(conn, username, is_admin=is_admin)
    return jsonify({'users': users_overview})

@app.route('/api/search_users')
def api_search_users():
    if 'username' not in session:
        return jsonify({'error': 'Nicht eingeloggt'}), 401

    username = session['username']
    q = request.args.get('q', '').strip()
    if len(q) < 2:
        return jsonify({'users': []})

    conn = ta.init_db()
    try:
        api_token = ta.get_valid_access_token(conn, username)
        if not api_token:
            return jsonify({'error': 'Kein gültiges API-Token', 'users': []}), 401

        encoded_q = urllib.parse.quote(q, safe='')
        url = f'https://traewelling.de/api/v1/user/search/{encoded_q}'
        resp, _ = ta.traewelling_api_get(url, api_token, conn=conn, username=username, timeout=10)

        if resp.status_code == 429:
            return jsonify({'error': 'Rate-Limit erreicht. Bitte kurz warten.', 'users': []}), 429
        if resp.status_code != 200:
            ta.log_error(
                username=username,
                error_type=f"HTTP_{resp.status_code}",
                error_message=f"Träwelling API /user/search Fehler: {resp.text}",
                context=f"api_search_users (q={q})",
                conn=conn
            )
            return jsonify({'error': f'Träwelling-Suche nicht verfügbar (HTTP {resp.status_code}).', 'users': []}), resp.status_code

        cursor = conn.cursor()
        cursor.execute("SELECT followed_users, trusted_users, followers FROM user_login_data WHERE owner = ?", (username,))
        row = cursor.fetchone()
        all_friends = set()
        if row:
            for val in row:
                if val:
                    try:
                        all_friends.update(json.loads(val))
                    except Exception:
                        pass

        data = resp.json().get('data', [])
        users_list = []
        for item in data:
            if not isinstance(item, dict):
                continue
            u_name = item.get('username') or ''
            if not u_name:
                continue
            private_profile = bool(item.get('privateProfile', False))
            is_self = (u_name == username)
            is_friend = (u_name in all_friends) or bool(item.get('following', False))
            likely_invisible = bool(private_profile and not is_friend and not is_self)
            users_list.append({
                'username': u_name,
                'displayName': item.get('displayName') or u_name,
                'privateProfile': private_profile,
                'isFriend': is_friend,
                'isSelf': is_self,
                'likelyInvisible': likely_invisible
            })

        return jsonify({'users': users_list})
    except Exception as e:
        ta.log_error(username=username, error=e, context=f"api_search_users (q={q})", conn=conn)
        return jsonify({'error': f'Fehler bei der Nutzersuche: {e}', 'users': []}), 500
    finally:
        conn.close()

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



# ==========================================
# EXCEL-SETTINGS ROUTEN (user_login_data)
# ==========================================
@app.route('/settings')
def settings_page():
    if 'username' not in session:
        return redirect(url_for('login'))
    username = session['username']

    conn = ta.init_db()
    current_settings = ta.get_user_excel_settings(conn, username)
    default_settings = ta.get_default_excel_settings()
    conn.close()

    return render_template(
        'settings.html',
        username=username,
        current_settings=current_settings,
        default_settings=default_settings,
        available_group_cols=ta.AVAILABLE_GROUP_COLUMNS,
        available_temporal=ta.AVAILABLE_TEMPORAL_TYPES,
        available_geo=ta.AVAILABLE_GEO_COLUMNS,
        sheet_metrics_schema=ta.SHEET_METRICS_SCHEMA
    )


@app.route('/api/settings/save', methods=['POST'])
def api_save_settings():
    if 'username' not in session:
        return jsonify({'error': 'Nicht eingeloggt'}), 401
    username = session['username']

    data = request.get_json(silent=True)
    if not data or not isinstance(data, dict):
        return jsonify({'error': 'Ungültige JSON-Daten'}), 400

    if not isinstance(data.get('user_sheets'), list) or not isinstance(data.get('shared_sheets'), list):
        return jsonify({'error': 'JSON muss user_sheets und shared_sheets als Listen enthalten.'}), 400

    conn = ta.init_db()
    try:
        ta.save_user_excel_settings(conn, username, data)
        return jsonify({'success': True})
    except Exception as e:
        ta.log_error(username=username, error=e, context="api_save_settings")
        return jsonify({'error': str(e)}), 500
    finally:
        conn.close()


@app.route('/api/settings/reset', methods=['POST'])
def api_reset_settings():
    if 'username' not in session:
        return jsonify({'error': 'Nicht eingeloggt'}), 401
    username = session['username']

    conn = ta.init_db()
    try:
        defaults = ta.get_default_excel_settings()
        ta.save_user_excel_settings(conn, username, defaults)
        return jsonify({'success': True, 'settings': defaults})
    except Exception as e:
        ta.log_error(username=username, error=e, context="api_reset_settings")
        return jsonify({'error': str(e)}), 500
    finally:
        conn.close()


@app.route('/debug')
def debug_page():
    if 'username' not in session: return redirect(url_for('login'))
    if not is_admin_user(session['username']):
        abort(403)

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

    # 10. Fehlerprotokoll (Error Log)
    check10_errors = ta.get_error_logs(limit=100)

    # Sammle alle eindeutigen betroffenen User über alle Checks hinweg für den Filter
    all_items = check1 + check2 + check3 + check4 + check5 + check6 + check7 + check8 + check10_errors
    unique_users = sorted(list(set([r.get('username') for r in all_items if r.get('username')])))

    return render_template('debug.html',
                           check1=check1, check2=check2, check3=check3,
                           check4=check4, check5=check5, check6=check6,
                           check7=check7, check8=check8,
                           check9_empty=check9_empty, check9_orphaned=check9_orphaned,
                           check10_errors=check10_errors,
                           unique_users=unique_users)


@app.route('/debug/clear_errors', methods=['POST'])
def debug_clear_errors():
    if 'username' not in session: return jsonify({'error': 'Nicht eingeloggt'}), 401
    if not is_admin_user(session['username']):
        return jsonify({'error': 'Nicht autorisiert'}), 403

    deleted_count = ta.clear_error_logs()
    return jsonify({'success': True, 'message': f'{deleted_count} Fehlereinträge wurden gelöscht.'})


@app.route('/debug/queue')
def debug_queue():
    if 'username' not in session:
        return jsonify({'error': 'Nicht eingeloggt'}), 401
    if not is_admin_user(session['username']):
        return jsonify({'error': 'Nicht autorisiert'}), 403

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
    if not is_admin_user(session['username']):
        return jsonify({'error': 'Nicht autorisiert'}), 403
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
    if not is_admin_user(session['username']):
        return jsonify({'error': 'Nicht autorisiert'}), 403
    username = session['username']

    conn = ta.init_db()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    api_token = ta.get_valid_access_token(conn, username)
    if not api_token:
        conn.close()
        return jsonify({'error': 'Kein Token'}), 401

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
                'User-Agent': 'TraewellingAnalysis_WebApp v10'
            })
            if resp.status_code == 401:
                new_token = ta.refresh_oauth_token(conn, username, force=True)
                if new_token:
                    api_token = new_token
                    resp = requests.get(f'https://traewelling.de/api/v1/status/{item_id}', headers={
                        'accept': 'application/json',
                        'Authorization': f'Bearer {api_token}',
                        'User-Agent': 'TraewellingAnalysis_WebApp v10'
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
                'User-Agent': 'TraewellingAnalysis_WebApp v10'
            })
            if resp.status_code == 401:
                new_token = ta.refresh_oauth_token(conn, username, force=True)
                if new_token:
                    api_token = new_token
                    resp = requests.get(f'https://traewelling.de/api/v1/stations/{item_id}', headers={
                        'accept': 'application/json',
                        'Authorization': f'Bearer {api_token}',
                        'User-Agent': 'TraewellingAnalysis_WebApp v10'
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
        ta.log_error(username=username, error=e, context=f"debug_api_check ({check_type}/{item_id})")
        return jsonify({'success': False, 'message': str(e)})


@app.route('/debug/fix_item/<check_type>/<item_id>', methods=['POST'])
def debug_fix_item(check_type, item_id):
    if 'username' not in session: return jsonify({'error': 'Nicht eingeloggt'}), 401
    if not is_admin_user(session['username']):
        return jsonify({'error': 'Nicht autorisiert'}), 403

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
        ta.log_error(username=username, error=e, context=f"debug_fix_item ({check_type}/{item_id})")
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

    # Berechtigungsprüfung: Gehört der Status zum eingeloggten User (oder ist User Admin)?
    cursor.execute("SELECT username FROM journeys WHERE statusId = ?", (status_id,))
    j_row = cursor.fetchone()
    if j_row and j_row['username'] != username and not is_admin_user(username):
        conn.close()
        return jsonify({
            'success': False,
            'message': f'Berechtigung verweigert: Fahrt #{status_id} gehört @{j_row["username"]}. Du kannst nur deine eigenen Fahrten neu berechnen.'
        }), 403

    api_token = ta.get_valid_access_token(conn, username)
    if not api_token:
        conn.close()
        return jsonify({'error': 'Kein Träwelling-Token für deinen Account vorhanden.'}), 401

    success, diffs = ta.reload_single_status(conn, api_token, status_id)
    conn.close()

    col_labels = {
        'distance': 'Distanz (m)',
        'distanceStraightLine': 'Luftlinie (m)',
        'distanceToPrevJourney': 'Distanz zur Vorfahrt (m)',
        'durationReal': 'Reale Fahrzeit (min)',
        'durationPlanned': 'Geplante Fahrzeit (min)',
        'departureDelay': 'Abfahrtsverspätung (min)',
        'arrivalDelay': 'Ankunftsverspätung (min)',
        'departureFinal': 'Abfahrt (real/final)',
        'arrivalFinal': 'Ankunft (real/final)',
        'departurePlanned': 'Abfahrt (Plan)',
        'arrivalPlanned': 'Ankunft (Plan)',
        'departureFromRealtime': 'Abfahrt (Echtzeit)',
        'arrivalFromRealtime': 'Ankunft (Echtzeit)',
        'realtimeAvailability': 'Echtzeit-Verfügbarkeit',
        'realtimeFromDataprovider': 'Echtzeit von Datenprovider',
        'interStops': 'Zwischenhalte',
        'interStopsNumber': 'Anzahl Zwischenhalte',
        'locomotiveClass': 'Baureihe',
        'vehicleNumber': 'Fahrzeug-/Triebzugnummer',
        'wagonClass': 'Wagenklasse',
        'travelClass': 'Klasse',
        'wagon': 'Wagennummer',
        'seat': 'Sitzplatz',
        'lineName': 'Linienname',
        'operatorName': 'Betreiber',
        'category': 'Kategorie',
        'body': 'Statustext',
        'ticket': 'Ticket',
        'price': 'Preis',
        'points': 'Punkte'
    }

    if success:
        if not diffs:
            message = f"Status #{status_id} neu berechnet: Es gab keine Änderungen (alle Daten waren bereits aktuell). Polyline wurde zurückgesetzt."
        else:
            diff_lines = []
            for col, vals in diffs.items():
                label = col_labels.get(col, col)
                old_val = '(leer)' if vals['old'] is None or vals['old'] == '' else str(vals['old'])
                new_val = '(leer)' if vals['new'] is None or vals['new'] == '' else str(vals['new'])
                if len(old_val) > 40: old_val = old_val[:37] + '...'
                if len(new_val) > 40: new_val = new_val[:37] + '...'
                diff_lines.append(f"• {label}: {old_val} ➔ {new_val}")

            message = f"Status #{status_id} erfolgreich neu berechnet!\n\nGeänderte Werte ({len(diffs)}):\n" + "\n".join(diff_lines)

        return jsonify({
            'success': True,
            'message': message,
            'differences': diffs,
            'has_changes': bool(diffs)
        })
    else:
        return jsonify({'success': False, 'message': f'Fehler beim Neuladen des Status {status_id}.'})


# ==========================================
# INTERAKTIVE AUSWERTUNG / DRILL-DOWN API
# ==========================================

@app.route('/interactive')
@app.route('/interactiv')
def interactive_page():
    if 'username' not in session:
        return redirect(url_for('login'))
    username = session['username']

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
    cursor.execute("SELECT username FROM users WHERE username IS NOT NULL")
    for r in cursor.fetchall():
        if r[0]:
            available_users.add(r[0])

    current_settings = ta.get_user_excel_settings(conn, username)
    conn.close()

    return render_template(
        'interactive.html',
        username=username,
        available_users=sorted(list(available_users)),
        current_settings=current_settings,
        available_group_cols=ta.AVAILABLE_GROUP_COLUMNS,
        available_temporal=ta.AVAILABLE_TEMPORAL_TYPES,
        available_geo=ta.AVAILABLE_GEO_COLUMNS,
        sheet_metrics_schema=ta.SHEET_METRICS_SCHEMA
    )


def _df_to_clean_records(df, use_index=False):
    if df is None or df.empty:
        return [], []
    df_out = df.copy()
    if use_index and df_out.index.name is not None:
        df_out = df_out.reset_index()
    df_out = df_out.loc[:, ~df_out.columns.duplicated()]

    for col in df_out.columns:
        if ta.pd.api.types.is_datetime64_any_dtype(df_out[col]):
            df_out[col] = df_out[col].dt.strftime('%d.%m.%Y %H:%M').fillna('')

    df_out = df_out.replace({ta.np.nan: None, ta.np.inf: '∞', -ta.np.inf: '-∞'})
    cols = [str(c) for c in df_out.columns]
    records = []
    for row in df_out.to_dict(orient='records'):
        clean_row = {}
        for k, v in row.items():
            if isinstance(v, (ta.datetime, ta.pd.Timestamp)):
                clean_row[str(k)] = v.strftime('%d.%m.%Y %H:%M')
            elif isinstance(v, float) and (ta.math.isnan(v) or ta.math.isinf(v)):
                clean_row[str(k)] = None
            else:
                clean_row[str(k)] = v
        records.append(clean_row)
    return cols, records


def _compute_filtered_summary(conn, usernames, start_date, end_date):
    if isinstance(usernames, str):
        usernames = [usernames]
    placeholders = ','.join(['?'] * len(usernames))
    params = list(usernames)
    d_sql = ""
    if start_date:
        d_sql += " AND arrivalFinal > ?"
        params.append(start_date)
    if end_date:
        e_str = end_date + 'T23:59:59' if len(end_date) == 10 else end_date
        d_sql += " AND departureFinal < ?"
        params.append(e_str)

    vis_sql = ta.get_shared_visibility_query()
    cur = conn.cursor()
    cur.execute(f"""
        SELECT
            COUNT(statusId),
            COALESCE(SUM(distance), 0),
            COALESCE(SUM(durationReal), 0),
            ROUND(AVG(arrivalDelay), 2),
            COUNT(DISTINCT operatorName),
            COUNT(DISTINCT operatorName || lineName),
            (COUNT(DISTINCT originId) + COUNT(DISTINCT CASE WHEN destinationId != originId THEN destinationId END))
        FROM journeys
        WHERE username IN ({placeholders}) {d_sql} {vis_sql}
    """, params)
    r = cur.fetchone() or (0, 0, 0, 0, 0, 0, 0)
    t_trips = r[0] or 0
    t_dist_km = round((r[1] or 0) / 1000.0, 2)
    t_dur = int(r[2] or 0)
    avg_speed = round(t_dist_km / (t_dur / 60.0), 2) if t_dur > 0 else 0.0
    return {
        'total_journeys': t_trips,
        'total_distance_km': t_dist_km,
        'total_duration_hhmm': f"{t_dur // 60:02d}:{t_dur % 60:02d}",
        'avg_speed_kmh': avg_speed,
        'avg_arr_delay': r[3] or 0.0,
        'unique_operators': r[4] or 0,
        'unique_lines': r[5] or 0,
        'unique_stops': r[6] or 0
    }


@app.route('/api/interactive/query', methods=['POST'])
def api_interactive_query():
    if 'username' not in session:
        return jsonify({'error': 'Nicht eingeloggt'}), 401
    creator = session['username']
    payload = request.get_json(silent=True) or {}

    mode = payload.get('mode', 'user')
    users = payload.get('users') or [creator]
    start_date = (payload.get('start_date') or '').strip()
    end_date = (payload.get('end_date') or '').strip()
    sheet_cfg = payload.get('sheet_config') or {'type': 'grouped', 'args': ['category'], 'sheet_name': 'Kategorie'}
    filters = payload.get('filters') or []

    conn = ta.init_db()
    try:
        cur = conn.cursor()
        cur.execute("SELECT followed_users, trusted_users, followers FROM user_login_data WHERE owner = ?", (creator,))
        r_priv = cur.fetchone()
        followed = set(json.loads(r_priv[0])) if r_priv and r_priv[0] else set()
        trusted = set(json.loads(r_priv[1])) if r_priv and r_priv[1] else set()
        followers = set(json.loads(r_priv[2])) if r_priv and r_priv[2] else set()
        ta.set_auth_user(creator)
        ta.set_followed_users(followed)
        ta.set_trusted_users(trusted)
        ta.set_followers_users(followers)

        stype = sheet_cfg.get('type', 'grouped')
        use_index = bool(sheet_cfg.get('index', stype in ('stops', 'stops_via', 'geo')))
        target_user = users[0] if users else creator
        kwargs = dict(conn=conn, username=target_user, start_date=start_date, end_date=end_date)

        group_args = []
        group_labels = []

        with ta.interactive_filter_context(conn, users if mode == 'shared' else [target_user], start_date, end_date, filters):
            summary = _compute_filtered_summary(conn, users if mode == 'shared' else [target_user], start_date, end_date)
            df = None

            if stype == 'meta':
                df = ta.create_user_meta_dataframe(conn, target_user)
            elif stype == 'journeys':
                df = ta.create_journeys_dataframe(**kwargs)
            elif stype == 'stats':
                df = ta.create_user_stats_dataframe(**kwargs)
            elif stype == 'stops':
                df = ta.create_user_stops_dataframe(**kwargs)
            elif stype == 'stops_via':
                df = ta.create_user_stops_dataframe_with_inter(**kwargs)
            elif stype == 'extended_stats':
                df = ta.create_user_extended_stats(**kwargs)
            elif stype == 'grouped':
                args = [a for a in sheet_cfg.get('args', ['category']) if a] or ['category']
                group_args = list(dict.fromkeys(args))
                group_labels = [ta.AVAILABLE_GROUP_COLUMNS.get(a, {}).get('label', a) for a in group_args]
                df = ta.create_user_grouped_frame(
                    **kwargs,
                    first_argument=group_args[0],
                    second_argument=group_args[1] if len(group_args) > 1 else "",
                    third_argument=group_args[2] if len(group_args) > 2 else "",
                    fourth_argument=group_args[3] if len(group_args) > 3 else "",
                    sort_by_distance=sheet_cfg.get('sort_by_distance', True),
                    exclude_empty=sheet_cfg.get('exclude_empty', False)
                )
            elif stype == 'temporal':
                g_type = sheet_cfg.get('grouping_type', 'time_of_day')
                t_base = ta.get_temporal_base_dataframe(**kwargs, create_picture=False)
                df = ta.create_temporal_stats_dataframe(t_base, grouping_type=g_type, username=target_user)
            elif stype == 'geo':
                g_col = sheet_cfg.get('geo_column', 'countryName')
                df = ta.create_geo_dataframe(conn, target_user, g_col, start_date, end_date)
            elif stype == 'shared_meta':
                df = ta.create_shared_meta_dataframe(conn, users, start_date, end_date)
            elif stype == 'shared_stats':
                df = ta.create_shared_stats_dataframe(conn, users, start_date, end_date)
            elif stype in ('shared_stops', 'shared_stops_filterable'):
                df_reg, df_flt = ta.create_shared_stops_dataframes(conn, users, start_date, end_date)
                df = df_reg if stype == 'shared_stops' else df_flt
            elif stype == 'shared_journeys':
                df = ta.create_shared_journeys_dataframe(conn, users, start_date, end_date)
            elif stype == 'shared_grouped':
                args = [a for a in sheet_cfg.get('args', ['category']) if a] or ['category']
                group_args = list(dict.fromkeys(args))
                group_labels = [ta.AVAILABLE_GROUP_COLUMNS.get(a, {}).get('label', a) for a in group_args]
                df, _ = ta.create_shared_grouped_dataframe(
                    conn, users, group_args, start_date=start_date, end_date=end_date,
                    min_users=int(sheet_cfg.get('min_users', 1) or 1),
                    sort_by_distance=sheet_cfg.get('sort_by_distance', True),
                    exclude_empty=sheet_cfg.get('exclude_empty', False)
                )

            if df is not None and not df.empty:
                df = ta.apply_sheet_metrics(df, stype, sheet_cfg.get('metrics'), num_group_args=max(1, len(group_args)))

        cols, rows = _df_to_clean_records(df, use_index=use_index)
        return jsonify({
            'columns': cols,
            'rows': rows,
            'sheet_type': stype,
            'group_args': group_args,
            'group_labels': group_labels,
            'summary': summary
        })
    except Exception as e:
        tb_str = traceback.format_exc()
        print(tb_str)
        ta.log_error(username=creator, error=e, traceback_str=tb_str, context=f"api_interactive_query (sheet_type: {stype})")
        return jsonify({'error': str(e)}), 500
    finally:
        ta.clear_user_context()
        conn.close()


@app.route('/api/interactive/filter_values', methods=['POST'])
def api_interactive_filter_values():
    if 'username' not in session:
        return jsonify({'error': 'Nicht eingeloggt'}), 401
    creator = session['username']
    payload = request.get_json(silent=True) or {}

    col = (payload.get('col') or '').strip()
    search_q = (payload.get('q') or '').strip()
    users = payload.get('users') or [creator]
    start_date = (payload.get('start_date') or '').strip()
    end_date = (payload.get('end_date') or '').strip()
    active_filters = payload.get('filters') or []

    if not col:
        return jsonify({'values': []})

    conn = ta.init_db()
    try:
        cur = conn.cursor()
        cur.execute("SELECT followed_users, trusted_users, followers FROM user_login_data WHERE owner = ?", (creator,))
        r_priv = cur.fetchone()
        followed = set(json.loads(r_priv[0])) if r_priv and r_priv[0] else set()
        trusted = set(json.loads(r_priv[1])) if r_priv and r_priv[1] else set()
        followers = set(json.loads(r_priv[2])) if r_priv and r_priv[2] else set()
        ta.set_auth_user(creator)
        ta.set_followed_users(followed)
        ta.set_trusted_users(trusted)
        ta.set_followers_users(followers)

        placeholders = ','.join(['?'] * len(users))
        base_params = list(users)
        date_sql = ""
        if start_date:
            date_sql += " AND COALESCE(j.arrivalFinal, j.arrivalPlanned) > ?"
            base_params.append(start_date)
        if end_date:
            e_val = end_date + 'T23:59:59' if len(end_date) == 10 else end_date
            date_sql += " AND COALESCE(j.departureFinal, j.departurePlanned) < ?"
            base_params.append(e_val)

        results = []

        with ta.interactive_filter_context(conn, users, start_date, end_date, active_filters):
            vis_sql = ta.get_shared_visibility_query('j')

            if col == 'visibility':
                cur.execute(f"""
                    SELECT CAST(j.visibility AS TEXT) AS val, COUNT(*) AS cnt
                    FROM journeys j
                    WHERE j.username IN ({placeholders}) {date_sql} {vis_sql}
                      AND j.visibility IS NOT NULL
                    GROUP BY val
                    ORDER BY cnt DESC
                """, base_params)
                vis_labels = {
                    '0': '🌐 Öffentlich',
                    '1': '🔒 Nicht gelistet',
                    '2': '👥 Nur Follower',
                    '3': '🔐 Privat',
                    '4': '👤 Angemeldete Nutzer',
                    '5': '⭐ Enge Freunde'
                }
                q_low = search_q.lower()
                for r in cur.fetchall():
                    if r[0] is not None and str(r[0]) != '':
                        v_str = str(r[0])
                        lbl = vis_labels.get(v_str, f"Sichtbarkeit {v_str}")
                        if not q_low or q_low in lbl.lower() or q_low in v_str.lower():
                            results.append({'value': v_str, 'label': lbl, 'count': r[1]})

            elif col in ('realtimeAvailability', 'realtimeFromDataprovider'):
                cur.execute(f"""
                    SELECT CAST(j.{col} AS TEXT) AS val, COUNT(*) AS cnt
                    FROM journeys j
                    WHERE j.username IN ({placeholders}) {date_sql} {vis_sql}
                      AND j.{col} IS NOT NULL
                    GROUP BY val
                    ORDER BY cnt DESC
                """, base_params)
                q_low = search_q.lower()
                for r in cur.fetchall():
                    if r[0] is not None and str(r[0]) != '':
                        v_str = str(r[0])
                        lbl = '✅ Ja' if v_str in ('1', 'True', 'true') else '❌ Nein'
                        if not q_low or q_low in lbl.lower() or q_low in v_str.lower():
                            results.append({'value': v_str, 'label': lbl, 'count': r[1]})

            elif col in ta.AVAILABLE_GROUP_COLUMNS:
                sql_expr = ta.AVAILABLE_GROUP_COLUMNS[col]['sql']
                like_sql = ""
                q_params = list(base_params)
                if search_q:
                    like_sql = f" AND CAST({sql_expr} AS TEXT) LIKE ?"
                    q_params.append(f"%{search_q}%")

                cur.execute(f"""
                    SELECT CAST({sql_expr} AS TEXT) AS val, COUNT(*) AS cnt
                    FROM journeys j
                    LEFT JOIN stops origin ON j.originId = origin.id
                    LEFT JOIN stops destination ON j.destinationId = destination.id
                    WHERE j.username IN ({placeholders}) {date_sql} {vis_sql}
                      AND {sql_expr} IS NOT NULL AND CAST({sql_expr} AS TEXT) != '' AND CAST({sql_expr} AS TEXT) != 'N/A'
                      {like_sql}
                    GROUP BY val
                    ORDER BY cnt DESC, val ASC
                    LIMIT 60
                """, q_params)
                results = [{'value': str(r[0]), 'label': str(r[0]), 'count': r[1]} for r in cur.fetchall() if r[0]]

            elif col == 'any_stop':
                like_sql = ""
                q_params = list(base_params) + list(base_params)
                if search_q:
                    like_sql = " WHERE val LIKE ?"
                    q_params.append(f"%{search_q}%")

                cur.execute(f"""
                    SELECT val, COUNT(*) AS cnt FROM (
                        SELECT origin.stopName AS val
                        FROM journeys j
                        JOIN stops origin ON j.originId = origin.id
                        WHERE j.username IN ({placeholders}) {date_sql} {vis_sql} AND origin.stopName IS NOT NULL
                        UNION ALL
                        SELECT destination.stopName AS val
                        FROM journeys j
                        JOIN stops destination ON j.destinationId = destination.id
                        WHERE j.username IN ({placeholders}) {date_sql} {vis_sql} AND destination.stopName IS NOT NULL
                    )
                    {like_sql}
                    GROUP BY val
                    ORDER BY cnt DESC, val ASC
                    LIMIT 60
                """, q_params)
                results = [{'value': str(r[0]), 'label': str(r[0]), 'count': r[1]} for r in cur.fetchall() if r[0]]

            elif col == 'any_stop_via':
                from collections import Counter
                stop_counter = Counter()
                cur.execute(f"""
                    SELECT origin.stopName, destination.stopName, j.interStops
                    FROM journeys j
                    LEFT JOIN stops origin ON j.originId = origin.id
                    LEFT JOIN stops destination ON j.destinationId = destination.id
                    WHERE j.username IN ({placeholders}) {date_sql} {vis_sql}
                """, base_params)
                q_low = search_q.lower()
                for o_name, d_name, inter_json in cur.fetchall():
                    for s in (o_name, d_name):
                        if s and (not q_low or q_low in s.lower()):
                            stop_counter[s] += 1
                    if inter_json:
                        try:
                            for s in json.loads(inter_json):
                                if s and (not q_low or q_low in s.lower()):
                                    stop_counter[s] += 1
                        except Exception:
                            pass
                results = [{'value': val, 'label': val, 'count': cnt} for val, cnt in stop_counter.most_common(60)]

            elif col == 'stop_via':
                from collections import Counter
                stop_counter = Counter()
                cur.execute(f"""
                    SELECT j.interStops
                    FROM journeys j
                    WHERE j.username IN ({placeholders}) {date_sql} {vis_sql}
                      AND j.interStops IS NOT NULL AND j.interStops != ''
                """, base_params)
                q_low = search_q.lower()
                for (inter_json,) in cur.fetchall():
                    if inter_json:
                        try:
                            stops_list = json.loads(inter_json)
                            if isinstance(stops_list, list):
                                for s in stops_list:
                                    if s and (not q_low or q_low in s.lower()):
                                        stop_counter[s] += 1
                        except Exception:
                            pass
                results = [{'value': val, 'label': val, 'count': cnt} for val, cnt in stop_counter.most_common(60)]

            elif col.startswith('geo_'):
                g_col = col.replace('geo_', '')
                if g_col in ta.AVAILABLE_GEO_COLUMNS:
                    like_sql = ""
                    q_params = list(base_params) + list(base_params)
                    if search_q:
                        like_sql = " WHERE val LIKE ?"
                        q_params.append(f"%{search_q}%")

                    cur.execute(f"""
                        SELECT val, COUNT(*) AS cnt FROM (
                            SELECT origin.{g_col} AS val
                            FROM journeys j
                            JOIN stops origin ON j.originId = origin.id
                            WHERE j.username IN ({placeholders}) {date_sql} {vis_sql} AND origin.{g_col} IS NOT NULL AND origin.{g_col} != ''
                            UNION ALL
                            SELECT destination.{g_col} AS val
                            FROM journeys j
                            JOIN stops destination ON j.destinationId = destination.id
                            WHERE j.username IN ({placeholders}) {date_sql} {vis_sql} AND destination.{g_col} IS NOT NULL AND destination.{g_col} != ''
                        )
                        {like_sql}
                        GROUP BY val
                        ORDER BY cnt DESC, val ASC
                        LIMIT 60
                    """, q_params)
                    results = [{'value': str(r[0]), 'label': str(r[0]), 'count': r[1]} for r in cur.fetchall() if r[0]]

        return jsonify({'values': results})
    except Exception as e:
        ta.log_error(username=creator, error=e, context="api_interactive_filter_values")
        return jsonify({'error': str(e), 'values': []}), 500
    finally:
        ta.clear_user_context()
        conn.close()


@app.route('/api/interactive/visual', methods=['POST'])
def api_interactive_visual():
    if 'username' not in session:
        return jsonify({'error': 'Nicht eingeloggt'}), 401
    creator = session['username']
    payload = request.get_json(silent=True) or {}

    vtype = payload.get('visual_type')
    users = payload.get('users') or [creator]
    start_date = (payload.get('start_date') or '').strip()
    end_date = (payload.get('end_date') or '').strip()
    filters = payload.get('filters') or []
    lite = bool(payload.get('lite_version', False))
    g_type = payload.get('grouping_type', 'time_of_day')

    conn = ta.init_db()
    try:
        cur = conn.cursor()
        cur.execute("SELECT access_token, followed_users, trusted_users, followers FROM user_login_data WHERE owner = ?", (creator,))
        row = cur.fetchone()
        api_token = row[0] if row else ''
        followed = set(json.loads(row[1])) if row and row[1] else set()
        trusted = set(json.loads(row[2])) if row and row[2] else set()
        followers = set(json.loads(row[3])) if row and row[3] else set()
        ta.set_auth_user(creator)
        ta.set_followed_users(followed)
        ta.set_trusted_users(trusted)
        ta.set_followers_users(followers)

        filename = None
        folder = None

        with ta.interactive_filter_context(conn, users, start_date, end_date, filters):
            if vtype == 'polyline':
                folder = 'polylines'
                filename = ta.create_polyline(conn, api_token, usernames=users, start_date=start_date, end_date=end_date, lite_version=lite, creator_username=creator)
            elif vtype == 'stops_map':
                folder = 'polylines'
                filename = ta.create_stops_map(conn, usernames=users, start_date=start_date, end_date=end_date, lite_version=lite, creator_username=creator)
            elif vtype == 'timeline':
                folder = 'timelines'
                filename = ta.create_timeline(conn, usernames=users, start_date=start_date, end_date=end_date, creator_username=creator)
            elif vtype == 'ganglinie':
                folder = 'gang_linien'
                filename = ta.create_ganglinie_plot(username=users[0], grouping_type=g_type, conn=conn, start_date=start_date, end_date=end_date, creator_username=creator)
            elif vtype == 'delay_plot':
                folder = 'verspaetungs_plots'
                files = ta.create_verspaetungs_verteilungen_plot(users[0], conn=conn, start_date=start_date, end_date=end_date, creator_username=creator)
                filename = files[0] if files else None

        if not filename or not folder:
            return jsonify({'error': 'Für diese Auswahl konnten keine Daten visualisiert werden.'}), 404

        return jsonify({
            'filename': filename,
            'view_url': f"/files/view/{folder}/{filename}",
            'download_url': f"/files/download/{folder}/{filename}"
        })
    except Exception as e:
        ta.log_error(username=creator, error=e, context=f"api_interactive_visual ({vtype})")
        return jsonify({'error': str(e)}), 500
    finally:
        ta.clear_user_context()
        conn.close()


@app.route('/api/interactive/status_detail/<int:status_id>')
def api_interactive_status_detail(status_id):
    if 'username' not in session:
        return jsonify({'error': 'Nicht eingeloggt'}), 401
    conn = ta.init_db()
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("""
        SELECT j.*, o.stopName as originName, d.stopName as destName
        FROM journeys j
        LEFT JOIN stops o ON j.originId = o.id
        LEFT JOIN stops d ON j.destinationId = d.id
        WHERE j.statusId = ?
    """, (status_id,))
    row = cur.fetchone()
    conn.close()
    if not row:
        return jsonify({'error': 'Fahrt nicht in Datenbank gefunden.'}), 404
    d = dict(row)
    inter = []
    if d.get('interStops'):
        try:
            inter = json.loads(d['interStops'])
        except Exception:
            pass

    logged_user = session.get('username')
    is_own_status = bool(logged_user and (d.get('username') == logged_user or is_admin_user(logged_user)))

    return jsonify({
        'journey': d,
        'inter_stops': inter,
        'is_own_status': is_own_status,
        'logged_in_user': logged_user
    })


@app.route('/api/interactive/export_table', methods=['POST'])
def api_interactive_export_table():
    if 'username' not in session:
        return jsonify({'error': 'Nicht eingeloggt'}), 401
    creator = session['username']
    payload = request.get_json(silent=True) or {}
    mode = payload.get('mode', 'user')
    users = payload.get('users') or [creator]
    start_date = (payload.get('start_date') or '').strip()
    end_date = (payload.get('end_date') or '').strip()
    sheet_cfg = payload.get('sheet_config') or {'type': 'grouped', 'args': ['category'], 'sheet_name': 'Export'}
    filters = payload.get('filters') or []

    conn = ta.init_db()
    try:
        with ta.user_auth_context_from_db(conn, creator):
            with ta.interactive_filter_context(conn, users, start_date, end_date, filters):
                if mode == 'shared':
                    fname = ta.create_shared_excel_file(conn, usernames=users, start_date=start_date, end_date=end_date, custom_settings={'shared_sheets': [sheet_cfg]}, creator_username=creator)
                    folder = 'shared'
                else:
                    fname = ta.create_user_excel_file(conn, username=users[0], start_date=start_date, end_date=end_date, create_picture=False, custom_settings={'user_sheets': [sheet_cfg]}, creator_username=creator)
                    folder = 'user'
        return jsonify({'download_url': f"/files/download/{folder}/{fname}"})
    except Exception as e:
        ta.log_error(username=creator, error=e, context=f"api_interactive_export_table ({mode})")
        return jsonify({'error': str(e)}), 500
    finally:
        conn.close()


# ==========================================
# KI-DATENBANK-ASSISTENT ROUTEN (/ai)
# ==========================================

@app.route('/ai')
def ai_page():
    if 'username' not in session:
        return redirect(url_for('login'))
    username = session['username']

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
    cursor.execute("SELECT username FROM users WHERE username IS NOT NULL")
    for r in cursor.fetchall():
        if r[0]:
            available_users.add(r[0])
    cursor.execute("SELECT gemini_api_key FROM user_login_data WHERE owner = ?", (username,))
    byok_row = cursor.fetchone()
    has_byok = bool(byok_row and byok_row[0] and str(byok_row[0]).strip())
    is_admin = is_admin_user(username)
    rate_status = ta.get_rate_limit_status(conn, username, is_admin=is_admin)

    conn.close()

    has_gemini_key = bool(get_env_val('GEMINI_API_KEY'))

    return render_template(
        'ai.html',
        username=username,
        available_users=sorted(list(available_users)),
        has_gemini_key=has_gemini_key,
        has_byok=has_byok,
        is_admin=is_admin,
        rate_status=rate_status
    )


@app.route('/api/ai/set_key', methods=['POST'])
def api_ai_set_key():
    if 'username' not in session:
        return jsonify({'error': 'Nicht eingeloggt'}), 401
    if not is_admin_user(session['username']):
        return jsonify({'error': 'Nur Administratoren dürfen den globalen Instanz-Key in der .env ändern.'}), 403
    payload = request.get_json(silent=True) or {}
    key = (payload.get('gemini_api_key') or '').strip()
    if not key:
        return jsonify({'success': False, 'error': 'Kein API-Key angegeben.'}), 400
    save_env_val('GEMINI_API_KEY', key)
    return jsonify({'success': True})


@app.route('/api/ai/byok/save', methods=['POST'])
def api_ai_byok_save():
    if 'username' not in session:
        return jsonify({'error': 'Nicht eingeloggt'}), 401
    username = session['username']
    payload = request.get_json(silent=True) or {}
    key = (payload.get('gemini_api_key') or '').strip()
    if not key:
        return jsonify({'success': False, 'error': 'Kein API-Key angegeben.'}), 400

    conn = ta.init_db()
    cursor = conn.cursor()
    cursor.execute("UPDATE user_login_data SET gemini_api_key = ? WHERE owner = ?", (key, username))
    conn.commit()
    conn.close()
    return jsonify({'success': True})


@app.route('/api/ai/byok/delete', methods=['POST'])
def api_ai_byok_delete():
    if 'username' not in session:
        return jsonify({'error': 'Nicht eingeloggt'}), 401
    username = session['username']

    conn = ta.init_db()
    cursor = conn.cursor()
    cursor.execute("UPDATE user_login_data SET gemini_api_key = NULL WHERE owner = ?", (username,))
    conn.commit()
    conn.close()
    return jsonify({'success': True})


@app.route('/api/ai/byok/status')
def api_ai_byok_status():
    if 'username' not in session:
        return jsonify({'error': 'Nicht eingeloggt'}), 401
    username = session['username']
    is_admin = is_admin_user(username)

    conn = ta.init_db()
    status = ta.get_rate_limit_status(conn, username, is_admin=is_admin)
    conn.close()
    return jsonify({'success': True, **status})


@app.route('/api/ai/chats')
def api_ai_chats():
    if 'username' not in session:
        return jsonify({'error': 'Nicht eingeloggt'}), 401
    username = session['username']
    conn = ta.init_db()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT id, title, selected_users, start_date, end_date, created_at, updated_at FROM ai_chats WHERE owner = ? ORDER BY updated_at DESC",
        (username,)
    )
    chats = []
    for r in cursor.fetchall():
        chats.append({
            'id': r[0],
            'title': r[1],
            'selected_users': json.loads(r[2]) if r[2] else [username],
            'start_date': r[3] or '',
            'end_date': r[4] or '',
            'created_at': r[5],
            'updated_at': r[6]
        })
    conn.close()
    return jsonify({'chats': chats})


@app.route('/api/ai/chats/<chat_id>')
def api_ai_get_chat(chat_id):
    if 'username' not in session:
        return jsonify({'error': 'Nicht eingeloggt'}), 401
    username = session['username']
    conn = ta.init_db()
    cursor = conn.cursor()
    cursor.execute("SELECT id, owner, title, selected_users, start_date, end_date, created_at, updated_at FROM ai_chats WHERE id = ?", (chat_id,))
    chat_row = cursor.fetchone()
    if not chat_row or chat_row[1] != username:
        conn.close()
        return jsonify({'error': 'Chat nicht gefunden oder keine Berechtigung'}), 404

    cursor.execute(
        "SELECT id, role, content, sql_query, row_count, sql_result, created_at FROM ai_messages WHERE chat_id = ? ORDER BY id ASC",
        (chat_id,)
    )
    messages = []
    for m in cursor.fetchall():
        sql_res = None
        if m[5]:
            try:
                sql_res = json.loads(m[5])
            except Exception:
                sql_res = m[5]
        messages.append({
            'id': m[0],
            'role': m[1],
            'content': m[2],
            'sql_query': m[3],
            'row_count': m[4],
            'sql_result': sql_res,
            'created_at': m[6]
        })

    meta = {
        'id': chat_row[0],
        'title': chat_row[2],
        'selected_users': json.loads(chat_row[3]) if chat_row[3] else [username],
        'start_date': chat_row[4] or '',
        'end_date': chat_row[5] or '',
        'created_at': chat_row[6] or '',
        'updated_at': chat_row[7] or ''
    }
    conn.close()
    return jsonify({'chat': meta, 'messages': messages})


@app.route('/api/ai/chat/backup_latest', methods=['GET'])
def api_ai_chat_backup_latest():
    """Backup-Endpunkt: Gibt NUR den Antworttext der letzten KI-Nachricht aus der DB zurück (ohne SQL-Query/Result)."""
    if 'username' not in session:
        return jsonify({'error': 'Nicht eingeloggt'}), 401
    username = session['username']
    chat_id = (request.args.get('chat_id') or '').strip()

    conn = ta.init_db()
    cursor = conn.cursor()
    try:
        # Falls currentChatId beim ersten Request im Frontend noch null war, neuesten Chat des Users holen
        if not chat_id:
            cursor.execute(
                "SELECT id, title FROM ai_chats WHERE owner = ? ORDER BY updated_at DESC LIMIT 1",
                (username,)
            )
            chat_row = cursor.fetchone()
            if not chat_row:
                return jsonify({'success': False, 'error': 'Kein Chat gefunden.'}), 404
            chat_id, chat_title = chat_row[0], chat_row[1]
        else:
            cursor.execute(
                "SELECT id, title FROM ai_chats WHERE id = ? AND owner = ?",
                (chat_id, username)
            )
            chat_row = cursor.fetchone()
            if not chat_row:
                return jsonify({'success': False, 'error': 'Chat nicht gefunden oder keine Berechtigung.'}), 404
            chat_id, chat_title = chat_row[0], chat_row[1]

        # Nur den reinen Text (content) der neuesten Model-Antwort laden (kein sql_query, kein sql_result!)
        cursor.execute(
            """SELECT id, content, created_at 
               FROM ai_messages 
               WHERE chat_id = ? AND role = 'model' 
               ORDER BY id DESC LIMIT 1""",
            (chat_id,)
        )
        msg_row = cursor.fetchone()
        if not msg_row:
            return jsonify({'success': False, 'error': 'Noch keine KI-Antwort in der Datenbank gespeichert.'}), 404

        return jsonify({
            'success': True,
            'is_backup': True,
            'chat_id': chat_id,
            'title': chat_title,
            'message_id': msg_row[0],
            'answer': msg_row[1] or '',
            'created_at': msg_row[2]
        })
    except Exception as e:
        ta.log_error(username=username, error=e, context="api_ai_chat_backup_latest", conn=conn)
        return jsonify({'success': False, 'error': str(e)}), 500
    finally:
        conn.close()


@app.route('/api/ai/chats/<chat_id>/backup_messages', methods=['GET'])
def api_ai_get_chat_backup(chat_id):
    """Backup-Endpunkt für ganze Chats: Lädt alle Nachrichten eines Chats OHNE SQL-Queries und SQL-Ergebnisse."""
    if 'username' not in session:
        return jsonify({'error': 'Nicht eingeloggt'}), 401
    username = session['username']
    conn = ta.init_db()
    cursor = conn.cursor()
    try:
        cursor.execute(
            "SELECT id, owner, title, selected_users, start_date, end_date, created_at, updated_at FROM ai_chats WHERE id = ?",
            (chat_id,)
        )
        chat_row = cursor.fetchone()
        if not chat_row or chat_row[1] != username:
            return jsonify({'error': 'Chat nicht gefunden oder keine Berechtigung'}), 404

        cursor.execute(
            "SELECT id, role, content, created_at FROM ai_messages WHERE chat_id = ? ORDER BY id ASC",
            (chat_id,)
        )
        messages = [
            {
                'id': m[0],
                'role': m[1],
                'content': m[2] or '',
                'sql_query': None,
                'row_count': None,
                'sql_result': None,
                'created_at': m[3]
            }
            for m in cursor.fetchall()
        ]
        meta = {
            'id': chat_row[0],
            'title': chat_row[2],
            'selected_users': json.loads(chat_row[3]) if chat_row[3] else [username],
            'start_date': chat_row[4] or '',
            'end_date': chat_row[5] or '',
            'created_at': chat_row[6] or '',
            'updated_at': chat_row[7] or ''
        }
        return jsonify({'chat': meta, 'messages': messages, 'is_backup': True})
    finally:
        conn.close()


@app.route('/api/ai/messages/<int:message_id>/sql_details', methods=['GET'])
def api_ai_message_sql_details(message_id):
    """Versucht, die SQL-Abfrage und das SQL-Ergebnis einer einzelnen Nachricht bereinigt nachzuladen."""
    if 'username' not in session:
        return jsonify({'error': 'Nicht eingeloggt'}), 401
    username = session['username']

    conn = ta.init_db()
    cursor = conn.cursor()
    try:
        cursor.execute(
            """SELECT m.sql_query, m.row_count, m.sql_result, c.owner
               FROM ai_messages m
               JOIN ai_chats c ON m.chat_id = c.id
               WHERE m.id = ?""",
            (message_id,)
        )
        row = cursor.fetchone()
        if not row or row[3] != username:
            return jsonify({'success': False, 'error': 'Nachricht nicht gefunden oder keine Berechtigung.'}), 404

        sql_query, row_count, raw_sql_result, _ = row

        def deep_sanitize(obj):
            """Bereinigt rekursiv alle Strings, NaN/Inf-Floats und kürzt überlange Texte für sicheres JSON."""
            if obj is None:
                return None
            if isinstance(obj, float):
                if math.isnan(obj) or math.isinf(obj):
                    return None
                return round(obj, 4)
            if isinstance(obj, str):
                cleaned = "".join(ch for ch in obj if ord(ch) >= 32 or ch in "\t\n\r")
                return cleaned[:1000] + ("..." if len(cleaned) > 1000 else "")
            if isinstance(obj, list):
                return [deep_sanitize(x) for x in obj]
            if isinstance(obj, dict):
                return {str(k): deep_sanitize(v) for k, v in obj.items()}
            return obj

        parsed_result = None
        result_recovered = False
        if raw_sql_result:
            try:
                raw_parsed = json.loads(raw_sql_result)
                parsed_result = deep_sanitize(raw_parsed)
                result_recovered = True
            except Exception as parse_err:
                ta.log_error(
                    username=username,
                    error=parse_err,
                    context=f"api_ai_message_sql_details (message_id={message_id} sql_result defekt, liefere nur sql_query)",
                    conn=conn
                )
                parsed_result = None

        return jsonify({
            'success': True,
            'message_id': message_id,
            'sql': deep_sanitize(sql_query) if sql_query else None,
            'row_count': row_count,
            'sql_result': parsed_result,
            'only_query_recovered': bool(sql_query and not result_recovered and raw_sql_result)
        })
    except Exception as e:
        ta.log_error(username=username, error=e, context=f"api_ai_message_sql_details ({message_id})", conn=conn)
        return jsonify({'success': False, 'error': str(e)}), 500
    finally:
        conn.close()


@app.route('/api/ai/chats/<chat_id>/export')
def api_ai_export_chat(chat_id):
    if 'username' not in session:
        return jsonify({'error': 'Nicht eingeloggt'}), 401
    username = session['username']
    conn = ta.init_db()
    cursor = conn.cursor()
    cursor.execute("SELECT id, owner, title, selected_users, start_date, end_date, created_at, updated_at FROM ai_chats WHERE id = ?", (chat_id,))
    chat_row = cursor.fetchone()
    if not chat_row or chat_row[1] != username:
        conn.close()
        return jsonify({'error': 'Chat nicht gefunden oder keine Berechtigung'}), 404

    cursor.execute(
        "SELECT id, role, content, sql_query, row_count, sql_result, created_at FROM ai_messages WHERE chat_id = ? ORDER BY id ASC",
        (chat_id,)
    )
    messages = []
    for m in cursor.fetchall():
        sql_res = None
        if m[5]:
            try:
                sql_res = json.loads(m[5])
            except Exception:
                sql_res = m[5]
        messages.append({
            'id': m[0],
            'role': m[1],
            'content': m[2],
            'sql_query': m[3],
            'row_count': m[4],
            'sql_result': sql_res,
            'created_at': m[6]
        })

    export_data = {
        'export_version': '1.0',
        'exported_at': datetime.now().isoformat(),
        'chat': {
            'id': chat_row[0],
            'owner': chat_row[1],
            'title': chat_row[2],
            'selected_users': json.loads(chat_row[3]) if chat_row[3] else [username],
            'start_date': chat_row[4] or '',
            'end_date': chat_row[5] or '',
            'created_at': chat_row[6] or '',
            'updated_at': chat_row[7] or ''
        },
        'messages': messages
    }
    conn.close()

    clean_title = re.sub(r"[^a-zA-Z0-9_\-]", "_", chat_row[2] or 'Chat')[:35]
    filename = f"traewelling_chat_{clean_title}_{chat_id[:8]}.json"

    response = jsonify(export_data)
    response.headers['Content-Disposition'] = f'attachment; filename="{filename}"'
    return response


@app.route('/api/ai/chats/new', methods=['POST'])
def api_ai_new_chat():
    if 'username' not in session:
        return jsonify({'error': 'Nicht eingeloggt'}), 401
    username = session['username']
    chat_id = str(uuid.uuid4())
    now_iso = datetime.now().isoformat()
    conn = ta.init_db()
    cursor = conn.cursor()
    cursor.execute(
        """INSERT INTO ai_chats (id, owner, title, selected_users, start_date, end_date, created_at, updated_at)
           VALUES (?, ?, ?, ?, NULL, NULL, ?, ?)""",
        (chat_id, username, "Neuer Chat", json.dumps([username]), now_iso, now_iso)
    )
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'chat_id': chat_id, 'title': 'Neuer Chat'})


@app.route('/api/ai/chats/<chat_id>', methods=['DELETE'])
def api_ai_delete_chat(chat_id):
    if 'username' not in session:
        return jsonify({'error': 'Nicht eingeloggt'}), 401
    username = session['username']
    conn = ta.init_db()
    cursor = conn.cursor()
    cursor.execute("SELECT owner FROM ai_chats WHERE id = ?", (chat_id,))
    row = cursor.fetchone()
    if not row or row[0] != username:
        conn.close()
        return jsonify({'error': 'Chat nicht gefunden oder keine Berechtigung'}), 404

    cursor.execute("DELETE FROM ai_messages WHERE chat_id = ?", (chat_id,))
    cursor.execute("DELETE FROM ai_chats WHERE id = ?", (chat_id,))
    conn.commit()
    conn.close()
    return jsonify({'success': True})


@app.route('/api/ai/chat', methods=['POST'])
def api_ai_chat():
    if 'username' not in session:
        return jsonify({'error': 'Nicht eingeloggt'}), 401
    username = session['username']
    payload = request.get_json(silent=True) or {}
    message = (payload.get('message') or '').strip()
    if not message:
        return jsonify({'success': False, 'error': 'Keine Nachricht eingegeben.'}), 400

    if message.lower().rstrip('!?.') in ('koch kaffee', 'koche kaffee', 'mach kaffee', 'brew coffee', '/coffee', 'kaffee'):
        err_msg = "418 I'm a teapot: Ich bin eine Bahn-Datenbank auf tee_beutel-Basis, keine Kaffeemaschine! 🫖"
        ta.log_error(username=username, error_type="HTTP_418", error_message=err_msg, context=f"Fast-Path Teapot ({message})")
        return jsonify({
            'success': False,
            'is_teapot': True,
            'error': err_msg
        }), 418

    chat_id = (payload.get('chat_id') or '').strip()
    users = payload.get('users') or [username]
    start_date = (payload.get('start_date') or '').strip()
    end_date = (payload.get('end_date') or '').strip()

    conn = ta.init_db()
    cursor = conn.cursor()

    if not chat_id:
        chat_id = str(uuid.uuid4())
        now_iso = datetime.now().isoformat()
        cursor.execute(
            """INSERT INTO ai_chats (id, owner, title, selected_users, start_date, end_date, created_at, updated_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (chat_id, username, "Neuer Chat", json.dumps(users), start_date or None, end_date or None, now_iso, now_iso)
        )
        conn.commit()
    else:
        cursor.execute("SELECT owner FROM ai_chats WHERE id = ?", (chat_id,))
        row = cursor.fetchone()
        if not row:
            now_iso = datetime.now().isoformat()
            cursor.execute(
                """INSERT INTO ai_chats (id, owner, title, selected_users, start_date, end_date, created_at, updated_at)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                (chat_id, username, "Neuer Chat", json.dumps(users), start_date or None, end_date or None, now_iso, now_iso)
            )
            conn.commit()
        elif row[0] != username:
            conn.close()
            return jsonify({'error': 'Keine Berechtigung für diesen Chat'}), 403

    is_admin = is_admin_user(username)
    try:
        res = ta.process_ai_chat_turn(
            conn=conn,
            creator=username,
            chat_id=chat_id,
            user_message=message,
            selected_users=users,
            start_date=start_date,
            end_date=end_date,
            is_admin=is_admin
        )
        if res.get('is_teapot'):
            return jsonify(res), 418
        return jsonify(res)
    except ta.RateLimitExceededError as rle:
        return jsonify({'success': False, 'error': str(rle), 'rate_limited': True}), 429
    except Exception as e:
        tb_str = traceback.format_exc()
        ta.log_error(
            username=username,
            error=e,
            traceback_str=tb_str,
            context=f"api_ai_chat route (message: {message[:50]})",
            conn=conn
        )
        return jsonify({'success': False, 'error': 'Interner Serverfehler: ' + str(e)}), 500
    finally:
        conn.close()

#tee_beutel

@app.route('/wp-admin')
@app.route('/wp-login.php')
@app.route('/.env')
def teapot_trap():
    return jsonify({'error': "I'm a teapot"}), 418

@app.route('/coffee')
@app.route('/kaffee')
def brew_coffee():
    try:
        username = session['username']
    except:
        username = 'Gast'

    ta.log_error(username=username, error_type="HTTP_418", error_message="418 I'm a teapot",
                 context=f"Entschuldigung, dieser Server läuft ausschließlich auf <b>tee_beutel</b>-Basis und kann keinen Kaffee kochen")

    return render_template_string("""
        <h1>418 I'm a teapot</h1>
        <p>Entschuldigung, dieser Server läuft ausschließlich auf <b>tee_beutel</b>-Basis und kann keinen Kaffee kochen.</p>
        <p><a href="/">Zurück zum Bahnhof</a></p>
    """), 418



if __name__ == '__main__':
    conn = ta.init_db()
    # Vorher: app.run(host='0.0.0.0', port=8080, debug=True, ssl_context= 'adhoc')
    app.run(host='127.0.0.1', port=8080, debug=False)