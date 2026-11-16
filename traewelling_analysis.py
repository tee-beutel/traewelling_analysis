import json
import os
import hashlib
from dotenv import load_dotenv

load_dotenv()
import sqlite3
import math
from pathlib import Path
from zoneinfo import ZoneInfo
import numpy as np
import pandas as pd
import requests
from time import sleep, time
from datetime import datetime, timedelta
from collections import Counter, defaultdict
from timezonefinder import TimezoneFinder
import reverse_geocoder as rgeo
import pycountry
import babel
import re
import matplotlib.pyplot as plt

# ==========================================
# GLOBALE VARIABLEN & CACHES
# ==========================================
tf = None
AUTH_USER = ""
FOLLOWED_USERS = set()
FOLLOWERS_USERS = set()
TRUSTED_USERS = set()
remaining = "Unbekannt"


# ==========================================
# HILFS- UND API-FUNKTIONEN
# ==========================================


#Debug-Funktionen
def reload_single_status(conn: sqlite3.Connection, api_token: str, status_id: int) -> tuple[bool, dict]:
    """Lädt einen einzelnen Status neu, berechnet alle Werte, korrigiert DB
    und gibt alle geänderten Spalten und Werte für das Popup zurück."""
    cursor = conn.cursor()

    # 1. Vorherige Werte für alle Spalten auslesen
    cursor.execute("PRAGMA table_info(journeys)")
    columns = [col[1] for col in cursor.fetchall()]

    cursor.execute(f"SELECT {', '.join(columns)} FROM journeys WHERE statusId = ?", (status_id,))
    old_row = cursor.fetchone()
    old_values = dict(zip(columns, old_row)) if old_row else {}

    # 2. Status von API laden
    resp = requests.get(f'https://traewelling.de/api/v1/status/{status_id}', headers={
        'accept': 'application/json',
        'Authorization': f'Bearer {api_token}',
        'User-Agent': 'TraewellingAnalysis_v9 tee_beutel'
    })
    if resp.status_code != 200:
        return False, {}

    data = resp.json().get('data', {})
    if not data:
        return False, {}

    username = data.get('user', {}).get('username')
    user_id = data.get('user', {}).get('id')
    trip_id = data.get('checkin', {}).get('trip')

    # 3. Zwischenhalte laden
    inter_stations = {}
    if trip_id:
        inter_resp = requests.get(f"https://traewelling.de/api/v1/stopovers/{trip_id}", headers={
            'accept': 'application/json',
            'Authorization': f'Bearer {api_token}',
            'User-Agent': 'TraewellingAnalysis_v9 tee_beutel'
        })
        if inter_resp.status_code == 200:
            inter_data = inter_resp.json().get('data', {})
            if isinstance(inter_data, dict):
                inter_stations = inter_data
            elif isinstance(inter_data, list) and inter_data:
                inter_stations = {str(trip_id): inter_data}

    # 4. Polyline löschen (wichtig: nur gelöscht, nicht neu geladen)
    cursor.execute("DELETE FROM polylines WHERE statusId = ?", (status_id,))

    # 5. Haltestellen prüfen/laden
    train_data = data.get('checkin', {})
    origin = train_data.get('origin', {})
    destination = train_data.get('destination', {})
    origin_id, destination_id = origin.get('id'), destination.get('id')

    origin_timezone, origin_name, origin_lat, origin_lon = get_or_fetch_stop_info(conn, api_token, origin_id,
                                                                                  origin.get('name'))
    destination_timezone, destination_name, destination_lat, destination_lon = get_or_fetch_stop_info(conn, api_token,
                                                                                                      destination_id,
                                                                                                      destination.get(
                                                                                                          'name'))

    # 6. Werte berechnen
    train_type_to_readable_dict = {
        'regional': 'Regionalzug', 'bus': 'Bus', 'ferry': 'Fähre',
        'national': 'Regionalzug', 'regionalExp': 'Fernverkehr',
        'nationalExpress': 'Fernverkehr', 'suburban': 'S-Bahn', 'subway': 'U-Bahn',
        'taxi': 'Taxi', 'tram': 'Tram', 'plane': 'Igitt ein Flugzeug'
    }

    body = data.get('body', '').replace('\n', ' ')
    body_mentions = json.dumps(data.get('bodyMentions', []), ensure_ascii=False)

    business = ''
    match data.get('business', ''):
        case 0:
            business = 'Privat'
        case 1:
            business = 'Geschäftlich'
        case 2:
            business = 'Arbeitsweg'

    visibility = data.get('visibility', '')
    likes = data.get('likes', 0)
    is_likeable = data.get('isLikable', False)
    hide_body = data.get('hide_body', '')
    moderation_notes = data.get('moderation_notes', '')
    lock_visibility = data.get('lock_visibility', '')
    created_at = data.get('createdAt', '')
    hafas_id = data.get('checkin', {}).get('hafasId')

    line_name = re.sub(r'\s*\(\d+\)', '', train_data.get('lineName', 'Unknown Line')).strip()
    operator_name = (train_data.get('operator') or {}).get('name', 'Unknown Operator')
    operator_id = (train_data.get('operator') or {}).get('id', None)
    client_name = (data.get('client') or {}).get('name', 'Unknown Client')
    client_id = (data.get('client') or {}).get('id', 0)

    journey_distance = train_data.get('distance', 0)
    line_number = train_data.get('number')
    route_color = train_data.get('routeColor')
    route_text_color = train_data.get('routeTextColor')

    given_train_type = train_type_to_readable_dict.get(train_data.get('category', None),
                                                       train_data.get('category', None))
    if re.match(r"^(REX|RE|RB|Stoptrein|FEX|Sprinter)", line_name) and given_train_type != 'Bus':
        train_type = 'Regionalzug'
    elif re.match(r"^(RS)", line_name):
        train_type = 'S-Bahn'
    else:
        train_type = given_train_type

    journey_points = train_data.get('points', 0)
    journey_number = train_data.get('journeyNumber')

    dist_straight_line = calculate_haversine(origin_lat, origin_lon, destination_lat, destination_lon)

    # Eigene Distanz zur Vorfahrt berechnen
    departure_planned = origin.get('departurePlanned', None)
    departure_from_realtime = origin.get('departureReal', None)
    departure_manual = train_data.get('manualDeparture', None)
    arrival_planned = destination.get('arrivalPlanned', None)
    arrival_from_realtime = destination.get('arrivalReal', None)
    arrival_manual = train_data.get('manualArrival', None)

    departure_final = (departure_manual or departure_from_realtime or departure_planned)
    arrival_final = (arrival_manual or arrival_from_realtime or arrival_planned)

    cursor.execute(
        "SELECT destinationId FROM journeys WHERE username = ? AND departureFinal < ? ORDER BY departureFinal DESC LIMIT 1",
        (username, departure_final or created_at))
    prev_row = cursor.fetchone()
    dist_to_prev = 0
    if prev_row:
        prev_dest_id = prev_row[0]
        cursor.execute("SELECT latitude, longitude FROM stops WHERE id = ?", (prev_dest_id,))
        p_lat_lon = cursor.fetchone()
        if p_lat_lon:
            dist_to_prev = calculate_haversine(p_lat_lon[0], p_lat_lon[1], origin_lat, origin_lon)

    realtime_availability = bool(departure_manual or arrival_manual or departure_from_realtime or arrival_from_realtime)
    realtime_from_dataprovider = bool(departure_from_realtime or arrival_from_realtime)

    dp_dt = datetime.fromisoformat(departure_planned) if departure_planned else None
    df_dt = datetime.fromisoformat(departure_final) if departure_final else None
    ap_dt = datetime.fromisoformat(arrival_planned) if arrival_planned else None
    af_dt = datetime.fromisoformat(arrival_final) if arrival_final else None

    departure_delay = int((df_dt - dp_dt).total_seconds() / 60) if df_dt and dp_dt else 0
    arrival_delay = int((af_dt - ap_dt).total_seconds() / 60) if af_dt and ap_dt else 0
    duration_planned = int((ap_dt - dp_dt).total_seconds() / 60) if ap_dt and dp_dt else 0
    duration_real = int((af_dt - df_dt).total_seconds() / 60) if af_dt and df_dt else 0
    time_delta = duration_real - duration_planned

    if not realtime_availability:
        departure_final = arrival_final = None

    event = (json.dumps(data.get('event'), ensure_ascii=False) if data.get('event') else None)

    # Tags verarbeiten
    vehicle_number = locomotive_class = wagon_class = price = travel_class = ticket = None
    passenger_rights = role = seat = social_status = train_utilization = wagon = None

    tags = data.get('tags', None)
    if tags:
        for tag in tags:
            match tag.get('key', None):
                case 'trwl:vehicle_number':
                    vehicle_number = tag.get('value', 'N/A')
                case 'trwl:journey_number':
                    journey_number = tag.get('value', 'N/A')
                case 'trwl:locomotive_class':
                    locomotive_class = tag.get('value', 'N/A')
                case 'trwl:ticket':
                    ticket = tag.get('value', 'N/A')
                case 'trwl:wagon_class':
                    wagon_class = tag.get('value', 'N/A')
                case "trwl:price":
                    price = tag.get('value', 'N/A')
                case "trwl:travel_class":
                    travel_class = tag.get('value', 'N/A')
                case 'trwl:passenger_rights':
                    passenger_rights = tag.get('value', 'N/A')
                case 'trwl:role':
                    role = tag.get('value', 'N/A')
                case 'trwl:seat':
                    seat = tag.get('value', 'N/A')
                case 'trwl:social_status':
                    social_status = tag.get('value', 'N/A')
                case 'trwl:train_utilization':
                    train_utilization = tag.get('value', 'N/A')
                case 'trwl:wagon':
                    wagon = tag.get('value', 'N/A')

        tracked_keys = {'trwl:vehicle_number', 'trwl:journey_number', 'trwl:locomotive_class',
                        'trwl:ticket', 'trwl:wagon_class', 'trwl:price', 'trwl:travel_class',
                        'trwl:passenger_rights', 'trwl:role', 'trwl:seat', 'trwl:social_status',
                        'trwl:train_utilization', 'trwl:wagon'}
        tags = [tag for tag in tags if tag.get('key') not in tracked_keys]
        tags = (json.dumps(tags, ensure_ascii=False) if tags else None)
    else:
        tags = None

    try:
        inter_stations_list = inter_stations.get(str(trip_id), [])
        inter_stations_ids = [d['id'] for d in inter_stations_list]
        origin_index = inter_stations_ids.index(origin_id)
        destination_index = inter_stations_ids.index(destination_id, origin_index + 1)
        inter_stops_number = destination_index - origin_index - 1
        inter_stops = json.dumps([(f'{d["name"]} ({d.get("rilIdentifier")})' if d.get("rilIdentifier") else d["name"])
                                  for d in inter_stations_list[origin_index + 1:destination_index]], ensure_ascii=False)
    except Exception:
        inter_stops_number = None
        inter_stops = None

    # In Datenbank speichern
    cursor.execute(
        '''REPLACE INTO journeys(statusId, userId, username, body, bodyMentions, business, visibility, likes,
                                 isLikable, tripId, hafasId, lineName, operatorName, operatorId, distance,
                                 distanceStraightLine, distanceToPrevJourney, routeColor, routeTextColor, category,
                                 points, journeyNumber, clientId, clientName, originId, destinationId, lineNumber,
                                 moderationNotes, lockVisibility, hideBody, createdAt, departurePlanned,
                                 departureFromRealtime, departureManual, departureFinal, arrivalPlanned,
                                 arrivalFromRealtime, arrivalManual, arrivalFinal, realtimeAvailability,
                                 realtimeFromDataprovider, departureDelay, arrivalDelay, durationPlanned,
                                 durationReal, timeDelta, locomotiveClass, vehicleNumber, wagonClass, price,
                                 travelClass, ticket, passengerRights, role, seat, socialStatus, trainUtilization,
                                 wagon, tags, event, interStops, interStopsNumber)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                   ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
        (status_id, user_id, username, body, body_mentions, business, visibility, likes, is_likeable, trip_id,
         hafas_id, line_name, operator_name, operator_id, journey_distance, dist_straight_line, dist_to_prev,
         route_color, route_text_color, train_type, journey_points, journey_number, client_id, client_name,
         origin_id, destination_id, line_number, moderation_notes, lock_visibility, hide_body, created_at,
         departure_planned, departure_from_realtime, departure_manual, departure_final, arrival_planned,
         arrival_from_realtime, arrival_manual, arrival_final, realtime_availability, realtime_from_dataprovider,
         departure_delay, arrival_delay, duration_planned, duration_real, time_delta, locomotive_class,
         vehicle_number, wagon_class, price, travel_class, ticket, passenger_rights, role, seat, social_status,
         train_utilization, wagon, tags, event, inter_stops, inter_stops_number))

    # Folgefahrt aktualisieren
    cursor.execute(
        "SELECT statusId, originId FROM journeys WHERE username = ? AND departureFinal > ? ORDER BY departureFinal ASC LIMIT 1",
        (username, departure_final or created_at))
    next_row = cursor.fetchone()
    if next_row:
        next_status_id, next_origin_id = next_row[0], next_row[1]
        cursor.execute("SELECT latitude, longitude FROM stops WHERE id = ?", (next_origin_id,))
        n_lat_lon = cursor.fetchone()
        if n_lat_lon:
            new_dist_to_prev = calculate_haversine(destination_lat, destination_lon, n_lat_lon[0], n_lat_lon[1])
            cursor.execute("UPDATE journeys SET distanceToPrevJourney = ? WHERE statusId = ?",
                           (new_dist_to_prev, next_status_id))

    conn.commit()

    # 7. Alle neuen Werte auslesen und mit alten Werten vergleichen
    cursor.execute(f"SELECT {', '.join(columns)} FROM journeys WHERE statusId = ?", (status_id,))
    new_row = cursor.fetchone()
    new_values = dict(zip(columns, new_row)) if new_row else {}

    differences = {}
    for col in columns:
        old_val = old_values.get(col)
        new_val = new_values.get(col)
        if old_val != new_val:
            differences[col] = {'old': old_val, 'new': new_val}

    return True, differences

def get_date_suffix(start_date: str = '', end_date: str = '') -> str:
    """Generiert einen sauberen Dateinamen-Suffix abhängig vom gewählten Zeitraum."""
    s = start_date.strip().split('T')[0] if start_date else ''
    e = end_date.strip().split('T')[0] if end_date else ''
    if s and e:
        return f"_{s}_bis_{e}"
    if s:
        return f"_ab_{s}"
    if e:
        return f"_bis_{e}"
    return ""


def safe_log(callback, msg: str, replace: bool = False):
    """Ruft das übergebene Logging-Callback sicher auf."""
    if not callback:
        return
    try:
        callback(msg, replace=replace)
    except TypeError:
        callback(msg)


def fetch_paginated_users(api_token: str, start_url: str) -> set:
    users = set()
    url = start_url
    while url:
        try:
            resp = requests.get(url, headers={
                'accept': 'application/json',
                'Authorization': f'Bearer {api_token}',
                'User-Agent': 'TraewellingAnalysis_v9 tee_beutel'
            })
            resp.raise_for_status()
            data = resp.json()

            for item in data.get('data', []):
                # Dynamisches Auslesen (egal ob flach oder verschachtelt)
                username = None
                if isinstance(item, dict):
                    if 'username' in item:
                        username = item['username']
                    elif 'user' in item and isinstance(item['user'], dict) and 'username' in item['user']:
                        username = item['user']['username']
                elif isinstance(item, str):
                    username = item

                if username:
                    users.add(username)

            url = data.get('links', {}).get('next')
        except Exception as e:
            if isinstance(e, requests.exceptions.HTTPError) and e.response.status_code == 429:
                print("[API] Rate Limit (429) erreicht. Warte 30 Sekunden...")
                sleep(30)
            else:
                print(f"[API WARNUNG] Fehler beim Abrufen der Liste ({start_url}): {e}")
                break
    return users


def log_onto_traewelling(conn: sqlite3.Connection) -> str:
    global AUTH_USER, FOLLOWED_USERS, TRUSTED_USERS, FOLLOWERS_USERS, remaining
    api_token = os.getenv('TRAEWELLING_API_TOKEN', '').strip()
    if not api_token:
        token_path = Path.cwd() / '.api_token.txt'
        try:
            with open(token_path, 'r', encoding='utf-8') as token_file:
                api_token = token_file.read().strip()
        except FileNotFoundError:
            pass

    if not api_token or api_token == "Bitte hier deinen API-Token eintragen":
        raise ValueError('[FEHLER] Kein API Token in .env (TRAEWELLING_API_TOKEN) oder .api_token.txt gefunden.')

    print("\n" + "=" * 50)
    print("[LOGIN] Verbinde mit Träwelling...")

    repeat_bool = True
    while repeat_bool:
        login_response = requests.get('https://traewelling.de/api/v1/auth/user',
                                      headers={'accept': 'application/json',
                                               'Authorization': f'Bearer {api_token}',
                                               'User-Agent': 'TraewellingAnalysis_v9 tee_beutel'})
        try:
            login_response.raise_for_status()
            login_json = login_response.json()

            remaining = login_response.headers.get('x-ratelimit-remaining')
            username = login_json.get('data', {}).get('username', None)

            if username:
                AUTH_USER = username
                print(f"[LOGIN] ✅ Erfolgreich angemeldet als: {username}")
                print("=" * 50)
            repeat_bool = False
        except Exception as e:
            if isinstance(e, requests.exceptions.HTTPError) and e.response.status_code == 429:
                print('[API] Rate Limit erreicht. Warte 30 Sekunden beim Login...')
                sleep(30)
            else:
                raise PermissionError(
                    f'[FEHLER] Fehler bei der Authentifizierung! Ist der Token noch gültig?\nDetails: {e}')

    cursor = conn.cursor()
    cursor.execute("SELECT followed_users, trusted_users, followers, last_updated FROM user_login_data WHERE owner = ?",
                   (AUTH_USER,))
    row = cursor.fetchone()

    needs_update = True
    if row:
        try:
            last_updated_dt = datetime.fromisoformat(row[3])
            if datetime.now() - last_updated_dt < timedelta(hours=24):
                FOLLOWED_USERS = set(json.loads(row[0])) if row[0] else set()
                TRUSTED_USERS = set(json.loads(row[1])) if row[1] else set()
                FOLLOWERS_USERS = set(json.loads(row[2])) if row[2] else set()
                needs_update = False
                print("[INFO] Sichtbarkeits-Rechte aus lokalem Cache geladen.")
        except Exception:
            pass

    if needs_update:
        print("[INFO] Aktualisiere Sichtbarkeits-Rechte (Follows/Followers/Trusted) von der API...")

        # Die korrekten Endpunkte gemäß API-Dokumentation
        TRUSTED_USERS = fetch_paginated_users(api_token, 'https://traewelling.de/api/v1/user/self/trusted-by')
        FOLLOWED_USERS = fetch_paginated_users(api_token, 'https://traewelling.de/api/v1/user/self/followings')
        FOLLOWERS_USERS = fetch_paginated_users(api_token, 'https://traewelling.de/api/v1/user/self/followers')

        cursor.execute(
            "REPLACE INTO user_login_data (owner, followed_users, trusted_users, followers, last_updated) VALUES (?, ?, ?, ?, ?)",
            (AUTH_USER, json.dumps(list(FOLLOWED_USERS)), json.dumps(list(TRUSTED_USERS)),
             json.dumps(list(FOLLOWERS_USERS)), datetime.now().isoformat()))
        conn.commit()
        print("[INFO] ✅ Rechte erfolgreich aktualisiert.")

    return api_token


# ==========================================
# DATENBANK FUNKTIONEN (SQLITE)
# ==========================================

def init_db() -> sqlite3.Connection:
    path = Path.cwd() / 'traewelling.db'
    conn = sqlite3.connect(path)
    cursor = conn.cursor()

    cursor.execute('''CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY,
    username TEXT,
    displayname TEXT,
    bio TEXT,
    mastodonUrl TEXT,
    totalDistance INTEGER,
    totalDuration INTEGER,
    profilepicture TEXT,
    lastUpdated TEXT,
    highestPrivilege INTEGER)''')

    cursor.execute('''CREATE TABLE IF NOT EXISTS user_login_data (
    owner TEXT PRIMARY KEY,
    followed_users TEXT,
    trusted_users TEXT,
    followers TEXT,
    last_updated TEXT,
    access_token TEXT,
    refresh_token TEXT,
    expires_at REAL)''')

    try:
        cursor.execute("ALTER TABLE user_login_data ADD COLUMN followers TEXT")
    except sqlite3.OperationalError:
        pass

    try:
        cursor.execute("ALTER TABLE user_login_data ADD COLUMN access_token TEXT")
        cursor.execute("ALTER TABLE user_login_data ADD COLUMN refresh_token TEXT")
        cursor.execute("ALTER TABLE user_login_data ADD COLUMN expires_at REAL")
    except sqlite3.OperationalError:
        pass

    cursor.execute('''CREATE TABLE IF NOT EXISTS stops (
                      id INTEGER PRIMARY KEY,
                      stopName TEXT,
                      rilIdentifier TEXT,
                      latitude TEXT,
                      longitude TEXT,
                      timezone TEXT,
                      countryName TEXT,
                      admin1State TEXT,
                      admin2Dist TEXT,
                      city TEXT,
                      officialName TEXT,
                      countryCode TEXT
                      )''')

    cursor.execute('''CREATE TABLE IF NOT EXISTS journeys (
                      statusId INTEGER PRIMARY KEY,
                      userId INTEGER,
                      username TEXT,
                      body TEXT,
                      bodyMentions TEXT,
                      business TEXT,
                      visibility INTEGER,
                      likes INTEGER,
                      isLikable BOOLEAN,
                      tripId INTEGER,
                      hafasId TEXT,
                      category TEXT,
                      lineName TEXT,
                      lineNumber Text,
                      routeColor TEXT,
                      routeTextColor TEXT,
                      journeyNumber TEXT,
                      distance INTEGER,
                      distanceStraightLine INTEGER,
                      distanceToPrevJourney INTEGER,
                      departureDelay INTEGER,
                      arrivalDelay INTEGER,
                      durationPlanned INTEGER,
                      durationReal INTEGER,
                      timeDelta INTEGER,
                      points INTEGER,
                      departurePlanned TEXT,
                      departureFromRealtime TEXT,
                      departureManual TEXT,
                      departureFinal TEXT,
                      arrivalPlanned TEXT,
                      arrivalFromRealtime TEXT,
                      arrivalManual TEXT,
                      arrivalFinal TEXT,
                      realtimeAvailability BOOLEAN,
                      realtimeFromDataprovider BOOLEAN,
                      originId INTEGER,
                      destinationId INTEGER,
                      interStopsNumber INTEGER,
                      interStops TEXT,
                      operatorId INTEGER,
                      operatorName TEXT,
                      clientId INTEGER,
                      clientName TEXT,
                      event TEXT,
                      locomotiveClass TEXT,
                      vehicleNumber TEXT,
                      wagonClass TEXT,
                      price TEXT,
                      travelClass TEXT,
                      ticket TEXT,
                      passengerRights TEXT,
                      role TEXT,
                      seat TEXT,
                      socialStatus TEXT,
                      trainUtilization TEXT,
                      wagon TEXT,
                      tags TEXT,
                      moderationNotes TEXT,
                      lockVisibility BOOLEAN,
                      hideBody BOOLEAN,
                      createdAt TEXT,
                      FOREIGN KEY(userId) REFERENCES users(id),
                      FOREIGN KEY(destinationId) REFERENCES stops(id),
                      FOREIGN KEY(originId) REFERENCES stops(id))''')

    cursor.execute('''CREATE TABLE IF NOT EXISTS polylines (statusId INTEGER PRIMARY KEY, geometry TEXT)''')
    cursor.execute('''CREATE INDEX IF NOT EXISTS idx_journeys_username ON journeys(username)''')
    cursor.execute('''CREATE INDEX IF NOT EXISTS idx_journeys_dates ON journeys(arrivalFinal, departureFinal)''')
    cursor.execute('''CREATE INDEX IF NOT EXISTS idx_journeys_tripId ON journeys(tripId)''')

    conn.commit()
    return conn


def get_privilege_level(username: str) -> int:
    if username == AUTH_USER:
        return 3
    if username in TRUSTED_USERS:
        return 2
    if username in FOLLOWED_USERS:
        return 1
    return 0


def get_visibility_clause(username: str, table_prefix: str = "") -> str:
    if username == AUTH_USER:
        return ""
    prefix = f"{table_prefix}." if table_prefix else ""
    allowed_visibilities = [0, 4]
    if username in FOLLOWED_USERS:
        allowed_visibilities.append(2)
    if username in TRUSTED_USERS:
        allowed_visibilities.append(5)
    return f" AND {prefix}visibility IN ({','.join(map(str, allowed_visibilities))})"


def get_shared_visibility_query(table_prefix: str = "") -> str:
    prefix = f"{table_prefix}." if table_prefix else ""
    auth_cond = f"{prefix}username = '{AUTH_USER}'"
    trusted_list = ",".join(f"'{u}'" for u in TRUSTED_USERS) if TRUSTED_USERS else "NULL"
    followed_list = ",".join(f"'{u}'" for u in FOLLOWED_USERS) if FOLLOWED_USERS else "NULL"
    return f" AND ({auth_cond} OR ({prefix}username IN ({trusted_list}) AND {prefix}visibility IN (0, 4, 2, 5)) OR ({prefix}username IN ({followed_list}) AND {prefix}visibility IN (0, 4, 2)) OR ({prefix}visibility IN (0, 4)))"


def already_downloaded_in_db(conn: sqlite3.Connection, username) -> pd.DataFrame:
    return pd.read_sql('SELECT statusId FROM journeys WHERE username = ?', conn, params=(username,))


# ==========================================
# DATENBESCHAFFUNG MIT HOCHZÄHLERN & COUNTDOWNS
# ==========================================

def fetch_statuses_from_user(api_token, username, already_downloaded: pd.DataFrame | None = None, log_callback=None) -> list:
    if already_downloaded is None or 'statusId' not in already_downloaded.columns:
        already_ids = set()
    else:
        already_ids = set(already_downloaded['statusId'])
    global remaining

    current_user_data = []
    number = 1
    url = f'https://traewelling.de/api/v1/user/{username}/statuses'
    
    msg = f"[API] Hole Statusse für {username}..."
    print(msg)
    safe_log(log_callback, msg)
    
    while url is not None:
        try:
            travel_response = requests.get(url,
                                           headers={'accept': 'application/json',
                                                    'Authorization': f'Bearer {api_token}',
                                                    'User-Agent': 'TraewellingAnalysis_v9 tee_beutel'})
            travel_response.raise_for_status()
            travel_json = travel_response.json()
            travel_data = travel_json.get('data', {})
            url = travel_json.get('links', {}).get('next')

            current_user_data.extend(travel_data)
            current_ids = [data.get('id', None) for data in travel_data]

            if set(current_ids) & already_ids:
                url = None

            remaining = travel_response.headers.get('x-ratelimit-remaining')
            counter_msg = f'   📥 Statusse werden geladen: Seite {number} (Bisher {len(current_user_data)} geladen, Limit: {remaining})'
            print(counter_msg)
            safe_log(log_callback, counter_msg, replace=(number > 1))
            number += 1
            
        except Exception as e:
            if isinstance(e, requests.exceptions.HTTPError) and e.response.status_code == 429:
                warn = '[API] Rate Limit (429) erreicht. Warte 30 Sekunden...'
                print(warn)
                safe_log(log_callback, warn)
                sleep(30)
            else:
                err = f'[FEHLER] Konnte Daten von {username} nicht abrufen: {e}'
                print(err)
                safe_log(log_callback, err)
                url = None
                
    done_msg = f'   ✅ Insgesamt {len(current_user_data)} neue Statusse für {username} geladen.'
    print(done_msg)
    if len(current_user_data) > 0:
        safe_log(log_callback, done_msg)
    return current_user_data


def fetch_information_from_user(conn: sqlite3.Connection, api_token, username: str, highest_privilege: int = 0) -> str:
    global remaining
    user_info = {}
    time_of_download_start = datetime.now().isoformat()
    repeat_bool = True
    
    while repeat_bool:
        user_response = requests.get(f'https://traewelling.de/api/v1/user/{username}',
                                     headers={'accept': 'application/json',
                                              'Authorization': f'Bearer {api_token}',
                                              'User-Agent': 'TraewellingAnalysis_v9 tee_beutel'})
        try:
            user_response.raise_for_status()
            user_json = user_response.json()
            user_info = user_json.get('data', {})
            username = user_info.get('username', '')
            repeat_bool = False
            remaining = user_response.headers.get('x-ratelimit-remaining')
        except Exception as e:
            if isinstance(e, requests.exceptions.HTTPError) and e.response.status_code == 429:
                print('[API] Rate Limit erreicht. Warte 30 Sekunden...')
                sleep(30)
            else:
                repeat_bool = False
                print(f'[FEHLER] Profil {username} nicht gefunden.')

    user_id = user_info.get('id', None)
    if user_id is not None:
        cursor = conn.cursor()
        cursor.execute("SELECT username FROM users WHERE id = ?", (user_id,))
        existing_user = cursor.fetchone()

        if existing_user and existing_user[0] != username:
            old_username = existing_user[0]
            print(f"[MIGRATION] 🔄 Namensänderung: '{old_username}' -> '{username}'. Aktualisiere Historie...")
            cursor.execute("UPDATE journeys SET username = ? WHERE userId = ?", (username, user_id))
            cursor.execute("UPDATE user_login_data SET owner = ? WHERE owner = ?", (username, old_username))
            conn.commit()

        cursor.execute('''REPLACE INTO users(id, username, displayName, bio, mastodonUrl, totalDistance, totalDuration, profilePicture, lastUpdated, highestPrivilege) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                       (user_id, username, user_info.get('displayName', None), user_info.get('bio', None),
                        user_info.get('mastodonUrl', None), user_info.get('totalDistance', None),
                        user_info.get('totalDuration', None), user_info.get('profilePicture', None),
                        time_of_download_start, highest_privilege))
        conn.commit()

    return username


def get_or_fetch_stop_info(conn: sqlite3.Connection, api_token, stop_id: int, stop_name: str = None,
                           on_download=None) -> tuple:
    if stop_id is None or str(stop_id).strip().lower() in ('', 'none'):
        return 'Europe/Berlin', stop_name or 'Unbekannt', None, None

    cursor = conn.cursor()
    cursor.execute('SELECT timezone, stopName, latitude, longitude FROM stops WHERE id=?', (stop_id,))
    row = cursor.fetchone()

    if row is None:
        global remaining, tf
        if not stop_name:
            stop_name = str(stop_id)
        timezone = 'Europe/Berlin'
        latitude, longitude = None, None
        repeat_bool = True
        error_404_count = 0
        server_error_count = 0

        while repeat_bool:
            try:
                station_response = requests.get(f'https://traewelling.de/api/v1/stations/{stop_id}',
                                                headers={'accept': 'application/json',
                                                         'Authorization': f'Bearer {api_token}',
                                                         'User-Agent': 'TraewellingAnalysis_v9 tee_beutel'})
                station_response.raise_for_status()

                remaining = station_response.headers.get('x-ratelimit-remaining')
                station_json = station_response.json().get('data', {})
                stop_name = station_json['name']
                station_ril = station_json.get('rilIdentifier')

                stop_name = (f'{stop_name} ({station_ril})' if station_ril else (
                    stop_name if stop_name != 'Braunschweig Hbf ZOB' else 'Braunschweig Hbf (HBS)'))
                latitude = station_json['latitude']
                longitude = station_json['longitude']

                try:
                    if tf is None:
                        tf = TimezoneFinder()
                    timezone = tf.timezone_at(lat=latitude, lng=longitude) or 'Europe/Berlin'
                except Exception:
                    timezone = 'Europe/Berlin'

                sprich_deutsch = babel.Locale('de')
                geo_results = rgeo.search([(latitude, longitude)])[0]
                country_code = geo_results['cc']
                country_object = pycountry.countries.get(alpha_2=country_code)

                country_flag = country_object.flag if hasattr(country_object, 'flag') else ""
                country_name = f'{sprich_deutsch.territories.get(country_object.alpha_2, country_object.name)} {country_flag}'
                admin1 = f'{geo_results.get("admin1")} {country_flag}'
                admin2 = f'{geo_results.get("admin2")} {country_flag}'
                city = geo_results.get('name')
                official_name = f'{getattr(country_object, "official_name", country_name)} {country_flag}'

                cursor.execute(
                    '''REPLACE INTO stops(id, stopName, rilIdentifier, latitude, longitude, timezone, countryName,
                                          admin1State, admin2Dist, city, officialName, countryCode) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                    (station_json['id'], stop_name, station_ril, latitude, longitude, timezone, country_name, admin1, admin2, city, official_name, country_code))
                conn.commit()

                if on_download:
                    on_download(stop_name)
                repeat_bool = False

            except requests.exceptions.HTTPError as e:
                status = e.response.status_code
                if status == 429:
                    print('[API] Rate Limit erreicht (429). Warte 30 Sekunden...')
                    sleep(30)
                elif status in (500, 502, 503, 504):
                    server_error_count += 1
                    print(f'[API] Server überlastet ({status}) bei Haltestelle {stop_id}. Versuch {server_error_count}/3')
                    if server_error_count >= 3:
                        print(f"\n[API WARNUNG] Haltestelle '{stop_name}' (ID: {stop_id}) konnte nach 3 Versuchen wegen Serverüberlastung ({status}) nicht geladen werden. Verwende Dummy-Werte.")
                        return 'Europe/Berlin', stop_name, None, None
                    print('Warte 30 Sekunden...')
                    sleep(30)
                elif status == 404:
                    error_404_count += 1
                    print(f"\n[API FEHLER] Haltestelle '{stop_name}' (ID: {stop_id}) nicht gefunden (404). Versuch {error_404_count}/3")
                    if error_404_count >= 3:
                        print(f"\n[API WARNUNG] Haltestelle '{stop_name}' (ID: {stop_id}) konnte nach 3 Versuchen nicht gefunden werden (404). Verwende Dummy-Werte.")
                        return 'Europe/Berlin', stop_name, None, None
                    sleep(5)
                else:
                    print(f"\n[API FEHLER] Unerwarteter HTTP-Fehler {status} bei '{stop_name}' (ID: {stop_id}).")
                    raise e
            except (requests.exceptions.ConnectionError, requests.exceptions.Timeout) as e:
                print(f"\n[VERBINDUNGSFEHLER] Timeout/Abbruch bei Haltestelle '{stop_name}' (ID: {stop_id}). Warte 30 Sekunden...")
                sleep(30)
            except KeyboardInterrupt:
                conn.commit()
                conn.close()
                raise SystemExit(0)
            except Exception as e:
                print(f"\n[API FEHLER] Allgemeiner Fehler bei '{stop_name}': {e}")
                raise e

        return timezone, stop_name, latitude, longitude
    else:
        return row


def get_inter_stops_for_journeys(api_token, statuses_dicts: list, log_callback=None) -> dict:
    global remaining
    statuses = [data['checkin']['trip'] for data in statuses_dicts]
    intermediate_data = {}
    current_index = 0
    total = len(statuses)

    if statuses:
        start_msg = f"[API] Hole Zwischenhalte für {total} Fahrten..."
        print(start_msg)
        safe_log(log_callback, start_msg)

    while total > current_index:
        joined_ids = "%2C".join(str(trip_id) for trip_id in statuses[current_index:current_index + 50])
        url = f"https://traewelling.de/api/v1/stopovers/{joined_ids}"

        repeat_bool = True
        while repeat_bool:
            try:
                intermediate_response = requests.get(url,
                                                     headers={'accept': 'application/json',
                                                              'Authorization': f'Bearer {api_token}',
                                                              'User-Agent': 'TraewellingAnalysis_v9 tee_beutel'})
                intermediate_response.raise_for_status()
                remaining = intermediate_response.headers.get('x-ratelimit-remaining')
                intermediate_data.update(intermediate_response.json().get('data', {}))

                repeat_bool = False
                current_index += 50
                noch_zu_laden = max(0, total - current_index)
                cd_msg = f"   🚉 Zwischenhalte-Countdown: Noch {noch_zu_laden} von {total} Fahrten verbleibend... (Limit: {remaining})"
                print(cd_msg)
                safe_log(log_callback, cd_msg, replace=(current_index > 50))
            except Exception as e:
                if isinstance(e, requests.exceptions.HTTPError) and e.response.status_code == 429:
                    sleep(30)
                else:
                    repeat_bool = False
                    current_index = total

    if statuses:
        safe_log(log_callback, f"   ✅ Zwischenhalte für alle {total} Fahrten vollständig synchronisiert.")
    return intermediate_data


def calculate_haversine(lat1, lon1, lat2, lon2) -> int:
    if lat1 is None or lon1 is None or lat2 is None or lon2 is None:
        return 0
    try:
        lat1, lon1, lat2, lon2 = float(lat1), float(lon1), float(lat2), float(lon2)
    except ValueError:
        return 0
    
    R = 6371000
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    
    a = math.sin(dphi / 2.0)**2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlambda / 2.0)**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return int(R * c)


def calculate_and_insert(api_token, conn, user_data: list, all_inter_stations: dict, log_callback=None) -> None:
    global remaining
    cursor = conn.cursor()

    train_type_to_readable_dict = {
        'regional': 'Regionalzug', 'bus': 'Bus', 'ferry': 'Fähre',
        'national': 'Regionalzug', 'regionalExp': 'Fernverkehr',
        'nationalExpress': 'Fernverkehr', 'suburban': 'S-Bahn', 'subway': 'U-Bahn',
        'taxi': 'Taxi', 'tram': 'Tram', 'plane': 'Igitt ein Flugzeug'
    }

    prev_dest_lat, prev_dest_lon = None, None
    if user_data:
        username_query = user_data[0]['user']['username']
        cursor.execute("SELECT destinationId FROM journeys WHERE username=? ORDER BY departureFinal DESC LIMIT 1",
                       (username_query,))
        row = cursor.fetchone()
        if row:
            cursor.execute("SELECT latitude, longitude FROM stops WHERE id=?", (row[0],))
            lat_lon_row = cursor.fetchone()
            if lat_lon_row:
                prev_dest_lat, prev_dest_lon = lat_lon_row

    # COUNTDOWN: Noch herunterzuladende Haltestellen ermitteln
    all_stop_ids = set()
    missing_stops_dict = {}
    for d in user_data:
        t = d.get('checkin', {})
        if t.get('origin', {}).get('id'):
            s_id = t['origin']['id']
            all_stop_ids.add(s_id)
            missing_stops_dict[s_id] = t['origin'].get('name')
        if t.get('destination', {}).get('id'):
            s_id = t['destination']['id']
            all_stop_ids.add(s_id)
            missing_stops_dict[s_id] = t['destination'].get('name')

    missing_stop_ids = []
    if all_stop_ids:
        placeholders = ','.join(['?'] * len(all_stop_ids))
        cursor.execute(f"SELECT id FROM stops WHERE id IN ({placeholders})", list(all_stop_ids))
        existing_ids = {r[0] for r in cursor.fetchall()}
        missing_stop_ids = list(all_stop_ids - existing_ids)

    total_missing_stops = len(missing_stop_ids)
    stops_downloaded = [0]

    def on_stop_download(s_name):
        stops_downloaded[0] += 1
        rem = max(0, total_missing_stops - stops_downloaded[0])
        s_msg = f"   🚏 Lade Haltestelle: {s_name} | Noch {rem} von {total_missing_stops} übrig (Limit: {remaining})"
        print(s_msg)
        safe_log(log_callback, s_msg, replace=(stops_downloaded[0] > 1))

    if total_missing_stops > 0:
        init_s_msg = f"   🚏 Haltestellen-Countdown: Noch {total_missing_stops} neue Haltestellen zu laden... (Limit: {remaining})"
        print(init_s_msg)
        safe_log(log_callback, init_s_msg)

        # Vorab-Download aller fehlenden Haltestellen vor den Berechnungen
        for s_id in missing_stop_ids:
            get_or_fetch_stop_info(conn, api_token, s_id, missing_stops_dict.get(s_id), on_download=on_stop_download)

    total_rides = len(user_data)
    current_ride = 0

    for data in reversed(user_data):
        current_ride += 1
        if current_ride % 10 == 0 or current_ride == total_rides:
            rem_rides = total_rides - current_ride
            ride_msg = f"   ⚙️ Fahrten-Verarbeitung: Noch {rem_rides} von {total_rides} Fahrten zu berechnen..."
            print(ride_msg)
            safe_log(log_callback, ride_msg, replace=(current_ride > 10))

        status_id = data['id']
        user_id = data['user']['id']
        user_name = data['user']['username']
        body = data.get('body', '').replace('\n', ' ')
        body_mentions = json.dumps(data.get('bodyMentions', []), ensure_ascii=False)

        business = ''
        match data.get('business', ''):
            case 0:
                business = 'Privat'
            case 1:
                business = 'Geschäftlich'
            case 2:
                business = 'Arbeitsweg'

        visibility = data.get('visibility', '')
        likes = data.get('likes', 0)
        is_likeable = data.get('isLikable', False)
        hide_body = data.get('hide_body', '')
        moderation_notes = data.get('moderation_notes', '')
        lock_visibility = data.get('lock_visibility', '')
        created_at = data.get('createdAt', '')
        trip_id = data.get('checkin', {}).get('trip')
        hafas_id = data.get('checkin', {}).get('hafasId')

        train_data = data.get('checkin', {})
        line_name = re.sub(r'\s*\(\d+\)', '', train_data.get('lineName', 'Unknown Line')).strip()
        operator_name = (train_data.get('operator') or {}).get('name', 'Unknown Operator')
        operator_id = (train_data.get('operator') or {}).get('id', None)
        client_name = (data.get('client') or {}).get('name', 'Unknown Client')
        client_id = (data.get('client') or {}).get('id', 0)

        journey_distance = train_data.get('distance', 0)
        line_number = train_data.get('number')
        route_color = train_data.get('routeColor')
        route_text_color = train_data.get('routeTextColor')

        given_train_type = train_type_to_readable_dict.get(train_data.get('category', None),
                                                           train_data.get('category', None))
        if re.match(r"^(REX|RE|RB|Stoptrein|FEX|Sprinter)", line_name) and given_train_type != 'Bus':
            train_type = 'Regionalzug'
        elif re.match(r"^(RS)", line_name):
            train_type = 'S-Bahn'
        else:
            train_type = given_train_type

        journey_points = train_data.get('points', 0)
        journey_number = train_data.get('journeyNumber')

        origin, destination = train_data.get('origin', {}), train_data.get('destination', {})
        origin_id, destination_id = origin.get('id'), destination.get('id')

        origin_timezone, origin_name, origin_lat, origin_lon = get_or_fetch_stop_info(conn, api_token, origin_id,
                                                                                      origin.get('name'))
        destination_timezone, destination_name, destination_lat, destination_lon = get_or_fetch_stop_info(conn,
                                                                                                          api_token,
                                                                                                          destination_id,
                                                                                                          destination.get(
                                                                                                              'name'))

        dist_straight_line = calculate_haversine(origin_lat, origin_lon, destination_lat, destination_lon)
        dist_to_prev = calculate_haversine(prev_dest_lat, prev_dest_lon, origin_lat,
                                           origin_lon) if prev_dest_lat is not None else 0
        prev_dest_lat, prev_dest_lon = destination_lat, destination_lon

        departure_planned = origin.get('departurePlanned', None)
        departure_from_realtime = origin.get('departureReal', None)
        departure_manual = train_data.get('manualDeparture', None)
        arrival_planned = destination.get('arrivalPlanned', None)
        arrival_from_realtime = destination.get('arrivalReal', None)
        arrival_manual = train_data.get('manualArrival', None)

        departure_final = (departure_manual or departure_from_realtime or departure_planned)
        arrival_final = (arrival_manual or arrival_from_realtime or arrival_planned)
        realtime_availability = bool(
            departure_manual or arrival_manual or departure_from_realtime or arrival_from_realtime)
        realtime_from_dataprovider = bool(departure_from_realtime or arrival_from_realtime)

        dp_dt = datetime.fromisoformat(departure_planned) if departure_planned else None
        df_dt = datetime.fromisoformat(departure_final) if departure_final else None
        ap_dt = datetime.fromisoformat(arrival_planned) if arrival_planned else None
        af_dt = datetime.fromisoformat(arrival_final) if arrival_final else None

        departure_delay = int((df_dt - dp_dt).total_seconds() / 60) if df_dt and dp_dt else 0
        arrival_delay = int((af_dt - ap_dt).total_seconds() / 60) if af_dt and ap_dt else 0
        duration_planned = int((ap_dt - dp_dt).total_seconds() / 60) if ap_dt and dp_dt else 0
        duration_real = int((af_dt - df_dt).total_seconds() / 60) if af_dt and df_dt else 0
        time_delta = duration_real - duration_planned

        if not realtime_availability:
            departure_final = arrival_final = None

        event = (json.dumps(data.get('event'), ensure_ascii=False) if data.get('event') else None)

        # Variablen initialisieren
        vehicle_number = locomotive_class = wagon_class = price = travel_class = ticket = None
        passenger_rights = role = seat = social_status = train_utilization = wagon = None

        tags = data.get('tags', None)
        if tags:
            for tag in tags:
                match tag.get('key', None):
                    case 'trwl:vehicle_number':
                        vehicle_number = tag.get('value', 'N/A')
                    case 'trwl:journey_number':
                        journey_number = tag.get('value', 'N/A')
                    case 'trwl:locomotive_class':
                        locomotive_class = tag.get('value', 'N/A')
                    case 'trwl:ticket':
                        ticket = tag.get('value', 'N/A')
                    case 'trwl:wagon_class':
                        wagon_class = tag.get('value', 'N/A')
                    case "trwl:price":
                        price = tag.get('value', 'N/A')
                    case "trwl:travel_class":
                        travel_class = tag.get('value', 'N/A')
                    # Neu hinzugefügte Tags
                    case 'trwl:passenger_rights':
                        passenger_rights = tag.get('value', 'N/A')
                    case 'trwl:role':
                        role = tag.get('value', 'N/A')
                    case 'trwl:seat':
                        seat = tag.get('value', 'N/A')
                    case 'trwl:social_status':
                        social_status = tag.get('value', 'N/A')
                    case 'trwl:train_utilization':
                        train_utilization = tag.get('value', 'N/A')
                    case 'trwl:wagon':
                        wagon = tag.get('value', 'N/A')

            tracked_keys = {'trwl:vehicle_number', 'trwl:journey_number', 'trwl:locomotive_class',
                            'trwl:ticket', 'trwl:wagon_class', 'trwl:price', 'trwl:travel_class',
                            'trwl:passenger_rights', 'trwl:role', 'trwl:seat', 'trwl:social_status',
                            'trwl:train_utilization', 'trwl:wagon'}
            tags = [tag for tag in tags if tag.get('key') not in tracked_keys]
            tags = (json.dumps(tags, ensure_ascii=False) if tags else None)
        else:
            tags = None

        try:
            inter_stations = all_inter_stations.get(str(trip_id), [])
            inter_stations_ids = [d['id'] for d in inter_stations]
            origin_index = inter_stations_ids.index(origin_id)
            destination_index = inter_stations_ids.index(destination_id, origin_index + 1)
            inter_stops_number = destination_index - origin_index - 1
            inter_stops = json.dumps([(f'{d["name"]} ({d.get("rilIdentifier")})' if d.get("rilIdentifier") else (
                d["name"] if d["name"] != 'Braunschweig Hbf ZOB' else 'Braunschweig Hbf (HBS)'))
                                      for d in inter_stations[origin_index + 1:destination_index]], ensure_ascii=False)
        except Exception:
            inter_stops_number = None
            inter_stops = None

        cursor.execute(
            '''REPLACE INTO journeys(statusId, userId, username, body, bodyMentions, business, visibility, likes,
                                     isLikable, tripId, hafasId, lineName, operatorName, operatorId, distance,
                                     distanceStraightLine, distanceToPrevJourney, routeColor, routeTextColor, category,
                                     points, journeyNumber, clientId, clientName, originId, destinationId, lineNumber,
                                     moderationNotes, lockVisibility, hideBody, createdAt, departurePlanned,
                                     departureFromRealtime, departureManual, departureFinal, arrivalPlanned,
                                     arrivalFromRealtime, arrivalManual, arrivalFinal, realtimeAvailability,
                                     realtimeFromDataprovider, departureDelay, arrivalDelay, durationPlanned,
                                     durationReal, timeDelta, locomotiveClass, vehicleNumber, wagonClass, price,
                                     travelClass, ticket, passengerRights, role, seat, socialStatus, trainUtilization,
                                     wagon, tags, event, interStops, interStopsNumber)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                       ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
            (status_id, user_id, user_name, body, body_mentions, business, visibility, likes, is_likeable, trip_id,
             hafas_id, line_name, operator_name, operator_id, journey_distance, dist_straight_line, dist_to_prev,
             route_color, route_text_color, train_type, journey_points, journey_number, client_id, client_name,
             origin_id, destination_id, line_number, moderation_notes, lock_visibility, hide_body, created_at,
             departure_planned, departure_from_realtime, departure_manual, departure_final, arrival_planned,
             arrival_from_realtime, arrival_manual, arrival_final, realtime_availability, realtime_from_dataprovider,
             departure_delay, arrival_delay, duration_planned, duration_real, time_delta, locomotive_class,
             vehicle_number, wagon_class, price, travel_class, ticket, passenger_rights, role, seat, social_status,
             train_utilization, wagon, tags, event, inter_stops, inter_stops_number))

    conn.commit()
    if user_data:
        safe_log(log_callback, f"   ✅ Alle {total_rides} Fahrten berechnet und gespeichert.")


def fetch_data_to_db(username: str, reload_all: bool = False, conn: sqlite3.Connection | None = None,
                     api_token: str | None = None, log_callback=None, update_db: bool = True) -> sqlite3.Connection:
    if conn is None:
        conn = init_db()
    if api_token is None:
        api_token = log_onto_traewelling(conn)

    current_privilege = get_privilege_level(username)
    cursor = conn.cursor()
    cursor.execute("SELECT highestPrivilege FROM users WHERE username = ?", (username,))
    row = cursor.fetchone()

    highest_privilege = row[0] if row else None
    if highest_privilege is not None:
        if current_privilege > highest_privilege:
            reload_all = True
            highest_privilege = current_privilege
    else:
        highest_privilege = current_privilege

    # Profil des Users immer updaten/prüfen
    username = fetch_information_from_user(conn, api_token, username, highest_privilege)

    # Abbrechen, falls keine neuen Statusse geladen werden sollen
    if not update_db:
        msg = f"[INFO] Update der Statusse übersprungen: Nutze bestehende Daten für {username}."
        print(msg)
        safe_log(log_callback, msg)
        return conn

    downloaded_frame = pd.DataFrame() if reload_all else already_downloaded_in_db(conn, username)
    user_statuses = fetch_statuses_from_user(api_token, username, downloaded_frame, log_callback=log_callback)

    if not user_statuses:
        msg = f"[INFO] Keine neuen Fahrten für {username} gefunden. System ist up-to-date."
        print(msg)
        safe_log(log_callback, msg)
        return conn

    inter_stations = get_inter_stops_for_journeys(api_token, user_statuses, log_callback=log_callback)
    calculate_and_insert(api_token, conn, user_statuses, inter_stations, log_callback=log_callback)
    return conn

# ==========================================
# DATENAUSWERTUNG / PANDAS DATAFRAMES
# ==========================================

def format_utc_to_datatime(utc: str, tz: str) -> datetime | None:
    if not utc:
        return None
    dt_utc = datetime.fromisoformat(utc)
    dt_local = dt_utc.astimezone(ZoneInfo(tz))
    return dt_local.replace(tzinfo=None)


def params_and_date_query(username: str, start_date: str = '', end_date: str = '', table_prefix: str = "") -> tuple:
    prefix = f"{table_prefix}." if table_prefix else ""
    date_query = f"WHERE {prefix}username = ?"
    params = [username]
    
    if start_date:
        date_query += f" AND {prefix}arrivalFinal > ?"
        params.append(start_date)
    if end_date:
        date_query += f" AND {prefix}departureFinal < ?"
        if len(end_date) == 10:
            end_date += 'T23:59:59'
        params.append(end_date)
        
    date_query += get_visibility_clause(username, table_prefix)
    return params, date_query


def create_journeys_dataframe(conn: sqlite3.Connection, username: str, start_date: str = '',
                              end_date: str = '') -> pd.DataFrame:
    params, date_query = params_and_date_query(username, start_date, end_date, "j")
    query = f"""
    select 
        j.lineName AS 'Linie',
        j.category AS 'Kategorie',
        j.journeyNumber AS 'Fahrtnummer',
        j.operatorName AS 'Betreiber',
        j.business AS 'Grund',
        j.vehicleNumber AS 'Fahrzeugnummer',
        origin.stopName AS 'Abfahrtshaltestelle',
        origin.timezone AS 'origin_tz',
        j.departurePlanned AS 'Abfahrt geplant',
        j.departureFinal AS 'Abfahrt real',
        j.departureDelay AS 'Abw ab',
        j.interStopsNumber AS 'Zwischenhaltestellen',
        destination.stopName AS 'Zielstation',
        destination.timezone AS 'destination_tz',
        j.arrivalPlanned AS 'Ankunft geplant',
        j.arrivalFinal AS 'Ankunft real',
        j.arrivalDelay AS 'Abw an',
        j.durationPlanned AS 'Reisezeit plan (min)',
        j.durationReal AS 'Reisezeit (min)',
        j.timeDelta AS 'Delta',
        j.distance AS 'Entfernung (m)',
        j.distanceStraightLine AS 'Luftlinie Fahrt (m)',
        j.distanceToPrevJourney AS 'Luftlinie zur Vorfahrt (m)',
        j.points AS 'Punkte',
        j.likes AS 'Likes',
        j.body AS 'Statustext',
        origin.countryName || ' -> ' || destination.countryName AS 'Besuchte Länder',
        '=HYPERLINK("https://traewelling.de/status/'||j.statusId||'", "https://traewelling.de/status/'||j.statusId||'")' AS 'Link'
    from journeys j
    LEFT JOIN stops origin ON j.originId = origin.id
    LEFT JOIN stops destination ON j.destinationId = destination.id {date_query}
    """
    df = pd.read_sql_query(query, conn, params=params)

    df['Abfahrt geplant'] = [format_utc_to_datatime(dt, tz) for dt, tz in zip(df['Abfahrt geplant'], df['origin_tz'])]
    df['Abfahrt real'] = [format_utc_to_datatime(dt, tz) for dt, tz in zip(df['Abfahrt real'], df['origin_tz'])]
    df['Ankunft geplant'] = [format_utc_to_datatime(dt, tz) for dt, tz in zip(df['Ankunft geplant'], df['destination_tz'])]
    df['Ankunft real'] = [format_utc_to_datatime(dt, tz) for dt, tz in zip(df['Ankunft real'], df['destination_tz'])]

    duration_real_hh_mm = df['Reisezeit (min)'].apply(lambda x: f'{int(x) // 60:02}:{x % 60:02}' if pd.notna(x) else None)
    df.insert(loc=int(df.columns.get_loc('Reisezeit (min)') + 1), column='Reisezeit (hh:mm)', value=duration_real_hh_mm)

    for col in ['Entfernung (m)', 'Luftlinie Fahrt (m)', 'Luftlinie zur Vorfahrt (m)', 'Reisezeit (min)', 'Delta', 'Reisezeit plan (min)']:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    df.insert(loc=int(df.columns.get_loc('Entfernung (m)') + 1), column='Entfernung (km)', value=(df['Entfernung (m)'] / 1000).round(2))
    df.insert(loc=int(df.columns.get_loc('Entfernung (km)') + 1), column='Luftlinie Fahrt (km)', value=(df['Luftlinie Fahrt (m)'] / 1000).round(2))
    df.insert(loc=int(df.columns.get_loc('Luftlinie Fahrt (km)') + 1), column='Luftlinie zur Vorfahrt (km)', value=(df['Luftlinie zur Vorfahrt (m)'] / 1000).round(2))

    temp_umweg = (df['Entfernung (m)'] / df['Luftlinie Fahrt (m)'].replace(0, np.nan)).round(2)
    df.insert(loc=int(df.columns.get_loc('Luftlinie zur Vorfahrt (km)') + 1), column='Umweg Faktor', value=np.where(df['Luftlinie Fahrt (m)'] == 0, "∞", temp_umweg))

    temp_geschwindigkeit = ((df['Entfernung (m)'] / df['Reisezeit (min)'].replace(0, np.nan)) * 0.06).round(2)
    df.insert(loc=int(df.columns.get_loc('Umweg Faktor') + 1), column='Geschwindigkeit (km/h)', value=np.where(df['Reisezeit (min)'] == 0, "∞", temp_geschwindigkeit))
    df.insert(loc=int(df.columns.get_loc('Delta') + 1), column='Delta pro Planzeit', value=((df['Delta'] / df['Reisezeit plan (min)']) * 100).round(2))

    df.drop(columns=['origin_tz', 'destination_tz', 'Luftlinie Fahrt (m)', 'Luftlinie zur Vorfahrt (m)'], inplace=True)
    return df


def create_user_meta_dataframe(conn: sqlite3.Connection, username: str) -> pd.DataFrame:
    cursor = conn.cursor()
    cursor.execute("SELECT lastUpdated, displayname, profilepicture, mastodonUrl, bio, totalDistance, totalDuration FROM users WHERE username = ?", (username,))
    row = cursor.fetchone()
    if row:
        lastUpdated, displayname, profilepicture, mastodon, bio, totalDist, totalDur = row
        totalDur = totalDur or 0
        totalDist = totalDist or 0
        return pd.DataFrame({
            "Label": ["Datenstand", "Anzeigename", "Profilbild", "Mastodonlink", "Bio", "Gesamtdistanz", "Gesamtzeit", "Programmversion"],
            "Wert": [
                lastUpdated, 
                displayname or 'N/A', 
                f'=HYPERLINK("{profilepicture}", "{profilepicture}")' if profilepicture else 'Kein Link',
                f'=HYPERLINK("{mastodon}", "{mastodon}")' if mastodon else 'Kein Link',
                bio or 'N/A',
                f"{totalDist / 1000} km",
                f"{int(totalDur/60)}:{totalDur%60:02d}",
                "traewelling_analysis_v9.py"
            ]
        })
    return pd.DataFrame()


def create_shared_meta_dataframe(conn: sqlite3.Connection, usernames: list[str], start_date: str = '',
                                 end_date: str = '') -> pd.DataFrame:
    users_with_journeys = []
    users_without_journeys = []

    # Prüfen, ob der jeweilige User im angegebenen Zeitraum Fahrten hat
    for u in usernames:
        if check_for_user_journeys(conn, u, start_date, end_date):
            users_with_journeys.append(u)
        else:
            users_without_journeys.append(u)

    return pd.DataFrame({
        "Label": [
            "Datenstand (Export)",
            "Programmversion",
            "Startdatum des Exports",
            "Enddatum des Exports",
            "Anzahl berücksichtigter User",
            "User mit Fahrten",
            "User ohne Fahrten"
        ],
        "Wert": [
            datetime.now().strftime("%d.%m.%Y %H:%M:%S"),
            "traewelling_analysis_v9.py",
            start_date if start_date else "Kein Startdatum (Alle)",
            end_date if end_date else "Kein Enddatum (Alle)",
            f'{len(usernames)}',
            ", ".join(users_with_journeys) if users_with_journeys else "-",
            ", ".join(users_without_journeys) if users_without_journeys else "-"
        ]
    })


def create_geo_dataframe(conn: sqlite3.Connection, username: str, geo_column: str, start_date: str = '', end_date: str = '') -> pd.DataFrame:
    params, date_query = params_and_date_query(username, start_date, end_date)
    beautiful_names = {'countryName': 'Land', 'admin1State': 'Region/Bundesland', 'admin2Dist': 'Bezirk/Kreis', 'city': 'Stadt'}
    beautiful_name = beautiful_names.get(geo_column, geo_column)
    
    query = f"""
    WITH filtered_journeys AS (
        SELECT originId, destinationId
        FROM journeys {date_query}
    ),
    all_visits AS (
        SELECT originId AS stopId FROM filtered_journeys
        UNION ALL
        SELECT destinationId AS stopId FROM filtered_journeys
    )
    SELECT s.{geo_column} AS '{beautiful_name}', COUNT(*) AS Besuche
    FROM all_visits v
    JOIN stops s ON v.stopId = s.id
    WHERE s.{geo_column} IS NOT NULL AND s.{geo_column} != ''
    GROUP BY s.{geo_column}
    ORDER BY Besuche DESC
    """
    df = pd.read_sql_query(query, conn, params=params)
    df.index = df['Besuche'].rank(method='min', ascending=False).astype(int)
    df = df.rename(columns={beautiful_name: f'{beautiful_name} ({len(df)})'})
    df.index.name = 'Platz'
    return df


def create_user_stats_dataframe(conn: sqlite3.Connection, username: str, start_date: str = '',
                                end_date: str = '') -> pd.DataFrame:
    cursor = conn.cursor()
    params, date_query = params_and_date_query(username, start_date, end_date)

    cursor.execute(f"SELECT DISTINCT date(arrivalFinal) FROM journeys {date_query} ORDER BY date(arrivalFinal)", params)
    dates = [datetime.strptime(row[0], '%Y-%m-%d').date() for row in cursor.fetchall() if row[0]]

    if dates:
        first_day, last_day = dates[0], dates[-1]
        gaps = ((d2 - d1).days - 1 for d1, d2 in zip(dates[:-1], dates[1:]))
        total_time_skipped = sum(gap if gap > 30 else 0 for gap in gaps)
        active_days = (last_day - first_day).days + 1 - total_time_skipped
    else:
        active_days = 1

    query_main = f"""
    SELECT 
        COUNT(statusId) as total_journeys,
        SUM(distance) as total_distance,
        SUM(durationReal) as total_duration,
        SUM(points) as total_points,
        SUM(likes) as total_likes,
        SUM(CASE WHEN realtimeAvailability = 1 THEN departureDelay ELSE 0 END) as total_dep_delay,
        SUM(CASE WHEN realtimeAvailability = 1 THEN arrivalDelay ELSE 0 END) as total_arr_delay,
        COUNT(DISTINCT operatorName || lineName) as unique_lines,
        COUNT(DISTINCT operatorName) as unique_operators,
        SUM(CASE WHEN arrivalDelay >= 3 AND realtimeAvailability = 1 THEN 1 ELSE 0 END) as delayed_3,
        SUM(CASE WHEN arrivalDelay >= 6 AND realtimeAvailability = 1 THEN 1 ELSE 0 END) as delayed_6,
        SUM(CASE WHEN arrivalDelay >= 15 AND realtimeAvailability = 1 THEN 1 ELSE 0 END) as delayed_15,
        SUM(CASE WHEN arrivalDelay >= 60 AND realtimeAvailability = 1 THEN 1 ELSE 0 END) as delayed_60,
        SUM(CASE WHEN arrivalDelay < 0 AND realtimeAvailability = 1 THEN 1 ELSE 0 END) as early,
        SUM(CASE WHEN realtimeAvailability = 1 THEN 1 ELSE 0 END) as realtime_count,
        AVG(interStopsNumber) as avg_inter_stops,
        SUM(CASE WHEN arrivalDelay = 0 AND realtimeAvailability = 1 THEN 1 ELSE 0 END) as on_time,
        SUM(distanceStraightLine) as total_straight_dist,
        SUM(distanceToPrevJourney) as total_prev_dist
    FROM journeys {date_query}
    """
    cursor.execute(query_main, params)
    res = cursor.fetchone()

    total_journeys = res[0] or 0
    total_distance = res[1] or 0
    total_duration = res[2] or 0
    total_points = res[3] or 0
    total_likes = res[4] or 0
    rt_dep_delay = res[5] or 0
    rt_arr_delay = res[6] or 0
    unique_lines = res[7] or 0
    unique_operators = res[8] or 0
    del_3, del_6, del_15, del_60 = res[9] or 0, res[10] or 0, res[11] or 0, res[12] or 0
    early, realtime_count = res[13] or 0, res[14] or 0
    avg_inter_stops = res[15] or 0.0
    on_time = res[16] or 0
    total_straight_dist = res[17] or 0
    total_prev_dist = res[18] or 0

    params_stations = params + params
    query_stations = f"""
    SELECT COUNT(DISTINCT stopName) FROM (
        SELECT s.stopName FROM journeys j JOIN stops s ON j.originId = s.id {date_query.replace('username', 'j.username')}
        UNION
        SELECT s.stopName FROM journeys j JOIN stops s ON j.destinationId = s.id {date_query.replace('username', 'j.username')}
    )
    """
    cursor.execute(query_stations, params_stations)
    unique_stations = cursor.fetchone()[0] or 0

    j_div = total_journeys if total_journeys > 0 else 1
    d_div = active_days
    rt_div = realtime_count if realtime_count > 0 else 1

    total_dep_delay = (rt_dep_delay / rt_div) * j_div
    total_arr_delay = (rt_arr_delay / rt_div) * j_div

    total_hours, total_minutes = divmod(total_duration, 60)
    per_day_hours, per_day_minutes = divmod(total_duration / d_div, 60)
    per_journey_hours, per_journey_minutes = divmod(total_duration / j_div, 60)

    dist_between_ratio = (total_prev_dist / total_distance * 100) if total_distance > 0 else 0

    stats_data = [
        ['Fahrten', f'{total_journeys} Fahrten', f'{total_journeys / d_div:.3f} Fahrten', f'{total_journeys / j_div:.3f} Fahrten'],
        ['Tage', f'{active_days} Tage', f'{active_days / d_div:.3f} Tage', f'{active_days / j_div:.3f} Tage'],
        ['Linien', f'{unique_lines} verschiedene Linien', f'{unique_lines / d_div:.3f} verschiedene Linien', f'{unique_lines / j_div:.3f} verschiedene Linien'],
        ['Betreiber', f'{unique_operators} verschiedene Betreiber', f'{unique_operators / d_div:.3f} verschiedene Betreiber', f'{unique_operators / j_div:.3f} verschiedene Betreiber'],
        ['Haltestellen', f'{unique_stations} einmalige Haltestellen', f'{unique_stations / d_div:.3f} einmalige Haltestellen', f'{unique_stations / j_div:.3f} einmalige Haltestellen'],
        ['Abfahrtsverspätung', f'{total_dep_delay:.0f} Minuten', f'{total_dep_delay / d_div:.3f} Minuten', f'{total_dep_delay / j_div:.3f} Minuten'],
        ['Ankunftsverspätung', f'{total_arr_delay:.0f} Minuten', f'{total_arr_delay / d_div:.3f} Minuten', f'{total_arr_delay / j_div:.3f} Minuten'],
        ['Zeit (min)', f'{total_duration:.0f} Minuten', f'{total_duration / d_div:.3f} Minuten', f'{total_duration / j_div:.3f} Minuten'],
        ['Zeit (hh:mm)', f"{total_hours:02.0f}:{int(round(total_minutes, 0)):02}", f"{per_day_hours:02.0f}:{round(per_day_minutes, 0):02.0f}", f"{per_journey_hours:02.0f}:{round(per_journey_minutes, 0):02.0f}"],
        ['Distanz (m)', f'{total_distance} Meter', f'{total_distance / d_div:.3f} Meter', f'{total_distance / j_div:.3f} Meter'],
        ['Distanz (km)', f'{total_distance / 1000:.2f} Kilometer', f'{total_distance / 1000 / d_div:.2f} Kilometer', f'{total_distance / 1000 / j_div:.2f} Kilometer'],
        ['Luftlinie Fahrten (km)', f'{total_straight_dist / 1000:.2f} Kilometer', f'{total_straight_dist / 1000 / d_div:.2f} Kilometer', f'{total_straight_dist / 1000 / j_div:.2f} Kilometer'],
        ['Luftlinie zwischen Fahrten (km)', f'{total_prev_dist / 1000:.2f} Kilometer', f'{total_prev_dist / 1000 / d_div:.2f} Kilometer', f'{total_prev_dist / 1000 / j_div:.2f} Kilometer'],
        ['Punkte', f'{total_points} Punkte', f'{total_points / d_div:.3f} Punkte', f'{total_points / j_div:.3f} Punkte'],
        ['Likes', f'{total_likes} Likes', f'{total_likes / d_div:.3f} Likes', f'{total_likes / j_div:.3f} Likes'],
        ['Tracking-Lücke', f'{dist_between_ratio:.2f}%', '', ''],
        ['Umweg Faktor', f'{(total_distance / total_straight_dist):.2f}' if total_straight_dist else '0', '', ''],
        ['Geschwindigkeit', f'{total_distance / (total_duration / 60 * 1000):.2f} km/h' if total_duration else '0 km/h', '', ''],
        ['Echtzeitquote', f'{(realtime_count / j_div) * 100:.2f}%', '', ''],
        ['Verfrüht', f'{(early / rt_div) * 100:.2f}%', '', ''],
        ['Pünktlich', f'{(on_time / rt_div) * 100:.2f}%', '', ''],
        ['3 Minuten oder mehr verspätet', f'{(del_3 / rt_div) * 100:.2f}%', '', ''],
        ['6 Minuten oder mehr verspätet', f'{(del_6 / rt_div) * 100:.2f}%', '', ''],
        ['15 Minuten oder mehr verspätet', f'{(del_15 / rt_div) * 100:.2f}%', '', ''],
        ['60 Minuten oder mehr verspätet', f'{(del_60 / rt_div) * 100:.2f}%', '', ''],
        ['Ø Zwischenhalte', f'{avg_inter_stops:.3f} Zwischenhalte', '', '']
    ]

    return pd.DataFrame(stats_data, columns=['Metrik', 'Gesamtwert', 'Durchschnitt pro Tag', 'Durchschnitt pro Fahrt'])


def create_user_stops_dataframe(conn: sqlite3.Connection, username: str, start_date: str = '',
                                end_date: str = '') -> pd.DataFrame:
    params, date_query = params_and_date_query(username, start_date, end_date)
    query = f"""
            WITH filtered_journeys AS (
                SELECT originId, destinationId, departurePlanned, arrivalPlanned
                FROM journeys {date_query}
            ),
            all_visits AS (
                SELECT originId AS stopId, departurePlanned AS visitTime FROM filtered_journeys
                UNION ALL
                SELECT destinationId AS stopId, arrivalPlanned AS visitTime FROM filtered_journeys
            )
            SELECT s.stopName, s.latitude, s.longitude, COUNT(*) AS Besuche, MAX(v.visitTime) AS Letzter_Besuch
            FROM all_visits v
            JOIN stops s ON v.stopId = s.id
            GROUP BY s.stopName
            ORDER BY Besuche DESC
        """
    df = pd.read_sql_query(query, conn, params=params)

    if not df.empty:
        # Als echtes datetime-Objekt belassen (ohne Zeitzone, damit Excel es als Datum versteht)
        # und erst dann mit 'Nie' auffüllen.
        df['Letzter Besuch'] = pd.to_datetime(df['Letzter_Besuch'], utc=True, errors='coerce').dt.tz_convert(
            'Europe/Berlin').dt.tz_localize(None).fillna('Nie')

        lat_most, lon_most = float(df.iloc[0]['latitude']), float(df.iloc[0]['longitude'])
        lat = pd.to_numeric(df['latitude'], errors='coerce')
        lon = pd.to_numeric(df['longitude'], errors='coerce')

        lat1_rad, lon1_rad = np.radians(lat_most), np.radians(lon_most)
        lat2_rad, lon2_rad = np.radians(lat), np.radians(lon)

        dlat = lat2_rad - lat1_rad
        dlon = lon2_rad - lon1_rad

        a = np.sin(dlat / 2.0) ** 2 + np.cos(lat1_rad) * np.cos(lat2_rad) * np.sin(dlon / 2.0) ** 2
        dist = 6371000 * (2 * np.arctan2(np.sqrt(a), np.sqrt(1 - a)))

        relevance = (df['Besuche'] * np.sqrt(dist) / 1000).round(3).astype(object)
        relevance.iloc[0] = "∞"
        df['Hub-Distanz-Wert'] = relevance

        # Unnötige Spalten entfernen
        df = df.drop(columns=['latitude', 'longitude', 'Letzter_Besuch'])

        # 'Letzter Besuch' hinter 'Hub-Distanz-Wert' schieben
        cols = list(df.columns)
        cols.remove('Letzter Besuch')
        cols.append('Letzter Besuch')
        df = df[cols]

    df.index = df['Besuche'].rank(method='min', ascending=False).astype(int)
    df = df.rename(columns={'stopName': f'Haltestelle ({len(df)})'})
    df.index.name = 'Platz'
    return df


def create_user_stops_dataframe_with_inter(conn: sqlite3.Connection, username: str, start_date: str = '',
                                           end_date: str = '', base_dataframe: pd.DataFrame = None) -> pd.DataFrame:
    base_df = base_dataframe.copy() if base_dataframe is not None and not base_dataframe.empty else create_user_stops_dataframe(
        conn, username, start_date, end_date)

    # "Letzter Besuch (Ohne Via)" soll nicht in die neue Liste übernommen werden
    base_df.drop(columns=['Hub-Distanz-Wert', 'Letzter Besuch'], inplace=True, errors='ignore')
    cursor = conn.cursor()
    base_df = base_df.rename(columns={base_df.columns[0]: 'Haltestelle', 'Besuche': 'Ohne Via'}, copy=True)

    params, date_query = params_and_date_query(username, start_date, end_date)

    # Für Via-Halte ebenfalls die Planzeit verwenden
    cursor.execute(f"SELECT departurePlanned, interStops FROM journeys {date_query}", params)
    rows = cursor.fetchall()

    inter_stops_list = []
    inter_stops_last_visit = {}

    for row in rows:
        dep_time = row[0]
        try:
            stops = json.loads(row[1])
            inter_stops_list.extend(stops)
            for stop in stops:
                if stop not in inter_stops_last_visit or (dep_time and dep_time > inter_stops_last_visit[stop]):
                    inter_stops_last_visit[stop] = dep_time
        except (json.decoder.JSONDecodeError, TypeError):
            continue

    inter_counts = Counter(inter_stops_list)
    inter_df_data = []
    for stop, count in inter_counts.items():
        inter_df_data.append(
            {'Haltestelle': stop, 'Nur Via': count, 'Letzter_Besuch_Via': inter_stops_last_visit.get(stop)})

    inter_df = pd.DataFrame(inter_df_data)
    if not inter_df.empty:
        # Als echtes datetime-Objekt (ohne Zeitzone) belassen
        inter_df['Letzter Besuch'] = pd.to_datetime(inter_df['Letzter_Besuch_Via'], utc=True,
                                                    errors='coerce').dt.tz_convert('Europe/Berlin').dt.tz_localize(None)
        inter_df.drop(columns=['Letzter_Besuch_Via'], inplace=True)
    else:
        inter_df = pd.DataFrame(columns=['Haltestelle', 'Nur Via', 'Letzter Besuch'])

    merged_df = pd.merge(base_df, inter_df, on='Haltestelle', how='outer')

    merged_df['Nur Via'] = pd.to_numeric(merged_df['Nur Via'], errors='coerce').fillna(0)
    merged_df['Ohne Via'] = pd.to_numeric(merged_df['Ohne Via'], errors='coerce').fillna(0)

    # Leere Felder bei nie durchfahrenen Haltestellen jetzt mit "Nie" füllen.
    # Da wir vorher echte datetime-Objekte hatten, ergibt dieser Mix in Pandas
    # eine "Object"-Spalte, die in Excel korrekt als Mix aus Daten und Text geschrieben wird.
    if 'Letzter Besuch' in merged_df.columns:
        merged_df['Letzter Besuch'] = merged_df['Letzter Besuch'].fillna('Nie')

    merged_df['insgesamt'] = merged_df['Nur Via'] + merged_df['Ohne Via']
    merged_df['Umsteigequote (%)'] = ((merged_df['Ohne Via'] / merged_df['insgesamt']) * 100).astype(float).round(2)
    merged_df['realative Umsteigequote (%)'] = (((merged_df['Ohne Via'] / 2) / (
            merged_df['insgesamt'] - merged_df['Ohne Via']).clip(lower=1)) * 100).astype(float).round(2)

    # Letzter Besuch ganz nach rechts schieben
    if 'Letzter Besuch' in merged_df.columns:
        cols = list(merged_df.columns)
        cols.remove('Letzter Besuch')
        cols.append('Letzter Besuch')
        merged_df = merged_df[cols]

    merged_df.index = merged_df['insgesamt'].rank(method='min', ascending=False).astype(int)
    merged_df = merged_df.rename(columns={'Haltestelle': f'Haltestelle ({len(merged_df)})'})
    merged_df.index.name = 'Platz'
    merged_df.sort_values(by=['insgesamt'], ascending=False, inplace=True)
    return merged_df

def create_user_excel_file(conn: sqlite3.Connection, username: str, start_date: str = '', end_date: str = '', create_picture: bool = True) -> str:
    """Erstellt das Excel-Fahrtenbuch mit dynamischem Datumszusatz und gibt den Dateinamen zurück."""
    kwargs = dict(conn=conn, username=username, start_date=start_date, end_date=end_date)
    temporal_base_df = get_temporal_base_dataframe(**kwargs, create_picture=create_picture)
    
    all_dataframes = [
        (create_user_meta_dataframe(conn, username), 'Meta', False, 0),
        (create_journeys_dataframe(**kwargs), 'Fahrtenliste', False, 1),
        (create_user_stats_dataframe(**kwargs), 'Statistiken', False, 0),
        (stop_dataframe := create_user_stops_dataframe(**kwargs), 'Haltestelle', True, 0),
        (create_user_stops_dataframe_with_inter(**kwargs, base_dataframe=stop_dataframe), 'Haltestellen (mit Via)', True, 0),
        (create_user_extended_stats(**kwargs, base_dataframe=stop_dataframe), 'erweiterte Statistik', False, 0),
        (create_user_grouped_frame(**kwargs, first_argument='category'), 'Kategorie', False, 1),
        (create_user_grouped_frame(**kwargs, first_argument='operatorName'), 'Betreiber', False, 1),
        (create_user_grouped_frame(**kwargs, first_argument='business'), 'Fahrtgrund', False, 1),
        (create_user_grouped_frame(**kwargs, first_argument='operatorName', second_argument='lineName'), 'Linien', False, 2),
        (create_user_grouped_frame(**kwargs, first_argument='origin.stopName', second_argument='destination.stopName'), 'Verbindungen', False, 2),
        (create_user_grouped_frame(**kwargs, first_argument='locomotiveClass'), 'Baureihe', False, 1),
        (create_user_grouped_frame(**kwargs, first_argument='vehicleNumber'), 'Fahrzeugnummer', False, 1),
        (create_user_grouped_frame(**kwargs, first_argument='clientName'), 'Check-In-Tool', False, 1),
        (create_temporal_stats_dataframe(temporal_base_df, grouping_type='time_of_day', username=username), 'Tageszeiten', False, 1),
        (create_temporal_stats_dataframe(temporal_base_df, grouping_type='weekday', username=username), 'Wochentage', False, 1),
        (create_temporal_stats_dataframe(temporal_base_df, grouping_type='month', username=username), 'Monat', False, 1),
        (create_temporal_stats_dataframe(temporal_base_df, grouping_type='month_year', username=username), 'Monate im Jahr', False, 1),
        (create_temporal_stats_dataframe(temporal_base_df, grouping_type='date', username=username), 'Tag', False, 1),
        (create_temporal_stats_dataframe(temporal_base_df, grouping_type='year', username=username), 'Jahr', False, 1),
        (create_geo_dataframe(conn, username, 'countryName', start_date, end_date), 'Länder', True, 0),
        (create_geo_dataframe(conn, username, 'city', start_date, end_date), 'Städte', True, 0),
        (create_geo_dataframe(conn, username, 'admin1State', start_date, end_date), 'Regionen_Bundesländer', True, 0),
        (create_geo_dataframe(conn, username, 'admin2Dist', start_date, end_date), 'Bezirke_Kreise', True, 0)
    ]

    output_dir = Path('finished_exports/user')
    output_dir.mkdir(parents=True, exist_ok=True)
    
    date_suffix = get_date_suffix(start_date, end_date)
    file_path = output_dir / f"{username}{date_suffix}_data.xlsx"
    
    print(f"[EXCEL] Speichere Datei: {file_path.name}")
    
    with pd.ExcelWriter(file_path, engine='xlsxwriter', datetime_format='dd.mm.yyyy hh:mm') as writer:
        for df, sheet_name, index, numb_of_freeze_row in all_dataframes:
            
            # WORKAROUND FÜR FILTERBARKEIT:
            if index and df.index.name is not None:
                df = df.reset_index()
                index = False
                if numb_of_freeze_row == 0:
                    numb_of_freeze_row = 1
            
            df.to_excel(writer, sheet_name=sheet_name, index=index, header=True)
            current_worksheet = writer.sheets[sheet_name]

            if numb_of_freeze_row > 0:
                current_worksheet.freeze_panes(numb_of_freeze_row, 0)
            if not df.empty:
                max_row = len(df)
                max_col = len(df.columns) - 1 + (1 if index else 0)
                current_worksheet.autofilter(0, 0, max_row, max_col)

            for runner, col in enumerate(df.columns):
                max_length = df[col].dropna().map(lambda x: len(str(x))).max()
                final_width = max((max_length if pd.notna(max_length) else 0), len(str(col)) + 2) + 2
                col_idx = (1 if index else 0) + runner
                current_worksheet.set_column(col_idx, col_idx, final_width)

    print(f'[ERFOLG] Excel gespeichert: {file_path.name}')
    return file_path.name


def create_user_extended_stats(conn: sqlite3.Connection, username: str, start_date: str = '',
                               end_date: str = '', base_dataframe: pd.DataFrame = None) -> pd.DataFrame:
    base_df = base_dataframe.copy() if base_dataframe is not None and not base_dataframe.empty else create_user_stops_dataframe(conn, username, start_date, end_date)
    actionradius = (((base_df['Hub-Distanz-Wert'][1:].sum() / base_df['Besuche'][1:].sum()) ** 2) * 1000).round(3) if not base_df.empty else 0

    params, date_query = params_and_date_query(username, start_date, end_date)
    query_errors = f"SELECT departureManual, departureFromRealtime, departurePlanned, arrivalManual, arrivalFromRealtime, arrivalPlanned, arrivalDelay FROM journeys {date_query}"
    df_err = pd.read_sql_query(query_errors, conn, params=params)

    for col in ['departureManual', 'departureFromRealtime', 'departurePlanned', 'arrivalManual', 'arrivalFromRealtime', 'arrivalPlanned']:
        df_err[col] = pd.to_datetime(df_err[col], errors='coerce', utc=True)

    dep_diff = (df_err['departureManual'] - df_err['departureFromRealtime']).dt.total_seconds() / 60
    arr_diff = (df_err['arrivalManual'] - df_err['arrivalFromRealtime']).dt.total_seconds() / 60
    all_diffs = pd.concat([dep_diff, arr_diff]).dropna()

    rmse_manual = np.sqrt((all_diffs ** 2).mean()) if not all_diffs.empty else 0.0
    sigma_delay = df_err['arrivalDelay'].std() if df_err['arrivalDelay'].count() > 1 else 0.0

    query_coords = f"""
        WITH filtered_journeys AS (
            SELECT originId, destinationId
            FROM journeys {date_query}
        ),
        all_visits AS (
            SELECT originId AS stopId FROM filtered_journeys
            UNION ALL
            SELECT destinationId AS stopId FROM filtered_journeys
        )
        SELECT DISTINCT s.stopName, s.latitude, s.longitude
        FROM all_visits v
        JOIN stops s ON v.stopId = s.id
        WHERE s.latitude IS NOT NULL AND s.longitude IS NOT NULL
    """
    df_coords = pd.read_sql_query(query_coords, conn, params=params)

    nord, sued, ost, west = 'N/A', 'N/A', 'N/A', 'N/A'
    max_dist_stop = 'N/A'
    max_dist_val = 0.0

    if not df_coords.empty:
        df_coords['latitude'] = pd.to_numeric(df_coords['latitude'], errors='coerce')
        df_coords['longitude'] = pd.to_numeric(df_coords['longitude'], errors='coerce')
        df_coords = df_coords.dropna(subset=['latitude', 'longitude'])

    if not df_coords.empty:
        nord = df_coords.loc[df_coords['latitude'].idxmax(), 'stopName']
        sued = df_coords.loc[df_coords['latitude'].idxmin(), 'stopName']
        ost = df_coords.loc[df_coords['longitude'].idxmax(), 'stopName']
        west = df_coords.loc[df_coords['longitude'].idxmin(), 'stopName']

        if not base_df.empty:
            hub_name = base_df.iloc[0, 0]
            hub_row = df_coords[df_coords['stopName'] == hub_name]

            if not hub_row.empty:
                hub_lat, hub_lon = hub_row['latitude'].values[0], hub_row['longitude'].values[0]
                lat1, lon1 = np.radians(hub_lat), np.radians(hub_lon)
                lat2, lon2 = np.radians(df_coords['latitude']), np.radians(df_coords['longitude'])

                dlat = lat2 - lat1
                dlon = lon2 - lon1
                a = np.sin(dlat / 2.0) ** 2 + np.cos(lat1) * np.cos(lat2) * np.sin(dlon / 2.0) ** 2
                dist = 6371 * (2 * np.arctan2(np.sqrt(a), np.sqrt(1 - a)))

                max_idx = dist.argmax()
                max_dist_stop = df_coords.iloc[max_idx]['stopName']
                max_dist_val = dist.iloc[max_idx].round(2)

    return pd.DataFrame({
        'Beschreibung': [
            'Aktionsradius (km)', 'RMSE der Abweichung von Realtimeanbieter', 'Sigma Ankunftsverspätung (σ)',
            'Nördlichste Haltestelle', 'Östlichste Haltestelle', 'Südlichste Haltestelle',
            'Westlichste Haltestelle', 'Größte Distanz zum Hub'
        ],
        'Wert': [
            actionradius, f'{rmse_manual:.2f} Minuten', f'{sigma_delay:.2f} Minuten',
            nord, ost, sued, west, f'{max_dist_stop} ({max_dist_val} km)' if max_dist_stop != 'N/A' else 'N/A'
        ]
    })


def create_user_grouped_frame(conn: sqlite3.Connection, username: str, first_argument: str, second_argument: str = "",
                              third_argument: str = "", fourth_argument: str = "", start_date: str = '',
                              end_date: str = '', sort_by_distance: bool = True) -> pd.DataFrame:
    # Wichtig: "j" als table_prefix übergeben, da wir jetzt Tabellen joinen
    params, date_query = params_and_date_query(username, start_date, end_date, "j")
    
    beautiful_names = {
        'category': 'Kategorie', 
        'operatorName': 'Betreiber', 
        'lineName': 'Linie', 
        'business': 'Fahrtgrund', 
        'vehicleNumber': 'Fahrzeugnummer',
        'origin.stopName': 'Starthaltestelle',
        'destination.stopName': 'Zielhaltestelle'
    }

    select_query = f"{first_argument} AS '{beautiful_names.get(first_argument, first_argument)}'"
    group_query = f"GROUP BY {first_argument}"
    for arg in [second_argument, third_argument, fourth_argument]:
        if arg:
            select_query += f",\n{arg} AS '{beautiful_names.get(arg, arg)}'"
            group_query += f", {arg}"

    query = f""" select {select_query},
        COUNT(j.statusId) AS Fahrten,
        SUM(j.distance) / 1000.0 AS 'Distanz (km)',
        ROUND(SUM(j.distance) * 1.0 / NULLIF(SUM(j.distanceStraightLine), 0), 2) AS 'Umweg Faktor',
        SUM(j.durationReal) AS 'Zeit (min)',
        ROUND(AVG(j.arrivalDelay),2) AS '∅ Verspätung (min)',
        ROUND((SUM(CASE WHEN j.arrivalDelay >= 3 THEN 1 ELSE 0 END) * 100.0) / COUNT(j.statusId),2) AS '3 Minuten und mehr (%)',
        ROUND((SUM(CASE WHEN j.arrivalDelay >= 6 THEN 1 ELSE 0 END) * 100.0) / COUNT(j.statusId),2) AS '6 Minuten und mehr (%)',
        ROUND((SUM(CASE WHEN j.arrivalDelay >= 15 THEN 1 ELSE 0 END) * 100.0) / COUNT(j.statusId),2) AS '15 Minuten und mehr (%)',
        SUM(j.points) AS Punkte,
        ROUND(AVG(j.interStopsNumber),2) AS '∅ Zwischenhalte'
        from journeys j
        LEFT JOIN stops origin ON j.originId = origin.id
        LEFT JOIN stops destination ON j.destinationId = destination.id
        {date_query} {group_query}
        {"ORDER BY [Distanz (km)] DESC, [Fahrten] DESC" if sort_by_distance else "ORDER BY [Fahrten] DESC, [Distanz (km)] DESC"}"""

    df = pd.read_sql_query(query, conn, params=params)
    total_distance = df['Distanz (km)'].sum()
    total_time = df['Zeit (min)'].sum()

    df.insert(loc=df.columns.get_loc('Distanz (km)') + 1, column='Anteil an Gesamtdistanz (%)', value=((df['Distanz (km)']/total_distance)*100).round(2))
    df.insert(loc=df.columns.get_loc('Zeit (min)') + 1, column='Anteil an Gesamtzeit (%)', value=((df['Zeit (min)'] / total_time) * 100).round(2))
    df.insert(loc=df.columns.get_loc('Anteil an Gesamtzeit (%)') + 1, column='∅ Geschwindigkeit', value=(df['Distanz (km)'] / (df['Zeit (min)'] / 60)).round(2))
    return df


def get_temporal_base_dataframe(conn: sqlite3.Connection, username: str, start_date: str = '', end_date: str = '',
                                create_picture: bool = True) -> pd.DataFrame:
    params, date_query = params_and_date_query(username, start_date, end_date, "j")
    query = f"""
    SELECT j.statusId, j.departureFinal, origin.timezone AS origin_tz, j.durationReal, j.distance, j.arrivalDelay
    FROM journeys j
    LEFT JOIN stops origin ON j.originId = origin.id
    {date_query} AND j.departureFinal IS NOT NULL AND j.durationReal > 0
    """
    df_raw = pd.read_sql_query(query, conn, params=params)
    if df_raw.empty:
        return pd.DataFrame()
    df_raw['departure_local'] = [format_utc_to_datatime(dt, tz) for dt, tz in zip(df_raw['departureFinal'], df_raw['origin_tz'])]
    return df_raw


def create_temporal_stats_dataframe(df_raw: pd.DataFrame, grouping_type: str, create_picture: bool = False, username: str = 'N/A') -> pd.DataFrame:
    if df_raw is None or df_raw.empty:
        return pd.DataFrame()

    fragments = []
    unique_dates = set()

    for row in df_raw.itertuples(index=False):
        if pd.isna(row.departure_local):
            continue
        dep = row.departure_local
        duration = int(row.durationReal) if pd.notna(row.durationReal) else 0
        if duration <= 0:
            continue

        dist = float(row.distance) if pd.notna(row.distance) else 0.0
        delay = float(row.arrivalDelay) if pd.notna(row.arrivalDelay) else 0.0
        status_id = row.statusId

        unique_dates.add(dep.date())
        remaining_duration = duration
        current_time = dep

        while remaining_duration > 0:
            if grouping_type == 'time_of_day':
                mins_to_next_boundary = 30 - (current_time.minute % 30)
                mins_to_allocate = min(remaining_duration, mins_to_next_boundary)
            else:
                mins_left_today = 1440 - (current_time.hour * 60 + current_time.minute)
                mins_to_allocate = min(remaining_duration, mins_left_today)

            dist_share = (mins_to_allocate / duration) * dist
            fragment = {'statusId': status_id, 'Minuten': mins_to_allocate, 'Distanz (m)': dist_share, 'Delay': delay if remaining_duration == duration else 0.0}

            if grouping_type == 'time_of_day':
                start_m = (current_time.minute // 30) * 30
                end_time = current_time.replace(minute=start_m, second=0) + timedelta(minutes=30)
                fragment['Bucket'] = f"{current_time.hour:02d}:{start_m:02d} - {end_time.hour % 24:02d}:{end_time.minute:02d}"
            elif grouping_type == 'weekday':
                weekdays = ['Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag', 'Sonntag']
                fragment['Bucket'] = weekdays[current_time.weekday()]
            elif grouping_type == 'date':
                fragment['Bucket'] = current_time.strftime('%Y-%m-%d')
            elif grouping_type == 'month_year':
                fragment['Bucket'] = current_time.strftime('%Y-%m')
            elif grouping_type == 'month':
                fragment['Bucket'] = current_time.strftime('%m')
            elif grouping_type == 'year':
                fragment['Bucket'] = current_time.strftime('%Y')

            fragments.append(fragment)
            remaining_duration -= mins_to_allocate
            current_time += timedelta(minutes=mins_to_allocate)

    df_frag = pd.DataFrame(fragments)
    total_mins_all = df_frag['Minuten'].sum()
    total_dist_all = df_frag['Distanz (m)'].sum()
    total_active_days = len(unique_dates) if len(unique_dates) > 0 else 1

    grouped = df_frag.groupby('Bucket', as_index=False).agg(
        Fahrten=('statusId', 'nunique'),
        Minuten=('Minuten', 'sum'),
        Distanz_m=('Distanz (m)', 'sum'),
        Avg_Delay=('Delay', lambda x: x[x != 0.0].mean() if len(x[x != 0.0]) > 0 else 0.0)
    )

    grouped['Zeit (min)'] = grouped['Minuten'].round(2)
    grouped['Anteil Zeit (%)'] = ((grouped['Minuten'] / total_mins_all) * 100).round(2)
    grouped['Distanz (km)'] = (grouped['Distanz_m'] / 1000).round(2)
    grouped['Anteil Distanz (%)'] = ((grouped['Distanz_m'] / total_dist_all) * 100).round(2)
    grouped['∅ Geschwindigkeit (km/h)'] = (grouped['Distanz (km)'] / (grouped['Minuten'] / 60)).round(2)
    grouped['∅ Ankunftsverspätung (min)'] = grouped['Avg_Delay'].round(2)

    if grouping_type in ['time_of_day', 'weekday']:
        grouped['∅ Minuten pro aktivem Tag'] = (grouped['Minuten'] / total_active_days).round(2)
        grouped['∅ Distanz pro aktivem Tag (km)'] = (grouped['Distanz (km)'] / total_active_days).round(2)

    name_mapping = {
        'time_of_day': 'Zeitfenster', 'weekday': 'Wochentag', 'date': 'Datum',
        'month_year': 'Monat (JJJJ-MM)', 'month': 'Monat', 'year': 'Jahr'
    }
    grouped = grouped.rename(columns={'Bucket': name_mapping.get(grouping_type, 'Bucket')}).drop(columns=['Minuten', 'Distanz_m', 'Avg_Delay'])

    if grouping_type == 'weekday':
        weekdays_order = ['Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag', 'Sonntag']
        grouped['Wochentag'] = pd.Categorical(grouped['Wochentag'], categories=weekdays_order, ordered=True)
        grouped = grouped.sort_values('Wochentag').reset_index(drop=True)

    return grouped


# ==========================================
# PLOTS UND KARTEN (VISUALISIERUNGEN)
# ==========================================

def create_ganglinie_plot(username: str, grouping_type: str, conn: sqlite3.Connection = None,
                          start_date: str = '', end_date: str = '', given_df: pd.DataFrame = None) -> str | None:
    if conn is None:
        conn = init_db()
    if given_df is None or given_df.empty:
        temporal_base_df = get_temporal_base_dataframe(conn=conn, username=username, start_date=start_date, end_date=end_date)
        df = create_temporal_stats_dataframe(temporal_base_df, grouping_type=grouping_type)
    else: 
        df = given_df.copy()

    if df.empty:
        return None

    export_dir = Path('finished_exports/gang_linien')
    export_dir.mkdir(parents=True, exist_ok=True)

    x_col = df.columns[0]
    num_bars = len(df[x_col])

    fig_width = min(max(16, num_bars * 0.15), 50)
    fig, ax = plt.subplots(figsize=(fig_width, 5))
    x = np.arange(num_bars)
    width = 0.35

    ax.bar(x - width / 2, df['Zeit (min)'], width, label='Zeit (min)', color='#004488')
    ax.bar(x + width / 2, df['Distanz (km)'], width, label='Distanz (km)', color='#e63900')

    step = max(1, num_bars // 40)
    ax.set_xticks(x[::step])
    ax.set_xticklabels(df[x_col].iloc[::step], rotation=45, ha='right')

    ax.legend(loc='center left', bbox_to_anchor=(1, 0.5), frameon=False)
    ax.yaxis.grid(True, linestyle='-', which='major', color='lightgrey', alpha=0.7)
    ax.set_axisbelow(True)
    plt.tight_layout()

    date_suffix = get_date_suffix(start_date, end_date)
    file_path = export_dir / f"{username}_{grouping_type}{date_suffix}_ganglinie.png"
    plt.savefig(file_path, dpi=300, bbox_inches='tight')
    plt.close()
    return file_path.name


def create_verspaetungs_verteilungen_plot(username: str | list, conn: sqlite3.Connection = None,
                                          start_date: str = '', end_date: str = '') -> list[str]:
    from scipy.stats import norm, poisson, gaussian_kde
    if conn is None:
        conn = init_db()
    usernames = [username] if isinstance(username, str) else username
    created_files = []

    for user in usernames:
        params, date_query = params_and_date_query(user, start_date, end_date)
        query = f"SELECT arrivalDelay FROM journeys {date_query} AND arrivalDelay IS NOT NULL"
        df = pd.read_sql_query(query, conn, params=params)
        if df.empty:
            continue

        delays = df['arrivalDelay'].dropna()
        mu, sigma = delays.mean(), delays.std()

        delays_poisson = delays[delays >= 0]
        lambda_val = delays_poisson.mean() if not delays_poisson.empty else 0

        kde = gaussian_kde(delays)
        kde.set_bandwidth(bw_method=0.1)

        export_dir = Path('finished_exports/verspaetungs_plots')
        export_dir.mkdir(parents=True, exist_ok=True)

        fig, ax = plt.subplots(figsize=(10, 6))
        xlim_min = max(delays.min() - 2, -10)
        xlim_max = min(delays.quantile(0.99) + 5, 60)
        bins = np.arange(xlim_min, xlim_max + 2, 1) - 0.5
        
        ax.hist(delays, bins=bins, density=True, alpha=0.5, color='#004488', edgecolor='black', label='Reale Daten')

        x_cont = np.linspace(xlim_min, xlim_max, 500)
        ax.plot(x_cont, norm.pdf(x_cont, mu, sigma), 'r-', linewidth=2.5, label=f'Gauß (μ={mu:.2f}, σ={sigma:.2f})')

        x_poisson = np.arange(0, int(xlim_max) + 1)
        ax.plot(x_poisson, poisson.pmf(x_poisson, lambda_val), 'g--', linewidth=2.5, label=f'Poisson (λ={lambda_val:.2f})')
        ax.plot(x_cont, kde.evaluate(x_cont), color='purple', linestyle='-', linewidth=2.5, label='KDE')

        ax.axvline(mu, color='red', linestyle='--', linewidth=1.5, label=f'Mittelwert ({mu:.2f} min)')
        ax.axvline(mu + sigma, color='orange', linestyle=':', linewidth=1.5)
        ax.axvline(mu - sigma, color='orange', linestyle=':', linewidth=1.5)

        ax.set_title(f'Ankunftsverspätung: Verteilungs-Vergleich ({user})', fontsize=12, fontweight='bold')
        ax.set_xlabel('Verspätung (Minuten)', fontsize=10)
        ax.set_ylabel('Dichte', fontsize=10)
        ax.set_xlim(xlim_min, xlim_max)

        ax.legend(loc='upper right')
        ax.yaxis.grid(True)
        ax.set_axisbelow(True)
        plt.tight_layout()

        date_suffix = get_date_suffix(start_date, end_date)
        file_path = export_dir / f"{user}_verspaetungs_verteilungen_plot{date_suffix}.png"
        plt.savefig(file_path, dpi=300, bbox_inches='tight')
        plt.close()
        created_files.append(file_path.name)

    return created_files


def fetch_polylines(conn: sqlite3.Connection, api_token: str, status_ids: list, log_callback=None) -> dict:
    global remaining
    polyline_data = {}
    current_index = 0
    total = len(status_ids)
    cursor = conn.cursor()

    while total > current_index:
        batch = status_ids[current_index:current_index + 50]
        joined_ids = "%2C".join(str(sid) for sid in batch)
        url = f"https://traewelling.de/api/v1/polyline/{joined_ids}"

        repeat_bool = True
        while repeat_bool:
            try:
                response = requests.get(url, headers={
                    'accept': 'application/json',
                    'Authorization': f'Bearer {api_token}',
                    'User-Agent': 'TraewellingAnalysis_v9 tee_beutel'
                })
                response.raise_for_status()
                remaining = response.headers.get('x-ratelimit-remaining')

                features = response.json().get('data', {}).get('features', [])

                # Verarbeite und schreibe die Daten direkt in die Datenbank
                batch_geom = {}
                for feature in features:
                    sid = feature.get('properties', {}).get('statusId')
                    geometry = feature.get('geometry')
                    if sid and geometry:
                        geom_str = json.dumps(geometry, ensure_ascii=False)
                        batch_geom[sid] = geom_str
                        polyline_data[sid] = geom_str

                for sid in batch:
                    geom_str = batch_geom.get(sid, "")
                    cursor.execute("REPLACE INTO polylines(statusId, geometry) VALUES(?, ?)", (sid, geom_str))

                conn.commit()  # Speichert direkt nach jedem 50er-Block ab

                repeat_bool = False
                current_index += 50
                rem_lines = max(0, total - current_index)
                cd_msg = f'   🗺️ Polylines-Countdown: Noch {rem_lines} von {total} Linien verbleibend... (Limit: {remaining})'
                print(cd_msg)
                safe_log(log_callback, cd_msg, replace=(current_index > 50))
            except Exception as e:
                if isinstance(e, requests.exceptions.HTTPError) and e.response.status_code == 429:
                    sleep(30)
                else:
                    repeat_bool = False
                    current_index += 50

                    # Schreibe auch im Fehlerfall (vermutlich Leerdaten) in die DB,
                    # um endlose Retries bei fehlenden Polylines zu vermeiden.
                    for sid in batch:
                        cursor.execute("REPLACE INTO polylines(statusId, geometry) VALUES(?, ?)", (sid, ""))
                    conn.commit()

    if total > 0:
        safe_log(log_callback, f"   ✅ Alle {total} Streckengrafiken (Polylines) erfolgreich heruntergeladen.")
    return polyline_data

def create_stops_map(conn: sqlite3.Connection, usernames: bool | str | list = True, start_date: str = '',
                     end_date: str = '', log_callback=None, lite_version: bool = False) -> str | None:
    if usernames is True:
        usernames = pd.read_sql('SELECT username FROM users', conn)['username'].tolist()
    elif isinstance(usernames, str):
        usernames = [usernames]

    cursor = conn.cursor()
    version_str = " (Lite-Version)" if lite_version else ""
    msg = f"\n[MAP] Erzeuge interaktive Haltestellenkarte{version_str} für {len(usernames)} User..."
    print(msg)
    safe_log(log_callback, msg)

    colors = [
        '#4363d8', '#3cb44b', '#911eb4', '#42d4f4', '#f032e6',
        '#469990', '#000075', '#9A6324', '#808000', '#800000',
        '#aaffc3', '#ffd8b1', '#a9a9a9', '#e6194B', '#f58231',
        '#ffe119', '#bfef45', '#fabed4', '#dcbeff', '#fffac8'
    ]
    user_color_map = {}
    is_single_user = len(usernames) == 1
    for i, u in enumerate(usernames):
        user_color_map[u] = '#4363d8' if is_single_user else colors[i % len(colors)]

    placeholders = ','.join(['?'] * len(usernames))
    date_query = ""
    params_date = []
    if start_date:
        date_query += " AND arrivalFinal > ?"
        params_date.append(start_date)
    if end_date:
        date_query += " AND departureFinal < ?"
        end_date_str = end_date + 'T23:59:59' if len(end_date) == 10 else end_date
        params_date.append(end_date_str)

    base_params = usernames + params_date

    query = f"""
        SELECT username, originId, departurePlanned 
        FROM journeys 
        WHERE username IN ({placeholders}) {date_query} {get_shared_visibility_query()}
        UNION ALL
        SELECT username, destinationId, arrivalPlanned 
        FROM journeys 
        WHERE username IN ({placeholders}) {date_query} {get_shared_visibility_query()}
    """
    cursor.execute(query, base_params + base_params)
    visits = cursor.fetchall()

    if not visits:
        return None

    cursor.execute(
        "SELECT id, stopName, latitude, longitude, timezone FROM stops WHERE latitude IS NOT NULL AND longitude IS NOT NULL")
    stops_dict = {row[0]: {'name': row[1], 'lat': float(row[2]), 'lon': float(row[3]), 'tz': row[4] or 'Europe/Berlin',
                           'visits': []} for row in cursor.fetchall() if row[2] and row[3]}

    all_years = set()
    all_months = set()
    all_users = set()

    for row in visits:
        u, s_id, t = row
        if s_id in stops_dict:
            if t:
                year = t[:4]
                month = t[5:7]
            else:
                year = 'Unbekannt'
                month = 'Unbekannt'

            stops_dict[s_id]['visits'].append({'u': u, 't': t})
            all_years.add(year)
            all_months.add(month)
            all_users.add(u)

    stops_with_visits = {k: v for k, v in stops_dict.items() if len(v['visits']) > 0}
    if not stops_with_visits:
        return None

    export_dir = Path('finished_exports/polylines')
    export_dir.mkdir(parents=True, exist_ok=True)

    names_joined = "_".join(sorted(usernames)[:6])
    if len(usernames) > 6:
        names_joined += f"_and_{len(usernames) - 6}_more"

    date_suffix = get_date_suffix(start_date, end_date)
    lite_suffix = "_lite" if lite_version else ""
    file_path = export_dir / f"{names_joined}{date_suffix}_stops_map{lite_suffix}.html"

    user_colors_json = json.dumps(user_color_map, ensure_ascii=False)

    if lite_version:
        # MASSIVE DATENKOMPRESSION FÜR LITE-VERSION: Berechnung im Backend
        user_stats = {u: {'hub_id': None, 'max_v': 0, 'h_lat': None, 'h_lon': None} for u in usernames}
        stop_agg = {}

        for s_id, s_data in stops_with_visits.items():
            u_v = Counter()
            u_l = {}
            for v in s_data['visits']:
                u, t = v['u'], v['t']
                u_v[u] += 1
                if u not in u_l or (t and str(t) > str(u_l.get(u, ""))):
                    u_l[u] = t

            stop_agg[s_id] = {
                'lat': s_data['lat'], 'lon': s_data['lon'], 'name': s_data['name'], 'tz': s_data['tz'],
                'u_v': dict(u_v), 'u_l': u_l, 'tot': sum(u_v.values())
            }

            for u, c in u_v.items():
                if c > user_stats[u]['max_v']:
                    user_stats[u]['max_v'] = c
                    user_stats[u]['hub_id'] = s_id

        for u, st in user_stats.items():
            if st['hub_id'] and st['hub_id'] in stop_agg:
                st['h_lat'] = stop_agg[st['hub_id']]['lat']
                st['h_lon'] = stop_agg[st['hub_id']]['lon']

        lite_markers = []
        for s_id, s_data in stop_agg.items():
            pop_users = []
            for u, visits_count in s_data['u_v'].items():
                hub_dist_str = '∞'
                if user_stats[u]['hub_id'] != s_id and user_stats[u]['h_lat'] is not None:
                    d = calculate_haversine(s_data['lat'], s_data['lon'], user_stats[u]['h_lat'],
                                            user_stats[u]['h_lon'])
                    val = visits_count * math.sqrt(d) / 1000
                    hub_dist_str = f"{val:.3f}"

                lvis = s_data['u_l'].get(u)
                if lvis:
                    try:
                        dt = datetime.fromisoformat(lvis.replace('Z', '+00:00')).astimezone(ZoneInfo(s_data['tz']))
                        lvis_str = dt.strftime('%d.%m.%Y, %H:%M Uhr')
                    except Exception:
                        lvis_str = str(lvis)[:16].replace('T', ' ') + ' Uhr'
                else:
                    lvis_str = 'Unbekannt'

                pop_users.append({'u': u, 'v': visits_count, 'hd': hub_dist_str, 'lv': lvis_str})

            pop_users.sort(key=lambda x: x['v'], reverse=True)
            lite_markers.append({
                'lat': s_data['lat'], 'lon': s_data['lon'], 'n': s_data['name'],
                't': s_data['tot'], 'ud': s_data['u_v'], 'pu': pop_users
            })

        lite_json_str = json.dumps(lite_markers, ensure_ascii=False, separators=(',', ':'))

        html_content = f"""<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
    <title>Haltestellen Karte (Lite) - {names_joined}</title>
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <style>
        body, html {{ margin: 0; padding: 0; height: 100%; font-family: Arial, sans-serif; }}
        #map {{ width: 100%; height: 100vh; }}
        .info.legend {{ background: white; padding: 10px; border-radius: 5px; box-shadow: 0 0 15px rgba(0,0,0,0.2); max-height: 45vh; display: flex; flex-direction: column; }}
        .legend-header {{ font-weight: bold; cursor: pointer; user-select: none; display: flex; justify-content: space-between; align-items: center; }}
        .legend-content {{ margin-top: 10px; overflow-y: auto; padding-right: 5px; display: block; }}
        .color-box {{ width: 15px; height: 15px; display: inline-block; margin-right: 8px; border-radius: 3px; vertical-align: middle; }}
        .legend-item {{ margin-bottom: 5px; font-size: 13px; }}
    </style>
</head>
<body>
    <div id="map"></div>
    <script>
        var markersData = {lite_json_str};
        var userColors = {user_colors_json};

        var map = L.map('map', {{ preferCanvas: true }}).setView([51.1657, 10.4515], 6);
        L.tileLayer('https://tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{ maxZoom: 19, attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors' }}).addTo(map);

        var controlBox = L.control({{position: 'topright'}});
        controlBox.onAdd = function () {{
            var div = L.DomUtil.create('div');
            if (Object.keys(userColors).length > 0) {{
                var html = '<div class="info legend">';
                html += '<div class="legend-header" id="legend-header" onclick="toggleLegend()">Legende ▼</div>';
                html += '<div class="legend-content" id="legend-content">';
                for (var user in userColors) {{
                    html += '<div class="legend-item"><span class="color-box" style="background:' + userColors[user] + '"></span>' + user + '</div>';
                }}
                html += '</div></div>';
                div.innerHTML = html;
            }}
            L.DomEvent.disableClickPropagation(div);
            return div;
        }};
        controlBox.addTo(map);

        window.toggleLegend = function() {{
            var content = document.getElementById('legend-content');
            var header = document.getElementById('legend-header');
            if (content.style.display === 'none') {{
                content.style.display = 'block'; header.innerHTML = 'Legende ▼';
            }} else {{
                content.style.display = 'none'; header.innerHTML = 'Legende ◀';
            }}
        }};

        var markersLayer = L.layerGroup();

        markersData.forEach(s => {{
            var gradientParts = [];
            var currentP = 0;

            for (var u in s.ud) {{
                var p = (s.ud[u] / s.t) * 100;
                var c = userColors[u] || '#004488';
                if (p > 0) {{
                    gradientParts.push(c + ' ' + currentP + '% ' + (currentP + p) + '%');
                    currentP += p;
                }}
            }}

            var bg = gradientParts.length > 0 ? 'conic-gradient(' + gradientParts.join(', ') + ')' : '#004488';
            var html = '<div style="background: ' + bg + '; border-radius: 50%; width: 30px; height: 30px; display: flex; align-items: center; justify-content: center; box-shadow: 0 0 4px rgba(0,0,0,0.6);">';
            html += '<div style="background: white; color: #333; border-radius: 50%; width: 22px; height: 22px; display: flex; align-items: center; justify-content: center; font-size: 9px; font-weight: bold;">' + s.t + '</div></div>';

            var icon = L.divIcon({{ html: html, className: '', iconSize: [30, 30], iconAnchor: [15, 15] }});
            var marker = L.marker([s.lat, s.lon], {{icon: icon}});

            var popStr = '<b>' + s.n + '</b><br><hr style="margin:5px 0;">';
            popStr += '<div style="max-height: 180px; overflow-y: auto; padding-right: 5px; margin-right: -5px;">';

            s.pu.forEach(function(pu) {{
                var uColor = userColors[pu.u] || '#004488';
                var distText = pu.hd === '∞' ? '∞' : pu.hd + ' km';
                popStr += '<div style="margin-bottom: 6px;"><span style="color:'+uColor+'">&#9608;</span> <b>' + pu.u + '</b>: ' + pu.v + ' Besuche / ' + distText + '<br><span style="font-size:10px; color:#555; margin-left:15px;">Letzter Besuch: ' + pu.lv + '</span></div>';
            }});
            popStr += '</div>';

            marker.bindPopup(popStr);
            markersLayer.addLayer(marker);
        }});

        markersLayer.addTo(map);
    </script>
</body>
</html>"""
    else:
        # Dies ist das Standard-HTML mit Filterfunktionen und allen Rohdaten
        stops_json = json.dumps(stops_with_visits, ensure_ascii=False)
        years_list = json.dumps(sorted(list(all_years), reverse=True))
        months_list = json.dumps(sorted(list(all_months)))
        users_list = json.dumps(sorted(list(all_users)))

        html_content = f"""<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <title>Haltestellen Karte - {names_joined}</title>
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <script src="https://unpkg.com/leaflet.heat@0.2.0/dist/leaflet-heat.js"></script>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/choices.js/public/assets/styles/choices.min.css" />
    <script src="https://cdn.jsdelivr.net/npm/choices.js/public/assets/scripts/choices.min.js"></script>
    <style>
        body, html {{ margin: 0; padding: 0; height: 100%; font-family: Arial, sans-serif; }}
        #map {{ width: 100%; height: 100vh; }}
        .info.legend, .filter-box {{ background: white; padding: 10px; border-radius: 5px; box-shadow: 0 0 15px rgba(0,0,0,0.2); max-height: 45vh; display: flex; flex-direction: column; }}
        .filter-box {{ margin-top: 10px; width: 300px; }}
        .filter-header, .legend-header {{ font-weight: bold; cursor: pointer; user-select: none; display: flex; justify-content: space-between; align-items: center; }}
        .filter-content, .legend-content {{ margin-top: 10px; overflow-y: auto; padding-right: 5px; }}
        .filter-content {{ display: block; }}
        .filter-box label {{ font-size: 13px; margin-top: 8px; display: block; font-weight: bold; color: #444; }}
        select {{ width: 100%; padding: 5px; margin-bottom: 10px; border: 1px solid #ccc; border-radius: 3px; font-size: 13px; }}
        .choices {{ margin-bottom: 5px; }}
        .choices__inner {{ min-height: 32px; padding: 2px 4px 0px 4px; font-size: 13px; border-radius: 4px; }}
        .choices__list--multiple .choices__item {{ font-size: 11px; padding: 2px 5px; margin: 2px; border-radius: 3px; }}
        .choices__list--dropdown {{ font-size: 13px; z-index: 9999; }}
        .info.legend {{ line-height: 1.5; color: #333; }}
        .color-box {{ width: 15px; height: 15px; display: inline-block; margin-right: 8px; border-radius: 3px; vertical-align: middle; }}
        .legend-item {{ margin-bottom: 5px; font-size: 13px; }}
    </style>
</head>
<body>
    <div id="map"></div>
    <script>
        var stopsData = {stops_json};
        var years = {years_list};
        var months = {months_list};
        var users = {users_list};
        var userColors = {user_colors_json};

        var map = L.map('map').setView([51.1657, 10.4515], 6);
        L.tileLayer('https://tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
            maxZoom: 19, attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
        }}).addTo(map);

        var markersLayer = L.layerGroup();
        var overlayLayer = L.layerGroup().addTo(map);
        var heatLayer = null;

        var controlBox = L.control({{position: 'topright'}});
        controlBox.onAdd = function () {{
            var div = L.DomUtil.create('div');
            var html = '';

            if (Object.keys(userColors).length > 0) {{
                html += '<div class="info legend" style="margin-bottom:10px;">';
                html += '<div class="legend-header" id="legend-header" onclick="toggleLegend()">Legende ▼</div>';
                html += '<div class="legend-content" id="legend-content">';
                for (var user in userColors) {{
                    html += '<div class="legend-item"><span class="color-box" style="background:' + userColors[user] + '"></span>' + user + '</div>';
                }}
                html += '</div></div>';
            }}

            html += '<div class="filter-box"><div class="filter-header" onclick="toggleFilter()">Karten-Filter ▼</div>';
            html += '<div class="filter-content" id="filter-content">';
            html += '<label>Darstellung</label><select id="mapType" onchange="renderMap()"><option value="pins">Pins mit Zahlen</option><option value="heatmap">Heatmap</option></select>';
            html += '<label>Wert</label><select id="valType" onchange="renderMap()"><option value="visits">Anzahl Besuche</option><option value="hub">Hub-Distanz-Wert</option></select>';
            html += '<label>Zusatz-Ebenen</label>';
            html += '<label style="font-weight:normal; margin-top: 2px;"><input type="checkbox" id="showHubs" checked onchange="renderMap()"> Haupthaltestellen umkreisen</label>';
            html += '<label style="font-weight:normal; margin-top: 2px;"><input type="checkbox" id="showRadii" checked onchange="renderMap()"> Aktionsradius einzeichnen</label>';
            html += '<label style="font-weight:normal; margin-top: 2px;"><input type="checkbox" id="showBounds" checked onchange="renderMap()"> Extrema-Box einzeichnen</label>';
            html += '<label>User</label><select id="filterUser" multiple>';
            users.forEach(u => html += '<option value="'+u+'">'+u+'</option>');
            html += '</select>';
            html += '<label>Jahr</label><select id="filterYear" multiple>';
            years.forEach(y => html += '<option value="'+y+'">'+y+'</option>');
            html += '</select>';
            html += '<label>Monat</label><select id="filterMonth" multiple>';
            months.forEach(m => html += '<option value="'+m+'">'+m+'</option>');
            html += '</select>';
            html += '</div></div>';
            div.innerHTML = html;
            L.DomEvent.disableClickPropagation(div);
            return div;
        }};
        controlBox.addTo(map);

        window.toggleFilter = function() {{
            var content = document.getElementById('filter-content');
            content.style.display = content.style.display === 'none' ? 'block' : 'none';
        }};

        window.toggleLegend = function() {{
            var content = document.getElementById('legend-content');
            var header = document.getElementById('legend-header');
            if (content.style.display === 'none') {{
                content.style.display = 'block'; header.innerHTML = 'Legende ▼';
            }} else {{
                content.style.display = 'none'; header.innerHTML = 'Legende ◀';
            }}
        }};

        document.querySelectorAll('.filter-box select[multiple]').forEach(el => {{
            new Choices(el, {{ removeItemButton: true, searchEnabled: true, placeholderValue: 'Alle (Klicken zum Filtern)', shouldSort: false }});
            el.addEventListener('change', renderMap);
        }});

        function getSelected(id) {{
            var el = document.getElementById(id);
            if(!el || el.selectedOptions.length === 0) return ['all'];
            return Array.from(el.selectedOptions).map(o => o.value);
        }}

        function getDist(lat1, lon1, lat2, lon2) {{
            var R = 6371000;
            var dLat = (lat2-lat1)*Math.PI/180;
            var dLon = (lon2-lon1)*Math.PI/180;
            var a = Math.sin(dLat/2)*Math.sin(dLat/2) + Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*Math.sin(dLon/2)*Math.sin(dLon/2);
            return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
        }}

        function renderMap() {{
            var uF = getSelected('filterUser');
            var yF = getSelected('filterYear');
            var mF = getSelected('filterMonth');
            var typeF = document.getElementById('mapType').value;
            var valF = document.getElementById('valType').value;

            var showHubs = document.getElementById('showHubs').checked;
            var showRadii = document.getElementById('showRadii').checked;
            var showBounds = document.getElementById('showBounds').checked;

            markersLayer.clearLayers();
            overlayLayer.clearLayers();
            if(heatLayer) {{ map.removeLayer(heatLayer); }}

            var userStats = {{}};
            users.forEach(u => {{
                if (uF.includes('all') || uF.includes(u)) {{
                    userStats[u] = {{
                        visitsByStop: {{}}, totalVisits: 0, uniqueStops: 0,
                        minLat: 90, maxLat: -90, minLon: 180, maxLon: -180,
                        minLatId: null, maxLatId: null, minLonId: null, maxLonId: null,
                        hubId: null, maxVisits: 0, actionRadius: 0, maxDistId: null, maxDistVal: 0
                    }};
                }}
            }});

            var processedStops = [];
            for (var id in stopsData) {{
                var stop = stopsData[id];
                var stopUserVisits = {{}};
                var stopUserLast = {{}};
                var stopTotal = 0;

                stop.visits.forEach(v => {{
                    if (userStats[v.u]) {{
                        var year = v.t ? v.t.substring(0,4) : 'Unbekannt';
                        var month = v.t ? v.t.substring(5,7) : 'Unbekannt';
                        if((yF.includes('all') || yF.includes(year)) && (mF.includes('all') || mF.includes(month))) {{
                            stopUserVisits[v.u] = (stopUserVisits[v.u] || 0) + 1;
                            stopTotal++;
                            userStats[v.u].totalVisits++;

                            if (stop.lat < userStats[v.u].minLat) {{ userStats[v.u].minLat = stop.lat; userStats[v.u].minLatId = id; }}
                            if (stop.lat > userStats[v.u].maxLat) {{ userStats[v.u].maxLat = stop.lat; userStats[v.u].maxLatId = id; }}
                            if (stop.lon < userStats[v.u].minLon) {{ userStats[v.u].minLon = stop.lon; userStats[v.u].minLonId = id; }}
                            if (stop.lon > userStats[v.u].maxLon) {{ userStats[v.u].maxLon = stop.lon; userStats[v.u].maxLonId = id; }}

                            if(!stopUserLast[v.u] || (v.t && v.t > stopUserLast[v.u])) {{
                                stopUserLast[v.u] = v.t;
                            }}
                        }}
                    }}
                }});

                if (stopTotal > 0) {{
                    processedStops.push({{ 
                        id: id, name: stop.name, lat: stop.lat, lon: stop.lon, tz: stop.tz, 
                        total: stopTotal, userVisits: stopUserVisits, userLast: stopUserLast 
                    }});

                    for (var u in stopUserVisits) {{
                        userStats[u].visitsByStop[id] = stopUserVisits[u];
                        userStats[u].uniqueStops++;
                        if (stopUserVisits[u] > userStats[u].maxVisits) {{
                            userStats[u].maxVisits = stopUserVisits[u];
                            userStats[u].hubId = id;
                        }}
                    }}
                }}
            }}

            for (var u in userStats) {{
                var stats = userStats[u];
                if (stats.hubId) {{
                    var hubStop = stopsData[stats.hubId];
                    stats.hubLat = hubStop.lat;
                    stats.hubLon = hubStop.lon;
                    var relevanceSum = 0;
                    var totalVisitsForRadius = 0;
                    stats.maxDistVal = 0;
                    stats.maxDistId = null;

                    for (var id in stats.visitsByStop) {{
                        if (id !== stats.hubId) {{
                            var v = stats.visitsByStop[id];
                            var d = getDist(hubStop.lat, hubStop.lon, stopsData[id].lat, stopsData[id].lon);
                            relevanceSum += v * Math.sqrt(d) / 1000;
                            totalVisitsForRadius += v;
                            if (d > stats.maxDistVal) {{ stats.maxDistVal = d; stats.maxDistId = id; }}
                        }}
                    }}

                    if (totalVisitsForRadius > 0) {{
                        var r = relevanceSum / totalVisitsForRadius;
                        stats.actionRadius = (r * r) * 1000000;
                    }} else {{
                        stats.actionRadius = 0;
                    }}
                }}
            }}

            var legendContent = document.getElementById('legend-content');
            if (legendContent) {{
                var newLegendHtml = '';
                var sortedLegendUsers = Object.keys(userStats).filter(function(u) {{ return userStats[u].totalVisits > 0; }});
                sortedLegendUsers.sort(function(a, b) {{ return userStats[b].uniqueStops - userStats[a].uniqueStops; }});

                sortedLegendUsers.forEach(function(u) {{
                    var ar = (userStats[u].actionRadius / 1000).toFixed(2);
                    var us = userStats[u].uniqueStops;
                    newLegendHtml += '<div class="legend-item" style="margin-bottom:8px;"><span class="color-box" style="background:' + (userColors[u]||'#004488') + '"></span>' + u + ' <br><span style="font-size:10px; color:#555; margin-left:23px;">' + us + ' Haltestellen | AR: ' + ar + ' km</span></div>';
                }});
                legendContent.innerHTML = newLegendHtml;
            }}

            var heatData = [];
            var maxPlotVal = 0;
            var drawnHubsCount = {{}};

            processedStops.forEach(s => {{
                var totalDisp = 0;
                var userDisp = {{}};
                var isHubForSomeone = false;

                for (var u in s.userVisits) {{
                    if (userStats[u].hubId === s.id) isHubForSomeone = true;
                    var uVal = s.userVisits[u];
                    if (valF === 'hub') {{
                        if (userStats[u].hubId === s.id) {{
                            uVal = 0;
                        }} else {{
                            var d = getDist(s.lat, s.lon, userStats[u].hubLat, userStats[u].hubLon);
                            uVal = s.userVisits[u] * Math.sqrt(d) / 1000;
                        }}
                    }}
                    userDisp[u] = uVal;
                    totalDisp += uVal;
                }}

                if (totalDisp > maxPlotVal) maxPlotVal = totalDisp;

                if (typeF === 'heatmap') {{
                    heatData.push([s.lat, s.lon, totalDisp]);
                }} else {{
                    var gradientParts = [];
                    var currentP = 0;
                    var displayStr = valF === 'hub' ? totalDisp.toFixed(1) : s.total;
                    if (valF === 'hub' && isHubForSomeone) displayStr = '∞';
                    var totalForColor = 0;
                    for (var u in s.userVisits) {{
                        var cv = s.userVisits[u];
                        if (valF === 'hub' && userStats[u].hubId !== s.id) cv = userDisp[u];
                        totalForColor += cv;
                    }}

                    for (var u in s.userVisits) {{
                        var cv = s.userVisits[u];
                        if (valF === 'hub' && userStats[u].hubId !== s.id) cv = userDisp[u];
                        var p = (cv / (totalForColor || 1)) * 100;
                        var c = userColors[u] || '#004488';
                        if (p > 0) {{
                            gradientParts.push(c + ' ' + currentP + '% ' + (currentP + p) + '%');
                            currentP += p;
                        }}
                    }}

                    var bg = gradientParts.length > 0 ? 'conic-gradient(' + gradientParts.join(', ') + ')' : '#004488';
                    var html = '<div style="background: ' + bg + '; border-radius: 50%; width: 30px; height: 30px; display: flex; align-items: center; justify-content: center; box-shadow: 0 0 4px rgba(0,0,0,0.6);">';
                    html += '<div style="background: white; color: #333; border-radius: 50%; width: 22px; height: 22px; display: flex; align-items: center; justify-content: center; font-size: 9px; font-weight: bold;">' + displayStr + '</div></div>';

                    var icon = L.divIcon({{ html: html, className: '', iconSize: [30, 30], iconAnchor: [15, 15] }});
                    var marker = L.marker([s.lat, s.lon], {{icon: icon}});

                    var popupUsers = [];
                    for (var u in s.userVisits) {{
                        var visits = s.userVisits[u];
                        var hubDistRaw = '∞';
                        if (userStats[u].hubId !== s.id) {{
                            var currentD = getDist(s.lat, s.lon, userStats[u].hubLat, userStats[u].hubLon);
                            hubDistRaw = visits * Math.sqrt(currentD) / 1000;
                        }}

                        var hubDistStr = hubDistRaw === '∞' ? '∞' : hubDistRaw.toFixed(3);
                        var sortVal = valF === 'hub' ? (hubDistRaw === '∞' ? Infinity : hubDistRaw) : visits;
                        var lVisRaw = s.userLast[u];
                        var lVisFormatted = 'Unbekannt';
                        if (lVisRaw) {{
                            try {{
                                var d_date = new Date(lVisRaw);
                                lVisFormatted = d_date.toLocaleString('de-DE', {{ timeZone: s.tz || 'Europe/Berlin', day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }}) + ' Uhr';
                            }} catch(e) {{
                                lVisFormatted = lVisRaw.replace('T', ' ').substring(0, 16) + ' Uhr';
                            }}
                        }}

                        var extrema = [];
                        if (userStats[u].maxLatId === s.id) extrema.push("Nördlichste");
                        if (userStats[u].minLatId === s.id) extrema.push("Südlichste");
                        if (userStats[u].maxLonId === s.id) extrema.push("Östlichste");
                        if (userStats[u].minLonId === s.id) extrema.push("Westlichste");

                        var extStr = "";
                        if (extrema.length > 0) extStr += extrema.join('/') + " Hst.";
                        if (userStats[u].hubId === s.id) {{
                            if (extStr !== "") extStr += " | ";
                            extStr += "Haupthaltestelle";
                        }}
                        if (userStats[u].maxDistId === s.id) {{
                            if (extStr !== "") extStr += " | ";
                            extStr += "Weiteste Entfernung (" + (userStats[u].maxDistVal/1000).toFixed(1) + " km)";
                        }}

                        popupUsers.push({{ u: u, visits: visits, hubDistStr: hubDistStr, sortVal: sortVal, lVisFormatted: lVisFormatted, extStr: extStr }});
                    }}

                    popupUsers.sort(function(a, b) {{ return b.sortVal - a.sortVal; }});

                    var popStr = '<b>' + s.name + '</b><br><hr style="margin:5px 0;">';
                    popStr += '<div style="max-height: 180px; overflow-y: auto; padding-right: 5px; margin-right: -5px;">';

                    popupUsers.forEach(function(pu) {{
                        var uColor = userColors[pu.u] || '#004488';
                        var valStr = valF === 'hub' ? pu.hubDistStr + (pu.hubDistStr === '∞' ? '' : ' km') + ' / ' + pu.visits + ' Besuche' : pu.visits + ' Besuche / ' + pu.hubDistStr + (pu.hubDistStr === '∞' ? '' : ' km');

                        popStr += '<div style="margin-bottom: 6px;"><span style="color:'+uColor+'">&#9608;</span> <b>' + pu.u + '</b>: ' + valStr + '<br><span style="font-size:10px; color:#555; margin-left:15px;">Letzter Besuch: ' + pu.lVisFormatted + '</span>';
                        if (pu.extStr) {{
                            popStr += '<br><span style="font-size:10px; color:#c0392b; margin-left:15px; font-weight:bold;">' + pu.extStr + '</span>';
                        }}
                        popStr += '</div>';
                    }});
                    popStr += '</div>';

                    marker.bindPopup(popStr);
                    markersLayer.addLayer(marker);
                }}
            }});

            if (typeF === 'heatmap') {{
                heatLayer = L.heatLayer(heatData, {{radius: 25, max: maxPlotVal, blur: 15}}).addTo(map);
            }} else {{
                markersLayer.addTo(map);
            }}

            for (var u in userStats) {{
                var stats = userStats[u];
                var color = userColors[u]||'#4363d8';

                if (stats.totalVisits > 0) {{
                    if (showHubs && stats.hubLat && stats.hubLon) {{
                        var hCount = drawnHubsCount[stats.hubId] || 0;
                        var dynRadius = 20 + (hCount * 6);
                        drawnHubsCount[stats.hubId] = hCount + 1;

                        L.circleMarker([stats.hubLat, stats.hubLon], {{
                            color: color, weight: 4, fillOpacity: 0, radius: dynRadius
                        }}).addTo(overlayLayer).bindPopup('<b>Haupthaltestelle von ' + u + '</b>');
                    }}

                    if (showRadii && stats.actionRadius > 0) {{
                        L.circle([stats.hubLat, stats.hubLon], {{
                            color: color, weight: 2, fillOpacity: 0.1, radius: stats.actionRadius, dashArray: '5, 5'
                        }}).addTo(overlayLayer).bindPopup('<b>Aktionsradius von ' + u + '</b><br>' + (stats.actionRadius/1000).toFixed(2) + ' km');
                    }}

                    if (showBounds && stats.minLat <= 90 && stats.maxLat >= -90) {{
                        var bounds = [[stats.minLat, stats.minLon], [stats.maxLat, stats.maxLon]];
                        L.rectangle(bounds, {{
                            color: color, weight: 2, fillOpacity: 0.05, dashArray: '10, 10'
                        }}).addTo(overlayLayer).bindPopup('<b>Extrema von ' + u + '</b>');
                    }}
                }}
            }}
        }}

        renderMap();
    </script>
</body>
</html>"""

    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(html_content)

    print(f"[ERFOLG] Interaktive Haltestellenkarte exportiert: {file_path.name}")
    return file_path.name


def create_polyline(conn: sqlite3.Connection, api_token: str, usernames: bool | str | list = True, start_date: str = '',
                    end_date: str = '', log_callback=None, lite_version: bool = False) -> str | None:
    if usernames is True:
        usernames = pd.read_sql('SELECT username FROM users', conn)['username'].tolist()
    elif isinstance(usernames, str):
        usernames = [usernames]

    cursor = conn.cursor()
    version_str = " (Lite-Version)" if lite_version else ""
    msg = f"\n[MAP] Erzeuge interaktive Streckenkarte{version_str} für {len(usernames)} User..."
    print(msg)
    safe_log(log_callback, msg)

    colors = [
        '#4363d8', '#3cb44b', '#911eb4', '#42d4f4', '#f032e6',
        '#469990', '#000075', '#9A6324', '#808000', '#800000',
        '#aaffc3', '#ffd8b1', '#a9a9a9', '#e6194B', '#f58231',
        '#ffe119', '#bfef45', '#fabed4', '#dcbeff', '#fffac8'
    ]
    business_map = {0: 'Privat', 1: 'Geschäftlich', 2: 'Arbeitsweg'}

    user_relevant_ids = {}
    all_relevant_ids = set()
    status_to_line, status_to_color, status_to_category = {}, {}, {}
    status_to_operator, status_to_business, status_to_year = {}, {}, {}
    status_to_weekday, status_to_tod, status_to_delay, status_to_loco = {}, {}, {}, {}

    for username in usernames:
        params, date_query = params_and_date_query(username, start_date, end_date)
        query = f"SELECT statusId, lineName, routeColor, category, operatorName, business, departurePlanned, arrivalDelay, locomotiveClass FROM journeys {date_query} AND distance > 0"
        cursor.execute(query, params)

        user_ids = []
        for row in cursor.fetchall():
            sid, line_name, route_color, category, operator = row[0], row[1] or "Unbekannt", row[2], row[
                3] or "Unbekannt", row[4] or "Unbekannt"
            business, dep_str, delay, loco = business_map.get(row[5], 'Privat'), row[6], row[7], row[8] or "Unbekannt"

            if dep_str:
                year = dep_str[:4]
                try:
                    dt = datetime.fromisoformat(dep_str.replace('Z', '+00:00')).astimezone(ZoneInfo('Europe/Berlin'))
                    weekdays = ['1. Montag', '2. Dienstag', '3. Mittwoch', '4. Donnerstag', '5. Freitag', '6. Samstag',
                                '7. Sonntag']
                    weekday = weekdays[dt.weekday()]
                    if 5 <= dt.hour < 11:
                        tod = "1. Morgens (05-11 Uhr)"
                    elif 11 <= dt.hour < 15:
                        tod = "2. Mittags (11-15 Uhr)"
                    elif 15 <= dt.hour < 21:
                        tod = "3. Abends (15-21 Uhr)"
                    else:
                        tod = "4. Nachts (21-05 Uhr)"
                except ValueError:
                    weekday, tod = "Unbekannt", "Unbekannt"
            else:
                year, weekday, tod = "Unbekannt", "Unbekannt", "Unbekannt"

            if delay is None:
                delay_cat = "Unbekannt"
            elif delay < 3:
                delay_cat = "1. Pünktlich (< 3 Min)"
            elif delay < 15:
                delay_cat = "2. Leicht verspätet (3-14 Min)"
            elif delay < 60:
                delay_cat = "3. Verspätet (15-59 Min)"
            else:
                delay_cat = "4. Stark verspätet (>= 60 Min)"

            user_ids.append(sid)
            all_relevant_ids.add(sid)
            status_to_line[sid] = line_name
            status_to_color[sid] = route_color
            status_to_category[sid] = category
            status_to_operator[sid] = operator
            status_to_business[sid] = business
            status_to_year[sid] = year
            status_to_weekday[sid] = weekday
            status_to_tod[sid] = tod
            status_to_delay[sid] = delay_cat
            status_to_loco[sid] = loco

        if user_ids:
            user_relevant_ids[username] = user_ids

    if not all_relevant_ids:
        return None

    all_relevant_ids_list = list(all_relevant_ids)
    chunk_size = 900
    existing_geometries = {}

    for j in range(0, len(all_relevant_ids_list), chunk_size):
        chunk = all_relevant_ids_list[j:j + chunk_size]
        placeholders = ','.join(['?'] * len(chunk))
        cursor.execute(f"SELECT statusId, geometry FROM polylines WHERE statusId IN ({placeholders})", chunk)
        for sid, geom in cursor.fetchall():
            existing_geometries[sid] = geom

    missing_ids = list(all_relevant_ids - set(existing_geometries.keys()))
    if missing_ids:
        new_polylines = fetch_polylines(conn, api_token, missing_ids, log_callback=log_callback)
        for sid in missing_ids:
            existing_geometries[sid] = new_polylines.get(sid, "")

    export_features = []
    user_color_map = {}
    segment_to_trips = defaultdict(list)
    is_single_user = len(usernames) == 1

    for i, username in enumerate(usernames):
        if username not in user_relevant_ids:
            continue
        legend_color = '#4363d8' if is_single_user else colors[i % len(colors)]
        user_color_map[username] = legend_color
        relevant_ids = user_relevant_ids[username]

        for sid in relevant_ids:
            geom_str = existing_geometries.get(sid)
            if not geom_str:
                continue
            try:
                geom = json.loads(geom_str)
            except json.JSONDecodeError:
                continue

            coords = geom.get("coordinates", [])
            if geom.get("type") == "LineString":
                lines = [coords]
            elif geom.get("type") == "MultiLineString":
                lines = coords
            else:
                continue

            if is_single_user:
                db_color = status_to_color.get(sid)
                if db_color and str(db_color).strip() and str(db_color).strip() != 'None':
                    trip_color = str(db_color).strip()
                    if not trip_color.startswith('#'):
                        trip_color = '#' + trip_color
                else:
                    trip_color = '#4363d8'
            else:
                trip_color = legend_color

            trip_info = {
                "statusId": sid, "username": username, "color": trip_color,
                "lineName": status_to_line.get(sid, "Unbekannt"),
                "category": status_to_category.get(sid, "Unbekannt"),
                "operator": status_to_operator.get(sid, "Unbekannt"),
                "business": status_to_business.get(sid, "Privat"), "year": status_to_year.get(sid, "Unbekannt"),
                "weekday": status_to_weekday.get(sid, "Unbekannt"), "tod": status_to_tod.get(sid, "Unbekannt"),
                "delay": status_to_delay.get(sid, "Unbekannt"), "loco": status_to_loco.get(sid, "Unbekannt")
            }

            for line in lines:
                for j in range(len(line) - 1):
                    p1, p2 = (round(line[j][0], 5), round(line[j][1], 5)), (round(line[j + 1][0], 5),
                                                                            round(line[j + 1][1], 5))
                    if p1 == p2:
                        continue
                    seg = tuple(sorted((p1, p2)))
                    if trip_info not in segment_to_trips[seg]:
                        segment_to_trips[seg].append(trip_info)

    trip_combo_to_segments = defaultdict(list)
    trip_combo_to_info = {}

    for seg, trips in segment_to_trips.items():
        trips_sorted = sorted(trips, key=lambda x: x["statusId"])
        combo_key = tuple(t["statusId"] for t in trips_sorted)
        trip_combo_to_segments[combo_key].append(seg)
        if combo_key not in trip_combo_to_info:
            trip_combo_to_info[combo_key] = trips_sorted

    segment_id_counter = 0
    for combo_key, segments in trip_combo_to_segments.items():
        segment_id_counter += 1
        trips = trip_combo_to_info[combo_key]
        all_segment_trips = []
        color_counts = {}
        for trip in trips:
            all_segment_trips.append({
                "statusId": trip['statusId'], "username": trip['username'], "color": trip['color'],
                "lineName": trip['lineName'], "category": trip['category'], "operator": trip['operator'],
                "business": trip['business'], "year": trip['year'], "weekday": trip['weekday'],
                "tod": trip['tod'], "delay": trip['delay'], "loco": trip['loco']
            })
            grouping_key = f"{trip['username']}_{trip['color']}"
            if grouping_key not in color_counts:
                color_counts[grouping_key] = {
                    "count": 0, "color": trip['color'], "statusIds": [], "username": trip['username'],
                    "categories": set(), "operators": set(), "businesses": set(), "years": set(),
                    "weekdays": set(), "tods": set(), "delays": set(), "locos": set()
                }
            color_counts[grouping_key]["count"] += 1
            color_counts[grouping_key]["statusIds"].append(trip['statusId'])
            color_counts[grouping_key]["categories"].add(trip['category'])
            color_counts[grouping_key]["operators"].add(trip['operator'])
            color_counts[grouping_key]["businesses"].add(trip['business'])
            color_counts[grouping_key]["years"].add(trip['year'])
            color_counts[grouping_key]["weekdays"].add(trip['weekday'])
            color_counts[grouping_key]["tods"].add(trip['tod'])
            color_counts[grouping_key]["delays"].add(trip['delay'])
            color_counts[grouping_key]["locos"].add(trip['loco'])

        multiline_coords = [[list(p1), list(p2)] for p1, p2 in segments]
        line_weight = 5 if len(color_counts) == 1 else 7

        for group_key, data in color_counts.items():
            props = {
                "sid": segment_id_counter,
                "c": data["count"],
                "sw": line_weight,
                "str": data["color"]
            }

            if lite_version:
                lite_trips = []
                for t in all_segment_trips:
                    if t['username'] == data['username'] and t['color'] == data['color']:
                        lite_trips.append({'u': t['username'], 'id': t['statusId'], 'l': t['lineName']})
                props['tr'] = lite_trips
            else:
                props.update({
                    "stroke": data["color"],
                    "stroke-width": line_weight,
                    "stroke-opacity": 0.9,
                    "segmentId": segment_id_counter,
                    "count": data["count"],
                    "tripDetails": all_segment_trips,
                    "statusIds": data["statusIds"],
                    "username": data["username"],
                    "categories": list(data["categories"]),
                    "operators": list(data["operators"]),
                    "businesses": list(data["businesses"]),
                    "years": list(data["years"]),
                    "weekdays": list(data["weekdays"]),
                    "tods": list(data["tods"]),
                    "delays": list(data["delays"]),
                    "locos": list(data["locos"])
                })

            export_features.append({
                "type": "Feature",
                "geometry": {"type": "MultiLineString", "coordinates": multiline_coords},
                "properties": props
            })

    if export_features:
        feature_collection = {"type": "FeatureCollection", "features": export_features}
        export_dir = Path('finished_exports/polylines')
        export_dir.mkdir(parents=True, exist_ok=True)

        names_joined = "_".join(sorted(usernames)[:6])
        if len(usernames) > 6:
            names_joined += f"_and_{len(usernames) - 6}_more"

        date_suffix = get_date_suffix(start_date, end_date)
        lite_suffix = "_lite" if lite_version else ""
        file_path = export_dir / f"{names_joined}{date_suffix}_map{lite_suffix}.html"
        geojson_string = json.dumps(feature_collection, ensure_ascii=False, separators=(',', ':'))
        user_colors_json = json.dumps(user_color_map if not is_single_user else {}, ensure_ascii=False)

        if lite_version:
            html_content = f"""<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
    <title>Traewelling Karte (Lite) - {names_joined}</title>
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <style>
        body, html {{ margin: 0; padding: 0; height: 100%; font-family: Arial, sans-serif; }}
        #map {{ width: 100%; height: 100vh; }}
        .info.legend {{ background: white; padding: 10px; border-radius: 5px; box-shadow: 0 0 15px rgba(0,0,0,0.2); line-height: 1.5; color: #333; }}
        .legend-header {{ font-weight: bold; cursor: pointer; user-select: none; display: flex; justify-content: space-between; align-items: center; }}
        .legend-content {{ margin-top: 10px; display: block; }}
        .color-box {{ width: 15px; height: 15px; display: inline-block; margin-right: 8px; border-radius: 3px; vertical-align: middle; }}
        .legend-item {{ margin-bottom: 5px; font-size: 13px; }}
    </style>
</head>
<body>
    <div id="map"></div>
    <script>
        var map = L.map('map', {{ preferCanvas: true }}).setView([51.1657, 10.4515], 6);

        L.tileLayer('https://tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
            maxZoom: 19, attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
        }}).addTo(map);

        var geojsonData = {geojson_string};
        var userColors = {user_colors_json};

        var segments = {{}};
        geojsonData.features.forEach(function(f) {{
            var sid = f.properties.sid;
            if (!segments[sid]) segments[sid] = [];
            segments[sid].push(f);
        }});

        var cycleLength = 50.0;
        for (var sid in segments) {{
            var feats = segments[sid];
            var totalTrips = 0;
            feats.forEach(function(f) {{ totalTrips += f.properties.c; }});
            var currentOffset = 0.0;
            feats.forEach(function(f) {{
                if (feats.length === 1) {{
                    f.properties.dynamicDashArray = null;
                    f.properties.dynamicDashOffset = null;
                }} else {{
                    var proportion = f.properties.c / totalTrips;
                    var lineLen = (proportion * cycleLength).toFixed(2);
                    var gapLen = (cycleLength - lineLen).toFixed(2);
                    f.properties.dynamicDashArray = lineLen + ", " + gapLen;
                    f.properties.dynamicDashOffset = "-" + currentOffset;
                    currentOffset += parseFloat(lineLen);
                }}
            }});
        }}

        var geojsonLayer = L.geoJSON(geojsonData, {{
            style: function (feature) {{
                return {{
                    color: feature.properties.str, weight: feature.properties.sw,
                    opacity: 0.9, dashArray: feature.properties.dynamicDashArray,
                    dashOffset: feature.properties.dynamicDashOffset
                }};
            }},
            onEachFeature: function (feature, layer) {{
                if (feature.properties && feature.properties.tr) {{
                    var popupHtml = "<div style='font-family: Arial, sans-serif; min-width: 220px;'>";
                    popupHtml += "<b style='font-size: 14px;'>" + feature.properties.tr.length + " Fahrt(en) auf dem Abschnitt</b><br>";
                    popupHtml += "<div style='max-height: 180px; overflow-y: auto; margin-top: 8px; border-top: 1px solid #ccc; padding-top: 5px;'>";
                    feature.properties.tr.forEach(function(trip) {{
                        popupHtml += "<span style='color:" + feature.properties.str + "; font-size: 16px;'>&#9608;</span> <b>" + trip.u + ":</b> <a href='https://traewelling.de/status/" + trip.id + "' target='_blank'>" + trip.id + "</a> (" + trip.l + ")<br>";
                    }});
                    popupHtml += "</div></div>";
                    layer.bindPopup(popupHtml);
                }}
            }}
        }}).addTo(map);

        var controlBox = L.control({{position: 'topright'}});
        controlBox.onAdd = function (map) {{
            var div = L.DomUtil.create('div');
            if (Object.keys(userColors).length > 0) {{
                var legendHtml = '<div class="info legend" style="margin-bottom:10px;">';
                legendHtml += '<div class="legend-header" id="legend-header" onclick="toggleLegend()">Legende ▼</div>';
                legendHtml += '<div class="legend-content" id="legend-content">';
                for (var user in userColors) {{
                    legendHtml += '<div class="legend-item"><span class="color-box" style="background:' + userColors[user] + '"></span>' + user + '</div>';
                }}
                legendHtml += '</div></div>';
                div.innerHTML += legendHtml;
            }}
            L.DomEvent.disableClickPropagation(div);
            return div;
        }};
        controlBox.addTo(map);

        window.toggleLegend = function() {{
            var content = document.getElementById('legend-content');
            var header = document.getElementById('legend-header');
            if (content.style.display === 'none') {{
                content.style.display = 'block'; header.innerHTML = 'Legende ▼';
            }} else {{
                content.style.display = 'none'; header.innerHTML = 'Legende ◀';
            }}
        }};

        if (geojsonData.features.length > 0) {{ map.fitBounds(geojsonLayer.getBounds()); }}
    </script>
</body>
</html>"""
        else:
            html_content = f"""<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <title>Traewelling Karte - {names_joined}</title>
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/choices.js/public/assets/styles/choices.min.css" />
    <script src="https://cdn.jsdelivr.net/npm/choices.js/public/assets/scripts/choices.min.js"></script>
    <style>
        body, html {{ margin: 0; padding: 0; height: 100%; font-family: Arial, sans-serif; }}
        #map {{ width: 100%; height: 100vh; }}
        .info.legend {{ background: white; padding: 10px; border-radius: 5px; box-shadow: 0 0 15px rgba(0,0,0,0.2); line-height: 1.5; color: #333; }}
        .filter-box {{ background: white; padding: 10px; border-radius: 5px; box-shadow: 0 0 15px rgba(0,0,0,0.2); margin-top: 10px; max-height: 80vh; overflow-y: auto; width: 300px; }}
        .filter-header {{ font-weight: bold; cursor: pointer; user-select: none; display: flex; justify-content: space-between; align-items: center; margin-bottom: 5px; }}
        .filter-content {{ display: none; margin-top: 10px; }}
        .filter-box label {{ font-size: 13px; margin-top: 8px; display: block; font-weight: bold; color: #444; }}
        .filter-hint {{ font-size: 11px; color: #777; display: block; margin-bottom: 8px; line-height: 1.3; }}
        .choices {{ margin-bottom: 5px; }}
        .choices__inner {{ min-height: 32px; padding: 2px 4px 0px 4px; font-size: 13px; border-radius: 4px; }}
        .choices__list--multiple .choices__item {{ font-size: 11px; padding: 2px 5px; margin: 2px; border-radius: 3px; }}
        .choices__list--dropdown {{ font-size: 13px; z-index: 9999; }}
        .legend-header {{ font-weight: bold; cursor: pointer; user-select: none; display: flex; justify-content: space-between; align-items: center; }}
        .legend-content {{ margin-top: 10px; display: block; }}
        .color-box {{ width: 15px; height: 15px; display: inline-block; margin-right: 8px; border-radius: 3px; vertical-align: middle; }}
        .legend-item {{ margin-bottom: 5px; font-size: 13px; }}
    </style>
</head>
<body>
    <div id="map"></div>
    <script>
        var map = L.map('map').setView([51.1657, 10.4515], 6);

        L.tileLayer('https://tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
            maxZoom: 19, attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
        }}).addTo(map);

        var geojsonData = {geojson_string};
        var userColors = {user_colors_json};
        var geojsonLayer;
        var currentHighlightId = null;

        var allUsers = new Set();
        var allCategories = new Set();
        var allOperators = new Set();
        var allBusinesses = new Set();
        var allYears = new Set();
        var allWeekdays = new Set();
        var allTods = new Set();
        var allDelays = new Set();
        var allLocos = new Set();

        geojsonData.features.forEach(f => {{
            if (f.properties.username) allUsers.add(f.properties.username);
            f.properties.categories.forEach(c => allCategories.add(c));
            f.properties.operators.forEach(o => allOperators.add(o));
            f.properties.businesses.forEach(b => allBusinesses.add(b));
            f.properties.years.forEach(y => allYears.add(y));
            f.properties.weekdays.forEach(w => allWeekdays.add(w));
            f.properties.tods.forEach(t => allTods.add(t));
            f.properties.delays.forEach(d => allDelays.add(d));
            f.properties.locos.forEach(l => allLocos.add(l));
        }});

        function getSelectedValues(selectId) {{
            var el = document.getElementById(selectId);
            if (!el) return ['all'];
            var values = Array.from(el.selectedOptions).map(opt => opt.value);
            if (values.length === 0) return ['all'];
            return values;
        }}

        function renderMap() {{
            var userFilter = getSelectedValues('filterUser');
            var catFilter = getSelectedValues('filterCategory');
            var opFilter = getSelectedValues('filterOperator');
            var busFilter = getSelectedValues('filterBusiness');
            var yearFilter = getSelectedValues('filterYear');
            var weekdayFilter = getSelectedValues('filterWeekday');
            var todFilter = getSelectedValues('filterTod');
            var delayFilter = getSelectedValues('filterDelay');
            var locoFilter = getSelectedValues('filterLoco');

            if (geojsonLayer) {{ map.removeLayer(geojsonLayer); }}

            var visibleFeatures = geojsonData.features.filter(function(feature) {{
                var p = feature.properties;
                return (userFilter.includes('all') || userFilter.includes(p.username)) &&
                       (catFilter.includes('all') || p.categories.some(c => catFilter.includes(c))) &&
                       (opFilter.includes('all') || p.operators.some(o => opFilter.includes(o))) &&
                       (busFilter.includes('all') || p.businesses.some(b => busFilter.includes(b))) &&
                       (yearFilter.includes('all') || p.years.some(y => yearFilter.includes(y))) &&
                       (weekdayFilter.includes('all') || p.weekdays.some(w => weekdayFilter.includes(w))) &&
                       (todFilter.includes('all') || p.tods.some(t => todFilter.includes(t))) &&
                       (delayFilter.includes('all') || p.delays.some(d => delayFilter.includes(d))) &&
                       (locoFilter.includes('all') || p.locos.some(l => locoFilter.includes(l)));
            }});

            var segments = {{}};
            visibleFeatures.forEach(function(f) {{
                var sid = f.properties.segmentId;
                if (!segments[sid]) segments[sid] = [];
                segments[sid].push(f);
            }});

            var cycleLength = 50.0;
            for (var sid in segments) {{
                var feats = segments[sid];
                var totalTrips = 0;
                feats.forEach(function(f) {{ totalTrips += f.properties.count; }});
                var currentOffset = 0.0;
                feats.forEach(function(f) {{
                    if (feats.length === 1) {{
                        f.properties.dynamicDashArray = null;
                        f.properties.dynamicDashOffset = null;
                    }} else {{
                        var proportion = f.properties.count / totalTrips;
                        var lineLen = (proportion * cycleLength).toFixed(2);
                        var gapLen = (cycleLength - lineLen).toFixed(2);
                        f.properties.dynamicDashArray = lineLen + ", " + gapLen;
                        f.properties.dynamicDashOffset = "-" + currentOffset;
                        currentOffset += parseFloat(lineLen);
                    }}
                }});
            }}

            geojsonLayer = L.geoJSON(visibleFeatures, {{
                style: function (feature) {{
                    return {{
                        color: feature.properties.stroke, weight: feature.properties["stroke-width"],
                        opacity: feature.properties["stroke-opacity"], dashArray: feature.properties.dynamicDashArray,
                        dashOffset: feature.properties.dynamicDashOffset
                    }};
                }},
                onEachFeature: function (feature, layer) {{
                    if (feature.properties && feature.properties.tripDetails) {{
                        var visibleTrips = feature.properties.tripDetails.filter(function(trip) {{
                            return (userFilter.includes('all') || userFilter.includes(trip.username)) &&
                                   (catFilter.includes('all') || catFilter.includes(trip.category)) &&
                                   (opFilter.includes('all') || opFilter.includes(trip.operator)) &&
                                   (busFilter.includes('all') || busFilter.includes(trip.business)) &&
                                   (yearFilter.includes('all') || yearFilter.includes(trip.year)) &&
                                   (weekdayFilter.includes('all') || weekdayFilter.includes(trip.weekday)) &&
                                   (todFilter.includes('all') || todFilter.includes(trip.tod)) &&
                                   (delayFilter.includes('all') || delayFilter.includes(trip.delay)) &&
                                   (locoFilter.includes('all') || locoFilter.includes(trip.loco));
                        }});
                        var popupHtml = "<div style='font-family: Arial, sans-serif; min-width: 220px;'>";
                        popupHtml += "<b style='font-size: 14px;'>" + visibleTrips.length + " Fahrt(en) auf dem Abschnitt</b><br>";
                        popupHtml += "<div style='max-height: 180px; overflow-y: auto; margin-top: 8px; border-top: 1px solid #ccc; padding-top: 5px;'>";
                        visibleTrips.forEach(function(trip) {{
                            popupHtml += "<span style='color:" + trip.color + "; font-size: 16px;'>&#9608;</span> <b>" + trip.username + ":</b> <a href='https://traewelling.de/status/" + trip.statusId + "' target='_blank'>" + trip.statusId + "</a> (<a href='#' onclick='highlightTrip(" + trip.statusId + "); return false;' title='Fahrt auf Karte hervorheben' style='text-decoration: underline; cursor: pointer;'>" + trip.lineName + "</a>)<br>";
                        }});
                        popupHtml += "</div></div>";
                        layer.bindPopup(popupHtml);
                    }}
                }}
            }}).addTo(map);
        }}

        window.highlightTrip = function(statusId) {{
            if (currentHighlightId === statusId) {{
                geojsonLayer.eachLayer(function(layer) {{
                    var props = layer.feature.properties;
                    layer.setStyle({{ opacity: props["stroke-opacity"], weight: props["stroke-width"], dashArray: props.dynamicDashArray }});
                }});
                currentHighlightId = null;
            }} else {{
                currentHighlightId = statusId;
                geojsonLayer.eachLayer(function(layer) {{
                    var props = layer.feature.properties;
                    if (props.statusIds && props.statusIds.includes(statusId)) {{
                        layer.setStyle({{ opacity: 1.0, weight: props["stroke-width"] + 5, dashArray: null }});
                        layer.bringToFront();
                    }} else {{
                        layer.setStyle({{ opacity: 0.15, weight: props["stroke-width"], dashArray: props.dynamicDashArray }});
                    }}
                }});
            }}
        }}

        map.on('click', function() {{ if (currentHighlightId !== null) highlightTrip(currentHighlightId); }});

        var controlBox = L.control({{position: 'topright'}});
        controlBox.onAdd = function (map) {{
            var div = L.DomUtil.create('div');

            if (Object.keys(userColors).length > 0) {{
                var legendHtml = '<div class="info legend" style="margin-bottom:10px;">';
                legendHtml += '<div class="legend-header" id="legend-header" onclick="toggleLegend()">Legende ▼</div>';
                legendHtml += '<div class="legend-content" id="legend-content">';
                for (var user in userColors) {{
                    legendHtml += '<div class="legend-item"><span class="color-box" style="background:' + userColors[user] + '"></span>' + user + '</div>';
                }}
                legendHtml += '</div></div>';
                div.innerHTML += legendHtml;
            }}

            var filterHtml = '<div class="filter-box">';
            filterHtml += '<div class="filter-header" id="filter-header" onclick="toggleFilter()">Karten-Filter ◀</div>';
            filterHtml += '<div class="filter-content" id="filter-content">';
            filterHtml += '<span class="filter-hint">Klicke in ein Feld, um Optionen auszuwählen.<br>Lässt du es leer, wird alles angezeigt.</span>';
            filterHtml += '<label>User</label><select id="filterUser" multiple>';
            Array.from(allUsers).sort().forEach(u => {{ filterHtml += '<option value="'+u+'">'+u+'</option>'; }});
            filterHtml += '</select><label>Jahr</label><select id="filterYear" multiple>';
            Array.from(allYears).sort().reverse().forEach(y => {{ if(y !== 'Unbekannt') filterHtml += '<option value="'+y+'">'+y+'</option>'; }});
            filterHtml += '</select><label>Monat/Wochentag</label><select id="filterWeekday" multiple>';
            Array.from(allWeekdays).sort().forEach(w => {{ if(w !== 'Unbekannt') filterHtml += '<option value="'+w+'">'+w.substring(3)+'</option>'; }});
            filterHtml += '</select><label>Tageszeit</label><select id="filterTod" multiple>';
            Array.from(allTods).sort().forEach(t => {{ if(t !== 'Unbekannt') filterHtml += '<option value="'+t+'">'+t.substring(3)+'</option>'; }});
            filterHtml += '</select><label>Pünktlichkeit</label><select id="filterDelay" multiple>';
            Array.from(allDelays).sort().forEach(d => {{ if(d !== 'Unbekannt') filterHtml += '<option value="'+d+'">'+d.substring(3)+'</option>'; }});
            filterHtml += '</select><label>Verkehrsmittel</label><select id="filterCategory" multiple>';
            Array.from(allCategories).sort().forEach(cat => {{ if(cat !== 'Unbekannt') filterHtml += '<option value="'+cat+'">'+cat+'</option>'; }});
            filterHtml += '</select><label>Fahrtgrund</label><select id="filterBusiness" multiple>';
            Array.from(allBusinesses).sort().forEach(b => {{ filterHtml += '<option value="'+b+'">'+b+'</option>'; }});
            filterHtml += '</select><label>Baureihe</label><select id="filterLoco" multiple>';
            Array.from(allLocos).sort().forEach(l => {{ if(l !== 'Unbekannt' && l !== 'N/A') filterHtml += '<option value="'+l+'">'+l+'</option>'; }});
            filterHtml += '</select><label>Betreiber</label><select id="filterOperator" multiple>';
            Array.from(allOperators).sort().forEach(op => {{ if(op !== 'Unbekannt' && op !== 'Unknown Operator') filterHtml += '<option value="'+op+'">'+op+'</option>'; }});
            filterHtml += '</select></div></div>';
            div.innerHTML += filterHtml;
            L.DomEvent.disableClickPropagation(div);
            return div;
        }};
        controlBox.addTo(map);

        var contentDiv = document.getElementById('filter-content');
        contentDiv.style.display = 'block'; 

        var selects = document.querySelectorAll('.filter-box select[multiple]');
        selects.forEach(function(el) {{
            new Choices(el, {{
                removeItemButton: true, searchEnabled: true, placeholder: true, placeholderValue: 'Alle (Klicken zum Filtern)',
                searchPlaceholderValue: 'Suchen...', itemSelectText: '', noResultsText: 'Keine Ergebnisse gefunden',
                noChoicesText: 'Keine weiteren Optionen', shouldSort: false
            }});
            el.addEventListener('change', renderMap);
        }});

        contentDiv.style.display = 'none';

        window.toggleFilter = function() {{
            var header = document.getElementById('filter-header');
            if (contentDiv.style.display === 'none' || contentDiv.style.display === '') {{
                contentDiv.style.display = 'block'; header.innerHTML = 'Karten-Filter ▼';
            }} else {{
                contentDiv.style.display = 'none'; header.innerHTML = 'Karten-Filter ◀';
            }}
        }}

        window.toggleLegend = function() {{
            var content = document.getElementById('legend-content');
            var header = document.getElementById('legend-header');
            if (content.style.display === 'none') {{
                content.style.display = 'block'; header.innerHTML = 'Legende ▼';
            }} else {{
                content.style.display = 'none'; header.innerHTML = 'Legende ◀';
            }}
        }}

        renderMap();
        if (geojsonData.features.length > 0) {{ map.fitBounds(geojsonLayer.getBounds()); }}
    </script>
</body>
</html>"""

        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(html_content)

        print(f"[ERFOLG] Interaktive Streckenkarte exportiert: {file_path.name}")
        return file_path.name


def get_trip_color(trip_id, shared_trips):
    """Generiert eine eindeutige, deterministische Farbe für gemeinsame Fahrten."""
    if pd.isna(trip_id) or trip_id not in shared_trips:
        return "#3498db"  # Standard-Blau für Einzelfahrten

    # Generiere eine konstante Farbe basierend auf der Trip-ID
    hex_color = hashlib.md5(str(trip_id).encode('utf-8')).hexdigest()[:6]
    return f"#{hex_color}"


def create_timeline(conn: sqlite3.Connection, usernames: list[str], start_date: str = '', end_date: str = '',
                    log_callback=None) -> str | None:
    if isinstance(usernames, str):
        usernames = [usernames]
    elif usernames is True:
        usernames = pd.read_sql('SELECT username FROM users', conn)['username'].tolist()

    if not usernames:
        return None

    msg = f"\n[TIMELINE] Erzeuge interaktive Timeline für {len(usernames)} User..."
    print(msg)
    safe_log(log_callback, msg)

    placeholders = ','.join(['?'] * len(usernames))
    params_date = []
    date_query_j = ""

    if start_date:
        date_query_j += " AND COALESCE(j.arrivalFinal, j.arrivalPlanned) > ?"
        params_date.append(start_date)
    if end_date:
        date_query_j += " AND COALESCE(j.departureFinal, j.departurePlanned) < ?"
        end_date_str = end_date + 'T23:59:59' if len(end_date) == 10 else end_date
        params_date.append(end_date_str)

    query = f"""
    SELECT 
        j.statusId, j.username, j.tripId, j.lineName, 
        j.departurePlanned, j.departureFinal, 
        j.arrivalPlanned, j.arrivalFinal,
        o.stopName as originName, d.stopName as destName
    FROM journeys j
    LEFT JOIN stops o ON j.originId = o.id
    LEFT JOIN stops d ON j.destinationId = d.id
    WHERE j.username IN ({placeholders}) {date_query_j} {get_shared_visibility_query('j')}
    AND j.departurePlanned IS NOT NULL 
    AND j.arrivalPlanned IS NOT NULL
    ORDER BY j.username, j.departurePlanned
    """

    df = pd.read_sql_query(query, conn, params=usernames + params_date)

    if df.empty:
        safe_log(log_callback, "⚠️ Keine passenden Fahrten für die Timeline im gewählten Zeitraum gefunden.")
        return None

    # Zeitstempel parsen und auf lokale Zeitzone umstellen
    for col in ['departurePlanned', 'departureFinal', 'arrivalPlanned', 'arrivalFinal']:
        df[col] = pd.to_datetime(df[col], utc=True).dt.tz_convert('Europe/Berlin')

    # Falls Final-Zeiten fehlen, Fallback auf Plan-Zeiten
    df['departureFinal'] = df['departureFinal'].fillna(df['departurePlanned'])
    df['arrivalFinal'] = df['arrivalFinal'].fillna(df['arrivalPlanned'])

    # Verspätungen in Minuten berechnen
    df['dep_delay'] = (df['departureFinal'] - df['departurePlanned']).dt.total_seconds() / 60
    df['arr_delay'] = (df['arrivalFinal'] - df['arrivalPlanned']).dt.total_seconds() / 60

    # Daten chronologisch sortieren, um Umstiege zu berechnen
    df.sort_values(by=['username', 'departureFinal'], inplace=True)

    # Identifiziere gemeinsame Fahrten (Trip IDs, die mehr als einmal vorkommen)
    trip_counts = df['tripId'].value_counts()
    shared_trips = trip_counts[trip_counts > 1].index.tolist()

    groups = []
    items = []

    # User als Spalten/Zeilen auf der linken Seite anlegen
    for user in usernames:
        groups.append({
            "id": user,
            "content": f"<b style='font-size: 14px;'>{user}</b>"
        })

    last_arrival = {}
    last_dest = {}

    for _, row in df.iterrows():
        user = row['username']
        trip = row['tripId']
        color = get_trip_color(trip, shared_trips)

        # Umsteigezeiten berechnen (Grauer Block zwischen zwei Fahrten)
        if user in last_arrival:
            transfer_duration = (row['departureFinal'] - last_arrival[user]).total_seconds() / 60
            # Nur Umstiege unter 6 Stunden (360 Min) und ohne Zeitreisen (> 0 Min) anzeigen
            if 0 <= transfer_duration < 360:
                items.append({
                    "id": f"transfer_{row['statusId']}",
                    "group": user,
                    "start": str(last_arrival[user]),
                    "end": str(row['departureFinal']),
                    "content": f"<div style='font-size: 10px; color: #555; text-align: center;'>{int(transfer_duration)} min</div>",
                    "style": "background-color: #e2e8f0; border: 1px dotted #94a3b8; z-index: 0; box-shadow: none;",
                    "title": f"Umstieg: {int(transfer_duration)} Minuten<br>Von: {last_dest.get(user, 'Unbekannt')}<br>Nach: {row['originName']}"
                })

        last_arrival[user] = row['arrivalFinal']
        last_dest[user] = row['destName']

        # Uhrzeiten und Verspätungen für die Detailanzeige formatieren
        dep_str = row['departureFinal'].strftime('%H:%M')
        arr_str = row['arrivalFinal'].strftime('%H:%M')

        dep_delay_str = f"+{int(row['dep_delay'])}" if row['dep_delay'] > 0 else f"{int(row['dep_delay'])}"
        arr_delay_str = f"+{int(row['arr_delay'])}" if row['arr_delay'] > 0 else f"{int(row['arr_delay'])}"

        # HTML-Inhalt der Box: Passt sich über CSS Container-Queries an die Breite an
        details_html = f"""
        <div class="trip-main"><b>{row['lineName']}</b></div>
        <div class="trip-details">
            <b>Start:</b> {row['originName']} ({dep_str} | {dep_delay_str} min)<br>
            <b>Ziel:</b> {row['destName']} ({arr_str} | {arr_delay_str} min)
        </div>
        """

        # 1. Geplanter Block (mit eigener CSS-Klasse plan-trip-item für Versatz nach oben)
        items.append({
            "id": f"{row['statusId']}_plan",
            "group": user,
            "className": "plan-trip-item",
            "start": str(row['departurePlanned']),
            "end": str(row['arrivalPlanned']),
            "content": f"<div style='font-size: 10px; opacity: 0.6;'>PLAN: {row['lineName']}</div>",
            "style": f"background-color: transparent; border: 2px dashed {color}; color: #333; z-index: 1;",
            "title": f"Planmäßige Fahrt: {row['originName']} ➔ {row['destName']}<br>Linie: {row['lineName']}"
        })

        # 2. Tatsächlicher Block (Der "Ist-Verschub")
        items.append({
            "id": f"{row['statusId']}_ist",
            "group": user,
            "className": "ist-trip-item",
            "start": str(row['departureFinal']),
            "end": str(row['arrivalFinal']),
            "content": details_html,
            "style": f"background-color: {color}; color: white; border: 1px solid #222; box-shadow: 2px 2px 4px rgba(0,0,0,0.3); z-index: 2;",
            "title": f"Tatsächliche Fahrt: {row['originName']} ➔ {row['destName']}<br>Linie: {row['lineName']}<br>Abfahrt: {dep_str} ({dep_delay_str})<br>Ankunft: {arr_str} ({arr_delay_str})"
        })

    # Start- und Endzeit für die Ansicht ermitteln (automatisch bei neuesten Fahrten starten)
    min_time = min(df['departurePlanned'].min(), df['departureFinal'].min())
    max_time = max(df['arrivalPlanned'].max(), df['arrivalFinal'].max())
    total_span_sec = (max_time - min_time).total_seconds()

    if total_span_sec <= 2 * 86400:
        window_start = min_time - pd.Timedelta(hours=2)
        window_end = max_time + pd.Timedelta(hours=3)
    else:
        # Standard: Die letzten 48 Stunden der Fahrten anzeigen, passend herausgezoomt
        window_start = max_time - pd.Timedelta(hours=48)
        window_end = max_time + pd.Timedelta(hours=4)

    min_iso = (min_time - pd.Timedelta(days=1)).strftime('%Y-%m-%dT%H:%M:%S')
    max_iso = (max_time + pd.Timedelta(days=1)).strftime('%Y-%m-%dT%H:%M:%S')
    w_start_iso = window_start.strftime('%Y-%m-%dT%H:%M:%S')
    w_end_iso = window_end.strftime('%Y-%m-%dT%H:%M:%S')

    export_dir = Path('finished_exports/timelines')
    export_dir.mkdir(parents=True, exist_ok=True)

    names_joined = "_".join(sorted(usernames)[:6])
    if len(usernames) > 6:
        names_joined += f"_and_{len(usernames) - 6}_more"

    date_suffix = get_date_suffix(start_date, end_date)
    file_path = export_dir / f"{names_joined}{date_suffix}_timeline.html"

    groups_json = json.dumps(groups, ensure_ascii=False)
    items_json = json.dumps(items, ensure_ascii=False)

    html_template = f"""<!DOCTYPE html>
<html lang="de">
<head>
  <meta charset="UTF-8">
  <title>Träwelling Timeline - {names_joined}</title>
  <script type="text/javascript" src="https://unpkg.com/vis-timeline@latest/standalone/umd/vis-timeline-graph2d.min.js"></script>
  <link href="https://unpkg.com/vis-timeline@latest/styles/vis-timeline-graph2d.min.css" rel="stylesheet" type="text/css" />
  <style type="text/css">
    body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f8f9fa; margin: 20px; }}
    h2 {{ color: #2c3e50; margin-bottom: 8px; }}
    .info-bar {{ font-size: 13px; color: #495057; background: #fff; padding: 12px 16px; border-radius: 6px; border: 1px solid #dee2e6; margin-bottom: 15px; line-height: 1.5; }}
    #visualization {{ width: 100%; height: 75vh; border: 1px solid #ced4da; background-color: white; border-radius: 5px; }}
    
    /* Container Queries Setup für zoomabhängige Details */
    .vis-item {{ container-type: inline-size; font-size: 12px; border-radius: 4px; overflow: visible; }}
    .vis-item.vis-range {{ padding: 2px; }}

    /* Vertikaler Versatz der geplanten Fahrtblöcke */
    .vis-item.plan-trip-item {{
        margin-top: -8px !important;
    }}
    
    /* Standardmäßig (ausgezoomt) sind Details versteckt */
    .trip-main {{ white-space: nowrap; overflow: hidden; text-overflow: ellipsis; padding: 2px; }}
    .trip-details {{ display: none; font-size: 10px; line-height: 1.3; margin-top: 4px; white-space: normal; background-color: rgba(0,0,0,0.15); padding: 4px; border-radius: 3px; }}
    
    /* Wenn das Item durch reinzoomen breiter als 160 Pixel wird, blende Details ein */
    @container (min-width: 160px) {{
        .trip-details {{ display: block; }}
    }}

    .vis-time-axis .vis-text {{ font-weight: bold; color: #495057; }}
    .vis-labelset .vis-label {{ border-bottom: 1px solid #e9ecef; }}

    /* Horizontaler Scrollbalken */
    .timeline-scroll-container {{
        width: 100%;
        height: 18px;
        overflow-x: scroll;
        overflow-y: hidden;
        margin-top: 8px;
        background: #e9ecef;
        border-radius: 4px;
        border: 1px solid #ced4da;
    }}
    .timeline-scroll-content {{
        height: 1px;
    }}
  </style>
</head>
<body>
<h2>⏱️ Träwelling Timeline ({names_joined})</h2>
<div class="info-bar">
  <b>Bedienung:</b> Mit dem Mausrad zoomen (oder Strg + Scrollen), klicken &amp; ziehen zum Verschieben oder den <b>horizontalen Scrollbalken</b> unten nutzen.<br>
  <b>Ansicht:</b> Beim Hineinzoomen erscheinen Details direkt im Block. Fährt man mit der Maus darüber, zeigt der Tooltip alle Zeiten &amp; Verspätungen.<br>
  <b>Legende:</b> Gestrichelter Rahmen = Geplante Zeit (nach oben versetzt). Ausgefüllter Block = Tatsächliche Fahrtzeit. Bunt markiert = Geteilte Fahrten (gleiche Trip-ID). Grau gepunktet = Umstiege.
</div>
<div id="visualization"></div>
<div class="timeline-scroll-container" id="timelineScroll">
    <div class="timeline-scroll-content" id="timelineScrollContent"></div>
</div>

<script type="text/javascript">
  var groups = new vis.DataSet({groups_json});
  var items = new vis.DataSet({items_json});

  var container = document.getElementById('visualization');
  var options = {{
    start: '{w_start_iso}',
    end: '{w_end_iso}',
    horizontalScroll: true,
    zoomKey: 'ctrlKey',
    groupOrder: 'content',
    orientation: 'top',
    stack: false,
    zoomMin: 1000 * 60 * 15,
    zoomMax: 1000 * 60 * 60 * 24 * 365,
    margin: {{ item: 10, axis: 5 }}
  }};

  var timeline = new vis.Timeline(container, items, groups, options);

  // Synchronisation des horizontalen Scrollbalkens
  var totalStart = new Date('{min_iso}').getTime();
  var totalEnd = new Date('{max_iso}').getTime();
  if (totalEnd <= totalStart) totalEnd = totalStart + 86400000;
  var totalRange = totalEnd - totalStart;

  var scrollWrapper = document.getElementById('timelineScroll');
  var scrollContent = document.getElementById('timelineScrollContent');
  var isScrollingFromBar = false;

  function updateScrollBar() {{
    if (isScrollingFromBar) return;
    var win = timeline.getWindow();
    var curDuration = win.end.getTime() - win.start.getTime();
    var ratio = Math.max(1, totalRange / curDuration);
    var contentWidth = Math.max(scrollWrapper.clientWidth, scrollWrapper.clientWidth * ratio);
    scrollContent.style.width = contentWidth + 'px';

    var scrollOffset = (win.start.getTime() - totalStart) / (totalRange - curDuration || 1);
    scrollOffset = Math.max(0, Math.min(1, scrollOffset));
    scrollWrapper.scrollLeft = scrollOffset * (contentWidth - scrollWrapper.clientWidth);
  }}

  scrollWrapper.addEventListener('scroll', function() {{
    isScrollingFromBar = true;
    var win = timeline.getWindow();
    var curDuration = win.end.getTime() - win.start.getTime();
    var maxScroll = scrollContent.scrollWidth - scrollWrapper.clientWidth;
    var scrollOffset = maxScroll > 0 ? (scrollWrapper.scrollLeft / maxScroll) : 0;
    var newStart = totalStart + scrollOffset * (totalRange - curDuration);
    var newEnd = newStart + curDuration;
    timeline.setWindow(newStart, newEnd, {{ animation: false }});
    setTimeout(function() {{ isScrollingFromBar = false; }}, 50);
  }});

  timeline.on('rangechange', updateScrollBar);
  setTimeout(updateScrollBar, 300);
</script>
</body>
</html>"""

    with open(file_path, "w", encoding="utf-8") as f:
        f.write(html_template)

    print(f"[ERFOLG] Interaktive Timeline exportiert: {file_path.name}")
    return file_path.name


# ==========================================
# SHARED / GEMEINSAME AUSWERTUNGEN
# ==========================================

def create_shared_stats_dataframe(conn: sqlite3.Connection, usernames: list[str], start_date: str = '',
                                  end_date: str = '') -> pd.DataFrame:
    if not usernames:
        return pd.DataFrame()
    cursor = conn.cursor()
    stats_list = []

    placeholders = ','.join(['?'] * len(usernames))
    date_query = ""
    params_date = []

    if start_date:
        date_query += " AND arrivalFinal > ?"
        params_date.append(start_date)
    if end_date:
        date_query += " AND departureFinal < ?"
        end_date_str = end_date + 'T23:59:59' if len(end_date) == 10 else end_date
        params_date.append(end_date_str)

    base_params = usernames + params_date

    query_dates = f"""
    SELECT username, date(arrivalFinal) 
    FROM journeys 
    WHERE username IN ({placeholders}) {date_query} 
      AND arrivalFinal IS NOT NULL {get_shared_visibility_query()}
    ORDER BY username, date(arrivalFinal)
    """
    cursor.execute(query_dates, base_params)

    user_dates = defaultdict(list)
    for row in cursor.fetchall():
        if row[1]:
            user_dates[row[0]].append(datetime.strptime(row[1], '%Y-%m-%d').date())

    active_days_dict = {}
    for user in usernames:
        dates = sorted(list(set(user_dates[user])))
        if dates:
            first_day, last_day = dates[0], dates[-1]
            gaps = ((d2 - d1).days - 1 for d1, d2 in zip(dates[:-1], dates[1:]))
            total_time_skipped = sum(gap if gap > 30 else 0 for gap in gaps)
            active_days_dict[user] = (last_day - first_day).days + 1 - total_time_skipped
        else:
            active_days_dict[user] = 1

    query_main = f"""
    SELECT 
        username,
        COUNT(statusId) as total_journeys,
        SUM(distance) as total_distance,
        SUM(durationReal) as total_duration,
        SUM(points) as total_points,
        SUM(likes) as total_likes,
        SUM(CASE WHEN realtimeAvailability = 1 THEN departureDelay ELSE 0 END) as total_dep_delay,
        SUM(CASE WHEN realtimeAvailability = 1 THEN arrivalDelay ELSE 0 END) as total_arr_delay,
        COUNT(DISTINCT operatorName || lineName) as unique_lines,
        COUNT(DISTINCT operatorName) as unique_operators,
        SUM(CASE WHEN realtimeAvailability = 1 THEN 1 ELSE 0 END) as realtime_count,
        AVG(interStopsNumber) as avg_inter_stops,
        SUM(distanceStraightLine) as total_straight_dist,
        SUM(distanceToPrevJourney) as total_prev_dist,
        SUM(CASE WHEN arrivalDelay < 0 AND realtimeAvailability = 1 THEN 1 ELSE 0 END) as early,
        SUM(CASE WHEN arrivalDelay = 0 AND realtimeAvailability = 1 THEN 1 ELSE 0 END) as on_time,
        SUM(CASE WHEN arrivalDelay >= 3 AND realtimeAvailability = 1 THEN 1 ELSE 0 END) as delayed_3,
        SUM(CASE WHEN arrivalDelay >= 6 AND realtimeAvailability = 1 THEN 1 ELSE 0 END) as delayed_6,
        SUM(CASE WHEN arrivalDelay >= 15 AND realtimeAvailability = 1 THEN 1 ELSE 0 END) as delayed_15,
        SUM(CASE WHEN arrivalDelay >= 60 AND realtimeAvailability = 1 THEN 1 ELSE 0 END) as delayed_60
    FROM journeys 
    WHERE username IN ({placeholders}) {date_query} {get_shared_visibility_query()}
    GROUP BY username 
    """
    cursor.execute(query_main, base_params)
    main_stats = {row[0]: row[1:] for row in cursor.fetchall()}

    station_params = base_params + base_params
    date_query_j = date_query.replace('arrivalFinal', 'j.arrivalFinal').replace('departureFinal', 'j.departureFinal')
    query_stations = f"""
    SELECT username, COUNT(DISTINCT stopName) FROM (
        SELECT j.username, s.stopName FROM journeys j JOIN stops s ON j.originId = s.id 
        WHERE j.username IN ({placeholders}) {date_query_j} {get_shared_visibility_query('j')}
        UNION
        SELECT j.username, s.stopName FROM journeys j JOIN stops s ON j.destinationId = s.id 
        WHERE j.username IN ({placeholders}) {date_query_j} {get_shared_visibility_query('j')}
    )
    GROUP BY username
    """
    cursor.execute(query_stations, station_params)
    station_stats = {row[0]: row[1] for row in cursor.fetchall()}

    for user in usernames:
        active_days = active_days_dict.get(user, 1)
        u_stations = station_stats.get(user, 0)
        res = main_stats.get(user, (0,) * 19)

        t_journeys, t_dist, t_dur, t_pts, t_likes = res[0] or 0, res[1] or 0, res[2] or 0, res[3] or 0, res[4] or 0
        rt_dep_del, rt_arr_del, u_lines, u_ops = res[5] or 0, res[6] or 0, res[7] or 0, res[8] or 0
        realtime_count, avg_inter_stops = res[9] or 0, res[10] or 0.0
        t_straight_dist, t_prev_dist = res[11] or 0, res[12] or 0
        early, on_time = res[13] or 0, res[14] or 0
        del_3, del_6, del_15, del_60 = res[15] or 0, res[16] or 0, res[17] or 0, res[18] or 0

        j_div = t_journeys if t_journeys > 0 else 1
        d_div = active_days if active_days > 0 else 1
        rt_div = realtime_count if realtime_count > 0 else 1

        stats_list.append({
            'User': user,
            'Fahrten': t_journeys,
            'Tage': active_days,
            'Fahrten pro Tag': round(t_journeys / d_div, 3),
            'Linien': u_lines,
            'Betreiber': u_ops,
            'Haltestellen': u_stations,
            'Zeit (min)': round(t_dur, 3),
            'Zeit (hh:mm)': f"{int(t_dur // 60):02d}:{int(t_dur % 60):02d}",
            'Distanz (km)': round(t_dist / 1000, 2),
            'Luftlinie Fahrten (km)': round(t_straight_dist / 1000, 2),
            'Luftlinie zwischen Fahrten (km)': round(t_prev_dist / 1000, 2),
            'Tracking-Lücke (%)': round((t_prev_dist / t_dist * 100) if t_dist > 0 else 0, 2),
            'Umweg Faktor': round((t_dist / t_straight_dist) if t_straight_dist > 0 else 0, 2),
            'Abfahrtsverspätung (min)': round((rt_dep_del / rt_div) * j_div, 0),
            'Ankunftsverspätung (min)': round((rt_arr_del / rt_div) * j_div, 0),
            'Punkte': t_pts,
            'Likes': t_likes,
            'Echtzeitquote (%)': round((realtime_count / j_div) * 100, 2),
            'Verfrüht (%)': round((early / rt_div) * 100, 2),
            'Pünktlich (%)': round((on_time / rt_div) * 100, 2),
            '3 Minuten oder mehr verspätet (%)': round((del_3 / rt_div) * 100, 2),
            '6 Minuten oder mehr verspätet (%)': round((del_6 / rt_div) * 100, 2),
            '15 Minuten oder mehr verspätet (%)': round((del_15 / rt_div) * 100, 2),
            '60 Minuten oder mehr verspätet (%)': round((del_60 / rt_div) * 100, 2),
            '∅ Zwischenhalte': round(avg_inter_stops, 3)
        })

    return pd.DataFrame(stats_list).sort_values('Fahrten', ascending=False)


def create_shared_stops_dataframes(conn: sqlite3.Connection, usernames: list[str], start_date: str = '',
                                   end_date: str = '') -> tuple[pd.DataFrame, pd.DataFrame]:
    if not usernames:
        return pd.DataFrame(), pd.DataFrame()
    cursor = conn.cursor()
    placeholders = ','.join(['?'] * len(usernames))
    date_query = ""
    params_date = []

    if start_date:
        date_query += " AND arrivalFinal > ?"
        params_date.append(start_date)
    if end_date:
        date_query += " AND departureFinal < ?"
        end_date_str = end_date + 'T23:59:59' if len(end_date) == 10 else end_date
        params_date.append(end_date_str)

    base_params = usernames + params_date
    user_stop_data = defaultdict(lambda: defaultdict(lambda: {'Ohne Via': 0, 'Nur Via': 0}))

    query_ohne_via = f"""
    WITH filtered_journeys AS (
        SELECT username, originId, destinationId
        FROM journeys 
        WHERE username IN ({placeholders}) {date_query} {get_shared_visibility_query()}
    ),
    all_visits AS (
        SELECT username, originId AS stopId FROM filtered_journeys
        UNION ALL
        SELECT username, destinationId AS stopId FROM filtered_journeys
    )
    SELECT v.username, s.stopName, COUNT(*) AS Besuche
    FROM all_visits v
    JOIN stops s ON v.stopId = s.id
    GROUP BY v.username, s.stopName
    """
    cursor.execute(query_ohne_via, base_params)
    for row in cursor.fetchall():
        uname, stop_name, besuche = row
        user_stop_data[uname][stop_name]['Ohne Via'] = besuche

    query_nur_via = f"""
    SELECT username, interStops 
    FROM journeys 
    WHERE username IN ({placeholders}) {date_query} AND interStops IS NOT NULL {get_shared_visibility_query()}
    """
    cursor.execute(query_nur_via, base_params)
    for row in cursor.fetchall():
        uname, inter_stops_json = row
        try:
            for stop in json.loads(inter_stops_json):
                user_stop_data[uname][stop]['Nur Via'] += 1
        except (json.JSONDecodeError, TypeError):
            continue

    station_user_counts = defaultdict(int)
    for uname, stations in user_stop_data.items():
        for stop_name, stats in stations.items():
            ohne, nur = stats['Ohne Via'], stats['Nur Via']
            insgesamt = ohne + nur
            stats['insgesamt'] = insgesamt
            stats['Quote'] = round((ohne / insgesamt) * 100, 2) if insgesamt > 0 else 0.0
            if ohne > 0:
                station_user_counts[stop_name] += 1

    threshold = min(max(2, len(usernames) - 1 if len(usernames) < 5 else len(usernames) - 2), 5)
    common_stations = sorted([stat for stat, count in station_user_counts.items() if count >= threshold])

    visits_data = []
    visits_data_filterable = []

    for station in common_stations:
        station_stats = []
        for user in usernames:
            stats = user_stop_data[user].get(station, {'Ohne Via': 0, 'Nur Via': 0, 'insgesamt': 0, 'Quote': 0.0})
            station_stats.append({
                'User': user, 'insgesamt': stats['insgesamt'], 'Ohne Via': stats['Ohne Via'],
                'Nur Via': stats['Nur Via'], 'Quote': stats['Quote']
            })

        station_stats.sort(key=lambda x: x['insgesamt'], reverse=True)
        current_station_filter_row = {'Haltestelle': station}
        first_user_added = False

        for st_stat in station_stats:
            user = st_stat['User']
            insgesamt = st_stat['insgesamt']
            current_station_filter_row[user] = insgesamt

            if insgesamt > 0:
                visits_data.append({
                    'Haltestelle': station if not first_user_added else '',
                    'User': user, 'Besuche insgesamt': insgesamt, 'Ohne Via': st_stat['Ohne Via'],
                    'Nur Via': st_stat['Nur Via'], 'Umstiegquote (%)': st_stat['Quote']
                })
                first_user_added = True

        visits_data.append({'Haltestelle': '', 'User': '', 'Besuche insgesamt': '', 'Ohne Via': '', 'Nur Via': '', 'Umstiegquote (%)': ''})
        visits_data_filterable.append(current_station_filter_row)

    df_regular = pd.DataFrame(visits_data)
    df_filterable = pd.DataFrame(visits_data_filterable)
    if not df_filterable.empty:
        df_filterable['Gesamt'] = df_filterable[usernames].sum(axis=1)
        df_filterable = df_filterable.sort_values('Gesamt', ascending=False).drop(columns=['Gesamt'])

    return df_regular, df_filterable


def create_shared_journeys_dataframe(conn: sqlite3.Connection, usernames: list[str], start_date: str = '',
                                     end_date: str = '') -> pd.DataFrame:
    placeholders = ','.join(['?'] * len(usernames))
    params_date = []
    date_query = ""

    if start_date:
        date_query += " AND arrivalFinal > ?"
        params_date.append(start_date)
    if end_date:
        date_query += " AND departureFinal < ?"
        end_date_str = end_date + 'T23:59:59' if len(end_date) == 10 else end_date
        params_date.append(end_date_str)

    params = usernames + params_date + usernames + params_date
    date_query_j = date_query.replace('arrivalFinal', 'j.arrivalFinal').replace('departureFinal', 'j.departureFinal')

    query = f"""
    SELECT 
        j.category AS 'Kategorie',
        j.lineName AS 'Linie',
        j.username AS 'User',
        o.stopName AS 'Starthaltestelle',
        j.departurePlanned AS 'Abfahrtszeit (plan)',
        j.departureFinal AS 'Abfahrtszeit (ist)',
        d.stopName AS 'Zielhaltestelle',
        j.arrivalPlanned AS 'Ankunftszeit (plan)',
        j.arrivalFinal AS 'Ankunftszeit (ist)',
        j.distance / 1000.0 AS 'Entfernung (km)',
        j.points AS 'Punkte',
        '=HYPERLINK("https://traewelling.de/status/'||j.statusId||'","https://traewelling.de/status/'||j.statusId||'")' AS 'Link',
        j.tripId, o.timezone AS 'origin_tz', d.timezone AS 'destination_tz'
    FROM journeys j
    LEFT JOIN stops o ON j.originId = o.id
    LEFT JOIN stops d ON j.destinationId = d.id
    WHERE j.username IN ({placeholders}) {date_query_j} {get_shared_visibility_query('j')}
      AND j.tripId IS NOT NULL 
      AND j.tripId IN (
          SELECT tripId 
          FROM journeys 
          WHERE username IN ({placeholders}) {date_query} {get_shared_visibility_query()}
          GROUP BY tripId 
          HAVING COUNT(DISTINCT username) > 1
      )
    ORDER BY j.tripId, j.departurePlanned
    """
    df = pd.read_sql_query(query, conn, params=params)
    if df.empty:
        return pd.DataFrame()

    time_columns = [
        ('Abfahrtszeit (plan)', 'origin_tz'), ('Abfahrtszeit (ist)', 'origin_tz'),
        ('Ankunftszeit (plan)', 'destination_tz'), ('Ankunftszeit (ist)', 'destination_tz')
    ]

    for time_col, tz_col in time_columns:
        df[time_col] = [format_utc_to_datatime(dt, tz) for dt, tz in zip(df[time_col], df[tz_col])]
        df[time_col] = df[time_col].apply(lambda x: x.strftime('%d.%m.%Y %H:%M') if pd.notna(x) else '')

    df = df.drop(columns=['origin_tz', 'destination_tz'])
    formatted_rows = []
    cols = [c for c in df.columns if c != 'tripId']

    for trip_id, group in df.groupby('tripId', sort=False):
        for _, row in group.iterrows():
            formatted_rows.append({c: row[c] for c in cols})
        formatted_rows.append({c: '' for c in cols})

    return pd.DataFrame(formatted_rows)


def create_shared_grouped_dataframe(conn: sqlite3.Connection, usernames: list[str], column_name: str,
                                    start_date: str = '', end_date: str = '') -> tuple[pd.DataFrame, str]:
    beautiful_names = {
        'category': 'Kategorie', 'operatorName': 'Betreiber', 'lineName': 'Linie', 'business': 'Fahrtgrund',
        'vehicleNumber': 'Fahrzeugnummer', 'locomotiveClass': 'Baureihe', 'clientName': 'Check-In-Tool'
    }
    beautiful_name = beautiful_names.get(column_name, column_name)
    placeholders = ','.join(['?'] * len(usernames))

    date_query = ""
    params_date = []
    if start_date:
        date_query += " AND arrivalFinal > ?"
        params_date.append(start_date)
    if end_date:
        date_query += " AND departureFinal < ?"
        if len(end_date) == 10:
            end_date += 'T23:59:59'
        params_date.append(end_date)

    params = usernames + params_date + usernames + params_date
    exclude_query = ""
    if column_name == 'operatorName':
        exclude_query = "AND operatorName != 'Unknown Operator'"
    elif column_name == 'clientName':
        exclude_query = "AND clientName != 'Unknown Client'"

    query = f"""
    SELECT 
        {column_name} AS '{beautiful_name}',
        username AS User,
        COUNT(statusId) AS Fahrten,
        SUM(distance) / 1000.0 AS 'Distanz (km)',
        ROUND(SUM(distance) * 1.0 / NULLIF(SUM(distanceStraightLine), 0), 2) AS 'Umweg Faktor',
        SUM(durationReal) AS 'Zeit (min)',
        ROUND(AVG(arrivalDelay), 2) AS '∅ Verspätung (min)',
        ROUND(SUM(distance) / (SUM(durationReal) / 60.0 * 1000.0), 2) AS 'Geschwindigkeit (km/h)'
    FROM journeys
    WHERE username IN ({placeholders}) {date_query} {exclude_query} {get_shared_visibility_query()}
      AND {column_name} IS NOT NULL AND {column_name} != '' AND {column_name} != 'N/A'
      AND {column_name} IN (
          SELECT {column_name} 
          FROM journeys 
          WHERE username IN ({placeholders}) {date_query} {get_shared_visibility_query()}
          GROUP BY {column_name} 
          HAVING COUNT(DISTINCT username) > 1
      )
    GROUP BY {column_name}, username
    ORDER BY {column_name} ASC, SUM(distance) DESC
    """
    df = pd.read_sql_query(query, conn, params=params)
    if df.empty:
        return pd.DataFrame(), beautiful_name

    df['Zeit (hh:mm)'] = df['Zeit (min)'].apply(lambda x: f"{int(x // 60):02d}:{int(x % 60):02d}" if pd.notna(x) else "")
    cols = [beautiful_name, 'User', 'Fahrten', 'Distanz (km)', 'Umweg Faktor', 'Zeit (hh:mm)', 'Zeit (min)', 'Geschwindigkeit (km/h)', '∅ Verspätung (min)']
    df = df[cols]

    records = df.to_dict('records')
    formatted_rows = []
    last_group_val = None

    for record in records:
        current_val = record[beautiful_name]
        if last_group_val is not None and current_val != last_group_val:
            formatted_rows.append({col: '' for col in cols})
        if current_val == last_group_val:
            record[beautiful_name] = ''
        formatted_rows.append(record)
        last_group_val = current_val

    if records:
        formatted_rows.append({col: '' for col in cols})
    return pd.DataFrame(formatted_rows), beautiful_name


def create_shared_excel_file(conn: sqlite3.Connection, usernames: list[str] = None, start_date: str = '',
                             end_date: str = '', columns_to_compare: list[str] = None) -> str | None:
    if usernames is None:
        usernames = pd.read_sql('SELECT username FROM users', conn)['username'].tolist()
    if not usernames:
        return None
    if columns_to_compare is None:
        columns_to_compare = ['category', 'operatorName', 'lineName', 'vehicleNumber', 'locomotiveClass', 'clientName']

    output_dir = Path('finished_exports/shared')
    output_dir.mkdir(parents=True, exist_ok=True)
    names_joined = "_".join(sorted(usernames)[:6])
    if len(usernames) > 6:
        names_joined += f"_and_{len(usernames) - 6}_more"
    
    date_suffix = get_date_suffix(start_date, end_date)
    file_path = output_dir / f"{names_joined}{date_suffix}_shared_data.xlsx"

    with pd.ExcelWriter(file_path, engine='xlsxwriter', datetime_format='dd.mm.yyyy hh:mm') as writer:
        def format_sheet(df, sheet_name, freeze_spalten=1):
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes(1, freeze_spalten)
            if not df.empty:
                worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
            for runner, col in enumerate(df.columns):
                max_length = df[col].dropna().map(lambda x: len(str(x))).max()
                final_width = max((max_length if pd.notna(max_length) else 0), len(str(col)) + 2) + 2
                worksheet.set_column(runner, runner, final_width)

        format_sheet(create_shared_meta_dataframe(conn, usernames, start_date, end_date), 'Meta', freeze_spalten=0)
        shared_stats_df = create_shared_stats_dataframe(conn, usernames, start_date, end_date)
        if not shared_stats_df.empty:
            format_sheet(shared_stats_df, 'Statistik', 1)

        df_regular_stops, df_filterable_stops = create_shared_stops_dataframes(conn, usernames, start_date, end_date)
        if not df_regular_stops.empty:
            format_sheet(df_regular_stops, 'Haltestellen', 2)
        if not df_filterable_stops.empty:
            format_sheet(df_filterable_stops, 'Haltestellen (filterbar)', 1)

        shared_journeys_df = create_shared_journeys_dataframe(conn, usernames, start_date, end_date)
        if not shared_journeys_df.empty:
            format_sheet(shared_journeys_df, 'Gemeinsame Fahrten', 3)

        for col_name in columns_to_compare:
            df, sheet_name = create_shared_grouped_dataframe(conn, usernames, col_name, start_date, end_date)
            if not df.empty:
                format_sheet(df, sheet_name, 2)

    print(f'[ERFOLG] Shared-Excel gespeichert: {file_path.name}')
    return file_path.name


def check_for_user_journeys(conn: sqlite3.Connection, username: str, start_date: str = '', end_date: str = '') -> bool:
    cursor = conn.cursor()
    params, date_query = params_and_date_query(username, start_date, end_date)
    cursor.execute(f"SELECT 1 FROM journeys {date_query} LIMIT 1", params)
    return cursor.fetchone() is not None


# ==========================================
# MAIN LOOP / AUSFÜHRUNG
# ==========================================

def fetch_and_analyse(usernames: list | str | None = None, include_stored_users: bool = True, reload_all: bool = False,
                      start_date: str = '', end_date: str = '', create_picture: bool = True) -> None:
    if usernames is None:
        usernames = []
    elif isinstance(usernames, str):
        usernames = [usernames]
    
    conn = init_db()
    api_token = log_onto_traewelling(conn)
    
    if include_stored_users:
        print("[INFO] Hole alle bereits gespeicherten User aus der DB...")
        usernames.extend(pd.read_sql('SELECT username FROM users', conn)['username'])

    for username in sorted(list(set(usernames))):
        fetch_data_to_db(username, conn=conn, api_token=api_token, reload_all=reload_all)
        if check_for_user_journeys(conn, username, start_date=start_date, end_date=end_date):
            create_user_excel_file(conn, username, start_date=start_date, end_date=end_date, create_picture=create_picture)
        else:
            print(f'[INFO] Überspringe Export: Keine Fahrten für {username} im Zeitraum vorhanden.')


if __name__ == '__main__':
    start = time()
    conn = init_db()
    api_token = log_onto_traewelling(conn)
    
    print("\n" + "="*50)
    print("🚂 STARTE TRÄWELLING ANALYSE (v9)")
    print("="*50)
    
    fetch_and_analyse(include_stored_users=True, start_date='2026-08-01', end_date='2026-08-31', usernames='zug_pegnitzpfeil_freak')

    print("\n" + "="*50)
    print(f'🏁 FERTIG! Requests übrig: {remaining} | Gesamtdauer: {time() - start:.2f} Sekunden')
    print("="*50)