import json
import os
import sys
import traceback
import hashlib
from dotenv import load_dotenv

load_dotenv()
import sqlite3
import math
from pathlib import Path
try:
    from google import genai
    from google.genai import types
except ImportError:
    genai = None
    types = None
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
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import threading
from contextlib import contextmanager

# ==========================================
# GLOBALE VARIABLEN & THREAD-LOKALER KONTEXT
# ==========================================
tf = None
_USER_CTX = threading.local()
_INTERACTIVE_FILTER_CTX = threading.local()


def get_auth_user() -> str:
    return getattr(_USER_CTX, 'auth_user', "") or ""


def set_auth_user(val: str) -> None:
    _USER_CTX.auth_user = str(val or "").strip()


def get_followed_users() -> set:
    val = getattr(_USER_CTX, 'followed_users', None)
    return set(val) if val is not None else set()


def set_followed_users(val) -> None:
    _USER_CTX.followed_users = set(val) if val else set()


def get_followers_users() -> set:
    val = getattr(_USER_CTX, 'followers_users', None)
    return set(val) if val is not None else set()


def set_followers_users(val) -> None:
    _USER_CTX.followers_users = set(val) if val else set()


def get_trusted_users() -> set:
    val = getattr(_USER_CTX, 'trusted_users', None)
    return set(val) if val is not None else set()


def set_trusted_users(val) -> None:
    _USER_CTX.trusted_users = set(val) if val else set()


def get_remaining() -> str:
    val = getattr(_USER_CTX, 'remaining', None)
    return "Unbekannt" if val is None else str(val)


def set_remaining(val) -> None:
    _USER_CTX.remaining = "Unbekannt" if val is None else str(val)


def clear_user_context() -> None:
    _USER_CTX.auth_user = ""
    _USER_CTX.followed_users = set()
    _USER_CTX.followers_users = set()
    _USER_CTX.trusted_users = set()
    _USER_CTX.remaining = "Unbekannt"


@contextmanager
def user_auth_context(auth_user: str = "", followed=None, followers=None, trusted=None, remaining=None):
    prev_auth = getattr(_USER_CTX, 'auth_user', "")
    prev_followed = getattr(_USER_CTX, 'followed_users', None)
    prev_followers = getattr(_USER_CTX, 'followers_users', None)
    prev_trusted = getattr(_USER_CTX, 'trusted_users', None)
    prev_remaining = getattr(_USER_CTX, 'remaining', "Unbekannt")
    try:
        set_auth_user(auth_user)
        set_followed_users(followed)
        set_followers_users(followers)
        set_trusted_users(trusted)
        if remaining is not None:
            set_remaining(remaining)
        yield
    finally:
        _USER_CTX.auth_user = prev_auth
        _USER_CTX.followed_users = prev_followed
        _USER_CTX.followers_users = prev_followers
        _USER_CTX.trusted_users = prev_trusted
        _USER_CTX.remaining = prev_remaining


@contextmanager
def user_auth_context_from_db(conn: sqlite3.Connection, creator: str):
    followed = set()
    trusted = set()
    followers = set()
    if creator:
        cur = conn.cursor()
        cur.execute("SELECT followed_users, trusted_users, followers FROM user_login_data WHERE owner = ?", (creator,))
        row = cur.fetchone()
        if row:
            try:
                followed = set(json.loads(row[0])) if row[0] else set()
            except Exception:
                followed = set()
            try:
                trusted = set(json.loads(row[1])) if row[1] else set()
            except Exception:
                trusted = set()
            try:
                followers = set(json.loads(row[2])) if row[2] else set()
            except Exception:
                followers = set()
    with user_auth_context(auth_user=creator, followed=followed, followers=followers, trusted=trusted):
        yield


def get_export_dir(folder_type: str, creator_username: str | None = None) -> Path:
    safe_folder = os.path.basename(folder_type)
    owner = (creator_username or get_auth_user() or "").strip()
    if owner:
        safe_owner = os.path.basename(owner)
        path = Path('finished_exports') / safe_owner / safe_folder
    else:
        path = Path('finished_exports') / safe_folder
    path.mkdir(parents=True, exist_ok=True)
    return path


def __getattr__(name: str):
    if name == 'AUTH_USER':
        return get_auth_user()
    if name == 'FOLLOWED_USERS':
        return get_followed_users()
    if name == 'FOLLOWERS_USERS':
        return get_followers_users()
    if name == 'TRUSTED_USERS':
        return get_trusted_users()
    if name == 'remaining':
        return get_remaining()
    raise AttributeError(f"module '{__name__}' has no attribute '{name}'")


class _TraewellingAnalysisModule(sys.modules[__name__].__class__):
    def __getattr__(self, name):
        if name == 'AUTH_USER':
            return get_auth_user()
        if name == 'FOLLOWED_USERS':
            return get_followed_users()
        if name == 'FOLLOWERS_USERS':
            return get_followers_users()
        if name == 'TRUSTED_USERS':
            return get_trusted_users()
        if name == 'remaining':
            return get_remaining()
        raise AttributeError(f"module '{__name__}' has no attribute '{name}'")

    def __setattr__(self, name, value):
        if name == 'AUTH_USER':
            set_auth_user(value)
        elif name == 'FOLLOWED_USERS':
            set_followed_users(value)
        elif name == 'FOLLOWERS_USERS':
            set_followers_users(value)
        elif name == 'TRUSTED_USERS':
            set_trusted_users(value)
        elif name == 'remaining':
            set_remaining(value)
        else:
            super().__setattr__(name, value)


sys.modules[__name__].__class__ = _TraewellingAnalysisModule



TOKEN_REFRESH_LOCK = threading.Lock()
OAUTH_TOKEN_URL = 'https://traewelling.de/oauth/token'

class StopTaskException(Exception):
    """Spezifische Exception um einen laufenden Thread sicher abzubrechen, ohne die DB zu zerstören."""
    pass


def refresh_oauth_token(conn: sqlite3.Connection | None, username: str, force: bool = False) -> str | None:
    if not username:
        return None
    with TOKEN_REFRESH_LOCK:
        needs_close = False
        db_conn = conn
        if db_conn is None:
            db_conn = init_db()
            needs_close = True
        try:
            cur = db_conn.cursor()
            cur.execute("SELECT access_token, refresh_token, expires_at FROM user_login_data WHERE owner = ?", (username,))
            row = cur.fetchone()
            if not row:
                return None
            cur_access, cur_refresh, cur_expires = row[0], row[1], row[2]

            # Double-Check innerhalb des Locks:
            # Falls nicht force und expires_at vorhanden und expires_at - time() > 300,
            # wurde der Token bereits durch einen anderen Thread erneuert!
            if not force and cur_expires is not None:
                try:
                    if (float(cur_expires) - time()) > 300:
                        return cur_access
                except (ValueError, TypeError):
                    pass

            if not cur_refresh or not str(cur_refresh).strip():
                return None

            load_dotenv()
            client_id = os.getenv('CLIENT_ID')
            client_secret = os.getenv('CLIENT_SECRET')
            if not client_id or not client_secret:
                log_error_to_db(
                    username=username,
                    error_type="OAUTH_CONFIG_MISSING",
                    error_message="CLIENT_ID oder CLIENT_SECRET fehlt in der Umgebung (.env)",
                    context="refresh_oauth_token",
                    conn=db_conn
                )
                return None

            data = {
                'grant_type': 'refresh_token',
                'refresh_token': cur_refresh,
                'client_id': client_id,
                'client_secret': client_secret
            }
            headers = {
                'Accept': 'application/json',
                'User-Agent': 'TraewellingAnalysis_v10 tee_beutel'
            }

            try:
                resp = requests.post(OAUTH_TOKEN_URL, data=data, headers=headers, timeout=20)
            except Exception as e:
                log_error_to_db(
                    username=username,
                    error=e,
                    error_type="OAUTH_REFRESH_NETWORK_ERROR",
                    error_message=f"Netzwerkfehler beim Token-Refresh: {e}",
                    context="refresh_oauth_token",
                    conn=db_conn
                )
                return None

            if resp.status_code == 200:
                token_json = resp.json()
                new_access_token = token_json.get('access_token')
                new_refresh_token = token_json.get('refresh_token') or cur_refresh
                expires_in = token_json.get('expires_in')
                new_expires_at = (time() + float(expires_in)) if expires_in else None

                cur.execute("""
                    UPDATE user_login_data 
                    SET access_token = ?, refresh_token = ?, expires_at = ? 
                    WHERE owner = ?
                """, (new_access_token, new_refresh_token, new_expires_at, username))
                db_conn.commit()
                return new_access_token
            else:
                log_error_to_db(
                    username=username,
                    error_type=f"OAUTH_REFRESH_{resp.status_code}",
                    error_message=f"Token-Refresh fehlgeschlagen: {resp.text}",
                    context="refresh_oauth_token",
                    conn=db_conn
                )
                if resp.status_code in (400, 401):
                    cur.execute("UPDATE user_login_data SET refresh_token = NULL WHERE owner = ?", (username,))
                    db_conn.commit()
                return None
        finally:
            if needs_close:
                db_conn.close()


def get_valid_access_token(conn: sqlite3.Connection | None, username: str, min_validity_sec: int = 600) -> str | None:
    if not username:
        return None
    needs_close = False
    db_conn = conn
    if db_conn is None:
        db_conn = init_db()
        needs_close = True
    try:
        cur = db_conn.cursor()
        cur.execute("SELECT access_token, refresh_token, expires_at FROM user_login_data WHERE owner = ?", (username,))
        row = cur.fetchone()
        if not row:
            return None
        access_token, refresh_token, expires_at = row[0], row[1], row[2]

        if refresh_token and expires_at is not None:
            try:
                exp_ts = float(expires_at)
                if time() + min_validity_sec >= exp_ts:
                    refreshed = refresh_oauth_token(db_conn, username)
                    if refreshed:
                        return refreshed
                    if time() < exp_ts:
                        return access_token
                    return None
            except (ValueError, TypeError):
                pass
        return access_token or None
    finally:
        if needs_close:
            db_conn.close()


def traewelling_api_get(url: str, api_token: str, conn: sqlite3.Connection | None = None,
                        username: str | None = None, timeout: int = 20) -> tuple[requests.Response, str]:
    headers = {
        'accept': 'application/json',
        'Authorization': f'Bearer {api_token}',
        'User-Agent': 'TraewellingAnalysis_v10 tee_beutel'
    }
    resp = requests.get(url, headers=headers, timeout=timeout)
    if resp.status_code == 401:
        target_owner = username or get_auth_user()
        if target_owner:
            new_token = refresh_oauth_token(conn, target_owner, force=True)
            if new_token:
                headers['Authorization'] = f'Bearer {new_token}'
                retry_resp = requests.get(url, headers=headers, timeout=timeout)
                return retry_resp, new_token
    return resp, api_token


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
    resp, api_token = traewelling_api_get(f'https://traewelling.de/api/v1/status/{status_id}', api_token, conn=conn)
    if resp.status_code != 200:
        log_error_to_db(error_type=f"HTTP_{resp.status_code}", error_message=f"Status {status_id} lieferte HTTP {resp.status_code}", context=f"reload_single_status (status_id: {status_id})")
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
        inter_resp, api_token = traewelling_api_get(f"https://traewelling.de/api/v1/stopovers/{trip_id}", api_token, conn=conn)
        if inter_resp.status_code == 200:
            inter_data = inter_resp.json().get('data', {})
            if isinstance(inter_data, dict):
                inter_stations = inter_data
            elif isinstance(inter_data, list) and inter_data:
                inter_stations = {str(trip_id): inter_data}

    # 4. Polyline löschen (wichtig: nur gelöscht, nicht neu geladen)
    cursor.execute("DELETE FROM polylines WHERE statusId = ?", (status_id,))
    conn.commit()

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
            resp, api_token = traewelling_api_get(url, api_token)
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
                log_error_to_db(username=None, error=e, context=f"fetch_paginated_users ({start_url})")
                break
    return users


def log_onto_traewelling(conn: sqlite3.Connection) -> str:
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
        login_response, api_token = traewelling_api_get('https://traewelling.de/api/v1/auth/user',
                                                        api_token, conn=conn)
        try:
            login_response.raise_for_status()
            login_json = login_response.json()

            set_remaining(login_response.headers.get('x-ratelimit-remaining'))
            username = login_json.get('data', {}).get('username', None)

            if username:
                set_auth_user(username)
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

    auth_u = get_auth_user()
    cursor = conn.cursor()
    cursor.execute("SELECT followed_users, trusted_users, followers, last_updated FROM user_login_data WHERE owner = ?",
                   (auth_u,))
    row = cursor.fetchone()

    needs_update = True
    if row:
        try:
            last_updated_dt = datetime.fromisoformat(row[3])
            if datetime.now() - last_updated_dt < timedelta(hours=24):
                set_followed_users(set(json.loads(row[0])) if row[0] else set())
                set_trusted_users(set(json.loads(row[1])) if row[1] else set())
                set_followers_users(set(json.loads(row[2])) if row[2] else set())
                needs_update = False
                print("[INFO] Sichtbarkeits-Rechte aus lokalem Cache geladen.")
        except Exception:
            pass

    if needs_update:
        print("[INFO] Aktualisiere Sichtbarkeits-Rechte (Follows/Followers/Trusted) von der API...")

        # Die korrekten Endpunkte gemäß API-Dokumentation
        trusted_set = fetch_paginated_users(api_token, 'https://traewelling.de/api/v1/user/self/trusted-by')
        followed_set = fetch_paginated_users(api_token, 'https://traewelling.de/api/v1/user/self/followings')
        followers_set = fetch_paginated_users(api_token, 'https://traewelling.de/api/v1/user/self/followers')
        set_trusted_users(trusted_set)
        set_followed_users(followed_set)
        set_followers_users(followers_set)

        now_iso = datetime.now().isoformat()
        cursor.execute(
            """INSERT INTO user_login_data (owner, followed_users, trusted_users, followers, last_updated, last_online)
               VALUES (?, ?, ?, ?, ?, ?)
               ON CONFLICT(owner) DO UPDATE SET
                   followed_users = excluded.followed_users,
                   trusted_users = excluded.trusted_users,
                   followers = excluded.followers,
                   last_updated = excluded.last_updated,
                   last_online = excluded.last_online""",
            (auth_u, json.dumps(list(get_followed_users())), json.dumps(list(get_trusted_users())),
             json.dumps(list(get_followers_users())), now_iso, now_iso))
        conn.commit()
        print("[INFO] ✅ Rechte erfolgreich aktualisiert.")

    return api_token


# ==========================================
# DATENBANK FUNKTIONEN (SQLITE)
# ==========================================

def init_db() -> sqlite3.Connection:
    path = Path.cwd() / 'traewelling.db'
    conn = sqlite3.connect(path, timeout=15.0)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA busy_timeout = 10000;")
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
    highestPrivilege INTEGER,
    privilegeCheckpoints TEXT)''')

    cursor.execute('''CREATE TABLE IF NOT EXISTS user_login_data (
    owner TEXT PRIMARY KEY,
    followed_users TEXT,
    trusted_users TEXT,
    followers TEXT,
    last_updated TEXT,
    access_token TEXT,
    refresh_token TEXT,
    expires_at REAL,
    last_online TEXT)''')

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

    try:
        cursor.execute("ALTER TABLE user_login_data ADD COLUMN excel_settings TEXT")
    except sqlite3.OperationalError:
        pass

    try:
        cursor.execute("ALTER TABLE user_login_data ADD COLUMN gemini_api_key TEXT")
    except sqlite3.OperationalError:
        pass

    try:
        cursor.execute("ALTER TABLE user_login_data ADD COLUMN ai_request_count INTEGER DEFAULT 0")
    except sqlite3.OperationalError:
        pass

    try:
        cursor.execute("ALTER TABLE user_login_data ADD COLUMN ai_window_start TEXT")
    except sqlite3.OperationalError:
        pass

    try:
        cursor.execute("ALTER TABLE user_login_data ADD COLUMN last_online TEXT")
    except sqlite3.OperationalError:
        pass

    try:
        cursor.execute("ALTER TABLE users ADD COLUMN privilegeCheckpoints TEXT")
        cursor.execute("SELECT username, highestPrivilege FROM users")
        for u_name, h_priv in cursor.fetchall():
            if u_name and h_priv is not None:
                cursor.execute("SELECT MAX(statusId) FROM journeys WHERE username = ?", (u_name,))
                max_sid_row = cursor.fetchone()
                max_sid = max_sid_row[0] if max_sid_row and max_sid_row[0] else 0
                checkpoints = {str(lvl): max_sid for lvl in range(int(h_priv) + 1)}
                cursor.execute(
                    "UPDATE users SET privilegeCheckpoints = ? WHERE username = ?",
                    (json.dumps(checkpoints), u_name)
                )
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

    cursor.execute('''CREATE TABLE IF NOT EXISTS ai_chats (
        id TEXT PRIMARY KEY,
        owner TEXT,
        title TEXT,
        selected_users TEXT,
        start_date TEXT,
        end_date TEXT,
        created_at TEXT,
        updated_at TEXT
    )''')

    cursor.execute('''CREATE TABLE IF NOT EXISTS ai_messages (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        chat_id TEXT,
        role TEXT,
        content TEXT,
        sql_query TEXT,
        row_count INTEGER,
        sql_result TEXT,
        created_at TEXT,
        FOREIGN KEY(chat_id) REFERENCES ai_chats(id) ON DELETE CASCADE
    )''')

    try:
        cursor.execute("ALTER TABLE ai_messages ADD COLUMN sql_result TEXT")
    except sqlite3.OperationalError:
        pass

    cursor.execute('''CREATE INDEX IF NOT EXISTS idx_ai_chats_owner ON ai_chats(owner)''')
    cursor.execute('''CREATE INDEX IF NOT EXISTS idx_ai_messages_chat_id ON ai_messages(chat_id)''')

    cursor.execute('''CREATE TABLE IF NOT EXISTS error_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        timestamp TEXT NOT NULL,
        username TEXT,
        error_type TEXT,
        error_message TEXT,
        traceback TEXT,
        context TEXT
    )''')
    cursor.execute('''CREATE INDEX IF NOT EXISTS idx_error_log_username ON error_log(username)''')
    cursor.execute('''CREATE INDEX IF NOT EXISTS idx_error_log_timestamp ON error_log(timestamp)''')

    cursor.execute('''CREATE TABLE IF NOT EXISTS download_telemetry (
        metric_key TEXT PRIMARY KEY,
        avg_sec_per_unit REAL NOT NULL,
        sample_count INTEGER DEFAULT 0,
        last_updated TEXT
    )''')

    conn.commit()
    return conn


def log_error_to_db(username: str | None = None,
                    error: Exception | str | None = None,
                    error_type: str | None = None,
                    error_message: str | None = None,
                    traceback_str: str | None = None,
                    context: str | None = None,
                    conn: sqlite3.Connection | None = None) -> int | None:
    """Protokolliert einen Fehler in die Tabelle 'error_log' in der SQLite-Datenbank.
    Arbeitet thread-sicher und isoliert, sodass ein Logging-Fehler niemals die aufrufende Anwendung crasht."""
    try:
        now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        # Fallback auf get_auth_user(), falls kein Username übergeben wurde
        if not username:
            username = get_auth_user()

        if isinstance(error, BaseException):
            if not error_type:
                error_type = type(error).__name__
            if not error_message:
                error_message = str(error)
            if not traceback_str:
                traceback_str = traceback.format_exc()
        elif isinstance(error, str):
            if not error_message:
                error_message = error
            if not error_type:
                error_type = "Error"

        if not error_type:
            error_type = "UnknownError"
        if not error_message:
            error_message = "Keine Fehlermeldung angegeben"

        # 1. Option: Falls eine bestehende Verbindung übergeben wurde, nutze diese direkt
        if conn is not None:
            try:
                cur = conn.cursor()
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS error_log (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        timestamp TEXT NOT NULL,
                        username TEXT,
                        error_type TEXT,
                        error_message TEXT,
                        traceback TEXT,
                        context TEXT
                    )
                """)
                cur.execute("""
                    INSERT INTO error_log (timestamp, username, error_type, error_message, traceback, context)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (now_str, username, error_type, error_message, traceback_str, context))
                conn.commit()
                return cur.lastrowid
            except Exception:
                pass  # Fallback auf lokale Verbindung

        # 2. Option: Dedizierte Verbindung mit WAL & busy_timeout
        db_path = Path.cwd() / 'traewelling.db'
        with sqlite3.connect(db_path, timeout=15.0) as local_conn:
            local_conn.execute("PRAGMA journal_mode=WAL;")
            local_conn.execute("PRAGMA busy_timeout = 10000;")
            cur = local_conn.cursor()
            cur.execute("""
                CREATE TABLE IF NOT EXISTS error_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    timestamp TEXT NOT NULL,
                    username TEXT,
                    error_type TEXT,
                    error_message TEXT,
                    traceback TEXT,
                    context TEXT
                )
            """)
            cur.execute("""
                INSERT INTO error_log (timestamp, username, error_type, error_message, traceback, context)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (now_str, username, error_type, error_message, traceback_str, context))
            local_conn.commit()
            return cur.lastrowid
    except Exception as log_err:
        print(f"[ERROR_LOG_FAILURE] Fehler beim Schreiben ins Errorlog: {log_err}", file=sys.stderr)
        return None

log_error = log_error_to_db


def get_error_logs(limit: int = 100, username: str | None = None, conn: sqlite3.Connection | None = None) -> list[dict]:
    """Liest die neuesten Fehlereinträge aus der Tabelle error_log aus."""
    close_after = False
    if conn is None:
        db_path = Path.cwd() / 'traewelling.db'
        conn = sqlite3.connect(db_path, timeout=10.0)
        close_after = True

    try:
        cur = conn.cursor()
        query = "SELECT id, timestamp, username, error_type, error_message, traceback, context FROM error_log"
        params = []
        if username:
            query += " WHERE username = ?"
            params.append(username)
        query += " ORDER BY id DESC"
        if limit and limit > 0:
            query += f" LIMIT {int(limit)}"

        cur.execute(query, params)
        rows = cur.fetchall()
        result = []
        for r in rows:
            result.append({
                'id': r[0],
                'timestamp': r[1],
                'username': r[2],
                'error_type': r[3],
                'error_message': r[4],
                'traceback': r[5],
                'context': r[6]
            })
        return result
    except sqlite3.OperationalError:
        return []
    finally:
        if close_after:
            conn.close()


def clear_error_logs(username: str | None = None, conn: sqlite3.Connection | None = None) -> int:
    """Löscht Einträge aus der Tabelle error_log (optional gefiltert nach User)."""
    close_after = False
    if conn is None:
        db_path = Path.cwd() / 'traewelling.db'
        conn = sqlite3.connect(db_path, timeout=10.0)
        close_after = True

    try:
        cur = conn.cursor()
        if username:
            cur.execute("DELETE FROM error_log WHERE username = ?", (username,))
        else:
            cur.execute("DELETE FROM error_log")
        conn.commit()
        return cur.rowcount
    except sqlite3.OperationalError:
        return 0
    finally:
        if close_after:
            conn.close()


def get_privilege_level(username: str) -> int:
    auth_user = get_auth_user()
    if username == auth_user:
        return 3
    if username in get_trusted_users():
        return 2
    if username in get_followed_users():
        return 1
    return 0


def get_visibility_clause(username: str, table_prefix: str = "") -> str:
    auth_user = get_auth_user()
    if username == auth_user:
        return ""
    prefix = f"{table_prefix}." if table_prefix else ""
    allowed_visibilities = [0, 4]
    if username in get_followed_users():
        allowed_visibilities.append(2)
    if username in get_trusted_users():
        allowed_visibilities.append(5)
    return f" AND {prefix}visibility IN ({','.join(map(str, allowed_visibilities))})"


def get_shared_visibility_query(table_prefix: str = "") -> str:
    prefix = f"{table_prefix}." if table_prefix else ""
    auth_user = get_auth_user()
    trusted_users = get_trusted_users()
    followed_users = get_followed_users()
    auth_cond = f"{prefix}username = '{auth_user}'"
    trusted_list = ",".join(f"'{u}'" for u in trusted_users) if trusted_users else "NULL"
    followed_list = ",".join(f"'{u}'" for u in followed_users) if followed_users else "NULL"
    base_vis = f" AND ({auth_cond} OR ({prefix}username IN ({trusted_list}) AND {prefix}visibility IN (0, 4, 2, 5)) OR ({prefix}username IN ({followed_list}) AND {prefix}visibility IN (0, 4, 2)) OR ({prefix}visibility IN (0, 4)))"
    extra_sql = getattr(_INTERACTIVE_FILTER_CTX, 'sql_clause', '')
    if extra_sql:
        base_vis += extra_sql
    return base_vis


def already_downloaded_in_db(conn: sqlite3.Connection, username: str, max_verified_status_id: int | None = None) -> pd.DataFrame:
    if max_verified_status_id is not None:
        return pd.read_sql(
            'SELECT statusId FROM journeys WHERE username = ? AND statusId <= ?',
            conn,
            params=(username, max_verified_status_id)
        )
    return pd.read_sql('SELECT statusId FROM journeys WHERE username = ?', conn, params=(username,))


# ==========================================
# DOWNLOAD TELEMETRIE & RESTZEITSCHÄTZUNG
# ==========================================

DEFAULT_TELEMETRY = {
    'sec_per_stop': 0.45,
    'sec_per_inter_batch': 0.65,
    'sec_per_ride_calc': 0.008
}


def get_telemetry_metrics(conn: sqlite3.Connection) -> tuple[dict, int]:
    """Liest historische Telemetrie-Werte und die Gesamtzahl der Samples."""
    cursor = conn.cursor()
    cursor.execute("SELECT metric_key, avg_sec_per_unit, sample_count FROM download_telemetry")
    rows = cursor.fetchall()
    metrics = dict(DEFAULT_TELEMETRY)
    total_samples = 0
    for key, avg_val, count in rows:
        if key in metrics and avg_val and avg_val > 0:
            metrics[key] = float(avg_val)
        if count:
            total_samples += int(count)
    return metrics, total_samples


def update_download_telemetry(conn: sqlite3.Connection, metric_key: str, elapsed_active_sec: float, processed_units: int) -> None:
    """Aktualisiert Telemetrie mit EMA (alpha = 0.3)."""
    if processed_units < 1 or elapsed_active_sec <= 0:
        return
    try:
        observed_sec_per_unit = float(elapsed_active_sec) / float(processed_units)
        cursor = conn.cursor()
        cursor.execute("SELECT avg_sec_per_unit, sample_count FROM download_telemetry WHERE metric_key = ?", (metric_key,))
        row = cursor.fetchone()
        now_iso = datetime.now().isoformat()
        if not row or not row[1]:
            new_avg = observed_sec_per_unit
            new_count = 1
        else:
            old_avg = float(row[0])
            old_count = int(row[1])
            new_avg = (0.7 * old_avg) + (0.3 * observed_sec_per_unit)
            new_count = old_count + 1
        cursor.execute(
            "REPLACE INTO download_telemetry (metric_key, avg_sec_per_unit, sample_count, last_updated) VALUES (?, ?, ?, ?)",
            (metric_key, new_avg, new_count, now_iso)
        )
        conn.commit()
    except Exception as e:
        print(f"[WARNUNG] Fehler beim Aktualisieren von download_telemetry ({metric_key}): {e}")


def get_users_sync_overview(conn: sqlite3.Connection, creator: str, is_admin: bool = False) -> list[dict]:
    """
    Ermittelt für den Creator und all seine Freunde den letzten sichtbaren Status
    gemäß eigener Freigabeklasse (get_shared_visibility_query).
    Normale User (is_admin=False) sehen nur sich selbst und ihre Freunde aus user_login_data.
    Admins (is_admin=True) sehen zusätzlich alle Accounts mit Daten in users oder journeys.
    """
    with user_auth_context_from_db(conn, creator):
        cursor = conn.cursor()
        all_friends = set().union(get_followed_users(), get_trusted_users(), get_followers_users())

        # 2. Wenn Admin: zusätzlich alle Accounts mit Daten in der Datenbank (users und journeys)
        if is_admin:
            cursor.execute("""
                SELECT username FROM users WHERE username IS NOT NULL AND username != ''
                UNION
                SELECT DISTINCT username FROM journeys WHERE username IS NOT NULL AND username != ''
            """)
            for u_row in cursor.fetchall():
                if u_row[0]:
                    all_friends.add(u_row[0])

        all_friends.discard(creator)
        sorted_friends = sorted(list(all_friends), key=lambda s: s.lower())
        all_users = [creator] + sorted_friends

        # 3. Frage für alle diese User aus journeys das Ankunftsdatum des jeweils letzten gespeicherten Status ab
        vis_clause = get_shared_visibility_query()
        query = f"""
        SELECT username,
               MAX(COALESCE(arrivalFinal, arrivalPlanned, departureFinal, departurePlanned)) AS last_arrival,
               COUNT(statusId) AS trip_count
        FROM journeys
        WHERE 1=1 {vis_clause}
        GROUP BY username
        """
        cursor.execute(query)
        user_stats = {}
        for uname, last_arr, count in cursor.fetchall():
            user_stats[uname] = {
                'last_arrival': last_arr,
                'trip_count': count or 0
            }

        # 4. & 5. Formatiere Ergebnisliste mit creator an 1. Stelle
        result = []
        for u in all_users:
            stats = user_stats.get(u, {'last_arrival': None, 'trip_count': 0})
            last_arr_raw = stats['last_arrival']
            german_date = None
            if last_arr_raw:
                try:
                    date_str = str(last_arr_raw).split('T')[0]
                    parts = date_str.split('-')
                    if len(parts) == 3:
                        german_date = f"{parts[2]}.{parts[1]}.{parts[0]}"
                except Exception:
                    german_date = None

            result.append({
                'username': u,
                'is_self': (u == creator),
                'last_arrival_date': german_date,
                'last_arrival_iso': last_arr_raw,
                'trip_count': stats['trip_count']
            })

        return result


# ==========================================
# DATENBESCHAFFUNG MIT HOCHZÄHLERN & COUNTDOWNS
# ==========================================

def fetch_statuses_from_user(api_token, username, already_downloaded: pd.DataFrame | None = None,
                             log_callback=None, progress_callback=None,
                             on_statuses_callback=None) -> list:
    if already_downloaded is None or 'statusId' not in already_downloaded.columns:
        already_ids = set()
    else:
        already_ids = set(already_downloaded['statusId'])

    current_user_data = []
    number = 1
    url = f'https://traewelling.de/api/v1/user/{username}/statuses'
    
    msg = f"[API] Hole Statusse für {username}..."
    print(msg)
    safe_log(log_callback, msg)
    
    while url is not None:
        try:
            travel_response, api_token = traewelling_api_get(url, api_token, username=username)
            travel_response.raise_for_status()
            travel_json = travel_response.json()
            travel_data = travel_json.get('data', [])
            if isinstance(travel_data, dict):
                travel_data = [travel_data]
            elif not isinstance(travel_data, list):
                travel_data = []
            url = travel_json.get('links', {}).get('next')

            current_user_data.extend(travel_data)
            current_ids = [data.get('id', None) for data in travel_data if isinstance(data, dict)]

            if on_statuses_callback and travel_data:
                snippets = []
                for data in travel_data:
                    if not isinstance(data, dict):
                        continue
                    body = (data.get('body') or '').strip()
                    if body:
                        checkin = data.get('checkin') or {}
                        origin = checkin.get('origin') or {}
                        destination = checkin.get('destination') or {}
                        dep_planned = origin.get('departurePlanned')
                        dep_final = checkin.get('manualDeparture') or origin.get('departureReal') or dep_planned
                        arr_planned = destination.get('arrivalPlanned')
                        arr_real = checkin.get('manualArrival') or destination.get('arrivalReal')
                        arr_final = arr_real or arr_planned
                        arrival_delay = None
                        if arr_planned and arr_real:
                            try:
                                t_real = datetime.fromisoformat(arr_real.replace('Z', '+00:00'))
                                t_plan = datetime.fromisoformat(arr_planned.replace('Z', '+00:00'))
                                arrival_delay = int((t_real - t_plan).total_seconds() / 60)
                            except Exception:
                                arrival_delay = None
                        snippets.append({
                            'statusId': data.get('id'),
                            'username': (data.get('user') or {}).get('username') or username,
                            'body': body,
                            'lineName': re.sub(r'\s*\(\d+\)', '', (checkin.get('lineName') or '')).strip(),
                            'category': checkin.get('category') or '',
                            'routeColor': checkin.get('routeColor') or '',
                            'origin': origin.get('name') or '',
                            'destination': destination.get('name') or '',
                            'distance_km': round((checkin.get('distance') or 0) / 1000.0, 1),
                            'departure': dep_final or data.get('createdAt') or '',
                            'arrival': arr_final or '',
                            'arrivalDelay': arrival_delay,
                            'likes': data.get('likes') or 0
                        })
                if snippets:
                    try:
                        on_statuses_callback(snippets)
                    except Exception as cb_err:
                        print(f"[WARN] on_statuses_callback Fehler: {cb_err}")

            if set(current_ids) & already_ids:
                url = None

            rem_hdr = travel_response.headers.get('x-ratelimit-remaining')
            set_remaining(rem_hdr)
            rem_val = 60
            try:
                rem_val = int(rem_hdr)
            except Exception:
                rem_val = 60

            counter_msg = f'   📥 Statusse werden geladen: Seite {number} (Bisher {len(current_user_data)} geladen, Limit: {get_remaining()})'
            print(counter_msg)
            safe_log(log_callback, counter_msg, replace=(number > 1))

            if progress_callback:
                progress_callback({
                    'step_index': 1,
                    'step_name': 'Statusse sammeln',
                    'step_current': number,
                    'step_total': 0,
                    'step_detail': f"Seite {number} ({len(current_user_data)} Statusse geladen)",
                    'rate_limit_remaining': rem_val,
                    'rate_limit_max': max(500, rem_val),
                    'rate_limit_waiting': False,
                    'rate_limit_sleep_remaining': 0
                })

            number += 1
            
        except Exception as e:
            if isinstance(e, requests.exceptions.HTTPError) and e.response.status_code == 429:
                warn = '[API] Rate Limit (429) erreicht. Warte 30 Sekunden...'
                print(warn)
                safe_log(log_callback, warn)
                for sec_left in range(30, 0, -1):
                    if progress_callback:
                        progress_callback({
                            'step_index': 1,
                            'step_name': 'Statusse sammeln',
                            'step_current': number,
                            'step_total': 0,
                            'step_detail': f"Rate-Limit (HTTP 429) – Pausiere {sec_left} Sekunden...",
                            'rate_limit_remaining': 0,
                            'rate_limit_max': max(500, rem_val),
                            'rate_limit_waiting': True,
                            'rate_limit_sleep_remaining': sec_left
                        })
                    sleep(1)
            else:
                err = f'[FEHLER] Konnte Daten von {username} nicht abrufen: {e}'
                print(err)
                safe_log(log_callback, err)
                log_error_to_db(username=username, error=e, context=f"fetch_statuses_from_user ({username}, Seite {number})")
                url = None
                
    done_msg = f'   ✅ Insgesamt {len(current_user_data)} neue Statusse für {username} geladen.'
    print(done_msg)
    if len(current_user_data) > 0:
        safe_log(log_callback, done_msg)
    return current_user_data


def fetch_information_from_user(conn: sqlite3.Connection, api_token, username: str, highest_privilege: int = 0) -> str:
    user_info = {}
    time_of_download_start = datetime.now().isoformat()
    repeat_bool = True
    
    while repeat_bool:
        user_response, api_token = traewelling_api_get(f'https://traewelling.de/api/v1/user/{username}',
                                                      api_token, conn=conn, username=username)
        try:
            user_response.raise_for_status()
            user_json = user_response.json()
            user_info = user_json.get('data', {})
            username = user_info.get('username', '')
            repeat_bool = False
            set_remaining(user_response.headers.get('x-ratelimit-remaining'))
        except Exception as e:
            if isinstance(e, requests.exceptions.HTTPError) and e.response.status_code == 429:
                print('[API] Rate Limit erreicht. Warte 30 Sekunden...')
                sleep(30)
            else:
                repeat_bool = False
                print(f'[FEHLER] Profil {username} nicht gefunden.')
                log_error_to_db(username=username, error=e, context=f"fetch_information_from_user ({username})")

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

        cursor.execute('''
            INSERT INTO users(id, username, displayName, bio, mastodonUrl, totalDistance, totalDuration, profilePicture, lastUpdated, highestPrivilege)
            VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(id) DO UPDATE SET
                username = excluded.username,
                displayName = excluded.displayName,
                bio = excluded.bio,
                mastodonUrl = excluded.mastodonUrl,
                totalDistance = excluded.totalDistance,
                totalDuration = excluded.totalDuration,
                profilePicture = excluded.profilePicture,
                lastUpdated = excluded.lastUpdated,
                highestPrivilege = MAX(COALESCE(users.highestPrivilege, 0), excluded.highestPrivilege)
        ''', (user_id, username, user_info.get('displayName', None), user_info.get('bio', None),
              user_info.get('mastodonUrl', None), user_info.get('totalDistance', None),
              user_info.get('totalDuration', None), user_info.get('profilePicture', None),
              time_of_download_start, highest_privilege))
        conn.commit()

    return username


def get_or_fetch_stop_info(conn: sqlite3.Connection, api_token, stop_id: int, stop_name: str = None,
                           on_download=None, rate_limit_callback=None) -> tuple:
    if stop_id is None or str(stop_id).strip().lower() in ('', 'none'):
        return 'Europe/Berlin', stop_name or 'Unbekannt', None, None

    cursor = conn.cursor()
    cursor.execute('SELECT timezone, stopName, latitude, longitude FROM stops WHERE id=?', (stop_id,))
    row = cursor.fetchone()

    if row is None:
        global tf
        if not stop_name:
            stop_name = str(stop_id)
        timezone = 'Europe/Berlin'
        latitude, longitude = None, None
        repeat_bool = True
        error_404_count = 0
        server_error_count = 0

        while repeat_bool:
            try:
                station_response, api_token = traewelling_api_get(f'https://traewelling.de/api/v1/stations/{stop_id}',
                                                                  api_token, conn=conn)
                station_response.raise_for_status()

                set_remaining(station_response.headers.get('x-ratelimit-remaining'))
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
                    if rate_limit_callback:
                        rate_limit_callback(30)
                    else:
                        sleep(30)
                elif status in (500, 502, 503, 504):
                    server_error_count += 1
                    print(f'[API] Server überlastet ({status}) bei Haltestelle {stop_id}. Versuch {server_error_count}/3')
                    if server_error_count >= 3:
                        msg = f"[API WARNUNG] Haltestelle '{stop_name}' (ID: {stop_id}) konnte nach 3 Versuchen wegen Serverüberlastung ({status}) nicht geladen werden. Verwende Dummy-Werte."
                        print(f"\n{msg}")
                        log_error_to_db(error=e, error_type=f"HTTP_{status}", error_message=msg, context=f"get_or_fetch_stop_info (ID: {stop_id})", conn=conn)
                        return 'Europe/Berlin', stop_name, None, None
                    print('Warte 30 Sekunden...')
                    if rate_limit_callback:
                        rate_limit_callback(30)
                    else:
                        sleep(30)
                elif status == 404:
                    error_404_count += 1
                    print(f"\n[API FEHLER] Haltestelle '{stop_name}' (ID: {stop_id}) nicht gefunden (404). Versuch {error_404_count}/3")
                    if error_404_count >= 3:
                        msg = f"[API WARNUNG] Haltestelle '{stop_name}' (ID: {stop_id}) konnte nach 3 Versuchen nicht gefunden werden (404). Verwende Dummy-Werte."
                        print(f"\n{msg}")
                        log_error_to_db(error=e, error_type=f"HTTP_{status}", error_message=msg, context=f"get_or_fetch_stop_info (ID: {stop_id})", conn=conn)
                        return 'Europe/Berlin', stop_name, None, None
                    sleep(5)
                else:
                    print(f"\n[API FEHLER] Unerwarteter HTTP-Fehler {status} bei '{stop_name}' (ID: {stop_id}).")
                    log_error_to_db(error=e, error_type=f"HTTP_{status}", context=f"get_or_fetch_stop_info (ID: {stop_id})", conn=conn)
                    raise e
            except (requests.exceptions.ConnectionError, requests.exceptions.Timeout) as e:
                print(f"\n[VERBINDUNGSFEHLER] Timeout/Abbruch bei Haltestelle '{stop_name}' (ID: {stop_id}). Warte 30 Sekunden...")
                if rate_limit_callback:
                    rate_limit_callback(30)
                else:
                    sleep(30)
            except KeyboardInterrupt:
                conn.commit()
                conn.close()
                raise SystemExit(0)
            except Exception as e:
                print(f"\n[API FEHLER] Allgemeiner Fehler bei '{stop_name}': {e}")
                log_error_to_db(error=e, context=f"get_or_fetch_stop_info (ID: {stop_id})", conn=conn)
                raise e

        return timezone, stop_name, latitude, longitude
    else:
        return row


def get_inter_stops_for_journeys(api_token, statuses_dicts: list, log_callback=None,
                                 progress_callback=None, rate_limit_callback=None) -> dict:
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
                intermediate_response, api_token = traewelling_api_get(url, api_token)
                intermediate_response.raise_for_status()
                set_remaining(intermediate_response.headers.get('x-ratelimit-remaining'))
                intermediate_data.update(intermediate_response.json().get('data', {}))

                repeat_bool = False
                current_index += 50
                noch_zu_laden = max(0, total - current_index)
                cd_msg = f"   🚉 Zwischenhalte-Countdown: Noch {noch_zu_laden} von {total} Fahrten verbleibend... (Limit: {get_remaining()})"
                print(cd_msg)
                safe_log(log_callback, cd_msg, replace=(current_index > 50))

                if progress_callback:
                    batches_done = math.ceil(min(current_index, total) / 50.0)
                    total_batches = math.ceil(total / 50.0)
                    progress_callback(batches_done, total_batches, min(current_index, total), total)
            except Exception as e:
                if isinstance(e, requests.exceptions.HTTPError) and e.response.status_code == 429:
                    if rate_limit_callback:
                        rate_limit_callback(30)
                    else:
                        sleep(30)
                else:
                    log_error_to_db(username=None, error=e, context=f"get_inter_stops_for_journeys ({url})")
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


def calculate_and_insert(api_token, conn, user_data: list, all_inter_stations: dict,
                         log_callback=None, progress_callback=None) -> None:
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
        oldest_item = user_data[-1]
        oldest_checkin = oldest_item.get('checkin', {})
        oldest_origin = oldest_checkin.get('origin', {})
        oldest_dep = (
            oldest_checkin.get('manualDeparture')
            or oldest_origin.get('departureReal')
            or oldest_origin.get('departurePlanned')
            or oldest_item.get('createdAt')
        )
        if oldest_dep:
            cursor.execute(
                "SELECT destinationId FROM journeys WHERE username=? AND COALESCE(departureFinal, departurePlanned, createdAt) < ? ORDER BY COALESCE(departureFinal, departurePlanned, createdAt) DESC LIMIT 1",
                (username_query, oldest_dep)
            )
        else:
            cursor.execute(
                "SELECT destinationId FROM journeys WHERE username=? ORDER BY departureFinal DESC LIMIT 1",
                (username_query,)
            )
        row = cursor.fetchone()
        if row:
            cursor.execute("SELECT latitude, longitude FROM stops WHERE id=?", (row[0],))
            lat_lon_row = cursor.fetchone()
            if lat_lon_row:
                prev_dest_lat, prev_dest_lon = lat_lon_row

    # Sicherheits-Check: Falls ausnahmsweise noch unbekannte Haltestellen fehlen
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

    for s_id in missing_stop_ids:
        get_or_fetch_stop_info(conn, api_token, s_id, missing_stops_dict.get(s_id))

    total_rides = len(user_data)
    current_ride = 0

    for data in reversed(user_data):
        current_ride += 1
        if progress_callback:
            progress_callback(current_ride, total_rides)

        if current_ride % 100 == 0 or current_ride == total_rides:
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

        if inter_stops is None and trip_id:
            cursor.execute(
                "SELECT interStops, interStopsNumber FROM journeys WHERE tripId = ? AND originId = ? AND destinationId = ? AND interStops IS NOT NULL LIMIT 1",
                (trip_id, origin_id, destination_id)
            )
            cached_inter = cursor.fetchone()
            if not cached_inter:
                cursor.execute(
                    "SELECT interStops, interStopsNumber FROM journeys WHERE tripId = ? AND interStops IS NOT NULL LIMIT 1",
                    (trip_id,)
                )
                cached_inter = cursor.fetchone()
            if cached_inter:
                inter_stops, inter_stops_number = cached_inter[0], cached_inter[1]

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

    # Folgefahrt nach der Schleife aktualisieren (falls durch eingefügte Fahrten eine neue direkte Vorfahrt entstanden ist)
    if user_data and prev_dest_lat is not None:
        newest_item = user_data[0]
        newest_checkin = newest_item.get('checkin', {})
        newest_origin = newest_checkin.get('origin', {})
        newest_dep = (
            newest_checkin.get('manualDeparture')
            or newest_origin.get('departureReal')
            or newest_origin.get('departurePlanned')
            or newest_item.get('createdAt')
        )
        username_query = user_data[0]['user']['username']
        if newest_dep:
            cursor.execute(
                """SELECT statusId, originId FROM journeys
                   WHERE username = ? AND COALESCE(departureFinal, departurePlanned, createdAt) > ?
                   ORDER BY COALESCE(departureFinal, departurePlanned, createdAt) ASC LIMIT 1""",
                (username_query, newest_dep)
            )
            next_row = cursor.fetchone()
            if next_row:
                next_status_id, next_origin_id = next_row[0], next_row[1]
                cursor.execute("SELECT latitude, longitude FROM stops WHERE id = ?", (next_origin_id,))
                n_lat_lon = cursor.fetchone()
                if n_lat_lon and n_lat_lon[0] and n_lat_lon[1]:
                    new_dist_to_prev = calculate_haversine(prev_dest_lat, prev_dest_lon, n_lat_lon[0], n_lat_lon[1])
                    cursor.execute("UPDATE journeys SET distanceToPrevJourney = ? WHERE statusId = ?",
                                   (new_dist_to_prev, next_status_id))

    conn.commit()
    if user_data:
        safe_log(log_callback, f"   ✅ Alle {total_rides} Fahrten berechnet und gespeichert.")


def fetch_data_to_db(username: str, reload_all: bool = False, conn: sqlite3.Connection | None = None,
                     api_token: str | None = None, log_callback=None, update_db: bool = True,
                     progress_callback=None, on_statuses_callback=None) -> sqlite3.Connection:
    if conn is None:
        conn = init_db()
    if api_token is None:
        api_token = log_onto_traewelling(conn)

    current_privilege = get_privilege_level(username)
    cursor = conn.cursor()
    cursor.execute("SELECT highestPrivilege, privilegeCheckpoints FROM users WHERE username = ?", (username,))
    row = cursor.fetchone()

    highest_privilege = row[0] if (row and row[0] is not None) else 0
    raw_checkpoints = row[1] if (row and len(row) > 1 and row[1]) else None
    checkpoints = {}
    if raw_checkpoints:
        try:
            checkpoints = json.loads(raw_checkpoints)
        except Exception:
            checkpoints = {}

    # Fallback für Altdaten ohne Checkpoints
    if not checkpoints and highest_privilege > 0:
        cursor.execute("SELECT MAX(statusId) FROM journeys WHERE username = ?", (username,))
        max_sid_row = cursor.fetchone()
        max_sid = max_sid_row[0] if max_sid_row and max_sid_row[0] else 0
        checkpoints = {str(lvl): max_sid for lvl in range(int(highest_privilege) + 1)}

    max_verified_id = int(checkpoints.get(str(current_privilege), 0) or 0) if not reload_all else 0

    # Profil des Users immer updaten/prüfen
    username = fetch_information_from_user(conn, api_token, username, max(highest_privilege, current_privilege))

    # Helper für Fortschrittsmeldung
    def report_progress(step_index, step_name, step_current, step_total, step_detail,
                        rate_limit_waiting=False, rate_limit_sleep_remaining=0,
                        eta_seconds=None, eta_based_on_history=False, eta_sample_count=0):
        if progress_callback:
            rem_val = 60
            try:
                rem_val = int(get_remaining())
            except Exception:
                rem_val = 60
            progress_callback({
                'step_index': step_index,
                'step_name': step_name,
                'step_current': step_current,
                'step_total': step_total,
                'step_detail': step_detail,
                'rate_limit_remaining': rem_val,
                'rate_limit_max': max(500, rem_val),
                'rate_limit_waiting': rate_limit_waiting,
                'rate_limit_sleep_remaining': rate_limit_sleep_remaining,
                'eta_seconds': eta_seconds,
                'eta_based_on_history': eta_based_on_history,
                'eta_sample_count': eta_sample_count
            })

    # Abbrechen, falls keine neuen Statusse geladen werden sollen
    if not update_db:
        msg = f"[INFO] Update der Statusse übersprungen: Nutze bestehende Daten für {username}."
        print(msg)
        safe_log(log_callback, msg)
        report_progress(1, 'Statusse sammeln', 0, 0, "Update übersprungen")
        report_progress(2, 'Haltestellen herunterladen', 0, 0, "Update übersprungen")
        report_progress(3, 'Zwischenhaltestellen laden', 0, 0, "Update übersprungen")
        report_progress(4, 'Fahrten berechnen', 0, 0, "Update übersprungen")
        return conn

    downloaded_frame = pd.DataFrame() if (reload_all or max_verified_id == 0) else already_downloaded_in_db(conn, username, max_verified_status_id=max_verified_id)

    # ==========================================
    # SCHRITT 1: Statusse sammeln
    # ==========================================
    report_progress(1, 'Statusse sammeln', 0, 0, f"Hole Statusse für {username}...")
    user_statuses = fetch_statuses_from_user(api_token, username, downloaded_frame,
                                             log_callback=log_callback,
                                             progress_callback=progress_callback,
                                             on_statuses_callback=on_statuses_callback)
    report_progress(1, 'Statusse sammeln', 1, 1, f"{len(user_statuses)} Statusse gefunden")

    if not user_statuses:
        msg = f"[INFO] Keine neuen Fahrten für {username} gefunden. System ist up-to-date."
        print(msg)
        safe_log(log_callback, msg)
        report_progress(2, 'Haltestellen herunterladen', 0, 0, "Keine neuen Haltestellen nötig (up-to-date)")
        report_progress(3, 'Zwischenhaltestellen laden', 0, 0, "Keine neuen Zwischenhalte nötig (up-to-date)")
        report_progress(4, 'Fahrten berechnen', 0, 0, "Alle Fahrten aktuell (up-to-date)")
    else:
        # Historische Telemetrie für ETA laden
        telemetry_metrics, telemetry_samples = get_telemetry_metrics(conn)
        hist_sec_per_stop = telemetry_metrics.get('sec_per_stop', 0.45)
        hist_sec_per_inter_batch = telemetry_metrics.get('sec_per_inter_batch', 0.65)
        hist_sec_per_ride_calc = telemetry_metrics.get('sec_per_ride_calc', 0.008)

        # Mengengerüste für Schritte 2, 3 und 4 ermitteln
        all_stop_ids = set()
        missing_stops_dict = {}
        for d in user_statuses:
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

        cursor.execute("SELECT DISTINCT tripId FROM journeys WHERE username = ? AND interStops IS NOT NULL", (username,))
        known_trip_ids = {r[0] for r in cursor.fetchall()}
        statuses_needing_inter = [
            s for s in user_statuses
            if (s.get('checkin') or {}).get('trip') not in known_trip_ids
        ]
        total_inter_trips = len(statuses_needing_inter)
        total_inter_batches = math.ceil(total_inter_trips / 50.0)
        total_rides = len(user_statuses)

        # ETA-Berechnungsfunktion (gemäß Abschnitt 5.2)
        def compute_eta(step_idx, current_done, step_active_time, waiting_sleep_sec=0):
            rem_int = 60
            try:
                rem_int = int(get_remaining())
            except Exception:
                rem_int = 60

            if step_idx == 2:
                rem_stops = max(0, total_missing_stops - current_done)
                if current_done >= 3 and step_active_time > 0:
                    live_rate = step_active_time / float(current_done)
                    eff_stop_rate = 0.5 * hist_sec_per_stop + 0.5 * live_rate
                else:
                    eff_stop_rate = hist_sec_per_stop
                work_eta = (rem_stops * eff_stop_rate) + (total_inter_batches * hist_sec_per_inter_batch) + (total_rides * hist_sec_per_ride_calc)
                rem_api_calls = rem_stops + total_inter_batches

            elif step_idx == 3:
                rem_batches = max(0, total_inter_batches - current_done)
                if current_done >= 3 and step_active_time > 0:
                    live_rate = step_active_time / float(current_done)
                    eff_batch_rate = 0.5 * hist_sec_per_inter_batch + 0.5 * live_rate
                else:
                    eff_batch_rate = hist_sec_per_inter_batch
                work_eta = (rem_batches * eff_batch_rate) + (total_rides * hist_sec_per_ride_calc)
                rem_api_calls = rem_batches

            elif step_idx == 4:
                rem_rides = max(0, total_rides - current_done)
                if current_done >= 3 and step_active_time > 0:
                    live_rate = step_active_time / float(current_done)
                    eff_ride_rate = 0.5 * hist_sec_per_ride_calc + 0.5 * live_rate
                else:
                    eff_ride_rate = hist_sec_per_ride_calc
                work_eta = rem_rides * eff_ride_rate
                rem_api_calls = 0
            else:
                return None

            if rem_api_calls > rem_int:
                expected_waits = math.ceil((rem_api_calls - rem_int) / 60.0)
                work_eta += expected_waits * 30.0

            if waiting_sleep_sec > 0:
                work_eta += waiting_sleep_sec

            return max(0, int(round(work_eta)))

        # ==========================================
        # SCHRITT 2: Haltestellen herunterladen
        # ==========================================
        if total_missing_stops == 0:
            report_progress(2, 'Haltestellen herunterladen', 0, 0, "Keine neuen Haltestellen nötig (alle gecacht)",
                            eta_seconds=compute_eta(2, 0, 0),
                            eta_based_on_history=(telemetry_samples > 0),
                            eta_sample_count=telemetry_samples)
        else:
            init_s_msg = f"   🚏 Haltestellen-Countdown: Noch {total_missing_stops} neue Haltestellen zu laden... (Limit: {get_remaining()})"
            print(init_s_msg)
            safe_log(log_callback, init_s_msg)

            t0_step2 = time()
            sleep_deduction_step2 = 0.0
            stops_downloaded = 0

            def on_stop_rate_limit(secs=30):
                nonlocal sleep_deduction_step2
                t_sl_start = time()
                for sl_left in range(secs, 0, -1):
                    act_sec = (time() - t0_step2) - sleep_deduction_step2
                    eta_val = compute_eta(2, stops_downloaded, act_sec, waiting_sleep_sec=sl_left)
                    report_progress(2, 'Haltestellen herunterladen', stops_downloaded, total_missing_stops,
                                    f"Rate-Limit (HTTP 429) – Pausiere {sl_left} Sekunden...",
                                    rate_limit_waiting=True, rate_limit_sleep_remaining=sl_left,
                                    eta_seconds=eta_val, eta_based_on_history=(telemetry_samples > 0),
                                    eta_sample_count=telemetry_samples)
                    sleep(1)
                sleep_deduction_step2 += (time() - t_sl_start)

            for idx, s_id in enumerate(missing_stop_ids):
                s_name_guess = missing_stops_dict.get(s_id)
                tz, fetched_name, lat, lon = get_or_fetch_stop_info(conn, api_token, s_id, s_name_guess,
                                                                    rate_limit_callback=on_stop_rate_limit)
                stops_downloaded += 1
                rem = max(0, total_missing_stops - stops_downloaded)
                s_msg = f"   🚏 Lade Haltestelle: {fetched_name} | Noch {rem} von {total_missing_stops} übrig (Limit: {get_remaining()})"
                print(s_msg)
                safe_log(log_callback, s_msg, replace=(stops_downloaded > 1))

                act_sec = (time() - t0_step2) - sleep_deduction_step2
                eta_val = compute_eta(2, stops_downloaded, act_sec)
                report_progress(2, 'Haltestellen herunterladen', stops_downloaded, total_missing_stops,
                                f"Lade Haltestelle: {fetched_name} ({stops_downloaded}/{total_missing_stops})",
                                eta_seconds=eta_val, eta_based_on_history=(telemetry_samples > 0),
                                eta_sample_count=telemetry_samples)

            active_sec_step2 = (time() - t0_step2) - sleep_deduction_step2
            update_download_telemetry(conn, 'sec_per_stop', active_sec_step2, total_missing_stops)

        # ==========================================
        # SCHRITT 3: Zwischenhaltestellen laden
        # ==========================================
        if total_inter_trips == 0:
            report_progress(3, 'Zwischenhaltestellen laden', 0, 0, "Keine neuen Zwischenhalte nötig",
                            eta_seconds=compute_eta(3, 0, 0),
                            eta_based_on_history=(telemetry_samples > 0),
                            eta_sample_count=telemetry_samples)
            inter_stations = {}
        else:
            t0_step3 = time()
            sleep_deduction_step3 = 0.0

            def on_inter_rate_limit(secs=30):
                nonlocal sleep_deduction_step3
                t_sl_start = time()
                for sl_left in range(secs, 0, -1):
                    act_sec = (time() - t0_step3) - sleep_deduction_step3
                    eta_val = compute_eta(3, 0, act_sec, waiting_sleep_sec=sl_left)
                    report_progress(3, 'Zwischenhaltestellen laden', 0, total_inter_trips,
                                    f"Rate-Limit (HTTP 429) – Pausiere {sl_left} Sekunden...",
                                    rate_limit_waiting=True, rate_limit_sleep_remaining=sl_left,
                                    eta_seconds=eta_val, eta_based_on_history=(telemetry_samples > 0),
                                    eta_sample_count=telemetry_samples)
                    sleep(1)
                sleep_deduction_step3 += (time() - t_sl_start)

            def on_inter_progress(batches_done, total_b, journeys_done, journeys_tot):
                act_sec = (time() - t0_step3) - sleep_deduction_step3
                eta_val = compute_eta(3, batches_done, act_sec)
                report_progress(3, 'Zwischenhaltestellen laden', journeys_done, journeys_tot,
                                f"Zwischenhalte: {journeys_done}/{journeys_tot} Fahrten",
                                eta_seconds=eta_val, eta_based_on_history=(telemetry_samples > 0),
                                eta_sample_count=telemetry_samples)

            inter_stations = get_inter_stops_for_journeys(api_token, statuses_needing_inter,
                                                          log_callback=log_callback,
                                                          progress_callback=on_inter_progress,
                                                          rate_limit_callback=on_inter_rate_limit)
            active_sec_step3 = (time() - t0_step3) - sleep_deduction_step3
            update_download_telemetry(conn, 'sec_per_inter_batch', active_sec_step3, total_inter_batches)

        # ==========================================
        # SCHRITT 4: Fahrten berechnen
        # ==========================================
        t0_step4 = time()

        def on_ride_calc_progress(current_ride, total_r):
            act_sec = time() - t0_step4
            eta_val = compute_eta(4, current_ride, act_sec)
            report_progress(4, 'Fahrten berechnen', current_ride, total_r,
                            f"Berechne Fahrt {current_ride}/{total_r}",
                            eta_seconds=eta_val, eta_based_on_history=(telemetry_samples > 0),
                            eta_sample_count=telemetry_samples)

        calculate_and_insert(api_token, conn, user_statuses, inter_stations,
                             log_callback=log_callback,
                             progress_callback=on_ride_calc_progress)

        active_sec_step4 = time() - t0_step4
        update_download_telemetry(conn, 'sec_per_ride_calc', active_sec_step4, total_rides)

    cursor.execute("SELECT MAX(statusId) FROM journeys WHERE username = ?", (username,))
    new_max_row = cursor.fetchone()
    new_max_sid = new_max_row[0] if new_max_row and new_max_row[0] else 0
    for lvl in range(current_privilege + 1):
        prev_lvl_sid = int(checkpoints.get(str(lvl), 0) or 0)
        if new_max_sid > prev_lvl_sid:
            checkpoints[str(lvl)] = new_max_sid

    cursor.execute(
        "UPDATE users SET highestPrivilege = ?, privilegeCheckpoints = ? WHERE username = ?",
        (max(highest_privilege, current_privilege), json.dumps(checkpoints), username)
    )
    conn.commit()
    return conn

# ==========================================
# DATENAUSWERTUNG / PANDAS DATAFRAMES
# ==========================================

def format_utc_to_datatime(utc: str, tz: str) -> datetime | None:
    if not utc or pd.isna(utc) or not isinstance(utc, str):
        return None
    if not tz or pd.isna(tz) or not isinstance(tz, str):
        tz = 'Europe/Berlin'
    try:
        dt_utc = datetime.fromisoformat(utc.replace('Z', '+00:00'))
        dt_local = dt_utc.astimezone(ZoneInfo(tz))
        return dt_local.replace(tzinfo=None)
    except Exception as e:
        print(f"[WARNUNG] Fehler beim Parsen des Zeitstempels '{utc}' (Zeitzone: '{tz}'): {e}")
        log_error_to_db(
            error=e,
            context=f"format_utc_to_datatime (utc={utc!r}, tz={tz!r})"
        )
        return None


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
    extra_sql = getattr(_INTERACTIVE_FILTER_CTX, 'sql_clause', '')
    if extra_sql:
        date_query += extra_sql
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
                "traewelling_analysis_v10.py"
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
            "traewelling_analysis_v10.py",
            start_date if start_date else "Kein Startdatum (Alle)",
            end_date if end_date else "Kein Enddatum (Alle)",
            f'{len(usernames)}',
            ", ".join(users_with_journeys) if users_with_journeys else "-",
            ", ".join(users_without_journeys) if users_without_journeys else "-"
        ]
    })


# ==========================================
# EXCEL-BAUKASTEN: PARAMETER & DEFAULT CONFIG
# ==========================================
AVAILABLE_GROUP_COLUMNS = {
    # Fahrzeug & Betrieb
    'category': {'sql': 'j.category', 'label': 'Kategorie', 'category': 'Fahrzeug & Betrieb'},
    'operatorName': {'sql': 'j.operatorName', 'label': 'Betreiber', 'category': 'Fahrzeug & Betrieb'},
    'lineName': {'sql': 'j.lineName', 'label': 'Linie', 'category': 'Fahrzeug & Betrieb'},
    'lineNumber': {'sql': 'j.lineNumber', 'label': 'Liniennummer', 'category': 'Fahrzeug & Betrieb'},
    'journeyNumber': {'sql': 'j.journeyNumber', 'label': 'Fahrtnummer', 'category': 'Fahrzeug & Betrieb'},
    'locomotiveClass': {'sql': 'j.locomotiveClass', 'label': 'Baureihe', 'category': 'Fahrzeug & Betrieb'},
    'vehicleNumber': {'sql': 'j.vehicleNumber', 'label': 'Fahrzeugnummer', 'category': 'Fahrzeug & Betrieb'},
    'wagonClass': {'sql': 'j.wagonClass', 'label': 'Wagenbauart', 'category': 'Fahrzeug & Betrieb'},
    'wagon': {'sql': 'j.wagon', 'label': 'Wagennummer', 'category': 'Fahrzeug & Betrieb'},
    'seat': {'sql': 'j.seat', 'label': 'Sitzplatz', 'category': 'Fahrzeug & Betrieb'},
    # Tarif, Komfort & Träwelling-Tags
    'business': {'sql': 'j.business', 'label': 'Fahrtgrund', 'category': 'Tarif & Träwelling-Tags'},
    'ticket': {'sql': 'j.ticket', 'label': 'Ticket', 'category': 'Tarif & Träwelling-Tags'},
    'travelClass': {'sql': 'j.travelClass', 'label': 'Reiseklasse', 'category': 'Tarif & Träwelling-Tags'},
    'price': {'sql': 'j.price', 'label': 'Fahrpreis', 'category': 'Tarif & Träwelling-Tags'},
    'trainUtilization': {'sql': 'j.trainUtilization', 'label': 'Auslastung', 'category': 'Tarif & Träwelling-Tags'},
    'role': {'sql': 'j.role', 'label': 'Rolle', 'category': 'Tarif & Träwelling-Tags'},
    'passengerRights': {'sql': 'j.passengerRights', 'label': 'Fahrgastrechte', 'category': 'Tarif & Träwelling-Tags'},
    'socialStatus': {'sql': 'j.socialStatus', 'label': 'Sozialstatus', 'category': 'Tarif & Träwelling-Tags'},
    'event': {'sql': 'j.event', 'label': 'Event', 'category': 'Tarif & Träwelling-Tags'},
    # Haltestellen & Geografie (Start / Ziel)
    'origin.stopName': {'sql': 'origin.stopName', 'label': 'Starthaltestelle', 'category': 'Haltestellen & Geografie'},
    'destination.stopName': {'sql': 'destination.stopName', 'label': 'Zielhaltestelle', 'category': 'Haltestellen & Geografie'},
    'origin.city': {'sql': 'origin.city', 'label': 'Start-Stadt', 'category': 'Haltestellen & Geografie'},
    'destination.city': {'sql': 'destination.city', 'label': 'Ziel-Stadt', 'category': 'Haltestellen & Geografie'},
    'origin.admin2Dist': {'sql': 'origin.admin2Dist', 'label': 'Start-Bezirk/Kreis', 'category': 'Haltestellen & Geografie'},
    'destination.admin2Dist': {'sql': 'destination.admin2Dist', 'label': 'Ziel-Bezirk/Kreis', 'category': 'Haltestellen & Geografie'},
    'origin.admin1State': {'sql': 'origin.admin1State', 'label': 'Start-Bundesland/Region', 'category': 'Haltestellen & Geografie'},
    'destination.admin1State': {'sql': 'destination.admin1State', 'label': 'Ziel-Bundesland/Region', 'category': 'Haltestellen & Geografie'},
    'origin.countryName': {'sql': 'origin.countryName', 'label': 'Start-Land', 'category': 'Haltestellen & Geografie'},
    'destination.countryName': {'sql': 'destination.countryName', 'label': 'Ziel-Land', 'category': 'Haltestellen & Geografie'},
    'origin.rilIdentifier': {'sql': 'origin.rilIdentifier', 'label': 'Start-RIL100', 'category': 'Haltestellen & Geografie'},
    'destination.rilIdentifier': {'sql': 'destination.rilIdentifier', 'label': 'Ziel-RIL100', 'category': 'Haltestellen & Geografie'},
    'origin.timezone': {'sql': 'origin.timezone', 'label': 'Start-Zeitzone', 'category': 'Haltestellen & Geografie'},
    'destination.timezone': {'sql': 'destination.timezone', 'label': 'Ziel-Zeitzone', 'category': 'Haltestellen & Geografie'},
    # System & Zeit-Ableitungen
    'clientName': {'sql': 'j.clientName', 'label': 'Check-In-Tool', 'category': 'System & Zeit'},
    'visibility': {'sql': 'j.visibility', 'label': 'Sichtbarkeit', 'category': 'System & Zeit'},
    'realtimeAvailability': {'sql': 'j.realtimeAvailability', 'label': 'Echtzeit verfügbar', 'category': 'System & Zeit'},
    'realtimeFromDataprovider': {'sql': 'j.realtimeFromDataprovider', 'label': 'Echtzeit (Anbieter)', 'category': 'System & Zeit'},
    'hafasId': {'sql': 'j.hafasId', 'label': 'HAFAS-ID', 'category': 'System & Zeit'},
    'year_planned': {'sql': "strftime('%Y', j.departurePlanned)", 'label': 'Jahr (Plan)', 'category': 'System & Zeit'},
    'month_year_planned': {'sql': "strftime('%Y-%m', j.departurePlanned)", 'label': 'Monat (JJJJ-MM)', 'category': 'System & Zeit'},
    'month_planned': {'sql': "strftime('%m', j.departurePlanned)", 'label': 'Kalendermonat', 'category': 'System & Zeit'},
    'date_planned': {'sql': "date(j.departurePlanned)", 'label': 'Kalendertag', 'category': 'System & Zeit'},
}

AVAILABLE_TEMPORAL_TYPES = {
    'time_of_day': 'Tageszeiten (30-Min-Fenster)',
    'weekday': 'Wochentage',
    'month': 'Monat (01–12)',
    'month_year': 'Monate im Jahr (JJJJ-MM)',
    'date': 'Tag (Datum)',
    'year': 'Jahr'
}

AVAILABLE_GEO_COLUMNS = {
    'countryName': 'Länder',
    'city': 'Städte',
    'admin1State': 'Regionen / Bundesländer',
    'admin2Dist': 'Bezirke / Kreise',
    'officialName': 'Offizielle Ländernamen',
    'countryCode': 'Ländercodes (ISO)',
    'timezone': 'Zeitzonen'
}

SHEET_METRICS_SCHEMA = {
    'meta': {
        'mode': 'rows',
        'key_col': 'Label',
        'label': 'Zeilen im Meta-Tabellenblatt',
        'items': [
            'Datenstand', 'Anzeigename', 'Profilbild', 'Mastodonlink',
            'Bio', 'Gesamtdistanz', 'Gesamtzeit', 'Programmversion'
        ]
    },
    'journeys': {
        'mode': 'columns',
        'fixed_cols': [],
        'label': 'Spalten in der Fahrtenliste',
        'items': [
            'Linie', 'Kategorie', 'Fahrtnummer', 'Betreiber', 'Grund', 'Fahrzeugnummer',
            'Abfahrtshaltestelle', 'Abfahrt geplant', 'Abfahrt real', 'Abw ab',
            'Zwischenhaltestellen', 'Zielstation', 'Ankunft geplant', 'Ankunft real', 'Abw an',
            'Reisezeit plan (min)', 'Reisezeit (min)', 'Reisezeit (hh:mm)', 'Delta', 'Delta pro Planzeit',
            'Entfernung (m)', 'Entfernung (km)', 'Luftlinie Fahrt (km)', 'Luftlinie zur Vorfahrt (km)',
            'Umweg Faktor', 'Geschwindigkeit (km/h)', 'Punkte', 'Likes', 'Statustext',
            'Besuchte Länder', 'Link'
        ]
    },
    'stats': {
        'mode': 'rows',
        'key_col': 'Metrik',
        'label': 'Statistik-Metriken (Zeilen)',
        'items': [
            'Fahrten', 'Tage', 'Linien', 'Betreiber', 'Haltestellen',
            'Abfahrtsverspätung', 'Ankunftsverspätung', 'Zeit (min)', 'Zeit (hh:mm)',
            'Distanz (m)', 'Distanz (km)', 'Luftlinie Fahrten (km)', 'Luftlinie zwischen Fahrten (km)',
            'Punkte', 'Likes', 'Tracking-Lücke', 'Umweg Faktor', 'Geschwindigkeit',
            'Echtzeitquote', 'Verfrüht', 'Pünktlich', '3 Minuten oder mehr verspätet',
            '6 Minuten oder mehr verspätet', '15 Minuten oder mehr verspätet',
            '60 Minuten oder mehr verspätet', 'Ø Zwischenhalte'
        ]
    },
    'stops': {
        'mode': 'columns',
        'fixed_cols': 1,
        'label': 'Metrik-Spalten (Haltestellen ohne Via)',
        'items': ['Besuche', 'Hub-Distanz-Wert', 'Letzter Besuch']
    },
    'stops_via': {
        'mode': 'columns',
        'fixed_cols': 1,
        'label': 'Metrik-Spalten (Haltestellen mit Via)',
        'items': ['Ohne Via', 'Nur Via', 'insgesamt', 'Umsteigequote (%)', 'realative Umsteigequote (%)', 'Letzter Besuch']
    },
    'extended_stats': {
        'mode': 'rows',
        'key_col': 'Beschreibung',
        'label': 'Metriken der erweiterten Statistik (Zeilen)',
        'items': [
            'Aktionsradius (km)', 'RMSE der Abweichung von Realtimeanbieter',
            'Sigma Ankunftsverspätung (σ)', 'Nördlichste Haltestelle',
            'Östlichste Haltestelle', 'Südlichste Haltestelle',
            'Westlichste Haltestelle', 'Größte Distanz zum Hub'
        ]
    },
    'grouped': {
        'mode': 'columns',
        'fixed_cols': 'dynamic_args',
        'label': 'Berechnete Statistik-Spalten (Gruppierung)',
        'default_items': [
            'Fahrten', 'Distanz (km)', 'Anteil an Gesamtdistanz (%)', 'Umweg Faktor',
            'Zeit (min)', 'Anteil an Gesamtzeit (%)', '∅ Geschwindigkeit', '∅ Verspätung (min)',
            '3 Minuten und mehr (%)', '6 Minuten und mehr (%)', '15 Minuten und mehr (%)',
            'Punkte', '∅ Zwischenhalte'
        ],
        'items': [
            'Fahrten', 'Tage', 'Linien', 'Betreiber', 'Haltestellen',
            'Distanz (km)', 'Distanz (m)', 'Anteil an Gesamtdistanz (%)',
            'Luftlinie Fahrten (km)', 'Luftlinie zwischen Fahrten (km)', 'Tracking-Lücke (%)', 'Umweg Faktor',
            'Zeit (min)', 'Zeit (hh:mm)', 'Anteil an Gesamtzeit (%)',
            '∅ Geschwindigkeit', 'Abfahrtsverspätung (min)', 'Ankunftsverspätung (min)',
            '∅ Abfahrtsverspätung (min)', '∅ Verspätung (min)',
            'Echtzeitquote (%)', 'Verfrüht (%)', 'Pünktlich (%)',
            '3 Minuten und mehr (%)', '6 Minuten und mehr (%)',
            '15 Minuten und mehr (%)', '60 Minuten und mehr (%)',
            'Punkte', 'Likes', '∅ Zwischenhalte'
        ]
    },
    'temporal': {
        'mode': 'columns',
        'fixed_cols': 1,
        'label': 'Berechnete Statistik-Spalten (Zeitreihen)',
        'default_items': [
            'Fahrten', 'Zeit (min)', 'Anteil Zeit (%)', 'Distanz (km)',
            'Anteil Distanz (%)', '∅ Geschwindigkeit (km/h)', '∅ Ankunftsverspätung (min)',
            '∅ Minuten pro aktivem Tag', '∅ Distanz pro aktivem Tag (km)'
        ],
        'items': [
            'Fahrten', 'Tage', 'Linien', 'Betreiber', 'Haltestellen',
            'Zeit (min)', 'Zeit (hh:mm)', 'Anteil Zeit (%)',
            'Distanz (km)', 'Distanz (m)', 'Anteil Distanz (%)',
            'Luftlinie Fahrten (km)', 'Luftlinie zwischen Fahrten (km)',
            'Tracking-Lücke (%)', 'Umweg Faktor',
            '∅ Geschwindigkeit (km/h)', '∅ Minuten pro aktivem Tag', '∅ Distanz pro aktivem Tag (km)',
            'Abfahrtsverspätung (min)', 'Ankunftsverspätung (min)',
            '∅ Abfahrtsverspätung (min)', '∅ Ankunftsverspätung (min)',
            'Echtzeitquote (%)', 'Verfrüht (%)', 'Pünktlich (%)',
            '3 Minuten und mehr (%)', '6 Minuten und mehr (%)',
            '15 Minuten und mehr (%)', '60 Minuten und mehr (%)',
            'Punkte', 'Likes', '∅ Zwischenhalte'
        ]
    },
    'geo': {
        'mode': 'columns',
        'fixed_cols': 1,
        'label': 'Spalten (Geografie-Ranking)',
        'items': ['Besuche']
    },
    'shared_meta': {
        'mode': 'rows',
        'key_col': 'Label',
        'label': 'Zeilen im Team-Meta-Blatt',
        'items': [
            'Datenstand (Export)', 'Programmversion', 'Startdatum des Exports',
            'Enddatum des Exports', 'Anzahl berücksichtigter User',
            'User mit Fahrten', 'User ohne Fahrten'
        ]
    },
    'shared_stats': {
        'mode': 'columns',
        'fixed_cols': ['User'],
        'label': 'Vergleichs-Metriken im Team-Statistik-Blatt (Spalten)',
        'items': [
            'Fahrten', 'Tage', 'Fahrten pro Tag', 'Linien', 'Betreiber', 'Haltestellen',
            'Zeit (min)', 'Zeit (hh:mm)', 'Distanz (km)', 'Luftlinie Fahrten (km)',
            'Luftlinie zwischen Fahrten (km)', 'Tracking-Lücke (%)', 'Umweg Faktor',
            'Abfahrtsverspätung (min)', 'Ankunftsverspätung (min)', 'Punkte', 'Likes',
            'Echtzeitquote (%)', 'Verfrüht (%)', 'Pünktlich (%)',
            '3 Minuten oder mehr verspätet (%)', '6 Minuten oder mehr verspätet (%)',
            '15 Minuten oder mehr verspätet (%)', '60 Minuten oder mehr verspätet (%)',
            '∅ Zwischenhalte'
        ]
    },
    'shared_stops': {
        'mode': 'columns',
        'fixed_cols': ['Haltestelle', 'User'],
        'label': 'Metrik-Spalten (Gemeinsame Haltestellen)',
        'items': ['Besuche insgesamt', 'Ohne Via', 'Nur Via', 'Umstiegquote (%)']
    },
    'shared_journeys': {
        'mode': 'columns',
        'fixed_cols': [],
        'label': 'Spalten (Gemeinsame Fahrten)',
        'items': [
            'Kategorie', 'Linie', 'User', 'Starthaltestelle',
            'Abfahrtszeit (plan)', 'Abfahrtszeit (ist)', 'Zielhaltestelle',
            'Ankunftszeit (plan)', 'Ankunftszeit (ist)', 'Entfernung (km)',
            'Punkte', 'Link'
        ]
    },
    'shared_grouped': {
        'mode': 'columns',
        'fixed_cols': 'dynamic_shared_args',
        'label': 'Vergleichs-Spalten (Gruppierter Team-Vergleich)',
        'default_items': [
            'Fahrten', 'Distanz (km)', 'Umweg Faktor', 'Zeit (hh:mm)',
            'Zeit (min)', 'Geschwindigkeit (km/h)', '∅ Verspätung (min)'
        ],
        'items': [
            'Fahrten', 'Tage', 'Linien', 'Betreiber', 'Haltestellen',
            'Distanz (km)', 'Distanz (m)', 'Luftlinie Fahrten (km)', 'Luftlinie zwischen Fahrten (km)',
            'Tracking-Lücke (%)', 'Umweg Faktor', 'Zeit (hh:mm)', 'Zeit (min)',
            'Geschwindigkeit (km/h)', 'Abfahrtsverspätung (min)', 'Ankunftsverspätung (min)',
            '∅ Abfahrtsverspätung (min)', '∅ Verspätung (min)',
            'Echtzeitquote (%)', 'Verfrüht (%)', 'Pünktlich (%)',
            '3 Minuten und mehr (%)', '6 Minuten und mehr (%)',
            '15 Minuten und mehr (%)', '60 Minuten und mehr (%)',
            'Punkte', 'Likes', '∅ Zwischenhalte'
        ]
    }
}


def apply_sheet_metrics(df: pd.DataFrame, sheet_type: str, selected_metrics: list | None, num_group_args: int = 1) -> pd.DataFrame:
    """Filtert und sortiert die Spalten bzw. Zeilen eines DataFrames anhand der vom User gewählten Metrik-Liste."""
    if df is None or df.empty or sheet_type not in SHEET_METRICS_SCHEMA:
        return df

    schema = SHEET_METRICS_SCHEMA[sheet_type]
    if not selected_metrics:
        selected_metrics = schema.get('default_items')
    if not selected_metrics:
        return df

    mode = schema.get('mode')

    if mode == 'rows':
        key_col = schema['key_col']
        existing_keys = set(df[key_col].tolist())
        valid_metrics = [m for m in selected_metrics if m in existing_keys]
        if not valid_metrics:
            return df
        df_filtered = df[df[key_col].isin(valid_metrics)].copy()
        df_filtered['__sort_order__'] = pd.Categorical(df_filtered[key_col], categories=valid_metrics, ordered=True)
        df_filtered = df_filtered.sort_values('__sort_order__').drop(columns=['__sort_order__']).reset_index(drop=True)
        return df_filtered

    elif mode == 'columns':
        fixed_spec = schema.get('fixed_cols', [])
        all_cols = list(df.columns)

        if isinstance(fixed_spec, int):
            prefix_cols = all_cols[:fixed_spec]
        elif fixed_spec == 'dynamic_args':
            prefix_cols = all_cols[:max(1, num_group_args)]
        elif fixed_spec == 'dynamic_shared_args':
            # Gruppierungsspalten + 'User'-Spalte bleiben vorne fixiert
            prefix_cols = all_cols[:max(1, num_group_args) + 1]
        elif isinstance(fixed_spec, list):
            prefix_cols = [c for c in fixed_spec if c in all_cols]
        else:
            prefix_cols = []

        metric_cols = [m for m in selected_metrics if m in all_cols and m not in prefix_cols]
        if not metric_cols and not prefix_cols:
            return df

        final_cols = prefix_cols + metric_cols
        return df[final_cols]

    return df


def resolve_filtered_status_ids(conn: sqlite3.Connection, usernames: list[str], start_date: str = '', end_date: str = '', extra_filters: list[dict] = None) -> list[int]:
    if not extra_filters:
        return []
    if isinstance(usernames, str):
        usernames = [usernames]
    placeholders = ','.join(['?'] * len(usernames))
    conds = [f"j_flt.username IN ({placeholders})"]
    params = list(usernames)

    if start_date:
        conds.append("COALESCE(j_flt.arrivalFinal, j_flt.arrivalPlanned) > ?")
        params.append(start_date)
    if end_date:
        e_val = end_date + 'T23:59:59' if len(end_date) == 10 else end_date
        conds.append("COALESCE(j_flt.departureFinal, j_flt.departurePlanned) < ?")
        params.append(e_val)

    temporal_filters = []
    for flt in extra_filters:
        col = flt.get('col', '')
        val = flt.get('val')
        if not col or val is None:
            continue
        # Aggregat- und Prozent-Spalten ignorieren
        if any(col.endswith(sfx) for sfx in ('(%)', '(km)', '(min)', '(km/h)')) or col in ('Umweg Faktor', 'Fahrten'):
            continue

        if col.startswith('temporal_'):
            t_type = col.replace('temporal_', '')
            t_val = str(val).strip()
            # Validierung, um fehlerhafte Werte (wie z. B. '4.25' bei month_year) abzufangen
            if t_type == 'time_of_day' and (':' not in t_val or '-' not in t_val):
                continue
            if t_type == 'weekday' and t_val not in ('Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag', 'Sonntag'):
                continue
            if t_type == 'month' and not re.match(r'^(0?[1-9]|1[0-2])$', t_val):
                continue
            if t_type == 'month_year' and not re.match(r'^\d{4}-(0[1-9]|1[0-2])$', t_val):
                continue
            if t_type == 'date' and not re.match(r'^\d{4}-\d{2}-\d{2}$', t_val):
                continue
            if t_type == 'year' and not re.match(r'^\d{4}$', t_val):
                continue
            temporal_filters.append((t_type, t_val))
        elif col == 'visibility':
            v_str = str(val).lower()
            vis_map = {
                '0': '0', 'öffentlich': '0',
                '1': '1', 'nicht gelistet': '1', 'unlisted': '1',
                '2': '2', 'nur follower': '2', 'follower': '2',
                '3': '3', 'privat': '3', 'private': '3',
                '4': '4', 'angemeldete nutzer': '4', 'authenticated': '4',
                '5': '5', 'enge freunde': '5'
            }
            clean_v = ''.join(c for c in v_str if c.isalnum() or c.isspace()).strip()
            target_code = vis_map.get(clean_v, vis_map.get(str(val).strip(), str(val).strip()))
            conds.append("CAST(j_flt.visibility AS TEXT) = ?")
            params.append(target_code)
        elif col in ('realtimeAvailability', 'realtimeFromDataprovider'):
            v_str = str(val).lower()
            b_val = 1 if ('ja' in v_str or '1' in v_str or 'true' in v_str) else 0
            conds.append(f"j_flt.{col} = ?")
            params.append(b_val)
        elif col in AVAILABLE_GROUP_COLUMNS:
            sql_expr = AVAILABLE_GROUP_COLUMNS[col]['sql'].replace('j.', 'j_flt.')
            if val is None or str(val) == '' or str(val) == 'None':
                conds.append(f"({sql_expr} IS NULL OR CAST({sql_expr} AS TEXT) = '')")
            else:
                conds.append(f"CAST({sql_expr} AS TEXT) = ?")
                params.append(str(val))
        elif col == 'any_stop':
            conds.append("(origin.stopName = ? OR destination.stopName = ?)")
            params.extend([str(val), str(val)])
        elif col == 'any_stop_via':
            conds.append("(origin.stopName = ? OR destination.stopName = ? OR j_flt.interStops LIKE ?)")
            params.extend([str(val), str(val), f'%"{val}"%'])
        elif col == 'stop_via':
            conds.append("j_flt.interStops LIKE ?")
            params.append(f'%"{val}"%')
        elif col in ('delay', 'arrivalDelay', 'departureDelay'):
            target_field = 'j_flt.departureDelay' if col == 'departureDelay' else 'j_flt.arrivalDelay'
            op = flt.get('op', '=')
            if op not in ('>', '>=', '=', '<=', '<'):
                op = '='
            try:
                num_val = int(round(float(val)))
                conds.append(f"(j_flt.realtimeAvailability = 1 AND {target_field} {op} ?)")
                params.append(num_val)
            except (ValueError, TypeError):
                pass
        elif col.startswith('geo_'):
            g_col = col.replace('geo_', '')
            if g_col in AVAILABLE_GEO_COLUMNS:
                conds.append(f"(origin.{g_col} = ? OR destination.{g_col} = ?)")
                params.extend([str(val), str(val)])
        elif col == 'delay_bucket':
            if val == 'early':
                conds.append("(j_flt.arrivalDelay < 0 AND j_flt.realtimeAvailability = 1)")
            elif val == 'ontime':
                conds.append("(j_flt.arrivalDelay = 0 AND j_flt.realtimeAvailability = 1)")
            elif val == 'del_3':
                conds.append("(j_flt.arrivalDelay >= 3)")
            elif val == 'del_6':
                conds.append("(j_flt.arrivalDelay >= 6)")
            elif val == 'del_15':
                conds.append("(j_flt.arrivalDelay >= 15)")
            elif val == 'del_60':
                conds.append("(j_flt.arrivalDelay >= 60)")

    query = f"""
        SELECT j_flt.statusId, j_flt.departureFinal, origin.timezone AS origin_tz, j_flt.durationReal
        FROM journeys j_flt
        LEFT JOIN stops origin ON j_flt.originId = origin.id
        LEFT JOIN stops destination ON j_flt.destinationId = destination.id
        WHERE {' AND '.join(conds)}
    """
    cursor = conn.cursor()
    cursor.execute(query, params)
    rows = cursor.fetchall()

    if not temporal_filters:
        return [r[0] for r in rows]

    # Zeitliche Bucket-Filterung unter Berücksichtigung lokaler Zeitzonen
    matched_ids = []
    weekdays = ['Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag', 'Sonntag']
    for sid, dep_utc, tz, dur in rows:
        if not dep_utc:
            continue
        loc_dt = format_utc_to_datatime(dep_utc, tz or 'Europe/Berlin')
        if not loc_dt:
            continue
        duration = int(dur) if dur and dur > 0 else 1
        # Prüfe alle 30-Min-Fragmente der Fahrt
        rem = duration
        cur = loc_dt
        buckets_for_trip = defaultdict(set)
        while rem > 0:
            start_m = (cur.minute // 30) * 30
            end_t = cur.replace(minute=start_m, second=0) + timedelta(minutes=30)
            buckets_for_trip['time_of_day'].add(f"{cur.hour:02d}:{start_m:02d} - {end_t.hour % 24:02d}:{end_t.minute:02d}")
            buckets_for_trip['weekday'].add(weekdays[cur.weekday()])
            buckets_for_trip['date'].add(cur.strftime('%Y-%m-%d'))
            buckets_for_trip['month_year'].add(cur.strftime('%Y-%m'))
            buckets_for_trip['month'].add(cur.strftime('%m'))
            buckets_for_trip['year'].add(cur.strftime('%Y'))
            step = min(rem, 30 - (cur.minute % 30))
            rem -= step
            cur += timedelta(minutes=step)

        if all(target_val in buckets_for_trip[t_type] for t_type, target_val in temporal_filters):
            matched_ids.append(sid)

    return matched_ids


@contextmanager
def interactive_filter_context(conn: sqlite3.Connection, usernames: list[str], start_date: str = '', end_date: str = '', extra_filters: list[dict] = None):
    prev_clause = getattr(_INTERACTIVE_FILTER_CTX, 'sql_clause', '')
    prev_stop_clause = getattr(_INTERACTIVE_FILTER_CTX, 'stop_clause', '')
    prev_stop_params = getattr(_INTERACTIVE_FILTER_CTX, 'stop_params', [])
    try:
        stop_conds = []
        stop_params = []
        if extra_filters:
            sids = resolve_filtered_status_ids(conn, usernames, start_date, end_date, extra_filters)
            if sids:
                sids_csv = ",".join(str(int(i)) for i in sids)
                _INTERACTIVE_FILTER_CTX.sql_clause = f" AND statusId IN ({sids_csv})"
            else:
                _INTERACTIVE_FILTER_CTX.sql_clause = " AND 1=0"

            for flt in extra_filters:
                col = flt.get('col', '')
                val = flt.get('val')
                if col.startswith('geo_'):
                    g_col = col.replace('geo_', '')
                    if g_col in AVAILABLE_GEO_COLUMNS and val is not None:
                        stop_conds.append(f"s.{g_col} = ?")
                        stop_params.append(str(val))
        else:
            _INTERACTIVE_FILTER_CTX.sql_clause = ""

        _INTERACTIVE_FILTER_CTX.stop_clause = (" AND " + " AND ".join(stop_conds)) if stop_conds else ""
        _INTERACTIVE_FILTER_CTX.stop_params = stop_params
        yield
    finally:
        _INTERACTIVE_FILTER_CTX.sql_clause = prev_clause
        _INTERACTIVE_FILTER_CTX.stop_clause = prev_stop_clause
        _INTERACTIVE_FILTER_CTX.stop_params = prev_stop_params


DEFAULT_EXCEL_SETTINGS = {
    'user_sheets': [
        {'sheet_name': 'Meta', 'type': 'meta', 'index': False, 'freeze_rows': 0, 'enabled': True},
        {'sheet_name': 'Fahrtenliste', 'type': 'journeys', 'index': False, 'freeze_rows': 1, 'enabled': True},
        {'sheet_name': 'Statistiken', 'type': 'stats', 'index': False, 'freeze_rows': 0, 'enabled': True},
        {'sheet_name': 'Haltestelle', 'type': 'stops', 'index': True, 'freeze_rows': 0, 'enabled': True},
        {'sheet_name': 'Haltestellen (mit Via)', 'type': 'stops_via', 'index': True, 'freeze_rows': 0, 'enabled': True},
        {'sheet_name': 'erweiterte Statistik', 'type': 'extended_stats', 'index': False, 'freeze_rows': 0, 'enabled': True},
        {'sheet_name': 'Kategorie', 'type': 'grouped', 'args': ['category'], 'sort_by_distance': True, 'exclude_empty': False, 'index': False, 'freeze_rows': 1, 'enabled': True},
        {'sheet_name': 'Betreiber', 'type': 'grouped', 'args': ['operatorName'], 'sort_by_distance': True, 'exclude_empty': False, 'index': False, 'freeze_rows': 1, 'enabled': True},
        {'sheet_name': 'Fahrtgrund', 'type': 'grouped', 'args': ['business'], 'sort_by_distance': True, 'exclude_empty': False, 'index': False, 'freeze_rows': 1, 'enabled': True},
        {'sheet_name': 'Linien', 'type': 'grouped', 'args': ['operatorName', 'lineName'], 'sort_by_distance': True, 'exclude_empty': False, 'index': False, 'freeze_rows': 2, 'enabled': True},
        {'sheet_name': 'Verbindungen', 'type': 'grouped', 'args': ['origin.stopName', 'destination.stopName'], 'sort_by_distance': True, 'exclude_empty': False, 'index': False, 'freeze_rows': 2, 'enabled': True},
        {'sheet_name': 'Baureihe', 'type': 'grouped', 'args': ['locomotiveClass'], 'sort_by_distance': True, 'exclude_empty': False, 'index': False, 'freeze_rows': 1, 'enabled': True},
        {'sheet_name': 'Fahrzeugnummer', 'type': 'grouped', 'args': ['vehicleNumber'], 'sort_by_distance': True, 'exclude_empty': False, 'index': False, 'freeze_rows': 1, 'enabled': True},
        {'sheet_name': 'Check-In-Tool', 'type': 'grouped', 'args': ['clientName'], 'sort_by_distance': True, 'exclude_empty': False, 'index': False, 'freeze_rows': 1, 'enabled': True},
        {'sheet_name': 'Tageszeiten', 'type': 'temporal', 'grouping_type': 'time_of_day', 'index': False, 'freeze_rows': 1, 'enabled': True},
        {'sheet_name': 'Wochentage', 'type': 'temporal', 'grouping_type': 'weekday', 'index': False, 'freeze_rows': 1, 'enabled': True},
        {'sheet_name': 'Monat', 'type': 'temporal', 'grouping_type': 'month', 'index': False, 'freeze_rows': 1, 'enabled': True},
        {'sheet_name': 'Monate im Jahr', 'type': 'temporal', 'grouping_type': 'month_year', 'index': False, 'freeze_rows': 1, 'enabled': True},
        {'sheet_name': 'Tag', 'type': 'temporal', 'grouping_type': 'date', 'index': False, 'freeze_rows': 1, 'enabled': True},
        {'sheet_name': 'Jahr', 'type': 'temporal', 'grouping_type': 'year', 'index': False, 'freeze_rows': 1, 'enabled': True},
        {'sheet_name': 'Länder', 'type': 'geo', 'geo_column': 'countryName', 'index': True, 'freeze_rows': 0, 'enabled': True},
        {'sheet_name': 'Städte', 'type': 'geo', 'geo_column': 'city', 'index': True, 'freeze_rows': 0, 'enabled': True},
        {'sheet_name': 'Regionen_Bundesländer', 'type': 'geo', 'geo_column': 'admin1State', 'index': True, 'freeze_rows': 0, 'enabled': True},
        {'sheet_name': 'Bezirke_Kreise', 'type': 'geo', 'geo_column': 'admin2Dist', 'index': True, 'freeze_rows': 0, 'enabled': True},
    ],
    'shared_sheets': [
        {'sheet_name': 'Meta', 'type': 'shared_meta', 'freeze_cols': 0, 'enabled': True},
        {'sheet_name': 'Statistik', 'type': 'shared_stats', 'freeze_cols': 1, 'enabled': True},
        {'sheet_name': 'Haltestellen', 'type': 'shared_stops', 'freeze_cols': 2, 'enabled': True},
        {'sheet_name': 'Haltestellen (filterbar)', 'type': 'shared_stops_filterable', 'freeze_cols': 1, 'enabled': True},
        {'sheet_name': 'Gemeinsame Fahrten', 'type': 'shared_journeys', 'freeze_cols': 3, 'enabled': True},
        {'sheet_name': 'Kategorie', 'type': 'shared_grouped', 'args': ['category'], 'min_users': 2, 'sort_by_distance': True, 'exclude_empty': False, 'freeze_cols': 2, 'enabled': True},
        {'sheet_name': 'Betreiber', 'type': 'shared_grouped', 'args': ['operatorName'], 'min_users': 2, 'sort_by_distance': True, 'exclude_empty': False, 'freeze_cols': 2, 'enabled': True},
        {'sheet_name': 'Linie', 'type': 'shared_grouped', 'args': ['lineName'], 'min_users': 2, 'sort_by_distance': True, 'exclude_empty': False, 'freeze_cols': 2, 'enabled': True},
        {'sheet_name': 'Fahrzeugnummer', 'type': 'shared_grouped', 'args': ['vehicleNumber'], 'min_users': 2, 'sort_by_distance': True, 'exclude_empty': False, 'freeze_cols': 2, 'enabled': True},
        {'sheet_name': 'Baureihe', 'type': 'shared_grouped', 'args': ['locomotiveClass'], 'min_users': 2, 'sort_by_distance': True, 'exclude_empty': False, 'freeze_cols': 2, 'enabled': True},
        {'sheet_name': 'Check-In-Tool', 'type': 'shared_grouped', 'args': ['clientName'], 'min_users': 2, 'sort_by_distance': True, 'exclude_empty': False, 'freeze_cols': 2, 'enabled': True},
    ]
}


def get_default_excel_settings() -> dict:
    return json.loads(json.dumps(DEFAULT_EXCEL_SETTINGS))


def get_user_excel_settings(conn: sqlite3.Connection, username: str) -> dict:
    if not username:
        return get_default_excel_settings()
    cursor = conn.cursor()
    cursor.execute("SELECT excel_settings FROM user_login_data WHERE owner = ?", (username,))
    row = cursor.fetchone()
    if row and row[0]:
        try:
            parsed = json.loads(row[0])
            if isinstance(parsed, dict):
                defaults = get_default_excel_settings()
                if 'user_sheets' not in parsed or not isinstance(parsed['user_sheets'], list):
                    parsed['user_sheets'] = defaults['user_sheets']
                if 'shared_sheets' not in parsed or not isinstance(parsed['shared_sheets'], list):
                    parsed['shared_sheets'] = defaults['shared_sheets']
                return parsed
        except Exception:
            pass
    return get_default_excel_settings()


def save_user_excel_settings(conn: sqlite3.Connection, username: str, settings_dict: dict) -> None:
    cursor = conn.cursor()
    settings_json = json.dumps(settings_dict, ensure_ascii=False)
    cursor.execute(
        """INSERT INTO user_login_data (owner, excel_settings, last_updated)
           VALUES (?, ?, ?)
           ON CONFLICT(owner) DO UPDATE SET
               excel_settings = excluded.excel_settings""",
        (username, settings_json, datetime.now().isoformat())
    )
    conn.commit()


def sanitize_sheet_name(raw_name: str, used_names: set) -> str:
    """Stellt sicher, dass Excel-Tabellenblattnamen <= 31 Zeichen, gültig und eindeutig sind."""
    clean = re.sub(r'[\[\]\:\*\?\/\\]', '_', str(raw_name or 'Sheet')).strip()
    if not clean:
        clean = 'Sheet'
    base = clean[:31]
    candidate = base
    counter = 2
    while candidate.lower() in used_names:
        suffix = f"_{counter}"
        candidate = base[:31 - len(suffix)] + suffix
        counter += 1
    used_names.add(candidate.lower())
    return candidate


def create_geo_dataframe(conn: sqlite3.Connection, username: str, geo_column: str, start_date: str = '', end_date: str = '') -> pd.DataFrame:
    if geo_column not in AVAILABLE_GEO_COLUMNS:
        geo_column = 'countryName'
    params, date_query = params_and_date_query(username, start_date, end_date)
    beautiful_name = AVAILABLE_GEO_COLUMNS.get(geo_column, geo_column)
    stop_clause = getattr(_INTERACTIVE_FILTER_CTX, 'stop_clause', '')
    stop_params = getattr(_INTERACTIVE_FILTER_CTX, 'stop_params', [])
    
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
    WHERE s.{geo_column} IS NOT NULL AND s.{geo_column} != '' {stop_clause}
    GROUP BY s.{geo_column}
    ORDER BY Besuche DESC
    """
    df = pd.read_sql_query(query, conn, params=params + stop_params)
    if not df.empty:
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
    stop_clause = getattr(_INTERACTIVE_FILTER_CTX, 'stop_clause', '')
    stop_params = getattr(_INTERACTIVE_FILTER_CTX, 'stop_params', [])
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
            WHERE 1=1 {stop_clause}
            GROUP BY s.stopName
            ORDER BY Besuche DESC
        """
    df = pd.read_sql_query(query, conn, params=params + stop_params)

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

    stop_clause = getattr(_INTERACTIVE_FILTER_CTX, 'stop_clause', '')
    stop_params = getattr(_INTERACTIVE_FILTER_CTX, 'stop_params', [])
    if stop_clause:
        cursor.execute(f"SELECT stopName FROM stops s WHERE 1=1 {stop_clause}", stop_params)
        allowed_stop_names = {r[0] for r in cursor.fetchall() if r[0]}
        inter_stops_list = [s for s in inter_stops_list if s in allowed_stop_names]
        inter_stops_last_visit = {s: t for s, t in inter_stops_last_visit.items() if s in allowed_stop_names}
        if not base_df.empty and 'Haltestelle' in base_df.columns:
            base_df = base_df[base_df['Haltestelle'].isin(allowed_stop_names)]

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

def create_user_excel_file(conn: sqlite3.Connection, username: str, start_date: str = '', end_date: str = '',
                           create_picture: bool = True, custom_settings: dict = None,
                           creator_username: str | None = None) -> str:
    """Erstellt das persönliche Excel-Fahrtenbuch dynamisch anhand der gespeicherten JSON-Konfiguration."""
    kwargs = dict(conn=conn, username=username, start_date=start_date, end_date=end_date)

    if custom_settings is None:
        settings_owner = creator_username or get_auth_user() or username
        custom_settings = get_user_excel_settings(conn, settings_owner)

    user_sheets = [s for s in custom_settings.get('user_sheets', []) if s.get('enabled', True) is not False]
    if not user_sheets:
        user_sheets = get_default_excel_settings()['user_sheets']

    needs_stops = any(s.get('type') in ('stops', 'stops_via', 'extended_stats') for s in user_sheets)
    needs_temporal = any(s.get('type') == 'temporal' for s in user_sheets)

    stop_dataframe = create_user_stops_dataframe(**kwargs) if needs_stops else None
    temporal_base_df = get_temporal_base_dataframe(**kwargs, create_picture=create_picture) if needs_temporal else None

    all_dataframes = []
    used_sheet_names = set()

    for sheet_cfg in user_sheets:
        stype = sheet_cfg.get('type', 'grouped')
        sname = sanitize_sheet_name(sheet_cfg.get('sheet_name', stype), used_sheet_names)
        use_index = bool(sheet_cfg.get('index', False))
        freeze_rows = int(sheet_cfg.get('freeze_rows', 0) or 0)
        df = None

        if stype == 'meta':
            df = create_user_meta_dataframe(conn, username)
        elif stype == 'journeys':
            df = create_journeys_dataframe(**kwargs)
        elif stype == 'stats':
            df = create_user_stats_dataframe(**kwargs)
        elif stype == 'stops':
            df = stop_dataframe.copy() if stop_dataframe is not None else create_user_stops_dataframe(**kwargs)
        elif stype == 'stops_via':
            df = create_user_stops_dataframe_with_inter(**kwargs, base_dataframe=stop_dataframe)
        elif stype == 'extended_stats':
            df = create_user_extended_stats(**kwargs, base_dataframe=stop_dataframe)
        elif stype == 'grouped':
            args = [a for a in sheet_cfg.get('args', ['category']) if a]
            if not args:
                args = ['category']
            df = create_user_grouped_frame(
                **kwargs,
                first_argument=args[0],
                second_argument=args[1] if len(args) > 1 else "",
                third_argument=args[2] if len(args) > 2 else "",
                fourth_argument=args[3] if len(args) > 3 else "",
                sort_by_distance=sheet_cfg.get('sort_by_distance', True),
                exclude_empty=sheet_cfg.get('exclude_empty', False)
            )
        elif stype == 'temporal':
            g_type = sheet_cfg.get('grouping_type', 'time_of_day')
            df = create_temporal_stats_dataframe(temporal_base_df, grouping_type=g_type, username=username)
        elif stype == 'geo':
            g_col = sheet_cfg.get('geo_column', 'countryName')
            df = create_geo_dataframe(conn, username, g_col, start_date, end_date)

        if df is not None:
            num_args = len([a for a in sheet_cfg.get('args', ['category']) if a]) if stype == 'grouped' else 1
            df = apply_sheet_metrics(df, stype, sheet_cfg.get('metrics'), num_group_args=num_args)
            all_dataframes.append((df, sname, use_index, freeze_rows))

    output_dir = get_export_dir('user', creator_username=creator_username)
    
    date_suffix = get_date_suffix(start_date, end_date)
    file_path = output_dir / f"{username}{date_suffix}_data.xlsx"
    
    print(f"[EXCEL] Speichere Datei: {file_path.name}")
    
    with pd.ExcelWriter(file_path, engine='xlsxwriter', datetime_format='dd.mm.yyyy hh:mm') as writer:
        for df, sheet_name, index, numb_of_freeze_row in all_dataframes:
            if index and df.index.name is not None:
                df = df.reset_index()
                index = False
                if numb_of_freeze_row == 0:
                    numb_of_freeze_row = 1
            
            df = df.loc[:, ~df.columns.duplicated()]
            df.to_excel(writer, sheet_name=sheet_name, index=index, header=True)
            current_worksheet = writer.sheets[sheet_name]

            if numb_of_freeze_row > 0:
                current_worksheet.freeze_panes(numb_of_freeze_row, 0)
            if not df.empty:
                max_row = len(df)
                max_col = len(df.columns) - 1 + (1 if index else 0)
                current_worksheet.autofilter(0, 0, max_row, max_col)

            for runner, col in enumerate(df.columns):
                series = df.iloc[:, runner].dropna()
                max_length = int(series.map(lambda x: len(str(x))).max()) if not series.empty else 0
                final_width = max(max_length, len(str(col)) + 2) + 2
                col_idx = (1 if index else 0) + runner
                current_worksheet.set_column(col_idx, col_idx, final_width)

    print(f'[ERFOLG] Excel gespeichert: {file_path.name}')
    return file_path.name


def create_user_extended_stats(conn: sqlite3.Connection, username: str, start_date: str = '',
                               end_date: str = '', base_dataframe: pd.DataFrame = None) -> pd.DataFrame:
    base_df = base_dataframe.copy() if base_dataframe is not None and not base_dataframe.empty else create_user_stops_dataframe(conn, username, start_date, end_date)
    actionradius = 0.0
    if not base_df.empty and len(base_df) > 1 and 'Besuche' in base_df.columns and 'Hub-Distanz-Wert' in base_df.columns:
        besuche_sum = float(base_df['Besuche'].iloc[1:].sum())
        if besuche_sum > 0:
            hub_sum = float(base_df['Hub-Distanz-Wert'].iloc[1:].sum())
            actionradius = round(((hub_sum / besuche_sum) ** 2) * 1000.0, 3)

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
                max_dist_val = round(float(dist.iloc[max_idx]), 2)

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
                              end_date: str = '', sort_by_distance: bool = True, exclude_empty: bool = False) -> pd.DataFrame:
    params, date_query = params_and_date_query(username, start_date, end_date, "j")

    raw_args = [a for a in [first_argument, second_argument, third_argument, fourth_argument] if a]
    valid_args = list(dict.fromkeys(raw_args))
    if not valid_args:
        valid_args = ['category']

    select_parts = []
    group_parts = []
    empty_filters = []
    used_labels = set()

    for arg in valid_args:
        col_meta = AVAILABLE_GROUP_COLUMNS.get(arg, {'sql': f"j.{arg}" if '.' not in arg else arg, 'label': arg})
        sql_expr = col_meta['sql']
        label = col_meta['label']
        while label in used_labels:
            label += "_"
        used_labels.add(label)

        select_parts.append(f"{sql_expr} AS '{label}'")
        group_parts.append(sql_expr)
        if exclude_empty:
            empty_filters.append(f"AND {sql_expr} IS NOT NULL AND CAST({sql_expr} AS TEXT) != '' AND CAST({sql_expr} AS TEXT) != 'N/A'")

    select_query = ",\n".join(select_parts)
    group_query = "GROUP BY " + ", ".join(group_parts)
    extra_where = " ".join(empty_filters)

    query = f""" SELECT {select_query},
        COUNT(j.statusId) AS 'Fahrten',
        COUNT(DISTINCT date(COALESCE(j.arrivalFinal, j.departureFinal, j.departurePlanned))) AS 'Tage',
        COUNT(DISTINCT j.operatorName || j.lineName) AS 'Linien',
        COUNT(DISTINCT j.operatorName) AS 'Betreiber',
        (COUNT(DISTINCT j.originId) + COUNT(DISTINCT CASE WHEN j.destinationId != j.originId THEN j.destinationId END)) AS 'Haltestellen',
        SUM(j.distance) AS 'Distanz (m)',
        ROUND(SUM(j.distance) / 1000.0, 2) AS 'Distanz (km)',
        ROUND(SUM(j.distanceStraightLine) / 1000.0, 2) AS 'Luftlinie Fahrten (km)',
        ROUND(SUM(j.distanceToPrevJourney) / 1000.0, 2) AS 'Luftlinie zwischen Fahrten (km)',
        ROUND(CASE WHEN SUM(j.distance) > 0 THEN (SUM(j.distanceToPrevJourney) * 100.0 / SUM(j.distance)) ELSE 0 END, 2) AS 'Tracking-Lücke (%)',
        ROUND(SUM(j.distance) * 1.0 / NULLIF(SUM(j.distanceStraightLine), 0), 2) AS 'Umweg Faktor',
        SUM(j.durationReal) AS 'Zeit (min)',
        SUM(CASE WHEN j.realtimeAvailability = 1 THEN j.departureDelay ELSE 0 END) AS 'Abfahrtsverspätung (min)',
        SUM(CASE WHEN j.realtimeAvailability = 1 THEN j.arrivalDelay ELSE 0 END) AS 'Ankunftsverspätung (min)',
        ROUND(AVG(j.departureDelay), 2) AS '∅ Abfahrtsverspätung (min)',
        ROUND(AVG(j.arrivalDelay), 2) AS '∅ Verspätung (min)',
        ROUND((SUM(CASE WHEN j.realtimeAvailability = 1 THEN 1 ELSE 0 END) * 100.0) / COUNT(j.statusId), 2) AS 'Echtzeitquote (%)',
        ROUND((SUM(CASE WHEN j.arrivalDelay < 0 AND j.realtimeAvailability = 1 THEN 1 ELSE 0 END) * 100.0) / NULLIF(SUM(CASE WHEN j.realtimeAvailability = 1 THEN 1 ELSE 0 END), 0), 2) AS 'Verfrüht (%)',
        ROUND((SUM(CASE WHEN j.arrivalDelay = 0 AND j.realtimeAvailability = 1 THEN 1 ELSE 0 END) * 100.0) / NULLIF(SUM(CASE WHEN j.realtimeAvailability = 1 THEN 1 ELSE 0 END), 0), 2) AS 'Pünktlich (%)',
        ROUND((SUM(CASE WHEN j.arrivalDelay >= 3 THEN 1 ELSE 0 END) * 100.0) / COUNT(j.statusId), 2) AS '3 Minuten und mehr (%)',
        ROUND((SUM(CASE WHEN j.arrivalDelay >= 6 THEN 1 ELSE 0 END) * 100.0) / COUNT(j.statusId), 2) AS '6 Minuten und mehr (%)',
        ROUND((SUM(CASE WHEN j.arrivalDelay >= 15 THEN 1 ELSE 0 END) * 100.0) / COUNT(j.statusId), 2) AS '15 Minuten und mehr (%)',
        ROUND((SUM(CASE WHEN j.arrivalDelay >= 60 THEN 1 ELSE 0 END) * 100.0) / COUNT(j.statusId), 2) AS '60 Minuten und mehr (%)',
        SUM(j.points) AS 'Punkte',
        SUM(j.likes) AS 'Likes',
        ROUND(AVG(j.interStopsNumber), 2) AS '∅ Zwischenhalte'
        FROM journeys j
        LEFT JOIN stops origin ON j.originId = origin.id
        LEFT JOIN stops destination ON j.destinationId = destination.id
        {date_query} {extra_where} {group_query}
        {"ORDER BY [Distanz (km)] DESC, [Fahrten] DESC" if sort_by_distance else "ORDER BY [Fahrten] DESC, [Distanz (km)] DESC"}"""

    df = pd.read_sql_query(query, conn, params=params)
    if df.empty:
        dist_share = pd.Series(dtype='float64')
        zeit_hh_mm = pd.Series(dtype='object')
        time_share = pd.Series(dtype='float64')
        speed_col = pd.Series(dtype='float64')
    else:
        total_distance = df['Distanz (km)'].sum()
        total_time = df['Zeit (min)'].sum()

        dist_share = ((df['Distanz (km)'] / total_distance) * 100).round(2) if total_distance > 0 else 0.0
        time_share = ((df['Zeit (min)'] / total_time) * 100).round(2) if total_time > 0 else 0.0
        speed_col = (df['Distanz (km)'] / (df['Zeit (min)'].replace(0, np.nan) / 60)).round(2)
        zeit_hh_mm = df['Zeit (min)'].apply(lambda x: f"{int(x // 60):02d}:{int(x % 60):02d}" if pd.notna(x) else "00:00")

    if 'Distanz (km)' in df.columns:
        df.insert(loc=df.columns.get_loc('Distanz (km)') + 1, column='Anteil an Gesamtdistanz (%)', value=dist_share)
    else:
        df['Anteil an Gesamtdistanz (%)'] = dist_share

    if 'Zeit (min)' in df.columns:
        df.insert(loc=df.columns.get_loc('Zeit (min)') + 1, column='Zeit (hh:mm)', value=zeit_hh_mm)
    else:
        df['Zeit (hh:mm)'] = zeit_hh_mm

    if 'Zeit (hh:mm)' in df.columns:
        df.insert(loc=df.columns.get_loc('Zeit (hh:mm)') + 1, column='Anteil an Gesamtzeit (%)', value=time_share)
    else:
        df['Anteil an Gesamtzeit (%)'] = time_share

    if 'Anteil an Gesamtzeit (%)' in df.columns:
        df.insert(loc=df.columns.get_loc('Anteil an Gesamtzeit (%)') + 1, column='∅ Geschwindigkeit', value=speed_col)
    else:
        df['∅ Geschwindigkeit'] = speed_col

    return df


def get_temporal_base_dataframe(conn: sqlite3.Connection, username: str, start_date: str = '', end_date: str = '',
                                create_picture: bool = True) -> pd.DataFrame:
    params, date_query = params_and_date_query(username, start_date, end_date, "j")
    query = f"""
    SELECT 
        j.statusId,
        j.departureFinal,
        origin.timezone AS origin_tz,
        j.durationReal,
        j.distance,
        j.distanceStraightLine,
        j.distanceToPrevJourney,
        j.departureDelay,
        j.arrivalDelay,
        j.realtimeAvailability,
        j.operatorName,
        j.lineName,
        j.originId,
        j.destinationId,
        j.points,
        j.likes,
        j.interStopsNumber
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
    trip_meta = {}

    for row in df_raw.itertuples(index=False):
        if pd.isna(row.departure_local):
            continue
        dep = row.departure_local
        duration = int(row.durationReal) if pd.notna(row.durationReal) else 0
        if duration <= 0:
            continue

        status_id = row.statusId
        dist = float(row.distance) if pd.notna(row.distance) else 0.0
        straight_dist = float(getattr(row, 'distanceStraightLine', 0.0)) if pd.notna(getattr(row, 'distanceStraightLine', None)) else 0.0
        prev_dist = float(getattr(row, 'distanceToPrevJourney', 0.0)) if pd.notna(getattr(row, 'distanceToPrevJourney', None)) else 0.0
        dep_delay = float(getattr(row, 'departureDelay', 0.0)) if pd.notna(getattr(row, 'departureDelay', None)) else 0.0
        arr_delay = float(row.arrivalDelay) if pd.notna(row.arrivalDelay) else 0.0
        rt_avail = bool(getattr(row, 'realtimeAvailability', False))
        op_name = str(getattr(row, 'operatorName', '') or '')
        line_name = str(getattr(row, 'lineName', '') or '')
        orig_id = getattr(row, 'originId', None)
        dest_id = getattr(row, 'destinationId', None)
        pts = float(getattr(row, 'points', 0.0)) if pd.notna(getattr(row, 'points', None)) else 0.0
        lks = int(getattr(row, 'likes', 0)) if pd.notna(getattr(row, 'likes', None)) else 0
        inter_stops = float(getattr(row, 'interStopsNumber', np.nan)) if pd.notna(getattr(row, 'interStopsNumber', None)) else np.nan

        trip_meta[status_id] = {
            'dep_delay': dep_delay,
            'arr_delay': arr_delay,
            'rt_avail': rt_avail,
            'op_name': op_name,
            'line_key': f"{op_name}||{line_name}",
            'orig_id': orig_id,
            'dest_id': dest_id,
            'likes': lks,
            'inter_stops': inter_stops
        }

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

            share = mins_to_allocate / duration
            fragment = {
                'statusId': status_id,
                'frag_date': current_time.date(),
                'Minuten': mins_to_allocate,
                'Distanz_m': share * dist,
                'Straight_m': share * straight_dist,
                'Prev_m': share * prev_dist,
                'Points_share': share * pts
            }

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

    if not fragments:
        return pd.DataFrame()

    df_frag = pd.DataFrame(fragments)
    total_mins_all = df_frag['Minuten'].sum()
    total_dist_all = df_frag['Distanz_m'].sum()
    total_active_days = len(unique_dates) if len(unique_dates) > 0 else 1

    name_mapping = {
        'time_of_day': 'Zeitfenster', 'weekday': 'Wochentag', 'date': 'Datum',
        'month_year': 'Monat (JJJJ-MM)', 'month': 'Monat', 'year': 'Jahr'
    }
    bucket_col_name = name_mapping.get(grouping_type, 'Bucket')

    rows_out = []
    for bucket_val, grp in df_frag.groupby('Bucket'):
        sids = grp['statusId'].unique()
        n_trips = len(sids)
        n_days = grp['frag_date'].nunique()

        mins_sum = grp['Minuten'].sum()
        dist_m_sum = grp['Distanz_m'].sum()
        dist_km_sum = dist_m_sum / 1000.0
        straight_km_sum = grp['Straight_m'].sum() / 1000.0
        prev_km_sum = grp['Prev_m'].sum() / 1000.0
        pts_sum = int(round(grp['Points_share'].sum()))

        lines_set = set()
        ops_set = set()
        stops_set = set()
        rt_count = 0
        dep_delay_rt_sum = 0.0
        arr_delay_rt_sum = 0.0
        dep_delays_all = []
        arr_delays_all = []
        early_cnt = 0
        ontime_cnt = 0
        del3_cnt = 0
        del6_cnt = 0
        del15_cnt = 0
        del60_cnt = 0
        likes_sum = 0
        inter_list = []

        for sid in sids:
            m = trip_meta.get(sid)
            if not m:
                continue
            if m['line_key']:
                lines_set.add(m['line_key'])
            if m['op_name']:
                ops_set.add(m['op_name'])
            if m['orig_id'] is not None:
                stops_set.add(m['orig_id'])
            if m['dest_id'] is not None:
                stops_set.add(m['dest_id'])

            likes_sum += m['likes']
            if pd.notna(m['inter_stops']):
                inter_list.append(m['inter_stops'])

            ad = m['arr_delay']
            dd = m['dep_delay']
            dep_delays_all.append(dd)
            arr_delays_all.append(ad)

            if m['rt_avail']:
                rt_count += 1
                dep_delay_rt_sum += dd
                arr_delay_rt_sum += ad
                if ad < 0:
                    early_cnt += 1
                elif ad == 0:
                    ontime_cnt += 1

            if ad >= 3:
                del3_cnt += 1
            if ad >= 6:
                del6_cnt += 1
            if ad >= 15:
                del15_cnt += 1
            if ad >= 60:
                del60_cnt += 1

        j_div = n_trips if n_trips > 0 else 1
        rt_div = rt_count if rt_count > 0 else 1

        non_zero_arr = [d for d in arr_delays_all if d != 0.0]
        avg_arr_delay = round(float(np.mean(non_zero_arr)), 2) if non_zero_arr else 0.0
        avg_dep_delay = round(float(np.mean(dep_delays_all)), 2) if dep_delays_all else 0.0

        mins_int = int(round(mins_sum))
        zeit_hh_mm = f"{mins_int // 60:02d}:{mins_int % 60:02d}"

        rows_out.append({
            bucket_col_name: bucket_val,
            'Fahrten': n_trips,
            'Tage': n_days,
            'Linien': len(lines_set),
            'Betreiber': len(ops_set),
            'Haltestellen': len(stops_set),
            'Zeit (min)': round(mins_sum, 2),
            'Zeit (hh:mm)': zeit_hh_mm,
            'Anteil Zeit (%)': round((mins_sum / total_mins_all) * 100, 2) if total_mins_all > 0 else 0.0,
            'Distanz (km)': round(dist_km_sum, 2),
            'Distanz (m)': int(round(dist_m_sum)),
            'Anteil Distanz (%)': round((dist_m_sum / total_dist_all) * 100, 2) if total_dist_all > 0 else 0.0,
            'Luftlinie Fahrten (km)': round(straight_km_sum, 2),
            'Luftlinie zwischen Fahrten (km)': round(prev_km_sum, 2),
            'Tracking-Lücke (%)': round((prev_km_sum / dist_km_sum) * 100, 2) if dist_km_sum > 0 else 0.0,
            'Umweg Faktor': round(dist_km_sum / straight_km_sum, 2) if straight_km_sum > 0 else 0.0,
            '∅ Geschwindigkeit (km/h)': round(dist_km_sum / (mins_sum / 60.0), 2) if mins_sum > 0 else 0.0,
            '∅ Minuten pro aktivem Tag': round(mins_sum / total_active_days, 2),
            '∅ Distanz pro aktivem Tag (km)': round(dist_km_sum / total_active_days, 2),
            'Abfahrtsverspätung (min)': int(round(dep_delay_rt_sum)),
            'Ankunftsverspätung (min)': int(round(arr_delay_rt_sum)),
            '∅ Abfahrtsverspätung (min)': avg_dep_delay,
            '∅ Ankunftsverspätung (min)': avg_arr_delay,
            'Echtzeitquote (%)': round((rt_count / j_div) * 100, 2),
            'Verfrüht (%)': round((early_cnt / rt_div) * 100, 2) if rt_count > 0 else 0.0,
            'Pünktlich (%)': round((ontime_cnt / rt_div) * 100, 2) if rt_count > 0 else 0.0,
            '3 Minuten und mehr (%)': round((del3_cnt / j_div) * 100, 2),
            '6 Minuten und mehr (%)': round((del6_cnt / j_div) * 100, 2),
            '15 Minuten und mehr (%)': round((del15_cnt / j_div) * 100, 2),
            '60 Minuten und mehr (%)': round((del60_cnt / j_div) * 100, 2),
            'Punkte': pts_sum,
            'Likes': likes_sum,
            '∅ Zwischenhalte': round(float(np.mean(inter_list)), 2) if inter_list else 0.0
        })

    grouped = pd.DataFrame(rows_out)

    if grouping_type == 'weekday':
        weekdays_order = ['Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag', 'Sonntag']
        grouped['Wochentag'] = pd.Categorical(grouped['Wochentag'], categories=weekdays_order, ordered=True)
        grouped = grouped.sort_values('Wochentag').reset_index(drop=True)
    else:
        grouped = grouped.sort_values(bucket_col_name).reset_index(drop=True)

    return grouped


# ==========================================
# PLOTS UND KARTEN (VISUALISIERUNGEN)
# ==========================================

def create_ganglinie_plot(username: str, grouping_type: str, conn: sqlite3.Connection = None,
                          start_date: str = '', end_date: str = '', given_df: pd.DataFrame = None,
                          creator_username: str | None = None) -> str | None:
    if conn is None:
        conn = init_db()
    if given_df is None or given_df.empty:
        temporal_base_df = get_temporal_base_dataframe(conn=conn, username=username, start_date=start_date, end_date=end_date)
        df = create_temporal_stats_dataframe(temporal_base_df, grouping_type=grouping_type)
    else: 
        df = given_df.copy()

    if df.empty:
        return None

    export_dir = get_export_dir('gang_linien', creator_username=creator_username)

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
                                          start_date: str = '', end_date: str = '',
                                          creator_username: str | None = None) -> list[str]:
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

        export_dir = get_export_dir('verspaetungs_plots', creator_username=creator_username)

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
                    'User-Agent': 'TraewellingAnalysis_v10 tee_beutel'
                })
                response.raise_for_status()
                set_remaining(response.headers.get('x-ratelimit-remaining'))

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
                cd_msg = f'   🗺️ Polylines-Countdown: Noch {rem_lines} von {total} Linien verbleibend... (Limit: {get_remaining()})'
                print(cd_msg)
                safe_log(log_callback, cd_msg, replace=(current_index > 50))
            except Exception as e:
                if isinstance(e, requests.exceptions.HTTPError) and e.response.status_code == 429:
                    sleep(30)
                else:
                    log_error_to_db(username=None, error=e, context=f"fetch_polylines ({url})", conn=conn)
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
                     end_date: str = '', log_callback=None, lite_version: bool = False,
                     creator_username: str | None = None) -> str | None:
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

    export_dir = get_export_dir('polylines', creator_username=creator_username)

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
                    end_date: str = '', log_callback=None, lite_version: bool = False,
                    creator_username: str | None = None) -> str | None:
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
        export_dir = get_export_dir('polylines', creator_username=creator_username)

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
                    log_callback=None, creator_username: str | None = None) -> str | None:
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

    export_dir = get_export_dir('timelines', creator_username=creator_username)

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

    stop_clause = getattr(_INTERACTIVE_FILTER_CTX, 'stop_clause', '')
    stop_params = getattr(_INTERACTIVE_FILTER_CTX, 'stop_params', [])

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
    WHERE 1=1 {stop_clause}
    GROUP BY v.username, s.stopName
    """
    cursor.execute(query_ohne_via, base_params + stop_params)
    for row in cursor.fetchall():
        uname, stop_name, besuche = row
        user_stop_data[uname][stop_name]['Ohne Via'] = besuche

    allowed_stop_names = None
    if stop_clause:
        cursor.execute(f"SELECT stopName FROM stops s WHERE 1=1 {stop_clause}", stop_params)
        allowed_stop_names = {r[0] for r in cursor.fetchall() if r[0]}

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
                if allowed_stop_names is not None and stop not in allowed_stop_names:
                    continue
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


def create_shared_grouped_dataframe(conn: sqlite3.Connection, usernames: list[str], column_name: str | list[str],
                                    start_date: str = '', end_date: str = '', min_users: int = 2,
                                    sort_by_distance: bool = True, exclude_empty: bool = False) -> tuple[pd.DataFrame, str]:
    raw_args = [column_name] if isinstance(column_name, str) else list(column_name or ['category'])
    valid_args = list(dict.fromkeys([a for a in raw_args if a]))
    if not valid_args:
        valid_args = ['category']

    select_parts = []
    group_sql_parts = []
    label_cols = []
    not_null_parts = []

    for arg in valid_args:
        col_meta = AVAILABLE_GROUP_COLUMNS.get(arg, {'sql': f"j.{arg}" if '.' not in arg else arg, 'label': arg})
        sql_expr = col_meta['sql']
        lbl = col_meta['label']
        while lbl in label_cols:
            lbl += "_"
        label_cols.append(lbl)
        select_parts.append(f"{sql_expr} AS '{lbl}'")
        group_sql_parts.append(sql_expr)
        not_null_parts.append(f"AND {sql_expr} IS NOT NULL AND CAST({sql_expr} AS TEXT) != '' AND CAST({sql_expr} AS TEXT) != 'N/A'")
        if arg == 'operatorName':
            not_null_parts.append("AND j.operatorName != 'Unknown Operator'")
        elif arg == 'clientName':
            not_null_parts.append("AND j.clientName != 'Unknown Client'")

    default_sheet_title = " & ".join(label_cols)[:31]
    placeholders = ','.join(['?'] * len(usernames))

    date_query_j = ""
    params_date = []
    if start_date:
        date_query_j += " AND j.arrivalFinal > ?"
        params_date.append(start_date)
    if end_date:
        date_query_j += " AND j.departureFinal < ?"
        end_date_val = end_date + 'T23:59:59' if len(end_date) == 10 else end_date
        params_date.append(end_date_val)

    combined_key_sql = " || '|||' || ".join([f"COALESCE(CAST({s} AS TEXT), '')" for s in group_sql_parts])
    group_by_sql = ", ".join(group_sql_parts)
    not_null_sql = " ".join(not_null_parts)
    order_metric_sql = "SUM(j.distance) DESC" if sort_by_distance else "COUNT(j.statusId) DESC"

    if min_users and min_users > 1:
        subquery_sql = f"""
          AND ({combined_key_sql}) IN (
              SELECT ({combined_key_sql})
              FROM journeys j
              LEFT JOIN stops origin ON j.originId = origin.id
              LEFT JOIN stops destination ON j.destinationId = destination.id
              WHERE j.username IN ({placeholders}) {date_query_j} {not_null_sql} {get_shared_visibility_query('j')}
              GROUP BY {group_by_sql}
              HAVING COUNT(DISTINCT j.username) >= {int(min_users)}
          )
        """
        params = usernames + params_date + usernames + params_date
    else:
        subquery_sql = ""
        params = usernames + params_date

    query = f"""
    SELECT 
        {", ".join(select_parts)},
        j.username AS User,
        COUNT(j.statusId) AS 'Fahrten',
        COUNT(DISTINCT date(COALESCE(j.arrivalFinal, j.departureFinal, j.departurePlanned))) AS 'Tage',
        COUNT(DISTINCT j.operatorName || j.lineName) AS 'Linien',
        COUNT(DISTINCT j.operatorName) AS 'Betreiber',
        (COUNT(DISTINCT j.originId) + COUNT(DISTINCT CASE WHEN j.destinationId != j.originId THEN j.destinationId END)) AS 'Haltestellen',
        SUM(j.distance) AS 'Distanz (m)',
        ROUND(SUM(j.distance) / 1000.0, 2) AS 'Distanz (km)',
        ROUND(SUM(j.distanceStraightLine) / 1000.0, 2) AS 'Luftlinie Fahrten (km)',
        ROUND(SUM(j.distanceToPrevJourney) / 1000.0, 2) AS 'Luftlinie zwischen Fahrten (km)',
        ROUND(CASE WHEN SUM(j.distance) > 0 THEN (SUM(j.distanceToPrevJourney) * 100.0 / SUM(j.distance)) ELSE 0 END, 2) AS 'Tracking-Lücke (%)',
        ROUND(SUM(j.distance) * 1.0 / NULLIF(SUM(j.distanceStraightLine), 0), 2) AS 'Umweg Faktor',
        SUM(j.durationReal) AS 'Zeit (min)',
        ROUND(SUM(j.distance) / (NULLIF(SUM(j.durationReal), 0) / 60.0 * 1000.0), 2) AS 'Geschwindigkeit (km/h)',
        SUM(CASE WHEN j.realtimeAvailability = 1 THEN j.departureDelay ELSE 0 END) AS 'Abfahrtsverspätung (min)',
        SUM(CASE WHEN j.realtimeAvailability = 1 THEN j.arrivalDelay ELSE 0 END) AS 'Ankunftsverspätung (min)',
        ROUND(AVG(j.departureDelay), 2) AS '∅ Abfahrtsverspätung (min)',
        ROUND(AVG(j.arrivalDelay), 2) AS '∅ Verspätung (min)',
        ROUND((SUM(CASE WHEN j.realtimeAvailability = 1 THEN 1 ELSE 0 END) * 100.0) / COUNT(j.statusId), 2) AS 'Echtzeitquote (%)',
        ROUND((SUM(CASE WHEN j.arrivalDelay < 0 AND j.realtimeAvailability = 1 THEN 1 ELSE 0 END) * 100.0) / NULLIF(SUM(CASE WHEN j.realtimeAvailability = 1 THEN 1 ELSE 0 END), 0), 2) AS 'Verfrüht (%)',
        ROUND((SUM(CASE WHEN j.arrivalDelay = 0 AND j.realtimeAvailability = 1 THEN 1 ELSE 0 END) * 100.0) / NULLIF(SUM(CASE WHEN j.realtimeAvailability = 1 THEN 1 ELSE 0 END), 0), 2) AS 'Pünktlich (%)',
        ROUND((SUM(CASE WHEN j.arrivalDelay >= 3 THEN 1 ELSE 0 END) * 100.0) / COUNT(j.statusId), 2) AS '3 Minuten und mehr (%)',
        ROUND((SUM(CASE WHEN j.arrivalDelay >= 6 THEN 1 ELSE 0 END) * 100.0) / COUNT(j.statusId), 2) AS '6 Minuten und mehr (%)',
        ROUND((SUM(CASE WHEN j.arrivalDelay >= 15 THEN 1 ELSE 0 END) * 100.0) / COUNT(j.statusId), 2) AS '15 Minuten und mehr (%)',
        ROUND((SUM(CASE WHEN j.arrivalDelay >= 60 THEN 1 ELSE 0 END) * 100.0) / COUNT(j.statusId), 2) AS '60 Minuten und mehr (%)',
        SUM(j.points) AS 'Punkte',
        SUM(j.likes) AS 'Likes',
        ROUND(AVG(j.interStopsNumber), 2) AS '∅ Zwischenhalte'
    FROM journeys j
    LEFT JOIN stops origin ON j.originId = origin.id
    LEFT JOIN stops destination ON j.destinationId = destination.id
    WHERE j.username IN ({placeholders}) {date_query_j} {not_null_sql} {get_shared_visibility_query('j')}
      {subquery_sql}
    GROUP BY {group_by_sql}, j.username
    ORDER BY {group_by_sql} ASC, {order_metric_sql}
    """
    df = pd.read_sql_query(query, conn, params=params)
    if df.empty:
        return pd.DataFrame(), default_sheet_title

    df['Zeit (hh:mm)'] = df['Zeit (min)'].apply(lambda x: f"{int(x // 60):02d}:{int(x % 60):02d}" if pd.notna(x) else "")
    all_metric_cols = SHEET_METRICS_SCHEMA['shared_grouped']['items']
    cols = list(dict.fromkeys(label_cols + ['User'] + [c for c in all_metric_cols if c in df.columns]))
    df = df.loc[:, ~df.columns.duplicated()][cols]

    records = df.to_dict('records')
    formatted_rows = []
    last_group_val = None

    for record in records:
        current_val = tuple(record[c] for c in label_cols)
        if last_group_val is not None and current_val != last_group_val:
            formatted_rows.append({col: '' for col in cols})
        if current_val == last_group_val:
            for c in label_cols:
                record[c] = ''
        formatted_rows.append(record)
        last_group_val = current_val

    if records:
        formatted_rows.append({col: '' for col in cols})
    return pd.DataFrame(formatted_rows), default_sheet_title


def create_shared_excel_file(conn: sqlite3.Connection, usernames: list[str] = None, start_date: str = '',
                             end_date: str = '', columns_to_compare: list[str] = None,
                             custom_settings: dict = None, creator_username: str | None = None) -> str | None:
    if usernames is None:
        usernames = pd.read_sql('SELECT username FROM users', conn)['username'].tolist()
    if not usernames:
        return None

    if custom_settings is None:
        settings_owner = creator_username or get_auth_user() or (usernames[0] if usernames else "")
        custom_settings = get_user_excel_settings(conn, settings_owner)

    shared_sheets = [s for s in custom_settings.get('shared_sheets', []) if s.get('enabled', True) is not False]
    if not shared_sheets:
        shared_sheets = get_default_excel_settings()['shared_sheets']

    output_dir = get_export_dir('shared', creator_username=creator_username)
    names_joined = "_".join(sorted(usernames)[:6])
    if len(usernames) > 6:
        names_joined += f"_and_{len(usernames) - 6}_more"
    
    date_suffix = get_date_suffix(start_date, end_date)
    file_path = output_dir / f"{names_joined}{date_suffix}_shared_data.xlsx"

    needs_shared_stops = any(s.get('type') in ('shared_stops', 'shared_stops_filterable') for s in shared_sheets)
    df_regular_stops, df_filterable_stops = (
        create_shared_stops_dataframes(conn, usernames, start_date, end_date)
        if needs_shared_stops else (pd.DataFrame(), pd.DataFrame())
    )

    used_sheet_names = set()

    with pd.ExcelWriter(file_path, engine='xlsxwriter', datetime_format='dd.mm.yyyy hh:mm') as writer:
        def format_sheet(df, sheet_name, freeze_spalten=1):
            df = df.loc[:, ~df.columns.duplicated()]
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes(1, freeze_spalten)
            if not df.empty:
                worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
            for runner, col in enumerate(df.columns):
                series = df.iloc[:, runner].dropna()
                max_length = int(series.map(lambda x: len(str(x))).max()) if not series.empty else 0
                final_width = max(max_length, len(str(col)) + 2) + 2
                worksheet.set_column(runner, runner, final_width)

        for sheet_cfg in shared_sheets:
            stype = sheet_cfg.get('type', 'shared_grouped')
            freeze_cols = int(sheet_cfg.get('freeze_cols', 1) or 0)
            raw_sname = sheet_cfg.get('sheet_name', stype)
            sel_metrics = sheet_cfg.get('metrics')

            if stype == 'shared_meta':
                sname = sanitize_sheet_name(raw_sname, used_sheet_names)
                df_meta = apply_sheet_metrics(create_shared_meta_dataframe(conn, usernames, start_date, end_date), stype, sel_metrics)
                format_sheet(df_meta, sname, freeze_spalten=freeze_cols)
            elif stype == 'shared_stats':
                shared_stats_df = create_shared_stats_dataframe(conn, usernames, start_date, end_date)
                if not shared_stats_df.empty:
                    shared_stats_df = apply_sheet_metrics(shared_stats_df, stype, sel_metrics)
                    sname = sanitize_sheet_name(raw_sname, used_sheet_names)
                    format_sheet(shared_stats_df, sname, freeze_cols)
            elif stype == 'shared_stops':
                if not df_regular_stops.empty:
                    df_stops_out = apply_sheet_metrics(df_regular_stops.copy(), stype, sel_metrics)
                    sname = sanitize_sheet_name(raw_sname, used_sheet_names)
                    format_sheet(df_stops_out, sname, freeze_cols)
            elif stype == 'shared_stops_filterable':
                if not df_filterable_stops.empty:
                    sname = sanitize_sheet_name(raw_sname, used_sheet_names)
                    format_sheet(df_filterable_stops, sname, freeze_cols)
            elif stype == 'shared_journeys':
                shared_journeys_df = create_shared_journeys_dataframe(conn, usernames, start_date, end_date)
                if not shared_journeys_df.empty:
                    shared_journeys_df = apply_sheet_metrics(shared_journeys_df, stype, sel_metrics)
                    sname = sanitize_sheet_name(raw_sname, used_sheet_names)
                    format_sheet(shared_journeys_df, sname, freeze_cols)
            elif stype == 'shared_grouped':
                args = [a for a in sheet_cfg.get('args', ['category']) if a]
                if not args:
                    args = ['category']
                df, default_title = create_shared_grouped_dataframe(
                    conn,
                    usernames,
                    args,
                    start_date=start_date,
                    end_date=end_date,
                    min_users=int(sheet_cfg.get('min_users', 2) or 2),
                    sort_by_distance=sheet_cfg.get('sort_by_distance', True),
                    exclude_empty=sheet_cfg.get('exclude_empty', False)
                )
                if not df.empty:
                    df = apply_sheet_metrics(df, stype, sel_metrics, num_group_args=len(args))
                    sname = sanitize_sheet_name(raw_sname or default_title, used_sheet_names)
                    format_sheet(df, sname, freeze_cols)

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


# ==========================================
# KI-DATENBANK-ASSISTENT (TEXT-TO-SQL)
# ==========================================

def call_gemini(client, contents: str, json_mode: bool = False, system_instruction: str = None) -> str | dict:
    models_to_try = [
        os.getenv('GEMINI_MODEL', 'gemini-3.8-flash'),
        'gemini-3.7-flash',
        'gemini-3.6-flash',
        'gemini-3.5-flash',
        'gemini-2.5-flash',
        'gemini-3.5-flash-lite'
    ]
    models_to_try = list(dict.fromkeys(models_to_try))

    cfg_kwargs = {}
    if json_mode:
        cfg_kwargs['response_mime_type'] = 'application/json'
    if system_instruction:
        cfg_kwargs['system_instruction'] = system_instruction

    config = types.GenerateContentConfig(**cfg_kwargs) if cfg_kwargs and types else None

    last_err = None
    for model_name in models_to_try:
        try:
            if config:
                resp = client.models.generate_content(model=model_name, contents=contents, config=config)
            else:
                resp = client.models.generate_content(model=model_name, contents=contents)
            text = (resp.text or '').strip()
            if json_mode:
                clean_json = text
                if clean_json.startswith("```"):
                    clean_json = clean_json.split("\n", 1)[1]
                    if clean_json.endswith("```"):
                        clean_json = clean_json.rsplit("```", 1)[0]
                clean_json = clean_json.strip()
                return json.loads(clean_json)
            return text
        except Exception as e:
            last_err = e
            continue
    raise RuntimeError(f"Alle Gemini-Modelle fehlgeschlagen. Letzter Fehler: {last_err}")


class RateLimitExceededError(Exception):
    pass


def check_and_update_rate_limit(conn: sqlite3.Connection, username: str, is_admin: bool = False, limit_per_hour: int = 10):
    if is_admin:
        return True, 999999

    cursor = conn.cursor()
    cursor.execute("SELECT gemini_api_key, ai_request_count, ai_window_start FROM user_login_data WHERE owner = ?", (username,))
    row = cursor.fetchone()
    if not row:
        return True, limit_per_hour

    byok_key, count, window_start_str = row[0], row[1] or 0, row[2]
    if byok_key and str(byok_key).strip():
        # Eigener Key hinterlegt -> vom Instanz-Rate-Limit befreit
        return True, 999999

    now = datetime.now()
    window_start = None
    if window_start_str:
        try:
            window_start = datetime.fromisoformat(window_start_str)
        except Exception:
            window_start = None

    if not window_start or (now - window_start).total_seconds() >= 3600:
        new_count = 1
        new_window_start = now.isoformat()
        cursor.execute("UPDATE user_login_data SET ai_request_count = ?, ai_window_start = ? WHERE owner = ?",
                       (new_count, new_window_start, username))
        conn.commit()
        return True, limit_per_hour - new_count

    if count >= limit_per_hour:
        remaining_seconds = int(3600 - (now - window_start).total_seconds())
        remaining_minutes = max(1, remaining_seconds // 60)
        raise RateLimitExceededError(
            f"Rate-Limit erreicht! Du hast deine {limit_per_hour} kostenlosen Anfragen für diese Stunde verbraucht. "
            f"Bitte warte noch ca. {remaining_minutes} Minuten oder trage oben deinen eigenen (kostenlosen) Google Gemini API-Key ein, um unbegrenzt weiterzufragen."
        )

    new_count = count + 1
    cursor.execute("UPDATE user_login_data SET ai_request_count = ? WHERE owner = ?", (new_count, username))
    conn.commit()
    return True, limit_per_hour - new_count


def get_rate_limit_status(conn: sqlite3.Connection, username: str, is_admin: bool = False, limit_per_hour: int = 10) -> dict:
    cursor = conn.cursor()
    cursor.execute("SELECT gemini_api_key, ai_request_count, ai_window_start FROM user_login_data WHERE owner = ?", (username,))
    row = cursor.fetchone()
    byok_key = row[0] if row else None
    count = (row[1] or 0) if row else 0
    window_start_str = row[2] if row else None
    has_byok = bool(byok_key and str(byok_key).strip())

    now = datetime.now()
    window_start = None
    if window_start_str:
        try:
            window_start = datetime.fromisoformat(window_start_str)
        except Exception:
            window_start = None

    if is_admin or has_byok:
        remaining = 999999
        rate_limited = False
    elif not window_start or (now - window_start).total_seconds() >= 3600:
        remaining = limit_per_hour
        rate_limited = False
    else:
        remaining = max(0, limit_per_hour - count)
        rate_limited = (remaining <= 0)

    result = {
        'is_admin': is_admin,
        'has_byok': has_byok,
        'rate_limited': rate_limited,
        'remaining': remaining,
        'limit': limit_per_hour
    }

    if is_admin:
        global_key = (os.getenv('GEMINI_API_KEY') or '').strip()
        has_global_key = bool(global_key)
        masked_global_key = (global_key[:6] + "..." + global_key[-4:]) if len(global_key) > 10 else ("Aktiv" if has_global_key else "Nicht hinterlegt")

        since_24h = (now - timedelta(hours=24)).strftime('%Y-%m-%d %H:%M:%S')
        since_1h = (now - timedelta(hours=1)).strftime('%Y-%m-%d %H:%M:%S')

        cursor.execute("""
            SELECT COUNT(*) 
            FROM ai_messages m
            JOIN ai_chats c ON m.chat_id = c.id
            LEFT JOIN user_login_data u ON c.owner = u.owner
            WHERE m.role = 'user' 
              AND m.created_at >= ?
              AND (u.gemini_api_key IS NULL OR TRIM(u.gemini_api_key) = '')
        """, (since_24h,))
        global_reqs_24h = cursor.fetchone()[0] or 0

        cursor.execute("""
            SELECT COUNT(*) 
            FROM ai_messages m
            JOIN ai_chats c ON m.chat_id = c.id
            LEFT JOIN user_login_data u ON c.owner = u.owner
            WHERE m.role = 'user' 
              AND m.created_at >= ?
              AND (u.gemini_api_key IS NULL OR TRIM(u.gemini_api_key) = '')
        """, (since_1h,))
        global_reqs_1h = cursor.fetchone()[0] or 0

        cursor.execute("""
            SELECT 
                COUNT(CASE WHEN gemini_api_key IS NOT NULL AND TRIM(gemini_api_key) != '' THEN 1 END),
                COUNT(CASE WHEN gemini_api_key IS NULL OR TRIM(gemini_api_key) = '' THEN 1 END)
            FROM user_login_data
        """)
        u_counts = cursor.fetchone()
        byok_users = u_counts[0] if u_counts else 0
        global_users = u_counts[1] if u_counts else 0

        model_name = os.getenv('GEMINI_MODEL', 'gemini-2.5-flash')

        result['admin_info'] = {
            'has_global_key': has_global_key,
            'masked_global_key': masked_global_key,
            'model_name': model_name,
            'google_rpm': 15,
            'google_rpd': 1500,
            'google_tpm': '1.000.000',
            'global_reqs_24h': global_reqs_24h,
            'global_reqs_1h': global_reqs_1h,
            'global_users_count': global_users,
            'byok_users_count': byok_users
        }

    return result


def sanitize_cell_value(val):
    """Entfernt unkontrollierte Steuerzeichen aus Textzellen, behält aber normale Lesbarkeit (Zeilenumbrüche, Tabs)."""
    if isinstance(val, str):
        return "".join(ch for ch in val if ord(ch) >= 32 or ch in "\t\n\r")
    return val


def run_dataframe_analysis(func_name: str, username: str, start_date: str = '', end_date: str = '',
                           conn: sqlite3.Connection = None, **kwargs) -> dict:
    """
    Sichere Router-Schnittstelle für komplexe Pandas-DataFrame- und Statistik-Funktionen.
    Unterstützt:
      - 'create_user_extended_stats': Aktionsradius, RMSE, Sigma, Extrempunkte (Nord/Süd/Ost/West), Max-Distanz zum Hub.
      - 'create_temporal_stats_dataframe': Ganglinien nach Zeitfenster (time_of_day, weekday, month, year, date).
      - 'create_user_stops_dataframe_with_inter': Haltestellen inkl. Via-Halte, Ein-/Ausstieg, Umsteigequote und letzter Besuch.
      - 'create_user_stats_dataframe': Globale Kernstatistiken eines Nutzers.
      - 'create_user_grouped_frame': Gruppierte Statistiken nach Attributen (z. B. category, operatorName).
    """
    close_conn = False
    local_conn = conn
    if local_conn is None:
        db_path = Path.cwd() / 'traewelling.db'
        local_conn = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
        close_conn = True

    try:
        preview_text = ""
        df = None

        if func_name == 'create_user_extended_stats':
            df = create_user_extended_stats(local_conn, username, start_date, end_date)
            preview_text = df.to_string(index=False) if not df.empty else "Keine Daten gefunden."

        elif func_name == 'create_temporal_stats_dataframe':
            grouping_type = kwargs.get('grouping_type') or kwargs.get('interval') or 'time_of_day'
            type_mapping = {
                'tageszeit': 'time_of_day',
                'uhrzeit': 'time_of_day',
                'wochentag': 'weekday',
                'monat': 'month',
                'jahr': 'year',
                'datum': 'date'
            }
            grouping_type = type_mapping.get(str(grouping_type).lower(), grouping_type)
            if grouping_type not in ['time_of_day', 'weekday', 'month', 'month_year', 'year', 'date']:
                grouping_type = 'time_of_day'

            df_raw = get_temporal_base_dataframe(local_conn, username, start_date, end_date)
            if df_raw.empty:
                df = pd.DataFrame()
                preview_text = f"Keine Fahrtdaten für '{username}' im Zeitraum gefunden."
            else:
                df = create_temporal_stats_dataframe(df_raw, grouping_type=grouping_type, create_picture=False, username=username)
                if not df.empty:
                    key_cols = [
                        c for c in [
                            'Zeitfenster', 'Bucket', 'Wochentag', 'Monat', 'Monat (JJJJ-MM)', 'Jahr', 'Datum',
                            'Fahrten', 'Tage', 'Distanz (km)', 'Anteil Distanz (%)',
                            'Zeit (hh:mm)', 'Anteil Zeit (%)', '∅ Geschwindigkeit (km/h)',
                            'Pünktlich (%)', '∅ Ankunftsverspätung (min)'
                        ] if c in df.columns
                    ]
                    preview_df = df[key_cols] if len(key_cols) >= 3 else df.iloc[:, :8]
                    preview_text = f"Ganglinien-Analyse (Intervall: {grouping_type}, {len(df)} Abschnitte):\n" + preview_df.to_string(index=False)
                else:
                    preview_text = f"Keine Gangliniendaten für '{username}' generiert."

        elif func_name == 'create_user_stops_dataframe_with_inter':
            df = create_user_stops_dataframe_with_inter(local_conn, username, start_date, end_date)
            if not df.empty:
                if df.index.name:
                    df = df.reset_index()
                preview_text = f"Haltestellen-Analyse inkl. Zwischenhalten (insgesamt {len(df)} Haltestellen, Top 25 nach Besuchen):\n" + df.head(25).to_string(index=False)
            else:
                preview_text = f"Keine Haltestellendaten für '{username}' gefunden."

        elif func_name == 'create_user_stats_dataframe':
            df = create_user_stats_dataframe(local_conn, username, start_date, end_date)
            preview_text = df.to_string(index=False) if not df.empty else "Keine Benutzerstatistiken verfügbar."

        elif func_name == 'create_user_grouped_frame':
            first_argument = kwargs.get('first_argument') or 'category'
            second_argument = kwargs.get('second_argument') or ''
            sort_by_distance = kwargs.get('sort_by_distance', True)
            exclude_empty = kwargs.get('exclude_empty', False)
            df = create_user_grouped_frame(
                conn=local_conn,
                username=username,
                first_argument=first_argument,
                second_argument=second_argument,
                start_date=start_date,
                end_date=end_date,
                sort_by_distance=sort_by_distance,
                exclude_empty=exclude_empty
            )
            if not df.empty:
                if df.index.name and df.index.name not in df.columns:
                    df = df.reset_index()
                label_str = f"{first_argument}" + (f", {second_argument}" if second_argument else "")
                preview_text = f"Gruppierte Auswertung ({label_str}, insgesamt {len(df)} Zeilen, Top 25):\n" + df.head(25).to_string(index=False)
            else:
                preview_text = f"Keine gruppierten Daten für '{username}' gefunden."

        else:
            return {
                "success": False,
                "error": f"Unbekannte Analyse-Funktion: '{func_name}'. Gültige Funktionen sind: create_user_extended_stats, create_temporal_stats_dataframe, create_user_stops_dataframe_with_inter, create_user_stats_dataframe, create_user_grouped_frame.",
                "columns": [],
                "data": [],
                "row_count": 0,
                "total_rows": 0,
                "summary_text": ""
            }

        if df is None or df.empty:
            return {
                "success": True,
                "func_name": func_name,
                "row_count": 0,
                "total_rows": 0,
                "columns": [],
                "data": [],
                "summary_text": preview_text or "Keine Daten vorhanden."
            }

        if df.index.name and df.index.name not in df.columns:
            df = df.reset_index()

        total_rows = len(df)
        preview_df = df.head(100).copy()
        preview_df = preview_df.replace([np.inf, -np.inf], None)

        rows = []
        for row in preview_df.itertuples(index=False):
            clean_row = []
            for val in row:
                if val is None or pd.isna(val):
                    clean_row.append(None)
                elif isinstance(val, (datetime, pd.Timestamp)):
                    clean_row.append(val.strftime('%Y-%m-%d %H:%M:%S'))
                elif isinstance(val, (np.integer, int)):
                    clean_row.append(int(val))
                elif isinstance(val, (np.floating, float)):
                    clean_row.append(round(float(val), 2))
                else:
                    clean_row.append(sanitize_cell_value(str(val)))
            rows.append(clean_row)

        return {
            "success": True,
            "func_name": func_name,
            "row_count": len(rows),
            "total_rows": total_rows,
            "columns": [str(c) for c in df.columns],
            "data": rows,
            "summary_text": preview_text
        }
    except Exception as e:
        log_error_to_db(username=username, error=e, context=f"run_dataframe_analysis ({func_name})")
        return {
            "success": False,
            "error": str(e),
            "columns": [],
            "data": [],
            "row_count": 0,
            "total_rows": 0,
            "summary_text": f"Fehler bei Ausführung von {func_name}: {e}"
        }
    finally:
        if close_conn and local_conn:
            local_conn.close()


def wrap_query_with_journeys_cte(sql_query: str, mandatory_j_where: str) -> str:
    """
    Schaltet automatisch eine CTE für `journeys` mit der Pflicht-Filterung vor das SQL-Statement.
    Verwendet main.journeys, um zirkuläre Referenzen im CTE zu vermeiden.
    Kombiniert sauber mit bestehenden WITH / WITH RECURSIVE Klauseln.
    """
    clean = sql_query.strip().rstrip(';')
    if not clean:
        return clean

    cte_body = f"journeys AS (SELECT * FROM main.journeys AS j WHERE {mandatory_j_where})"

    # Prüfe auf führende Kommentare und WITH RECURSIVE
    match_rec = re.match(r"^(\s*(?:--[^\n]*\n\s*)*)WITH\s+RECURSIVE\s+", clean, re.IGNORECASE)
    if match_rec:
        prefix = match_rec.group(1)
        rest = clean[match_rec.end():]
        return f"{prefix}WITH RECURSIVE {cte_body},\n{rest}"

    # Prüfe auf führende Kommentare und einfaches WITH
    match_with = re.match(r"^(\s*(?:--[^\n]*\n\s*)*)WITH\s+", clean, re.IGNORECASE)
    if match_with:
        prefix = match_with.group(1)
        rest = clean[match_with.end():]
        return f"{prefix}WITH {cte_body},\n{rest}"

    return f"WITH {cte_body}\n{clean}"


def _process_ai_chat_turn_impl(conn: sqlite3.Connection, creator: str, chat_id: str, user_message: str,
                              selected_users: list[str], start_date: str = '', end_date: str = '',
                              is_admin: bool = False) -> dict:
    with user_auth_context_from_db(conn, creator):
        return _process_ai_chat_turn_worker(conn, creator, chat_id, user_message, selected_users, start_date, end_date, is_admin)


def _process_ai_chat_turn_worker(conn: sqlite3.Connection, creator: str, chat_id: str, user_message: str,
                                selected_users: list[str], start_date: str = '', end_date: str = '',
                                is_admin: bool = False) -> dict:
    # 0. Key-Auswahl & Rate-Limit
    cursor = conn.cursor()
    cursor.execute("SELECT gemini_api_key FROM user_login_data WHERE owner = ?", (creator,))
    user_row = cursor.fetchone()
    user_byok = (user_row[0] or '').strip() if user_row and user_row[0] else ''

    active_key = user_byok if user_byok else (os.getenv('GEMINI_API_KEY') or '').strip()

    if not active_key:
        return {
            "success": False,
            "error": "Kein Gemini API-Key verfügbar. Bitte klicke oben auf '🔑 API-Key konfigurieren', um deinen eigenen Key einzutragen, oder hinterlege einen Instanz-Key in der .env."
        }

    if genai is None:
        return {
            "success": False,
            "error": "Das Paket 'google-genai' ist nicht installiert. Bitte 'pip install google-genai' ausführen."
        }

    if not user_byok:
        check_and_update_rate_limit(conn, creator, is_admin=is_admin, limit_per_hour=10)

    # 1. Sichtbarkeits- & Kontext-Setup (wird durch user_auth_context_from_db bereitgestellt)

    if selected_users:
        cleaned_users = [re.sub(r"[^a-zA-Z0-9_\-\.]", "", str(u)) for u in selected_users if str(u).strip()]
        user_filter = f"j.username IN ({','.join(f"'{u}'" for u in cleaned_users)})" if cleaned_users else f"j.username = '{creator}'"
        default_target_user = cleaned_users[0] if cleaned_users else creator
    else:
        cleaned_users = [creator]
        user_filter = f"j.username = '{creator}'"
        default_target_user = creator

    # 2. Query-Verschlankung bei Selbstabfragen: Wenn ausschließlich der eingeloggte User
    # abgefragt wird, darf er alle eigenen Fahrten sehen -> vis_sql entfällt vollständig.
    if cleaned_users == [creator] or not cleaned_users:
        vis_sql = ""
    else:
        vis_sql = get_shared_visibility_query("j")

    date_conds = []
    if start_date:
        clean_start = re.sub(r"[^0-9\-]", "", str(start_date))
        date_conds.append(f"COALESCE(j.arrivalFinal, j.arrivalPlanned) >= '{clean_start}'")
    if end_date:
        clean_end = re.sub(r"[^0-9\-]", "", str(end_date))
        date_conds.append(f"COALESCE(j.departureFinal, j.departurePlanned) <= '{clean_end}T23:59:59'")
    date_filter = (" AND " + " AND ".join(date_conds)) if date_conds else ""

    mandatory_j_where = f"({user_filter}{vis_sql}{date_filter})"

    # 2. Bisherigen Chat-Verlauf laden
    cursor.execute(
        """SELECT role, content, sql_query FROM (
               SELECT id, role, content, sql_query FROM ai_messages
               WHERE chat_id = ? ORDER BY id DESC LIMIT 10
           ) ORDER BY id ASC""",
        (chat_id,)
    )
    history_rows = cursor.fetchall()

    history_lines = []
    for h_role, h_content, h_sql in history_rows:
        if h_role == 'user':
            history_lines.append(f"User: {h_content}")
        else:
            line = f"Assistant: {h_content}"
            if h_sql:
                line += f"\n[Ausgeführtes SQL: {h_sql}]"
            history_lines.append(line)
    history_text = "\n\n".join(history_lines) if history_lines else "(Bisher keine Nachrichten in diesem Chat)"

    client = genai.Client(api_key=active_key, http_options={'timeout': 20000})

    # 3. Schritt 1 – SQL-Generierung (Gemini)
    schema_prompt = f"""Du bist der KI-Agent "Antigravity" – ein hochpräziser SQLite-Datenbankexperte für Träwelling-Fahrtauswertungen.
Ziel: Erzeuge basierend auf der Nutzerfrage, dem bisherigen Kontext und dem Datenbankschema ein extrem robustes, sauberes und fehlerfreies SQLite-SELECT-Statement.

DATENBANKSCHEMA & SPALTENBEDEUTUNG:
Tabelle journeys (Alias j):
  statusId INTEGER PRIMARY KEY,
  userId INTEGER,
  username TEXT,
  body TEXT (Kommentar/Notiz zur Fahrt),
  business TEXT ('Privat', 'Geschäftlich', 'Arbeitsweg'),
  visibility INTEGER (0=öffentlich, 2=Follower, 4=nicht gelistet, 5=Freunde),
  category TEXT ('Fernverkehr', 'Regionalzug', 'S-Bahn', 'U-Bahn', 'Tram', 'Bus', 'Fähre', 'Igitt ein Flugzeug'),
  lineName TEXT (z. B. 'ICE 574', 'RE 1', 'U2'),
  journeyNumber TEXT,
  operatorName TEXT (z. B. 'DB Fernverkehr AG', 'S-Bahn Berlin GmbH'),
  distance INTEGER (in METERN! Verwende SUM(j.distance)/1000.0 oder ROUND(j.distance/1000.0, 1) für Kilometer),
  distanceStraightLine INTEGER (Luftlinie in Metern),
  distanceToPrevJourney INTEGER (Distanz zur vorherigen Haltestelle in Metern),
  departureDelay INTEGER (Verspätung Abfahrt in Minuten),
  arrivalDelay INTEGER (Verspätung Ankunft in Minuten),
  durationPlanned INTEGER (Geplante Fahrzeit in Minuten),
  durationReal INTEGER (Reale Fahrzeit in Minuten),
  timeDelta INTEGER (durationReal - durationPlanned in Minuten),
  points INTEGER,
  likes INTEGER,
  departurePlanned TEXT, departureFromRealtime TEXT, departureManual TEXT, departureFinal TEXT (Abfahrt real ISO-8601),
  arrivalPlanned TEXT, arrivalFromRealtime TEXT, arrivalManual TEXT, arrivalFinal TEXT (Ankunft real ISO-8601),
  realtimeAvailability BOOLEAN (1 = Echtzeitdaten vorhanden, 0 = nur Plan-Zeiten),
  realtimeFromDataprovider BOOLEAN (1 = Echtzeitdaten direkt vom Betreiber/Provider, 0 = sonstig),
  originId INTEGER (FK zu stops.id),
  destinationId INTEGER (FK zu stops.id),
  interStopsNumber INTEGER,
  interStops TEXT (JSON-Array der Zwischenhalte als Strings, z. B. '["Station A", "Station B"]'),
  locomotiveClass TEXT (Baureihe z. B. '401', '412', '423'),
  vehicleNumber TEXT (Triebzugnummer),
  ticket TEXT, travelClass TEXT, price TEXT, passengerRights TEXT,
  createdAt TEXT

Tabelle stops (Alias s, o für origin, d für destination):
  id INTEGER PRIMARY KEY,
  stopName TEXT (z. B. 'Frankfurt(Main)Hbf', 'Hamburg Hbf'),
  rilIdentifier TEXT (z. B. 'FF', 'AA'),
  latitude TEXT, longitude TEXT,
  countryCode TEXT (2-stelliger ISO-Alpha-2-Code, z. B. 'DE', 'AT', 'CH', 'NL', 'FR'),
  countryName TEXT (WICHTIG: Enthält IMMER den deutschen Namen PLUS Flaggen-Emoji, z. B. 'Deutschland 🇩🇪', 'Österreich 🇦🇹'),
  admin1State TEXT (Bundesland/Region PLUS Flaggen-Emoji, z. B. 'Bayern 🇩🇪', 'Niedersachsen 🇩🇪'),
  admin2Dist TEXT (Bezirk/Kreis PLUS Flaggen-Emoji),
  city TEXT (Stadtname ohne Flagge, z. B. 'Braunschweig', 'Frankfurt')

WICHTIGE REGELN & BEST PRACTICES:
1. SICHERHEIT & SCOPE: Das Backend filtert die Tabelle `journeys` automatisch auf die berechtigten User und den gewählten Zeitraum. Du musst KEINE User- oder Datums-Filterlisten in deine WHERE-Klauseln einbauen, es sei denn, der Nutzer schränkt in seiner Frage explizit ein (z. B. nach einer bestimmten Linie, Stadt oder Kategorie). Schreibe einfach ganz normale SELECT-Abfragen auf `journeys`.
2. UMGANG MIT LÄNDERN & REGIONEN:
   - Da `countryName`, `admin1State` und `admin2Dist` in der DB ein Flaggen-Emoji enthalten (z. B. 'Deutschland 🇩🇪'), nutze für exakte Vergleiche idealerweise den ISO-Code `countryCode` (z. B. `o.countryCode = 'DE'` oder `o.countryCode != 'DE'`) ODER arbeite mit `LIKE 'Deutschland%'` bzw. `NOT LIKE 'Deutschland%'`.
   - Verwende bei Ländern oder Bundesländern NIEMALS einfache Vergleiche wie `o.countryName = 'Deutschland'` oder `originCountry = 'Deutschland'`.
3. KOMPLEXE ANALYSEN & MEHRERE REIHEN (z. B. erste/letzte Ausreise, Einreise, Grenzübertritte):
   - Vermeide unübersichtliche oder verschachtelte `WITH`-Ketten, die SQLite-Syntaxfehler erzeugen können.
   - Wenn du mehrere Abfragen kombinierst (z. B. "Letzte Fahrt raus" und "Erste Fahrt rein"), nutze saubere Subqueries mit `UNION ALL`:
     SELECT * FROM (
         SELECT 'Ausreise' AS typ, j.departureFinal AS zeitpunkt, o.stopName AS start, d.stopName AS ziel
         FROM journeys j JOIN stops o ON j.originId = o.id JOIN stops d ON j.destinationId = d.id
         WHERE o.countryCode = 'DE' AND d.countryCode != 'DE'
         ORDER BY j.departureFinal DESC LIMIT 1
     )
     UNION ALL
     SELECT * FROM (
         SELECT 'Einreise' AS typ, j.arrivalFinal AS zeitpunkt, o.stopName AS start, d.stopName AS ziel
         FROM journeys j JOIN stops o ON j.originId = o.id JOIN stops d ON j.destinationId = d.id
         WHERE o.countryCode != 'DE' AND d.countryCode = 'DE'
         ORDER BY j.arrivalFinal ASC LIMIT 1
     );
4. HALTESTELLEN & ORTE:
   - Haltestellenbesuche (entweder als Start ODER Ziel) zählst du immer über eine CTE mit `UNION ALL`:
     WITH stop_visits AS (
         SELECT j.originId AS stopId, j.statusId, j.username, j.distance FROM journeys j
         UNION ALL
         SELECT j.destinationId AS stopId, j.statusId, j.username, j.distance FROM journeys j
     )
     SELECT s.stopName, s.city, COUNT(*) AS besuche, ROUND(SUM(sv.distance)/1000.0, 1) AS km
     FROM stop_visits sv JOIN stops s ON sv.stopId = s.id
     GROUP BY s.id, s.stopName ORDER BY besuche DESC LIMIT 10;
   - Bei Stadt- oder Haltestellensuche immer flexibel mit `LIKE` suchen: `(s.city LIKE '%Name%' OR s.stopName LIKE '%Name%')`.
5. METRIKEN & AGGREGATIONEN:
   - Distanzen: Immer `SUM(j.distance)/1000.0` für Kilometer.
   - Echtzeitquoten / Provider-Verfügbarkeit: Nutze `SUM(CASE WHEN j.realtimeAvailability = 1 THEN 1 ELSE 0 END) * 100.0 / COUNT(*)` bzw. `j.realtimeFromDataprovider`.
   - Verspätungen: `AVG(j.arrivalDelay)` und `AVG(j.departureDelay)` in Minuten.
6. SQLITE SYNTAX-REGELN:
   - Nutze bei `UNION ALL` NIEMALS `ORDER BY` oder `LIMIT` direkt in den äußeren Selects ohne umgebende Klammern/Subqueries.
   - Vergib in allen `SELECT`-Teilen eines `UNION ALL` identische Spaltennamen und -typen.
7. SICHERHEIT: Nur lesende `SELECT`- oder `WITH`-Abfragen erzeugen! Keine schreibenden Befehle (kein INSERT, UPDATE, DROP, PRAGMA etc.).
8. VERFÜGBARE ERWEITERTE PYTHON/PANDAS-ANALYSEFUNKTIONEN (DATAFRAMES):
   Es stehen spezialisierte Python/Pandas-Analysefunktionen zur Verfügung, die komplexe mathematische, zeitliche oder Zwischenhalte-Berechnungen durchführen, welche in reinem SQL nur schwer oder gar nicht möglich sind:
   a) create_user_extended_stats:
      - Berechnet: Aktionsradius in km (Hub-gewichtete Distanzquadrate), RMSE der Abweichung von Realtimeanbietern (Minuten), Standardabweichung Sigma der Ankunftsverspätung (σ in Minuten), Extrempunkte (Nördlichste, Südlichste, Westlichste, Östlichste Haltestelle) sowie die Haltestelle mit der maximalen Distanz zum Hub.
      - Wann nutzen: Wenn der Nutzer nach Aktionsradius, RMSE, Sigma, Standardabweichung der Verspätung, Extrem-Haltestellen (nördlichster/südlichster Halt) oder maximaler Distanz fragt.
      - Argumente: tool_args: {{"username": "{default_target_user}"}}
   b) create_temporal_stats_dataframe:
      - Berechnet: Zeitreihen- und Ganglinien-Analysen mit anteiliger Minuten- und Distanzaufteilung, Reisegeschwindigkeit und Anteil an Gesamtzeit.
      - Wann nutzen: Wenn nach Fahrzeiten/Fahrten aufgeteilt nach Tageszeiten (z. B. Halbstundenintervalle), Wochentagen, Monaten oder Jahren gefragt wird.
      - Argumente: tool_args: {{"username": "{default_target_user}", "grouping_type": "time_of_day" | "weekday" | "month" | "year" | "date"}}
   c) create_user_stops_dataframe_with_inter:
      - Berechnet: Haltestellenauswertungen inklusive aller Zwischenhalte (Via-Halte aus interStops), getrennt nach 'Ohne Via' (Ein-/Ausstieg), 'Nur Via' (Durchfahrt), Gesamtbesuchen sowie echter und relativer Umsteigequote (%) und letztem Besuch.
      - Wann nutzen: Wenn nach Umsteigequoten, Via-Halten/Durchfahrten oder vollständigen Haltestellen-Rankings inkl. Zwischenhalten gefragt wird.
      - Argumente: tool_args: {{"username": "{default_target_user}"}}
   d) create_user_stats_dataframe:
      - Berechnet: Globale Kernstatistiken eines Nutzers (Gesamtkilometer, Reisezeit, Verspätungsmittelwert, Anzahl Fahrten, Pünktlichkeitsquote etc.).
      - Argumente: tool_args: {{"username": "{default_target_user}"}}

   e) create_user_grouped_frame:
      - Berechnet: Gruppierte Statistiken nach Attributen (z. B. Verkehrsmittel 'category', Betreiber 'operatorName', Bundesland 'admin1State', Land 'countryName', etc.) sortiert nach Kilometern oder Fahrten.
      - Wann nutzen: Für aggregierte Aufteilungen nach Verkehrsmittelkategorien, Betreibern, Bundesländern o. ä.
      - Argumente: tool_args: {{"username": "{default_target_user}", "first_argument": "category", "second_argument": "", "sort_by_distance": true, "exclude_empty": false}}

9. MULTI-ACTION-BATCHING & REGELN FÜR KOMPLEXE FRAGEN:
   - Du kannst in `actions` bis zu 6 Abfragen (sowohl `dataframe`-Tools als auch `sql`-Queries) gleichzeitig kombinieren!
   - WICHTIG bei offenen Analysefragen wie 'Charakterisiere den User', 'Erstelle ein Mobilitätsprofil' oder 'Gib mir einen Gesamtüberblick': Führe IMMER eine Kombination aus mehreren Abfragen aus, um ein vollständiges Bild zu erhalten:
     1. `create_user_stats_dataframe` (Gesamtfahrten, km, Fahrzeit, Pünktlichkeit, Tracking-Lücke, Umwegfaktor)
     2. `create_user_extended_stats` (Aktionsradius in km, Haupthaltestelle/Hub, Extrempunkte Nord/Süd/Ost/West, Sigma/RMSE)
     3. `create_temporal_stats_dataframe` mit `grouping_type='time_of_day'` oder `'weekday'` (Reisegewohnheiten nach Tageszeit/Wochentag – Frühaufsteher, Pendler, Nachtfahrer)
     4. `create_user_stops_dataframe_with_inter` (Meistbesuchte Haltestellen inkl. Via-Durchfahrten und Umsteigequote)
     5. Ergänzende `sql`-Abfragen auf `journeys j` für: Top-Verkehrsmittel (`category`), Top-Betreiber (`operatorName`), Lieblings-Linien (`lineName`) & Baureihen (`locomotiveClass`), Fahrtgründe (`business`), besuchte Länder (`o.countryName`, `d.countryName`) sowie ggf. häufige Tags (`ticket`, `travelClass`, `role`) oder markante Statustexte (`body`).
   - Bei einfachen, gezielten Einzelfragen (z. B. "Wie viele Kilometer bin ich gefahren?") genügt 1 gezielte Action in `actions`.
   - Bei reinen Begrüßungen oder Fragen ohne Datenbezug: Setze `"actions": []`, `"requires_followup": false`, `"direct_answer": "Hallo! ..."`

11. TEAPOT-EASTER-EGG (HTTP 418):
   - Wenn der Nutzer dich auffordert, Kaffee (oder Espresso, Cappuccino, Latte Macchiato, Café Crema etc.) zu kochen, zuzubereiten, zu holen oder zu servieren, setze `"is_teapot": true`, `"actions": []` und schreibe in `"direct_answer"` eine humorvolle Ablehnung im Stil von "418 I'm a teapot" (gerne mit Anspielung auf Eisenbahn/Bordbistro oder den Admin tee_beutel (Ich mag keinen Kaffee)).
   - WICHTIG: Fragt der Nutzer hingegen nach echten Daten aus seiner Datenbank (z. B. "Wie oft habe ich Kaffee in meinen Statustexten erwähnt?"), ist das KEIN 418-Fehler, sondern eine reguläre SQL-Abfrage auf `j.body`!
   
12. REGELN FÜR FREITEXT- UND STATUS-SUCHEN (j.body):
   - Wenn der Nutzer nach Begriffen im Statustext sucht (z. B. "Kaffee", "zuhause", bestimmte Getränke, Orte oder Aktivitäten), verwende in SQL immer eine flexible, case-insensitive Suche mit Wildcards: `LOWER(j.body) LIKE '%begriff%'`.
   - Bei Begriffen mit unterschiedlichen Schreibweisen oder Varianten (z. B. "zuhause" vs. "zu Hause" vs. "nach hause"), baue IMMER eine kombinierte OR-Klausel auf (`LOWER(j.body) LIKE '%zuhause%' OR LOWER(j.body) LIKE '%zu hause%' OR LOWER(j.body) LIKE '%nach hause%'`).

13. STRIKTES VERBOT VON EXAKTEN HALTESTELLEN-NAMEN OHNE VORHERIGEN LOOKUP (NAMENSAUFLÖSUNG & RELATIONEN):
   - WICHTIGSTE DB-EIGENSCHAFT: Haltestellen in `stops.stopName` enthalten in der Datenbank fast immer den Stadtnamen als Präfix (z. B. 'Braunschweig, Pockelsstraße', 'Braunschweig, Rathaus', 'Braunschweig, Mühlenpfortstraße') oder ein angehängtes RIL100-Kürzel (z. B. 'Braunschweig Hbf (HBS)', 'Frankfurt(Main)Hbf (FF)').
   - STRIKTES VERBOT FÜR RUNDE 1: Verwende in Runde 1 NIEMALS `stopName = '...'` oder `stopName IN ('Name1', 'Name2', ...)` mit vom Nutzer genannten Namen! Das führt wegen fehlender Städtenamen oder RIL100-Kürzel zu 0 Treffern!
   - GOLDENE REGEL FÜR VERBINDUNGS- UND RICHTUNGSFRAGEN ("Welche Haltestellen nutze ich VON X und welche ZU X?"):
     1. Verwende bei Fragen nach beiden Richtungen (ab X vs. nach X) NIEMALS eine einzige kombinierte `OR`-Abfrage über `o.stopName` und `d.stopName`, sondern lege ZWINGEND ZWEI GETRENNTE SQL-Actions in `"actions"` an!
     2. Gruppiere dabei jeweils NUR nach der Gegenhaltestelle, damit verschiedene Schreibweisen der Referenzhaltestelle X (wie 'Braunschweig Michelfelderplatz' und 'Michelfelderplatz, Braunschweig') automatisch zusammengezählt werden:
        - Action 1 (Fahrten AB X):
          SELECT d.stopName AS ziel_halt, COUNT(*) AS fahrten
          FROM journeys j
          JOIN stops o ON j.originId = o.id
          JOIN stops d ON j.destinationId = d.id
          WHERE LOWER(o.stopName) LIKE '%michelfelder%'
          GROUP BY d.stopName
          ORDER BY fahrten DESC;
        - Action 2 (Fahrten NACH X):
          SELECT o.stopName AS start_halt, COUNT(*) AS fahrten
          FROM journeys j
          JOIN stops o ON j.originId = o.id
          JOIN stops d ON j.destinationId = d.id
          WHERE LOWER(d.stopName) LIKE '%michelfelder%'
          GROUP BY o.stopName
          ORDER BY fahrten DESC;
     3. Halte `LIKE`-Bedingungen so einfach wie möglich: Nutze den markanten Wortstamm (z. B. `LOWER(o.stopName) LIKE '%michelfelder%'` oder `LOWER(s.stopName) LIKE '%mühlenpfor%'`), statt fehleranfällige, mehrfach verschachtelte `REPLACE`-Ketten zu bauen.
   - ZWISCHENHALTE (`interStops`) MITPRÜFEN:
     Wenn der Nutzer nach Haltestellen auf einer Strecke fragt oder bestimmte genannte Haltestellen nicht als Start/Ziel (`originId`/`destinationId`) auftauchen, prüfe zusätzlich (oder über `"requires_followup": true`), ob diese Haltestellen als Zwischenhalt in `j.interStops` durchfahren wurden:
     SELECT o.stopName AS start_halt, d.stopName AS ziel_halt, j.lineName, j.interStops, COUNT(*) AS fahrten
     FROM journeys j
     JOIN stops o ON j.originId = o.id
     JOIN stops d ON j.destinationId = d.id
     WHERE (
         LOWER(o.stopName) LIKE '%michelfelder%'
         OR LOWER(d.stopName) LIKE '%michelfelder%'
       )
     GROUP BY o.stopName, d.stopName, j.lineName, j.interStops;
   - ALLGEMEINER KANDIDATEN-LOOKUP (wenn nach mehreren losen Haltestellen oder Betreibern gefragt wird):
     Nutze `"requires_followup": true` und suche in Runde 1 IMMER mit `OR`-verknüpften `LIKE '%wortstamm%'`-Bedingungen (z. B. `LOWER(s.stopName) LIKE '%pockel%' OR LOWER(s.stopName) LIKE '%mühlenpfort%' OR LOWER(s.stopName) LIKE '%rathaus%'`), niemals mit `IN (...)`. Erst in Runde 2 darfst du die in Runde 1 gefundenen exakten `stopName`-Werte aus der Datenbank für `=` oder `IN (...)` verwenden.
   - RÜCKFRAGE BEI UNSICHERHEIT:
     Wenn der Lookup mehrere mehrdeutige Haltestellen/Betreiber liefert (z. B. 'Rathaus' in verschiedenen Städten) und aus dem Kontext nicht eindeutig hervorgeht, welche gemeint ist, gib in Runde 2 `"actions": []` zurück und frage beim Nutzer mit den gefundenen DB-Namen nach.

14. ANTWORTRÜCKGABE (JSON-FORMAT):
Gib deine Antwort IMMER als valides JSON-Objekt zurück:
{{
  "is_teapot": false,
  "actions": [
    {{
      "type": "dataframe",
      "label": "Kernstatistiken abrufen",
      "tool_name": "create_user_stats_dataframe",
      "tool_args": {{"username": "{default_target_user}"}}
    }},
    {{
      "type": "dataframe",
      "label": "Erweiterte Geo- und Verspätungsmetriken",
      "tool_name": "create_user_extended_stats",
      "tool_args": {{"username": "{default_target_user}"}}
    }},
    {{
      "type": "sql",
      "label": "Top Verkehrsmittel",
      "sql": "SELECT j.category, COUNT(*) AS fahrten, ROUND(SUM(j.distance)/1000.0, 1) AS km FROM journeys j GROUP BY j.category ORDER BY fahrten DESC"
    }}
  ],
  "requires_followup": false,
  "direct_answer": ""
}}

BISHERIGER KONVERSATIONSVERLAUF:
{history_text}

AKTUELLE NUTZERFRAGE:
{user_message}
"""

    sql_resp = call_gemini(client, schema_prompt, json_mode=True)

    if sql_resp.get("is_teapot"):
        teapot_msg = (sql_resp.get("direct_answer") or "").strip()
        if not teapot_msg:
            teapot_msg = "418 I'm a teapot: Ich bin eine Bahn-Datenbank auf tee_beutel-Basis, keine Kaffeemaschine! 🫖"
        elif "418" not in teapot_msg:
            teapot_msg = f"418 I'm a teapot: {teapot_msg}"

        log_error_to_db(
            username=creator,
            error_type="HTTP_418",
            error_message=teapot_msg,
            context=f"AI Teapot Trigger (Nachricht: {user_message})",
            conn=conn
        )
        return {
            "success": False,
            "is_teapot": True,
            "error": teapot_msg
        }

    raw_actions = sql_resp.get("actions")
    actions = []
    if isinstance(raw_actions, list) and raw_actions:
        actions = raw_actions
    else:
        # Abwärtskompatibilitäts-Fallback auf altes Format
        if sql_resp.get("use_dataframe_tool") and sql_resp.get("tool_name"):
            actions.append({
                "type": "dataframe",
                "label": f"DataFrame: {sql_resp.get('tool_name')}",
                "tool_name": sql_resp.get("tool_name"),
                "tool_args": sql_resp.get("tool_args") or {}
            })
        elif sql_resp.get("needs_sql") and sql_resp.get("sql"):
            actions.append({
                "type": "sql",
                "label": "SQL-Abfrage",
                "sql": sql_resp.get("sql")
            })

    requires_followup = bool(sql_resp.get("requires_followup", False))
    direct_answer = (sql_resp.get("direct_answer") or "").strip()

    forbidden_pattern = r"\b(INSERT|UPDATE|DELETE|DROP|ALTER|REPLACE\s+INTO|PRAGMA|ATTACH|DETACH|CREATE|VACUUM|user_login_data|ai_chats|ai_messages)\b"
    db_path = Path.cwd() / 'traewelling.db'

    def execute_single_action(action: dict, step_num: int) -> dict:
        a_type = (action.get("type") or "").strip().lower()
        label = (action.get("label") or f"Schritt {step_num}").strip()

        if a_type == "dataframe":
            tool_name = (action.get("tool_name") or "").strip()
            tool_args = action.get("tool_args") or {}
            target_username = tool_args.get("username") or default_target_user
            extra_tool_args = {
                k: v for k, v in tool_args.items()
                if k not in ("func_name", "username", "start_date", "end_date", "conn")
            }
            df_res = run_dataframe_analysis(
                func_name=tool_name,
                username=target_username,
                start_date=start_date,
                end_date=end_date,
                conn=conn,
                **extra_tool_args
            )

            arg_str = f"username='{target_username}'"
            if start_date:
                arg_str += f", start_date='{start_date}'"
            if end_date:
                arg_str += f", end_date='{end_date}'"
            for k, v in extra_tool_args.items():
                arg_str += f", {k}={repr(v)}"

            q_str = f"# Pandas-Analysefunktion: {tool_name}({arg_str})"
            r_count = df_res.get("row_count", 0)
            t_rows = df_res.get("total_rows", 0)
            preview = df_res.get("summary_text", "")
            if not df_res.get("success"):
                preview = f"⚠️ Fehler bei {tool_name}: {df_res.get('error', 'Unbekannter Fehler')}"

            return {
                "step_index": step_num,
                "label": label,
                "type": "dataframe",
                "query": q_str,
                "row_count": r_count,
                "total_rows": t_rows,
                "columns": df_res.get("columns", []),
                "data": df_res.get("data", []),
                "preview_text": preview,
                "success": df_res.get("success", False)
            }

        elif a_type == "sql":
            raw_sql = (action.get("sql") or "").strip()
            if not raw_sql:
                return {
                    "step_index": step_num,
                    "label": label,
                    "type": "sql",
                    "query": "-- Leere SQL-Abfrage",
                    "row_count": 0,
                    "total_rows": 0,
                    "columns": [],
                    "data": [],
                    "preview_text": "Keine SQL-Query angegeben.",
                    "success": False
                }

            clean_sql = raw_sql.strip().rstrip(';')
            clean_sql = re.sub(
                r"\b((?:[a-zA-Z0-9_]+\.)?(?:countryName|admin1State|admin2Dist|originCountry|destCountry|destinationCountry|country))\s*=\s*'([A-Za-zÄÖÜäöüß\s\-]+)'",
                r"\1 LIKE '\2%'",
                clean_sql,
                flags=re.IGNORECASE
            )
            clean_sql = re.sub(
                r"\b((?:[a-zA-Z0-9_]+\.)?(?:countryName|admin1State|admin2Dist|originCountry|destCountry|destinationCountry|country))\s*(?:!=|<>)\s*'([A-Za-zÄÖÜäöüß\s\-]+)'",
                r"\1 NOT LIKE '\2%'",
                clean_sql,
                flags=re.IGNORECASE
            )

            # Sicherheitsnetz: Wandelt exakte stopName IN ('A', 'B') Abfragen ohne Wildcards
            # automatisch in tolerante LIKE-Abfragen um (wichtig wegen Ortspräfixen & RIL100-Codes)
            def _replace_stopname_in(match):
                col_expr = match.group(1)
                inner_items = re.findall(r"'([^']+)'", match.group(2))
                if not inner_items:
                    return match.group(0)
                # Nur umwandeln, wenn keine RIL100-Klammern und keine Kommas (wie 'Stadt, Halt') enthalten sind
                if all(',' not in item and '(' not in item for item in inner_items):
                    like_parts = [
                        f"LOWER(REPLACE(REPLACE({col_expr}, ' ', ''), '-', '')) LIKE '%{re.sub(r'[\s\-]+', '', item.lower())}%'"
                        for item in inner_items
                    ]
                    return "(" + " OR ".join(like_parts) + ")"
                return match.group(0)

            clean_sql = re.sub(
                r"\b((?:[a-zA-Z0-9_]+\.)?stopName)\s+IN\s*\(([^)]+)\)",
                _replace_stopname_in,
                clean_sql,
                flags=re.IGNORECASE
            )

            # Sicherheitsprüfung: nur SELECT oder WITH
            if not re.match(r"^(SELECT|WITH)\b", clean_sql, re.IGNORECASE) or re.search(forbidden_pattern, clean_sql, re.IGNORECASE):
                return {
                    "step_index": step_num,
                    "label": label,
                    "type": "sql",
                    "query": clean_sql,
                    "row_count": 0,
                    "total_rows": 0,
                    "columns": [],
                    "data": [],
                    "preview_text": "Sicherheitsprüfung fehlgeschlagen: Nur lesende SELECT- oder WITH-Abfragen ohne verbotene Befehle sind zulässig.",
                    "success": False
                }

            # Backend-Wrapping: Automatische CTE für journeys vorschalten
            wrapped_sql = wrap_query_with_journeys_cte(clean_sql, mandatory_j_where)

            ro_conn = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
            df = None
            sql_error = None
            try:
                df = pd.read_sql_query(wrapped_sql, ro_conn)
            except Exception as sql_err:
                sql_error = sql_err
                fix_prompt = f"""Die folgende SQLite-Query hat einen Fehler geworfen:
Query:
{clean_sql}

Fehler:
{sql_err}

Hinweis: Die Tabelle `journeys` wird vom Backend automatisch auf die berechtigten Daten gefiltert. Du musst keine Pflicht-WHERE-Bedingungen für User/Datumsbereiche einbauen.

Bitte korrigiere die SQL-Abfrage für SQLite. Gib wieder nur das JSON zurück:
{{"needs_sql": true, "sql": "KORRIGIERTES SELECT ...", "direct_answer": ""}}
"""
                try:
                    retry_resp = call_gemini(client, fix_prompt, json_mode=True)
                    retry_sql = retry_resp.get("sql", "").strip().rstrip(';')
                    retry_sql = re.sub(
                        r"\b((?:[a-zA-Z0-9_]+\.)?(?:countryName|admin1State|admin2Dist|originCountry|destCountry|destinationCountry|country))\s*=\s*'([A-Za-zÄÖÜäöüß\s\-]+)'",
                        r"\1 LIKE '\2%'",
                        retry_sql,
                        flags=re.IGNORECASE
                    )
                    retry_sql = re.sub(
                        r"\b((?:[a-zA-Z0-9_]+\.)?(?:countryName|admin1State|admin2Dist|originCountry|destCountry|destinationCountry|country))\s*(?:!=|<>)\s*'([A-Za-zÄÖÜäöüß\s\-]+)'",
                        r"\1 NOT LIKE '\2%'",
                        retry_sql,
                        flags=re.IGNORECASE
                    )
                    retry_sql = re.sub(
                        r"\b((?:[a-zA-Z0-9_]+\.)?stopName)\s+IN\s*\(([^)]+)\)",
                        _replace_stopname_in,
                        retry_sql,
                        flags=re.IGNORECASE
                    )
                    if retry_sql and re.match(r"^(SELECT|WITH)\b", retry_sql, re.IGNORECASE) and not re.search(forbidden_pattern, retry_sql, re.IGNORECASE):
                        clean_sql = retry_sql
                        retry_wrapped_sql = wrap_query_with_journeys_cte(clean_sql, mandatory_j_where)
                        df = pd.read_sql_query(retry_wrapped_sql, ro_conn)
                        sql_error = None
                except Exception as retry_err:
                    sql_error = retry_err
            finally:
                ro_conn.close()

            if df is not None:
                r_count = len(df)
                t_rows = r_count
                preview = df.head(50).to_string(index=False) if r_count > 0 else "Keine Datensätze gefunden (0 Treffer)."
                preview_df = df.head(100)
                clean_records = [
                    [sanitize_cell_value(val) for val in row]
                    for row in preview_df.where(pd.notnull(preview_df), None).values.tolist()
                ]
                return {
                    "step_index": step_num,
                    "label": label,
                    "type": "sql",
                    "query": clean_sql,
                    "row_count": len(clean_records),
                    "total_rows": t_rows,
                    "columns": [str(c) for c in df.columns],
                    "data": clean_records,
                    "preview_text": preview,
                    "success": True
                }
            else:
                log_error_to_db(username=creator, error=sql_error or "SQL-Abfrage fehlgeschlagen", context=f"AI SQL ({label}): {clean_sql}", conn=conn)
                return {
                    "step_index": step_num,
                    "label": label,
                    "type": "sql",
                    "query": clean_sql,
                    "row_count": 0,
                    "total_rows": 0,
                    "columns": [],
                    "data": [],
                    "preview_text": f"SQL-Fehler: {sql_error}",
                    "success": False
                }
        else:
            return {
                "step_index": step_num,
                "label": label,
                "type": "unknown",
                "query": "-- Unbekannter Aktionstyp",
                "row_count": 0,
                "total_rows": 0,
                "columns": [],
                "data": [],
                "preview_text": f"Unbekannter Aktionstyp: '{a_type}'",
                "success": False
            }

    executed_steps = []
    current_actions = actions

    for round_idx in range(1, 3):
        if not current_actions:
            break

        for act in current_actions[:6]:
            step_res = execute_single_action(act, len(executed_steps) + 1)
            executed_steps.append(step_res)

        if round_idx == 1 and requires_followup and executed_steps:
            round1_summary_lines = []
            for s in executed_steps:
                round1_summary_lines.append(f"=== ERGEBNIS SCHRITT {s['step_index']}: {s['label']} ===\n{s['preview_text']}")
            round1_summary = "\n\n".join(round1_summary_lines)

            followup_prompt = f"""Du bist der KI-Agent "Antigravity". Du hast in Runde 1 folgende Teilanalysen ausgeführt:
{round1_summary}

URSPRÜNGLICHE NUTZERFRAGE:
{user_message}

WICHTIGE REGELN FÜR FOLGEABFRAGEN:
- Das Backend filtert die Tabelle `journeys` automatisch auf die berechtigten User und den gewählten Zeitraum. Du musst keine Pflicht-WHERE-Bedingungen für User/Datumsbereiche einbauen, sondern schreibst ganz normale SELECT-Abfragen auf `journeys`.
- Wenn Runde 1 eine Namenssuche (Haltestellen, Betreiber etc.) war und mehrere mehrdeutige Kandidaten gefunden wurden, bei denen unklar ist, welchen der Nutzer meint: Gib `"actions": []` zurück, damit im nächsten Schritt gezielt beim Nutzer nachgefragt wird.

Welche gezielten Folgeabfragen (maximal 4 weitere Actions) sind basierend auf diesen Daten noch notwendig, um die Nutzerfrage vollständig zu beantworten?
Falls bereits alle Daten vorliegen, gib eine leere actions-Liste zurück (`"actions": []`).

Gib deine Antwort wieder als JSON im Planungsformat zurück:
{{
  "actions": [
    {{
      "type": "sql" | "dataframe",
      "label": "Kurzbeschreibung der Folgeabfrage",
      "sql": "SELECT ...",
      "tool_name": "...",
      "tool_args": {{...}}
    }}
  ],
  "requires_followup": false,
  "direct_answer": ""
}}
"""
            followup_resp = call_gemini(client, followup_prompt, json_mode=True)
            current_actions = followup_resp.get("actions") or []
            requires_followup = False
        else:
            break

    # Synthese-Prompt (Finale Antwort)
    if executed_steps:
        steps_summary_parts = []
        for s in executed_steps:
            steps_summary_parts.append(
                f"=== SCHRITT {s['step_index']}: {s['label']} ({s['query']}) ===\n{s['preview_text']}"
            )
        combined_steps_summary = "\n\n".join(steps_summary_parts)

        answer_prompt = f"""Du bist der KI-Datenbank-Assistent für Träwelling-Fahrtauswertungen.
Beantworte die Frage des Nutzers präzise, freundlich und auf Deutsch in Markdown (mit Formatierungen wie Bullet Points, Fettdruck oder Markdown-Tabellen falls sinnvoll).

BISHERIGER KONVERSATIONSVERLAUF:
{history_text}

AKTUELLE FRAGE:
{user_message}

ERGEBNISSE DER AUSGEFÜHRTEN TEILABFRAGEN UND ANALYSEN ({len(executed_steps)} Schritte):
{combined_steps_summary}

WICHTIGE REGELN FÜR DIE ZUSAMMENFASSUNG VON HALTESTELLEN & RICHTUNGEN:
- In der Datenbank existieren für dieselbe physische Haltestelle oft unterschiedliche Schreibweisen (z. B. 'Braunschweig Mittelweg', 'Braunschweig, Mittelweg', 'Mittelweg, Braunschweig', 'BS, ...' oder 'Braunschweig Hbf', 'Braunschweig Hbf (HBS)', 'Hauptbahnhof, Braunschweig' oder Tippfehler-Varianten wie 'Mühlenpfordtstraße' vs. 'Mühlenpfortstraße').
- Addiere für jede Richtung ALLE Schreibweisen derselben physischen Haltestelle zu einer Gesamtsumme zusammen (z. B. 76 + 1 = 77 Fahrten) und weise in einer kurzen Fußnote transparent darauf hin, aus welchen DB-Schreibweisen sich die Summe zusammensetzt.
- Verwechsle NIEMALS eine andere Schreibweise derselben Haltestelle (wie 'Michelfelderplatz, Braunschweig' statt 'Braunschweig Michelfelderplatz') mit der Gegenrichtung! Maßgeblich für die Richtung ist ausschließlich, ob die Haltestelle als Start (`origin` / `start_halt`) oder als Ziel (`destination` / `ziel_halt`) abgefragt wurde.

WICHTIGER HINWEIS ZU ERGEBNISSEN:
- WICHTIG: Wenn eine SQL-Abfrage 'Keine Datensätze gefunden (0 Treffer).' zurückgibt, bedeutet das, dass die Abfrage syntaktisch und technisch FEHLERFREI ausgeführt wurde, aber schlicht keine passenden Fahrten in der Datenbank existieren. Erfinde NIEMALS einen Syntax- oder Datenbankfehler, sondern teile dem Nutzer freundlich mit, dass für diese Kriterien keine Fahrten gefunden wurden.

Formuliere eine fundierte, hervorragend strukturierte deutsche Antwort. Führe alle Datenquellen und statistischen Kennzahlen sinnvoll zusammen und interpretiere sie im Gesamtzusammenhang (z. B. Mobilitätsmuster, Reisezeiten, Geschwindigkeiten, Lieblingsstrecken, Verkehrsmittelwahl, Verspätungsstatistiken)."""

        final_answer = call_gemini(client, answer_prompt, json_mode=False)
    else:
        final_answer = direct_answer or "Ich konnte dazu leider keine passenden Daten in der Datenbank abfragen."

    # Formatierung für DB & Frontend
    clean_sql_parts = []
    for s in executed_steps:
        clean_sql_parts.append(f"-- [{s['step_index']}] {s['label']}\n{s['query']}")
    clean_sql = "\n\n".join(clean_sql_parts)

    total_row_count = sum(s.get("total_rows", 0) for s in executed_steps)

    if executed_steps:
        sql_result_dict = {
            "type": "multi",
            "steps": [
                {
                    "label": s["label"],
                    "query_type": s["type"],
                    "query": s["query"],
                    "row_count": s["row_count"],
                    "total_rows": s["total_rows"],
                    "columns": s["columns"],
                    "data": [
                        [sanitize_cell_value(val) for val in row]
                        for row in s.get("data", [])
                    ]
                }
                for s in executed_steps
            ]
        }
        sql_result_json = json.dumps(sql_result_dict, default=str)
    else:
        sql_result_dict = None
        sql_result_json = None

    # 4. Persistenz
    now_iso = datetime.now().isoformat()

    # Chat-Titel prüfen
    cursor.execute("SELECT title FROM ai_chats WHERE id = ?", (chat_id,))
    chat_row = cursor.fetchone()
    current_title = chat_row[0] if chat_row else 'Neuer Chat'
    if not current_title or current_title == 'Neuer Chat':
        clean_msg = re.sub(r"[\n\r]+", " ", user_message).strip()
        current_title = clean_msg[:36] + ("..." if len(clean_msg) > 36 else "")

    cursor.execute(
        """INSERT INTO ai_chats (id, owner, title, selected_users, start_date, end_date, created_at, updated_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?)
           ON CONFLICT(id) DO UPDATE SET
               title = excluded.title,
               selected_users = excluded.selected_users,
               start_date = excluded.start_date,
               end_date = excluded.end_date,
               updated_at = excluded.updated_at""",
        (chat_id, creator, current_title, json.dumps(selected_users or [creator]),
         start_date or None, end_date or None, now_iso, now_iso)
    )

    # User message speichern
    cursor.execute(
        """INSERT INTO ai_messages (chat_id, role, content, sql_query, row_count, sql_result, created_at)
           VALUES (?, 'user', ?, NULL, NULL, NULL, ?)""",
        (chat_id, user_message, now_iso)
    )

    # Model message speichern
    cursor.execute(
        """INSERT INTO ai_messages (chat_id, role, content, sql_query, row_count, sql_result, created_at)
           VALUES (?, 'model', ?, ?, ?, ?, ?)""",
        (chat_id, final_answer, clean_sql or None, total_row_count if executed_steps else None, sql_result_json, now_iso)
    )

    conn.commit()

    return {
        "success": True,
        "chat_id": chat_id,
        "title": current_title,
        "answer": final_answer,
        "sql": clean_sql,
        "row_count": total_row_count,
        "sql_result": sql_result_dict,
        "tool_used": "multi" if len(executed_steps) > 1 else (executed_steps[0]["label"] if executed_steps else None)
    }


def process_ai_chat_turn(conn: sqlite3.Connection, creator: str, chat_id: str, user_message: str,
                         selected_users: list[str], start_date: str = '', end_date: str = '',
                         is_admin: bool = False) -> dict:
    try:
        return _process_ai_chat_turn_impl(
            conn=conn,
            creator=creator,
            chat_id=chat_id,
            user_message=user_message,
            selected_users=selected_users,
            start_date=start_date,
            end_date=end_date,
            is_admin=is_admin
        )
    except RateLimitExceededError:
        raise
    except Exception as e:
        tb_str = traceback.format_exc()
        log_error_to_db(
            username=creator,
            error=e,
            traceback_str=tb_str,
            context=f"process_ai_chat_turn (chat_id: {chat_id}, message: {user_message[:50]})",
            conn=conn
        )
        return {
            "success": False,
            "error": f"Interner KI-Fehler: {str(e)}",
            "sql_result": None
        }


if __name__ == '__main__':
    start = time()
    conn = init_db()
    api_token = log_onto_traewelling(conn)
    
    print("\n" + "="*50)
    print("🚂 STARTE TRÄWELLING ANALYSE (v10)")
    print("="*50)
    
    fetch_and_analyse(include_stored_users=True, start_date='2026-08-01', end_date='2026-08-31', usernames='zug_pegnitzpfeil_freak')

    print("\n" + "="*50)
    print(f'🏁 FERTIG! Requests übrig: {get_remaining()} | Gesamtdauer: {time() - start:.2f} Sekunden')
    print("="*50)