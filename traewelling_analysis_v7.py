
#traewelling_analysis_v7.py

import csv
import json
import re
from itertools import chain
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from statistics import mean
from time import time, sleep
from zoneinfo import ZoneInfo
from babel import Locale
import pycountry
import requests
import pandas as pd
from timezonefinder import TimezoneFinder
import reverse_geocoder as rg
import traceback
import openpyxl
import jinja2
class Journey():
    def __init__(self, data, tf: TimezoneFinder, timezone_info: dict, coordinates_places_map: dict):
        train_type_to_readable_dict = {'regional': 'Regionalzug', 'bus': 'Bus', 'ferry': 'Fähre',
                                       'national': 'Regionalzug', 'regionalExp': 'Fernverkehr',
                                       'nationalExpress': 'Fernverkehr', 'suburban': 'S-Bahn', 'subway': 'U-Bahn',
                                       'taxi': 'Taxi', 'tram': 'Tram',
                                       'plane': 'Igitt ein Flugzeug'}
        self._raw_data = data
        self._status_id = data['id']
        self._trip_id = data['checkin'].get('trip', 'N/A')
        self._user_name = data['user'].get('username', 'Username not found')
        self._user_id = data['userDetails'].get('id', 'N/A')
        self._trip_reason = data.get('business', 'N/A')
        self._status_text = " ".join((data.get('body', '') or '').splitlines())
        self._status_text_mention = data.get('bodyMentions', [])
        self._number_of_likes = data.get('likes', 0)


        match self._trip_reason:
            case 0:
                self._trip_reason = 'Privat'
            case 1:
                self._trip_reason = 'Geschäftlich'
            case 2:
                self._trip_reason = 'Arbeitsweg'
        train = data.get('train', {})
        o, d = train.get('origin', {}), train.get('destination', {})
        self._origin_stop = (f'{o['name']} ({o['rilIdentifier']})' if o.get('rilIdentifier') else (o['name']
                                                                                                   if o[
                                                                                                          'name'] != 'Braunschweig Hbf ZOB' else 'Braunschweig Hbf (HBS)'))
        self._destination_stop = (f'{d['name']} ({d['rilIdentifier']})' if d.get('rilIdentifier') else (d['name']
                                                                                                        if d[
                                                                                                               'name'] != 'Braunschweig Hbf ZOB' else 'Braunschweig Hbf (HBS)'))

        self._origin_coordinates = (o.get('latitude', 'N/A'), o.get('longitude', 'N/A'))
        if 'N/A' not in self._origin_coordinates:
            self._origin_coordinates = (round(self._origin_coordinates[0], 5), round(self.origin_coordinates[1], 5))
        self._destination_coordinates = (d.get('latitude', 'N/A'), d.get('longitude', 'N/A'))
        if 'N/A' not in self._destination_coordinates:
            self._destination_coordinates = (round(self.destination_coordinates[0], 5),
                                             round(self.destination_coordinates[1], 5))

        self._line_name_raw = train.get('lineName', 'Unknown Line')
        self._line_name = re.sub(r'\s*\(\d+\)', '', self._line_name_raw).strip()

        try:
            self._via_stations_data = train.get('stopovers', [])
            self._via_stations = list(f'{d['name']} ({d['rilIdentifier']})' if d.get('rilIdentifier') else (
                d['name'] if d['name'] != 'Braunschweig Hbf ZOB' else 'Braunschweig Hbf (HBS)')
                                      for d in self._via_stations_data)

            number_origin = self._via_stations.index(self._origin_stop)
            number_destination = self._via_stations.index(self._destination_stop, number_origin + 1)
            self._via_stations = self._via_stations[number_origin + 1:number_destination]
            self._number_of_inter_stations = len(self._via_stations)
        except Exception as e:
            self._via_stations = []
            self._number_of_inter_stations = len(self._via_stations)
            print(f'Could not get inter stations for {self._line_name}: {e}')

        map_origin = coordinates_places_map.get(self.origin_coordinates, {})
        map_dest = coordinates_places_map.get(self.destination_coordinates, {})

        self._country_name_origin = map_origin.get('country_name', None)
        self._country_name_dest = map_dest.get('country_name', None)
        self._country_traveled = f'{self._country_name_origin} -> {self._country_name_dest}'
        self._cc_origin = map_origin.get('cc', None)
        self._cc_dest = map_dest.get('cc', None)
        self._city_origin = map_origin.get('name', None)
        self._city_dest = map_dest.get('name', None)
        self._admin1_origin = map_origin.get('admin1', None)
        self._admin1_dest = map_dest.get('admin1', None)
        self._admin2_origin = map_origin.get('admin2', None)
        self._admin2_dest = map_dest.get('admin2', None)

        if self._country_name_dest == self._country_name_origin:
            self._border_crossing = False
        else:
            self._border_crossing = True

        self._timezone_origin = timezone_info.get(self._origin_stop, None)
        self._timezone_destination = timezone_info.get(self._destination_stop, None)

        if self._timezone_origin is None:
            if 'N/A' not in self._origin_coordinates:
                self._timezone_origin = tf.timezone_at(lng=self._origin_coordinates[1], lat=self._origin_coordinates[0])
            else:
                self._timezone_origin = "Europe/Berlin"
        if self._timezone_destination is None:
            if 'N/A' not in self._destination_coordinates:
                self._timezone_destination = tf.timezone_at(lng=self._destination_coordinates[1],
                                                            lat=self._destination_coordinates[0])
            else:
                self._timezone_destination = "Europe/Berlin"

        self._operator_info = train.get('operator', None)
        if self._operator_info:
            self._operator_name = self._operator_info.get('name', 'Unknown Operator')
            self._operator_id = self._operator_info.get('id', None)
        else:
            self._operator_name = 'Unknown Operator'
            self._operator_id = None

        self._journey_distance = train.get('distance', 0)
        self._line_number = train.get('number')

        raw_train_type = train.get('category', None)
        given_train_type = train_type_to_readable_dict.get(raw_train_type, raw_train_type)
        if re.match(r"^(REX|RE|RB|Stoptrein|FEX|Sprinter)", self._line_name) and given_train_type != 'Bus':
            self._train_type = 'Regionalzug'
        elif re.match(r"^(RS)", self._line_name):
            self._train_type = 'S-Bahn'
        else:
            self._train_type = given_train_type

        self._journey_points = train.get('points', 0)
        self._journey_number = train.get('journeyNumber')

        self._planned_departure = datetime.fromisoformat(o.get('departurePlanned')).astimezone(
            ZoneInfo(self._timezone_origin))
        self._planned_arrival = datetime.fromisoformat(d.get('arrivalPlanned')).astimezone(
            ZoneInfo(self._timezone_destination))
        manualDeparture = train.get('manualDeparture')
        manualArrival = train.get('manualArrival')
        realtimeDeparture = o.get('departureReal')
        realtimeArrival = d.get('arrivalReal')

        self._realtime_availability = False
        if manualDeparture is not None:
            self._realtime_availability = True
            self._real_departure = datetime.fromisoformat(manualDeparture).astimezone(ZoneInfo(self._timezone_origin))
        elif realtimeDeparture is not None:
            self._realtime_availability = True
            self._real_departure = datetime.fromisoformat(realtimeDeparture).astimezone(ZoneInfo(self._timezone_origin))
        else:
            self._real_departure = self._planned_departure

        if manualArrival is not None:
            self._realtime_availability = True
            self._real_arrival = datetime.fromisoformat(manualArrival).astimezone(ZoneInfo(self._timezone_destination))
        elif realtimeArrival is not None:
            self._realtime_availability = True
            self._real_arrival = datetime.fromisoformat(realtimeArrival).astimezone(
                ZoneInfo(self._timezone_destination))
        else:
            self._real_arrival = self._planned_arrival

        self._departure_delay = (self._real_departure - self._planned_departure).total_seconds()
        self._arrival_delay = (self._real_arrival - self._planned_arrival).total_seconds()

        self._vehicle_number = 'N/A'
        self._ticket_used = 'N/A'
        self._wagon_class = 'N/A'
        self._locomotive_class = 'N/A'
        self._tags = data.get('tags', [])
        if len(self._tags) > 0:
            for tag in self._tags:
                match tag.get('key', None):
                    case 'trwl:vehicle_number':
                        self._vehicle_number = tag.get('value', 'N/A')
                    case 'trwl:journey_number':
                        self._journey_number = tag.get('value', 'N/A')
                    case 'trwl:locomotive_class':
                        self._locomotive_class = tag.get('value', 'N/A')
                    case 'trwl:ticket':
                        self._ticket_used = tag.get('value', 'N/A')
                    case 'trwl:wagon_class':
                        self._wagon_class = tag.get('value', 'N/A')

        self._journey_time_planned = (self._planned_arrival - self._planned_departure).total_seconds()
        self._journey_time_real = (self._real_arrival - self._real_departure).total_seconds()
        self._journey_time_delta = (self._journey_time_real - self._journey_time_planned)

    def __str__(self):
        string_base =(f'Fahrt {self._line_name} ({self._journey_number}) von {self._origin_stop} nach {self._destination_stop} '
                     f'am {self._planned_departure.strftime("%d.%m.%Y")} ' )
        match self.arrival_delay:
            case d if d>0:
                return string_base + f'mit {self._arrival_delay / 60:.2f} Minuten Verspätung, Fzg: {self._vehicle_number}'
            case d if d<0:
                return string_base + f'mit {-self._arrival_delay / 60:.2f} Minuten Verfrühung, Fzg: {self._vehicle_number}'
            case _:
                return string_base + f'war pünktlich, Fzg: {self._vehicle_number}'

    def __lt__(self, other):
        if isinstance(other, Journey):
            return self._real_arrival < other._real_arrival
        elif isinstance(other, datetime):
            return self._real_departure < other
        return None

    def delayed_by_standard(self, standard: int)->bool:
        return self._arrival_delay >= standard


    @property
    def user_id(self):
        return self._user_id
    @property
    def trip_id(self):
        return self._trip_id
    @property
    def status_id(self):
        return self._status_id
    @property
    def user_name(self):
        return self._user_name
    @property
    def operator_name(self):
        return self._operator_name
    @property
    def line_name(self):
        return self._line_name
    @property
    def journey_distance(self):
        return self._journey_distance
    @property
    def train_type(self):
        return self._train_type
    @property
    def arrival_delay(self):
        return self._arrival_delay
    @property
    def departure_delay(self):
        return self._departure_delay
    @property
    def arrival_planned(self):
        return self._planned_arrival
    @property
    def departure_planned(self):
        return self._planned_departure
    @property
    def departure_real(self):
        return self._real_departure
    @property
    def arrival_real(self):
        return self._real_arrival
    @property
    def realtime_availability(self):
        return self._realtime_availability
    @property
    def vehicle_number(self):
        return self._vehicle_number
    @property
    def origin_stop(self):
        return self._origin_stop
    @property
    def destination_stop(self):
        return self._destination_stop
    @property
    def journey_number(self):
        return self._journey_number
    @property
    def journey_time_real(self):
        return self._journey_time_real
    @property
    def journey_time_delta(self):
        return self._journey_time_delta
    @property
    def journey_points(self):
        return self._journey_points
    @property
    def origin_coordinates(self):
        return self._origin_coordinates
    @property
    def destination_coordinates(self):
        return self._destination_coordinates
    @property
    def timezone_origin(self):
        return self._timezone_origin
    @property
    def timezone_destination(self):
        return self._timezone_destination
    @property
    def via_stations(self):
        return self._via_stations
    @property
    def trip_reason(self):
        return self._trip_reason
    @property
    def status_text(self):
        return self._status_text
    @property
    def number_of_inter_stations(self):
        return self._number_of_inter_stations
    @property
    def origin_country(self):
        return self._origin_country
    @property
    def destination_country(self):
        return self._destination_country
    @property
    def journey_time_planned(self):
        return self._journey_time_planned
    @property
    def cc_origin(self):
        return self._cc_origin
    @property
    def cc_dest(self):
        return self._cc_dest
    @property
    def country_name_origin(self):
        return self._country_name_origin
    @property
    def country_name_dest(self):
        return self._country_name_dest
    @property
    def city_origin(self):
        return self._city_origin
    @property
    def city_dest(self):
        return self._city_dest
    @property
    def admin1_origin(self):
        return self._admin1_origin
    @property
    def admin1_dest(self):
        return self._admin1_dest
    @property
    def admin2_origin(self):
        return self._admin2_origin
    @property
    def admin2_dest(self):
        return self._admin2_dest
    @property
    def country_traveled(self):
        return self._country_traveled
    @property
    def border_crossing(self):
        return self._border_crossing
    @property
    def number_of_likes(self):
        return self._number_of_likes






class User:
    def __init__(self, id: int, delay_standard : int)->None:
        self.__delay_standard = delay_standard
        self.__id : int = id
        self.__name: str | None = None
        self.__meta: dict | None = None
        self.__journeys : list[Journey] = []
        self.__number_of_journeys: int = 0
        self.__exported_days:int = 0
        self.__total_distance: int = 0
        self.__total_journey_time: int = 0
        self.__distance_whole: dict|None= None
        self.__distance_type_sorted: dict|None= None
        self.__distance_operator_line_sorted = None
        self.__distance_operator_sorted = None
        self.__distance_reason_sorted: dict|None= None
        self.__average_delay: float | None = None
        self.__realtime_availability: float | None = None
        self.__cumulative_delay: int | None = None
        self.__delay_rate_standard: float | None = None
        self.__visited_stations: dict[str,int] = dict()
        self.__stations_with_via: dict[str, int] = dict()
        self.__used_vehicles: dict[tuple,int] = dict()
        self.__number_of_visited_stations: int = 0
        self.__number_of_visited_stations_with_via: int = 0
        self.__sorted_countries = {}
        self.__sorted_cities = {}
        self.__sorted_admin1 = {}
        self.__sorted_admin2 = {}


    def add_journey(self, journey: Journey)->None:
        self.__journeys.append(journey)

    def __repr__(self) -> str:
        return self.__name

    def __str__(self) -> str:
        return self.__name

    def __eq__(self, other)->bool:
        if isinstance(other,User):
            return self.__id == other.__id
        else:
            print(f'Kein Vergleich von {self.__name} mit {type(other).__name__} möglich')
            return NotImplemented

    def __hash__(self):
        return hash(self.__id)

    def __lt__(self, other)->bool:
        if isinstance(other,User):
            return len(self.__journeys) < len(other.__journeys)
        elif isinstance(other,int):
            return len(self.__journeys) < other
        else:
            print(f'Kein Vergleich von {self.__name} mit {type(other).__name__} möglich')
            return False

    def get_import_length(self, all_meta: dict) -> None:
        self.__name = self.__journeys[-1].user_name
        self.__meta = all_meta.get(self.__name, {})
        date_list = set(d for journey in self.__journeys for d in (journey.arrival_real.date(), journey.departure_real.date()))
        date_list = sorted(date_list)
        first_day = date_list[0]
        last_day = date_list[-1]
        gaps = ((d2-d1).days - 1 for d1 , d2 in zip(date_list[:-1], date_list[1:]))
        total_time_skipped = sum(gap if gap > 30 else 0 for gap in gaps)

        self.__exported_days = (last_day- first_day).days + 1 - total_time_skipped
        print(f'In {self.__exported_days} Tagen hat {self.__name} {len(self.__journeys)} Fahrten gemacht. '
              f'Das sind {len(self.__journeys) / self.__exported_days:.2f} Fahrten pro Tag!\n')
        if total_time_skipped > 0:
            print(f'Es wurden Fahrten von {first_day.strftime("%d.%m.%Y")} bis '
              f'{last_day.strftime("%d.%m.%Y")} berücksichtigt ({total_time_skipped} Tage mit über einer Inaktivität von über 30 Tagen übersprungen)')
        else:
            print(f'Es wurden Fahrten von {first_day.strftime("%d.%m.%Y")} bis '
                  f'{last_day.strftime("%d.%m.%Y")} berücksichtigt')
        print(f'\n{50 * "-"}\n')
        self.__number_of_journeys = len(self.__journeys)
        self.__total_distance = sum(journey.journey_distance for journey in self.__journeys)
        self.__total_journey_time = sum(journey.journey_time_real for journey in self.__journeys)

    def user_geo_analysis(self) -> None:
        numbers_of_countries = defaultdict(int)
        numbers_of_cities = defaultdict(int)
        numbers_of_admin1 = defaultdict(int)
        numbers_of_admin2 = defaultdict(int)

        for j in self.__journeys:
            numbers_of_countries[j.country_name_origin] += 1
            numbers_of_countries[j.country_name_dest] += 1
            numbers_of_cities[j.city_origin] += 1
            numbers_of_cities[j.city_dest] += 1
            numbers_of_admin1[j.admin1_origin] += 1
            numbers_of_admin1[j.admin1_dest] += 1
            numbers_of_admin2[j.admin2_origin] += 1
            numbers_of_admin2[j.admin2_dest] += 1
        self.__sorted_countries = dict(sorted(numbers_of_countries.items(), key=lambda x: x[1], reverse=True))
        self.__sorted_cities = dict(sorted(numbers_of_cities.items(), key=lambda x: x[1], reverse=True))
        self.__sorted_admin1 = dict(sorted(numbers_of_admin1.items(), key=lambda x: x[1], reverse=True))
        self.__sorted_admin2 = dict(sorted(numbers_of_admin2.items(), key=lambda x: x[1], reverse=True))




    def user_distance_time_analysis_execute(self):
        self.__distance_whole = {
            'distance': 0, 'time': 0, 'arrival_delay': 0,
            'realtime_availability': 0, 'sum': len(self.__journeys), 'delayed_by_standard': 0,
            'points': 0, 'all_delay_arr': [], 'inter_stops': 0}
        distance_type: dict = {}
        distance_operator: dict = {}
        distance_operator_line: dict = {}
        distance_reason : dict = {}


        for journey in self.__journeys:
            journey_distance = journey.journey_distance
            journey_time = journey.journey_time_real
            journey_arrival_delay = max(journey.arrival_delay,0)
            journey_realtime_availability = journey.realtime_availability
            journey_delayed_by_standard = journey.delayed_by_standard(self.__delay_standard)
            journey_points = journey.journey_points
            journey_stops = journey.number_of_inter_stations
            train_type = journey.train_type
            operator_name = journey.operator_name
            line_tuple = (operator_name, journey.line_name)
            trip_reason = journey.trip_reason

            self.__distance_whole['distance'] += journey_distance
            self.__distance_whole['time'] += journey_time
            self.__distance_whole['arrival_delay'] += journey_arrival_delay
            self.__distance_whole['realtime_availability'] += journey_realtime_availability
            self.__distance_whole['delayed_by_standard'] += journey_delayed_by_standard
            self.__distance_whole['points'] += journey_points
            self.__distance_whole['all_delay_arr'].append(journey_arrival_delay)
            self.__distance_whole['inter_stops'] += journey_stops

            distance_type[train_type] = {
                'distance': journey_distance + distance_type.get(train_type, {}).get('distance', 0),
                'time': journey_time + distance_type.get(train_type, {}).get('time', 0),
                'arrival_delay': journey_arrival_delay + distance_type.get(train_type, {}).get('arrival_delay', 0),
                'realtime_availability': journey_realtime_availability + distance_type.get(train_type, {}).get('realtime_availability', 0),
                'sum': 1 + distance_type.get(train_type, {}).get('sum', 0),
                'delayed_by_standard': journey_delayed_by_standard + distance_type.get(train_type, {}).get('delayed_by_standard', 0),
                'points': journey_points + distance_type.get(train_type, {}).get('points', 0),
                'all_delay_arr': distance_type.get(train_type, {}).get('all_delay_arr', []) + [journey_arrival_delay],
                'inter_stops': journey_stops + distance_type.get(train_type, {}).get('inter_stops', 0)}

            distance_operator_line[line_tuple] = {'distance': journey_distance + distance_operator_line.get(line_tuple, {}).get('distance', 0),
                                                  'time': journey_time + distance_operator_line.get(line_tuple, {}).get('time', 0),
                                                  'arrival_delay': journey_arrival_delay + distance_operator_line.get(line_tuple, {}).get('arrival_delay', 0),
                                                  'realtime_availability': journey_realtime_availability + distance_operator_line.get(line_tuple, {}).get('realtime_availability', 0),
                                                  'sum': 1 + distance_operator_line.get(line_tuple, {}).get('sum', 0),
                                                  'delayed_by_standard': journey_delayed_by_standard + distance_operator_line.get(line_tuple, {}).get('delayed_by_standard', 0),
                                                  'points': journey_points + distance_operator_line.get(line_tuple, {}).get('points', 0),
                                                  'all_delay_arr': distance_operator_line.get(line_tuple, {}).get('all_delay_arr', []) + [journey_arrival_delay],
                                                  'inter_stops': journey_stops + distance_operator_line.get(line_tuple, {}).get('inter_stops', 0)}

            distance_operator[operator_name] = {'distance': journey_distance + distance_operator.get(operator_name, {}).get('distance', 0),
                                                'time': journey_time + distance_operator.get(operator_name, {}).get('time', 0),
                                                'arrival_delay': journey_arrival_delay + distance_operator.get(operator_name, {}).get('arrival_delay', 0),
                                                'realtime_availability': journey_realtime_availability + distance_operator.get(operator_name, {}).get('realtime_availability', 0),
                                                'sum': 1 + distance_operator.get(operator_name, {}).get('sum', 0),
                                                'delayed_by_standard': journey_delayed_by_standard + distance_operator.get(operator_name, {}).get('delayed_by_standard', 0),
                                                'points': journey_points + distance_operator.get(operator_name, {}).get('points', 0),
                                                'all_delay_arr': distance_operator.get(operator_name, {}).get('all_delay_arr', []) + [journey_arrival_delay],
                                                'inter_stops': journey_stops + distance_operator.get(operator_name, {}).get('inter_stops', 0)}

            distance_reason[trip_reason] = {
                'distance': journey_distance + distance_reason.get(trip_reason, {}).get('distance', 0),
                'time': journey_time + distance_reason.get(trip_reason, {}).get('time', 0),
                'arrival_delay': journey_arrival_delay + distance_reason.get(trip_reason, {}).get('arrival_delay',0),
                'realtime_availability': journey_realtime_availability + distance_reason.get(trip_reason, {}).get('realtime_availability', 0),
                'sum': 1 + distance_reason.get(trip_reason, {}).get('sum', 0),
                'delayed_by_standard': journey_delayed_by_standard + distance_reason.get(trip_reason, {}).get('delayed_by_standard', 0),
                'points': journey_points + distance_reason.get(trip_reason, {}).get('points', 0),
                'all_delay_arr': distance_reason.get(trip_reason, {}).get('all_delay_arr', []) + [journey_arrival_delay],
                'inter_stops': journey_stops + distance_reason.get(trip_reason, {}).get('inter_stops', 0)}




        self.__distance_type_sorted = dict(sorted(distance_type.items(), key=lambda x: x[1].get('distance', 0), reverse=True))
        self.__distance_operator_line_sorted = dict(sorted(distance_operator_line.items(), key=lambda x: x[1].get('distance', 0), reverse=True))
        self.__distance_operator_sorted = dict(sorted(distance_operator.items(), key=lambda x: x[1].get('distance', 0), reverse=True))
        self.__distance_reason_sorted = dict(sorted(distance_reason.items(), key=lambda x: x[1].get('distance', 0), reverse=True))
        self.__cumulative_delay = self.__distance_whole['arrival_delay']
        self.__average_delay = self.__cumulative_delay / self.__distance_whole['sum']
        self.__realtime_availability = self.__distance_whole['realtime_availability'] / self.__distance_whole['sum']
        self.__delay_rate_standard = self.__distance_whole['delayed_by_standard'] / self.__distance_whole['sum']



    def visited_station_execution(self):
        for j in self.__journeys:
            self.__visited_stations[j.origin_stop] = self.__visited_stations.get(j.origin_stop, 0) + 1
            self.__visited_stations[j.destination_stop] = self.__visited_stations.get(j.destination_stop, 0) + 1
        self.__number_of_visited_stations = len(self.__visited_stations)
        self.__visited_stations = dict(sorted(self.__visited_stations.items(), key=lambda x: x[1], reverse=True))
        self.__stations_with_via = self.__visited_stations.copy()
        for j in self.__journeys:
            for name in j.via_stations:
                self.__stations_with_via[name] = self.__stations_with_via.get(name, 0) + 1
        self.__number_of_visited_stations_with_via = len(self.__stations_with_via)
        self.__stations_with_via = dict(sorted(self.__stations_with_via.items(), key=lambda x: x[0], reverse=False))
        self.__stations_with_via = dict(sorted(self.__stations_with_via.items(), key=lambda x: x[1], reverse=True))


    def vehicle_execution(self):
        for j in self.__journeys:
            vehicle_ident = (j.vehicle_number,j.operator_name)
            self.__used_vehicles[vehicle_ident] = self.__used_vehicles.get(vehicle_ident, 0) + 1
        self.__used_vehicles = dict(sorted(self.__used_vehicles.items(), key=lambda x: x[1], reverse=True))




    def create_excel(self) -> None:
        j_data: list[dict] = []
        for j in self.__journeys:
            journey_hours, journey_minutes = divmod(max(j.journey_time_real / 60, 0), 60)
            j_data.append({'Kategorie': j.train_type,
                          'Linie': j.line_name,
                          'Fahrtnummer': j.journey_number,
                          'Betreiber': j.operator_name,
                          'Grund': j.trip_reason,
                          'Fahrzeugnummer': j.vehicle_number,
                          'Abfahrtshaltestelle': j.origin_stop,
                          'Abfahrt geplant': j.departure_planned.strftime('%d.%m.%Y %H:%M'),
                          'Abfahrt real': j.departure_real.strftime('%d.%m.%Y %H:%M') if j.realtime_availability else 'N/A',
                          'Abw ab': int(round(j.departure_delay / 60, 0)),
                          'Anzahl Zwischenhalte': j.number_of_inter_stations,
                          'Zielstation': j.destination_stop,
                          'Ankunft geplant': j.arrival_planned.strftime('%d.%m.%Y %H:%M'),
                          'Ankunft real': j.arrival_real.strftime('%d.%m.%Y %H:%M') if j.realtime_availability else 'N/A',
                          'Abw an': int(round(j.arrival_delay / 60, 0)),
                          'Reisezeit plan (min)': int(round(j.journey_time_planned / 60, 0)),
                          'Reisezeit (min)': int(round(j.journey_time_real / 60, 0)),
                          'Reisezeit (hh:mm)': f"{int(journey_hours):02}:{int(journey_minutes):02}",
                          'Delta': int(round(j.journey_time_delta / 60, 0)),
                          'Entfernung (m)': j.journey_distance,
                          'Entfernung (km)': round(j.journey_distance / 1000, 3),
                          'Geschwindigkeit (km/h)': (round(j.journey_distance / (j.journey_time_real / 3.6), 3) if j.journey_time_real else 'N/A'),
                          'Punkte': j.journey_points,
                          'Likes': j.number_of_likes,
                          'Link': f'=HYPERLINK("https://traewelling.de/status/{j.status_id}", "https://traewelling.de/status/{j.status_id}")',
                          'Statustext': j.status_text,
                          'Bereiste Länder': j.country_traveled,
                          'Grenzübersreitend': j.border_crossing})

        journeys_dataframe = pd.DataFrame(j_data)


        total_hours, total_minutes = divmod(self.__total_journey_time / 60, 60)
        per_day_hours, per_day_minutes = divmod((self.__total_journey_time / self.__exported_days) / 60, 60)
        per_journey_hours, per_journey_minutes = divmod((self.__total_journey_time / self.__number_of_journeys) / 60, 60)
        stats_data = {'Metrik': ['Gesamtwert', 'Durchschnitt pro Tag', 'Durchschnitt pro Fahrt'],
                      'Fahrten': [f'{len(self.__journeys)} Fahrten',
                                  f'{len(self.__journeys) / self.__exported_days:.3f} Fahrten',
                                  f'{len(self.__journeys) / self.__number_of_journeys:.3f} Fahrten'],
                      'Tage':[f'{self.__exported_days} Tage',
                              f'{self.__exported_days / self.__exported_days:.3f} Tage',
                              f'{self.__exported_days / self.__number_of_journeys:.3f} Tage'],
                      'Linien':[f'{len(self.distance_operator_line_sorted)} verschiedene Linien',
                                f'{len(self.distance_operator_line_sorted) / self.__exported_days:.3f} verschiedene Linien',
                                f'{len(self.distance_operator_line_sorted) / self.__number_of_journeys:.3f} verschiedene Linien'],
                      'Betreiber':[f'{len(self.distance_operator_sorted)} verschiedene Betreiber',
                                   f'{len(self.distance_operator_sorted) / self.__exported_days:.3f} verschiedene Betreiber',
                                   f'{len(self.distance_operator_sorted) / self.__number_of_journeys:.3f} verschiedene Betreiber'],
                      'Haltestellen':[f'{self.number_of_visited_stations} einmalige Haltestellen',
                                      f'{self.number_of_visited_stations / self.__exported_days:.3f} einmalige Haltestellen',
                                      f'{self.number_of_visited_stations / self.__number_of_journeys:.3f} einmalige Haltestellen'],
                      'Abfahrtsverspätung':[f'{sum(j.departure_delay for j in self.__journeys) / 60:.0f} Minuten',
                                            f'{(sum(j.departure_delay for j in self.__journeys) / 60) / self.__exported_days :.3f} Minuten',
                                            f'{(sum(j.departure_delay for j in self.__journeys) / 60) / self.__number_of_journeys :.3f} Minuten'],
                      'Ankunftsverspätung':[f'{sum(j.arrival_delay for j in self.__journeys) / 60:.0f} Minuten',
                                            f'{(sum(j.arrival_delay for j in self.__journeys) / 60) / self.__exported_days :.3f} Minuten',
                                            f'{(sum(j.arrival_delay for j in self.__journeys) / 60) / self.__number_of_journeys :.3f} Minuten'],
                      'Zeit (min)':[f'{int(self.__total_journey_time / 60)} Minuten',
                                          f'{self.__total_journey_time / 60 / self.__exported_days :.3f} Minuten',
                                          f'{self.__total_journey_time / self.__number_of_journeys / 60:.3f} Minuten'],
                      'Zeit (hh:mm)':[f"{total_hours:02.0f}:{int(round(total_minutes, 0)):02}",
                                            f"{per_day_hours:02.0f}:{round(per_day_minutes, 0):02.0f}",
                                            f"{per_journey_hours:02.0f}:{round(per_journey_minutes, 0):02.0f}"],
                      'Distanz (m)':[f'{self.__total_distance} Meter',
                                           f'{self.__total_distance / self.__exported_days :.3f} Meter',
                                           f'{self.__total_distance / self.__number_of_journeys :.3f} Meter'],
                      'Distanz (km)':[f'{self.__total_distance / 1000:.2f} Kilometer',
                                            f'{self.__total_distance / 1000 / self.__exported_days:.2f} Kilometer',
                                            f'{self.__total_distance / 1000 / self.__number_of_journeys:.2f} Kilometer'],
                      'Punkte':[f'{sum(j.journey_points for j in self.__journeys)} Punkte',
                                f'{sum(j.journey_points for j in self.__journeys) / self.__exported_days :.3f} Punkte',
                                f'{sum(j.journey_points for j in self.__journeys) / self.__number_of_journeys :.3f} Punkte'],
                      'Likes':[f'{sum(j.number_of_likes for j in self.__journeys)} Likes',
                                f'{sum(j.number_of_likes for j in self.__journeys) / self.__exported_days :.3f} Likes',
                                f'{sum(j.number_of_likes for j in self.__journeys) / self.__number_of_journeys :.3f} Likes'],
                      'Geschwindigkeit': [f'{self.__total_distance / (self.__total_journey_time / 3.6):.2f} km/h', f'', f''],
                      'Echtzeitquote': [f'{self.realtime_availability * 100:.2f}%', f'', f''],
                      f'Mehr als {round(self.__delay_standard/60)} Minuten verspätet': [f'{self.delay_rate_standard*100 :.2f}%',
                                        f'',
                                        f'',],
                      '\u2205 Zwischenhalte':[f'{round(mean(j.number_of_inter_stations for j in self.journeys),3)} Zwischenhalte',f'',f'']}

        stats_dataframe = pd.DataFrame(stats_data)

        type_data = []

        for type_name, type_info in self.distance_type_sorted.items():
            type_data.append({f'Kategorie ({len(self.distance_type_sorted)})': type_name,
                              'Fahrten': type_info.get('sum', 0),
                              'Distanz (km)': round(type_info.get('distance', 0) / 1000, 2),
                              'Anteil an Gesamtdistanz': f'{type_info.get('distance', 0) / self.__total_distance * 100:.2f}%',
                              'Zeit (min)': int(type_info.get('time', 0) / 60),
                              'Zeit (hh:mm)': f'{int(divmod(type_info.get('time', 0) / 60, 60)[0]):02}:{int(divmod(type_info.get('time', 0) / 60, 60)[1]):02}',
                              'Anteil an Gesamtzeit': f'{type_info.get('time', 0) / self.__total_journey_time * 100:.2f}%',
                              '\u2205 Geschwindigkeit (km/h)': round((type_info.get('distance', 0) / 1000) / (
                                      type_info.get('time', 0) / 3600) if type_info.get('time', 0) else 0, 2),
                              '\u2205 Verspätung (min)': round(
                                  type_info.get('arrival_delay', 0) / type_info.get('sum', 0) / 60, 2),
                              'Pünktlichkeitsquote (%)': round(
                                  100 - (type_info.get('delayed_by_standard', 0) / type_info.get('sum', 0) * 100),
                                  2),
                              'Punkte': type_info.get('points', 0),
                              '\u2205 Zwischenhalte': round(type_info.get('inter_stops', 0) / type_info.get('sum', 0),3)})

        type_dataframe = pd.DataFrame(type_data)

        operator_data = []

        for operator_name, operator_info in self.distance_operator_sorted.items():
            operator_data.append({f'Betreiber ({len(self.distance_operator_sorted)})': operator_name,
                                  'Fahrten': operator_info.get('sum', 0),
                                  'Distanz (km)': round(operator_info.get('distance', 0) / 1000, 2),
                                  'Anteil an Gesamtdistanz': f'{operator_info.get('distance', 0) / self.__total_distance * 100:.2f}%',
                                  'Zeit (min)': int(operator_info.get('time', 0) / 60),
                                  'Zeit (hh:mm)': f'{int(divmod(operator_info.get('time', 0) / 60, 60)[0]):02}:{int(divmod(operator_info.get('time', 0) / 60, 60)[1]):02}',
                                  'Anteil an Gesamtzeit': f'{operator_info.get('time', 0) / self.__total_journey_time * 100:.2f}%',
                                  '\u2205 Geschwindigkeit (km/h)': round((operator_info.get('distance', 0) / 1000) / (operator_info.get('time', 0) / 3600) if operator_info.get('time', 0) else 0, 2),
                                  '\u2205 Verspätung (min)': round(operator_info.get('arrival_delay', 0) / operator_info.get('sum', 0) / 60, 2),
                                  'Pünktlichkeitsquote (%)': round(100 - (operator_info.get('delayed_by_standard', 0) / operator_info.get('sum', 0) * 100), 2),
                                  'Punkte': operator_info.get('points', 0),
                                  '\u2205 Zwischenhalte': round(operator_info.get('inter_stops', 0) / operator_info.get('sum', 0),3)})

        operator_dataframe = pd.DataFrame(operator_data)

        line_data = []

        for line_name, line_info in self.distance_operator_line_sorted.items():
            line_data.append({f'Linie ({len(self.distance_operator_line_sorted)})': line_name[1],
                              'Betreiber': line_name[0],
                              'Fahrten': line_info.get('sum', 0),
                              'Distanz (km)': round(line_info.get('distance', 0) / 1000, 2),
                              'Anteil an Gesamtdistanz': f'{line_info.get('distance', 0) / self.__total_distance * 100:.2f}%',
                              'Zeit (min)': int(line_info.get('time', 0) / 60),
                              'Zeit (hh:mm)': f'{int(divmod(line_info.get('time', 0) / 60, 60)[0]):02}:{int(divmod(line_info.get('time', 0) / 60, 60)[1]):02}',
                              'Anteil an Gesamtzeit': f'{line_info.get('time', 0) / self.__total_journey_time * 100:.2f}%',
                              '\u2205 Geschwindigkeit (km/h)': round((line_info.get('distance', 0) / 1000) / (
                                          line_info.get('time', 0) / 3600) if line_info.get('time', 0) else 0, 2),
                              '\u2205 Verspätung (min)': round(
                                  line_info.get('arrival_delay', 0) / line_info.get('sum', 0) / 60, 2),
                              'Pünktlichkeitsquote (%)': round(
                                  100 - (line_info.get('delayed_by_standard', 0) / line_info.get('sum', 0) * 100), 2),
                              'Punkte': line_info.get('points', 0),
                              '\u2205 Zwischenhalte': round(line_info.get('inter_stops', 0) / line_info.get('sum', 0),3)})

        line_dataframe = pd.DataFrame(line_data)

        reason_data = []

        for reason_name, reason_info in self.distance_reason_sorted.items():
            reason_data.append({f'Fahrtgrund ({len(self.distance_reason_sorted)})': reason_name,
                                'Fahrten': reason_info.get('sum', 0),
                                'Distanz (km)': round(reason_info.get('distance', 0) / 1000, 2),
                                'Anteil an Gesamtdistanz': f'{reason_info.get('distance', 0) / self.__total_distance * 100:.2f}%',
                                'Zeit (min)': int(reason_info.get('time', 0) / 60),
                                'Zeit (hh:mm)': f'{int(divmod(reason_info.get('time', 0) / 60, 60)[0]):02}:{int(divmod(reason_info.get('time', 0) / 60, 60)[1]):02}',
                                'Anteil an Gesamtzeit': f'{reason_info.get('time', 0) / self.__total_journey_time * 100:.2f}%',
                                '\u2205 Geschwindigkeit (km/h)': round((reason_info.get('distance', 0) / 1000) / (
                                        reason_info.get('time', 0) / 3600) if reason_info.get('time', 0) else 0, 2),
                                '\u2205 Verspätung (min)': round(
                                    reason_info.get('arrival_delay', 0) / reason_info.get('sum', 0) / 60, 2),
                                'Pünktlichkeitsquote (%)': round(
                                    100 - (reason_info.get('delayed_by_standard', 0) / reason_info.get('sum', 0) * 100),
                                    2),
                                'Punkte': reason_info.get('points', 0),
                                '\u2205 Zwischenhalte': round(reason_info.get('inter_stops', 0) / reason_info.get('sum', 0),3)})

        reason_dataframe = pd.DataFrame(reason_data)

        stop_data = []

        current_place, place_previous = 0, 0
        visited_previous = list(self.visited_stations.values())[0] + 1

        for stop, times_visited in self.visited_stations.items():
            current_place += 1
            if times_visited != visited_previous:
                place_previous = current_place
            visited_previous = times_visited
            stop_data.append({'Platz': place_previous,
                              f'Haltestelle ({self.__number_of_visited_stations})': stop,
                              'Besuche': times_visited})

        stop_dataframe = pd.DataFrame(stop_data)

        stop_with_via_data = []

        current_place, place_previous = 0, 0
        visited_previous = list(self.stations_with_via.values())[0] + 1

        for stop, times_visited in self.stations_with_via.items():
            current_place += 1
            if times_visited != visited_previous:
                place_previous = current_place
            visited_previous = times_visited
            stop_with_via_data.append({'Platz': place_previous,
                                        f'Haltestelle ({self.__number_of_visited_stations_with_via})': stop,
                                        'Anzahl ohne Via': self.visited_stations.get(stop,0),
                                        'Anzahl mit Via': times_visited,
                                        'Umstiegquote (%)': round(self.visited_stations.get(stop,0)/times_visited * 100, 2)})

        stop_with_via_dataframe = pd.DataFrame(stop_with_via_data)

        vehicle_data = []
        for (vehicle, operator), number_used in self.used_vehicles.items():
            vehicle_data.append({f'Fahrzeug ({len(self.used_vehicles)})': vehicle,
                                 'Betreiber': operator,
                                 'Fahrten': number_used})
        vehicle_dataframe = pd.DataFrame(vehicle_data)

        country_data = []
        for country, visited in self.sorted_countries.items():
            country_data.append({f'Land ({len(self.sorted_countries)})': country,
                                 'Besuche': visited})
        country_dataframe = pd.DataFrame(country_data)

        city_data = []
        for city, visited in self.sorted_cities.items():
            city_data.append({f'Stadt ({len(self.sorted_cities)})': city,
                              'Besuche': visited})
        city_dataframe = pd.DataFrame(city_data)

        admin1_data = []
        for admin1, visited in self.sorted_admin1.items():
            admin1_data.append({f'Region/Bundesland ({len(self.sorted_admin1)})': admin1,
                                'Besuche': visited})
        admin1_dataframe = pd.DataFrame(admin1_data)

        admin2_data = []
        for admin2, visited in self.sorted_admin2.items():
            admin2_data.append({f'Bezirk/Kreis ({len(self.sorted_admin2)})': admin2,
                                'Besuche': visited})
        admin2_dataframe = pd.DataFrame(admin2_data)

        try:
            state = datetime.fromisoformat(self.__meta['state']).strftime('%d.%m.%Y %H:%M:%S')
        except:
            state = 'N/A'

        user_info = self.__meta.get('user_profile_info', {})
        user_time = user_info.get('totalDuration', 0)

        meta_dataframe = pd.DataFrame({"Label": ["Datenstand", "Anzeigename", "Profilbild",
                                                 "Mastodonlink", "Bio", "Gesamtdistanz", "Gesamtzeit", "Programmversion"],
                                       "Wert": [state, user_info.get('displayName','N/A'),
                                                (f'=HYPERLINK("{user_info.get("profilePicture", "")}", '
                                                 f'"{user_info.get("profilePicture", "Link")}")'
                                                 if user_info.get('profilePicture', None) else 'Kein Link'),

                                                (f'=HYPERLINK("{user_info.get("mastodonUrl", "")}", '
                                                 f'"{user_info.get("mastodonUrl", "Link")}")'
                                                 if user_info.get('mastodonUrl', None) else 'Kein Link'),

                                                user_info.get('bio','N/A'),
                                                f"{user_info.get('totalDistance', 0) / 1000} km",
                                                f'{int(user_time/60)}:{user_time%60:02d}',f'traewelling_analysis_v{7}.py']})


        try:
            folder_name = 'finished exports/user'
            output_dir = Path(folder_name)
            output_dir.mkdir(parents=True, exist_ok=True)
            with pd.ExcelWriter(output_dir / f"{self.__name}'s_data.xlsx", engine='openpyxl') as writer:
                sheets_to_export = [
                    ('Meta', meta_dataframe, False),
                    ('Fahrtenliste', journeys_dataframe, False),
                    ('Statistiken', stats_dataframe.set_index('Metrik').T, True),
                    ('Kategorie', type_dataframe, False),
                    ('Betreiber', operator_dataframe, False),
                    ('Linie', line_dataframe, False),
                    ('Fahrtgrund', reason_dataframe, False),
                    ('Haltestelle (ohne Via)', stop_dataframe, False),
                    ('Haltestelle (mit Via)', stop_with_via_dataframe, False),
                    ('Fahrzeuge', vehicle_dataframe, False),
                    ('Länder', country_dataframe, False),
                    ('Städte', city_dataframe, False),
                    ('Regionen_Bundesländer', admin1_dataframe, False),
                    ('Bezirke_Kreise', admin2_dataframe, False),
                ]

                for sheet_name, df, keep_index in sheets_to_export:
                    if sheet_name != 'Meta':
                        df.to_excel(writer, sheet_name=sheet_name, index=keep_index)
                    else:
                        df.to_excel(writer, sheet_name=sheet_name, index=keep_index, header=True)
                    worksheet = writer.sheets[sheet_name]

                    # --- AB HIER: Breite anpassen | Achtung: KI generiert ---
                    for sheet_name in writer.sheets:
                        worksheet = writer.sheets[sheet_name]
                        for col in worksheet.columns:
                            max_length = 0
                            column = col[0].column_letter  # Den Buchstaben der Spalte (A, B, C...) holen

                            for cell in col:
                                try:
                                    if len(str(cell.value)) > max_length:
                                        max_length = len(str(cell.value))
                                except:
                                    pass

                            # Ein kleiner Puffer (ca. 2 Einheiten), damit es nicht zu eng ist
                            adjusted_width = (max_length + 2)
                            worksheet.column_dimensions[column].width = adjusted_width




                print(f"Created {folder_name}/{self.__name}'s_data.xlsx")

        except Exception as e:
            print(f"{self.__name}'s_data.xlsx could not be created:\n\n{e}")

        print(f'\n{50 * "-"}\n')



    @property
    def name(self) -> str | None:
        return self.__name
    @property
    def total_distance(self) -> float:
        return self.__total_distance
    @property
    def distance_type_sorted(self):
        if not self.__distance_type_sorted:
            self.user_distance_time_analysis_execute()
        return self.__distance_type_sorted
    @property
    def distance_operator_line_sorted(self):
        if not self.__distance_operator_line_sorted:
            self.user_distance_time_analysis_execute()
        return self.__distance_operator_line_sorted
    @property
    def distance_operator_sorted(self):
        if not self.__distance_operator_sorted:
            self.user_distance_time_analysis_execute()
        return self.__distance_operator_sorted
    @property
    def distance_reason_sorted(self):
        if not self.__distance_reason_sorted:
            self.user_distance_time_analysis_execute()
        return self.__distance_reason_sorted
    @property
    def exported_days(self) -> int:
        return self.__exported_days
    @property
    def visited_stations(self) -> dict:
        if not self.__visited_stations:
            self.visited_station_execution()
        return self.__visited_stations
    @property
    def used_vehicles(self) -> dict:
        if not self.__used_vehicles:
            self.vehicle_execution()
        return self.__used_vehicles
    @property
    def number_of_visited_stations(self) -> int:
        if 0 ==  self.__number_of_visited_stations:
            self.visited_station_execution()
        return self.__number_of_visited_stations
    @property
    def journeys(self) -> list:
        return self.__journeys
    @property
    def number_of_visited_stations_with_via(self) -> int:
        if 0 == self.__number_of_visited_stations_with_via:
            self.visited_station_execution()
        return self.__number_of_visited_stations_with_via
    @property
    def stations_with_via(self) -> dict:
        if not self.__stations_with_via:
            self.visited_station_execution()
        return self.__stations_with_via
    @property
    def realtime_availability(self) -> float | None:
        if not self.__realtime_availability:
            self.user_distance_time_analysis_execute()
        return self.__realtime_availability
    @property
    def delay_rate_standard(self) -> float | None:
        if not self.__delay_rate_standard:
            self.user_distance_time_analysis_execute()
        return self.__delay_rate_standard
    @property
    def number_of_journeys(self) -> int:
        return self.__number_of_journeys
    @property
    def total_journey_time(self) -> float:
        return self.__total_journey_time
    @property
    def sorted_countries(self):
        if not self.__sorted_countries:
            self.user_geo_analysis()
        return self.__sorted_countries
    @property
    def sorted_cities(self):
        if not self.__sorted_cities:
            self.user_geo_analysis()
        return self.__sorted_cities
    @property
    def sorted_admin1(self):
        if not self.__sorted_admin1:
            self.user_geo_analysis()
        return self.__sorted_admin1
    @property
    def sorted_admin2(self):
        if not self.__sorted_admin2:
            self.user_geo_analysis()
        return self.__sorted_admin2


class Traewelling:
    def __init__(self, users_for_api_get: list[str] | str | None = None, delay_standard: int = 300,
                 start_date : str = '2000-01-01', end_date : str = '3000-01-01',
                 search_for_json : bool = False, search_for_correct_usernames: bool = False,
                 update_past_downloaded_checkins: bool = False, update_all_downloaded_users: bool = False):
        self.__delay_standard = delay_standard
        self.__start_date = datetime.fromisoformat(start_date)
        self.__end_date = datetime.fromisoformat(end_date)
        self.__user_meta = {}
        users_for_api_get = ([users_for_api_get] if isinstance(users_for_api_get, str) else users_for_api_get)
        remaining = 500
        max_downloaded_checkin: dict = {}

        if users_for_api_get or update_all_downloaded_users:
            token_path = Path.cwd() / '.api_token.txt'
            try:
                with open(token_path, 'r') as token_file:
                    self.__api_token = token_file.read().strip()
            except FileNotFoundError:
                raise FileNotFoundError(f'{token_path} nicht gefunden.')
            if self.__api_token == "Bitte hier deinen API-Token eintragen":
                raise ValueError(f'Kein API in {token_path} eingetragen.')
            repeat_bool = True
            while repeat_bool:
                travel_response = requests.get('https://traewelling.de/api/v1/auth/user',
                                               headers={'accept': 'application/json',
                                                        'Authorization': f'Bearer {self.__api_token}'})
                try:
                    travel_response.raise_for_status()
                    travel_json = travel_response.json()
                    username = travel_json.get('data', {}).get('username', None)

                    if username:
                        print(f'Moin {username}')
                    repeat_bool = False
                    remaining = travel_response.headers.get('x-ratelimit-remaining')
                except Exception as e:
                    if isinstance(e, requests.exceptions.HTTPError) and e.response.status_code == 429:
                        print(f'Rate Limit (429) erreicht. Warte 30 Sekunden und versuche es erneut...')
                        sleep(30)
                    else:
                        raise PermissionError(
                            f'Fehler bei der Authentifizierung! Ist der Token noch gültig?\nDetails: {e}')

        print("Importing träwelling data...")

        self.__whole_travel_data = []
        path = Path.cwd()
        try:
            with open('._data_saver.json', encoding='utf-8') as f:
                coordinates_json = json.load(f)
            self.__coordinates_dict = coordinates_json.get('coordinates_data', {})
            self.__timezone_per_stop = coordinates_json.get('coordinates_timezone', {})
        except Exception as e:
            print(e)
            self.__coordinates_dict = {}
            self.__timezone_per_stop = {}

        if search_for_json:
            file_names = []
            for file in path.glob('*.json'):
                if file.name != '._data_saver.json':
                    file_names.append(file)

            if not file_names:
                 print(FileNotFoundError("No .json files found"))
            file_names.sort()

            for file_name in file_names:
                try:
                    with open(file_name, encoding='utf-8') as f:
                        travel = json.load(f)
                    travel_data = travel.get('data', {})

                    if not update_past_downloaded_checkins:
                        max_downloaded_checkin[travel_data[0]['userDetails']['username']] = ([data['id'] for data in travel_data], travel_data)
                    self.__user_meta[travel_data[0]['userDetails']['username']] = travel.get('meta', {})


                    self.__whole_travel_data += travel_data
                    print(f'Imported {len(travel_data)} travel data from {file_name}')
                except Exception as e:
                    print(f'Could not import {file_name} because {e}')

            print('Getting coordinates from given user and saver .json files...')
            for data in self.__whole_travel_data:
                trip_data = data.get('checkin', {})

                try:
                    origin_id = str(trip_data['origin'].get('id', None))
                    origin_coordinates = (round(trip_data['origin'].get('latitude', None),5),
                                          round(trip_data['origin'].get('longitude', None),5))
                    destination_id = str(trip_data['destination'].get('id', None))
                    destination_coordinates = (round(trip_data['destination'].get('latitude', None),5),
                                               round(trip_data['destination'].get('longitude', None),5))
                    if origin_id != 'None' and origin_coordinates[0] not in (None, 'N/A') and origin_coordinates[1] not in (None, 'N/A'):
                        self.__coordinates_dict[origin_id] = self.__coordinates_dict.get(origin_id, origin_coordinates)
                    if origin_id != 'None' and destination_coordinates[0] not in (None, 'N/A') and destination_coordinates[1] not in (None, 'N/A'):
                        self.__coordinates_dict[destination_id] = self.__coordinates_dict.get(destination_id, destination_coordinates)
                except:
                    pass

        print(f'\n{50 * "-"}\n')

        if update_all_downloaded_users and max_downloaded_checkin:
            if users_for_api_get is None:
                users_for_api_get = []
            users_for_api_get = list(set(max_downloaded_checkin.keys()) | set(users_for_api_get))


        if users_for_api_get:
            for user_number, user_name in enumerate(users_for_api_get):
                time_of_download_start = datetime.now().isoformat()
                user_profile_info = None
                user_not_found = False
                if not search_for_correct_usernames:
                    repeat_bool = True
                    while repeat_bool:
                        user_response = requests.get(f'https://traewelling.de/api/v1/user/{user_name}', headers={'accept': 'application/json',
                                                                   'Authorization': f'Bearer {self.__api_token}'})
                        try:
                            user_response.raise_for_status()
                            user_json = user_response.json()
                            userinfo: dict = user_json.get('data', {})

                            user_name = userinfo.get('username', None)
                            user_profile_info = {'bio': userinfo.get('bio', None),
                                                 'displayName': userinfo.get('displayName', None),
                                                 'mastodonUrl': userinfo.get('mastodonUrl', None),
                                                 'totalDistance': userinfo.get('totalDistance', None),
                                                 'totalDuration': userinfo.get('totalDuration', None),
                                                 'profilePicture': userinfo.get('profilePicture', None),}

                            repeat_bool = False
                            remaining = user_response.headers.get('x-ratelimit-remaining')

                        except Exception as e:
                            if isinstance(e, requests.exceptions.HTTPError) and e.response.status_code == 429:
                                print(f'Rate Limit (429) erreicht. Warte 30 Sekunden und versuche es erneut...')
                                sleep(30)
                            else:
                                user_not_found = True
                                repeat_bool = False
                                print(f'{user_name} not found.')


                if search_for_correct_usernames or user_not_found:
                    print(f'Searching for {user_name}')

                    url = f'https://traewelling.de/api/v1/user/search/{user_name}'

                    possible_users_list : list = []
                    while url is not None:
                        repeat_bool = True
                        while repeat_bool:
                            user_response = requests.get(url, headers={'accept': 'application/json',
                                                                'Authorization': f'Bearer {self.__api_token}'})
                            try:
                                user_response.raise_for_status()
                                user_json = user_response.json()
                                userlist : dict = user_json.get('data', {})

                                print(userlist)
                                possible_usernames = [item.get('username') for item in userlist]
                                possible_users_infos = [{'bio': userinfo.get('bio', None),
                                                 'displayName': userinfo.get('displayName', None),
                                                 'mastodonUrl': userinfo.get('mastodonUrl', None),
                                                 'totalDistance': userinfo.get('totalDistance', None),
                                                 'totalDuration': userinfo.get('totalDuration', None),
                                                 'profilePicture': userinfo.get('profilePicture', None),}
                                                        for userinfo in userlist]
                                possible_users_list.extend(possible_usernames)

                                repeat_bool = False
                                remaining = user_response.headers.get('x-ratelimit-remaining')

                                url = user_json.get('links', {}).get('next')
                            except Exception as e:
                                if isinstance(e, requests.exceptions.HTTPError) and e.response.status_code == 429:
                                    print(f'Rate Limit (429) erreicht. Warte 30 Sekunden und versuche es erneut...')
                                    sleep(30)
                                else:
                                    raise

                    for number, name in enumerate(possible_users_list):
                        print(f' {number + 1}. {name}')

                    if user_name in possible_users_list:
                        selected_user = input(f'Welcher User ist gesucht(oder Enter für nicht gefunden), {user_name} ist'
                                                f'Nummer {possible_users_list.index(user_name)+1}: ')
                    else:
                        selected_user = input('Welcher User ist gesucht(oder Enter für nicht gefunden): ')

                    try:
                        selected_user = int(selected_user) - 1
                    except ValueError:
                        print('Kein User ausgewählt')
                        continue

                    if 0 <= selected_user < len(possible_users_list):
                        user_name = possible_users_list[selected_user]
                        user_profile_info = possible_users_infos[selected_user]
                    else:
                        continue

                print(f"\nDownloading {user_name}'s data... ({user_number+1}/{len(users_for_api_get)}) {datetime.now().strftime("%H:%M:%S")}")
                existing_data = max_downloaded_checkin.get(user_name, None)

                current_user_data = []
                number = 1
                url = f'https://traewelling.de/api/v1/user/{user_name}/statuses'
                while url is not None:
                    travel_response = requests.get(url,
                                          headers = {'accept': 'application/json',
                                                     'Authorization': f'Bearer {self.__api_token}'})
                    try:
                        travel_response.raise_for_status()
                        travel_json = travel_response.json()
                        travel_data = travel_json.get('data',{})

                        if not update_past_downloaded_checkins and existing_data is not None:
                            for checkin in travel_data:
                                if checkin['id'] in existing_data[0]: #Ist die Aktuelle Check-in Nummer gleich der Maximalen
                                    current_user_data.extend(existing_data[1])
                                    url = None
                                    break
                                else:
                                    current_user_data.append(checkin)
                            if url is not None:
                                url = travel_json.get('links', {}).get('next')
                        else:
                            current_user_data += travel_data
                            url = travel_json.get('links',{}).get('next')

                        remaining = travel_response.headers.get('x-ratelimit-remaining')
                        print(f'Imported {len(travel_data)} travel data from {user_name}, Page: {number}->'
                              f'{len(current_user_data)} journeys, Requests Remaining: {remaining} {datetime.now().strftime("%H:%M:%S")}')
                        #if int(remaining) < 20:
                         #   print(f'ratelimit critical: slowing down Requests')
                          #  sleep(10)

                        number += 1
                    except Exception as e:
                        if isinstance(e, requests.exceptions.HTTPError) and e.response.status_code == 429:
                            print(f'Rate Limit (429) erreicht. Warte 30 Sekunden und versuche es erneut... {datetime.now().strftime("%H:%M:%S")}')
                            sleep(30)
                        else:
                            print(f'Could not import data from {user_name}: {e}')
                            url = None

                if current_user_data:
                    print(f'getting coordinates for {user_name}...')
                    file_name = f'{user_name}_API_export.json'
                    trip_ids = []
                    if not remaining:
                        remaining = 500
                    stops_to_get = [(data.get('train', {}).get('origin',{}).get('id'),
                                     data.get('train', {}).get('destination',{}).get('id'))
                                    for data in current_user_data]
                    stops_to_get = (set(map(str, chain.from_iterable(stops_to_get))) -
                                    (self.__coordinates_dict.keys()))
                    number_of_missing_coordinates = len(stops_to_get)
                    print(f'{number_of_missing_coordinates} missing coordinates from {user_name}')
                    current_runner = 0
                    for data in current_user_data:
                        trip_ids.append(data['checkin'].get('trip'))
                        train = data.get('train', {})
                        for stop_key in ['origin', 'destination']:
                            stop_data = train.get(stop_key, {})
                            stop_id = str(stop_data.get('id'))
                            if stop_id is not None:
                                coordinates = self.__coordinates_dict.get(stop_id)
                                if coordinates not in [None, ("N/A", "N/A")]:
                                    stop_data['latitude'] = coordinates[0]
                                    stop_data['longitude'] = coordinates[1]
                                else:
                                    current_runner += 1
                                    repeat_bool = True
                                    while repeat_bool:
                                        print(f'getting coordinates for stop {stop_data.get("name",''):.<30}'
                                              f'{current_runner}/{number_of_missing_coordinates}: Remaining: {int(remaining)} {datetime.now().strftime("%H:%M:%S")}')
                                        station_response = requests.get(f'https://traewelling.de/api/v1/stations/{stop_id}',
                                                                       headers={'accept': 'application/json',
                                                                                'Authorization': f'Bearer {self.__api_token}'})

                                        try:
                                            station_response.raise_for_status()

                                            remaining = station_response.headers.get('x-ratelimit-remaining')
                                            station_json = station_response.json().get('data',{})
                                            if str(station_json.get('id')) == stop_id:
                                                latitude = round(station_json.get('latitude', None),5)
                                                longitude = round(station_json.get('longitude', None),5)
                                                stop_data['latitude'] = latitude
                                                stop_data['longitude'] = longitude
                                                if (self.__coordinates_dict.get(stop_id, None) is None and
                                                        latitude is not None and longitude is not None):
                                                    self.__coordinates_dict[stop_id] = (latitude, longitude)
                                            else:
                                                raise Exception(f'Could not find stop {stop_data.get("name")}')

                                            repeat_bool = False
                                        except Exception as e:
                                            if isinstance(e, requests.exceptions.HTTPError) and e.response.status_code == 429:
                                                print(f'Rate Limit (429) erreicht. Warte 30 Sekunden und versuche es erneut... {datetime.now().strftime("%H:%M:%S")}')
                                                sleep(30)
                                            else:
                                                print(f'Could not import stop data from {stop_data.get("name")}: {e}')
                                                repeat_bool = False


                    trips_without_stopover_data = [data['checkin'].get('trip')
                                                   for data in current_user_data
                                                   if data['train'].get('stopovers') is None]

                    intermediate_data = {}
                    current_index = 0
                    while len(trips_without_stopover_data) > current_index:
                        print(f'getting intermediate stops from {min(current_index+50,len(trips_without_stopover_data))}/{len(trips_without_stopover_data)} '
                              f' journeys for {user_name}...: Remaining: {remaining} {datetime.now().strftime("%H:%M:%S")}')
                        joined_ids = "%2C".join(str(trip_id) for trip_id in trips_without_stopover_data[current_index:current_index+50])
                        url = f"https://traewelling.de/api/v1/stopovers/{joined_ids}"
                        repeat_bool = True
                        while repeat_bool:
                            try:
                                intermediate_response = requests.get(url,
                                                                     headers={'accept': 'application/json',
                                                                              'Authorization': f'Bearer {self.__api_token}'})
                                intermediate_response.raise_for_status()
                                remaining = intermediate_response.headers.get('x-ratelimit-remaining')
                                intermediate_json = intermediate_response.json()
                                intermediate_data.update(intermediate_json.get('data', {}))
                                repeat_bool = False
                                current_index += 50
                            except Exception as e:
                                if isinstance(e, requests.exceptions.HTTPError) and e.response.status_code == 429:
                                    print(f'Rate Limit (429) erreicht. Warte 30 Sekunden und versuche es erneut... {datetime.now().strftime("%H:%M:%S")}')
                                    sleep(30)
                                else:
                                    print(f'Could not import intermediate stop data from {user_name}: {e}')
                                    repeat_bool = False
                                    current_index = len(current_user_data)

                    for data in current_user_data:
                        if data['train'].get('stopovers',None) is None:
                            trip_id = str(data['checkin'].get('trip'))
                            data['train']['stopovers'] = intermediate_data.get(trip_id, [])

                    if user_name in max_downloaded_checkin:
                        old_id = {checkin['id'] for checkin in max_downloaded_checkin[user_name][1]}
                        self.__whole_travel_data = [data
                                                    for data in self.__whole_travel_data
                                                    if data['id'] not in old_id]

                    self.__whole_travel_data.extend(current_user_data)
                    current_user_meta = {'state': time_of_download_start,
                                         'user_profile_info': user_profile_info}
                    self.__user_meta[user_name] = current_user_meta
                    print(f'Imported {len(current_user_data)} travel data from {user_name}')


                    try:
                        with open(file_name, 'w', encoding='utf-8') as f:
                            json.dump({'data': current_user_data,
                                       'meta': current_user_meta}, f, ensure_ascii=False, indent=4)
                            print(f'Created {file_name}')
                    except Exception as e:
                        print(f'Could not create {file_name}: {e}')
            print(f'\n{50 * '-'}\n')

        if self.__whole_travel_data:
            sprich_deutsch = Locale('de')
            coordinates_set = set(tuple(c) for c in self.__coordinates_dict.values())
            coordinates_list = list(c for c in coordinates_set if isinstance(c, tuple) and c != ('N/A', 'N/A'))
            geo_results = rg.search(list(coordinates_list))
            self.__coordinates_places_map = {}
            for coordinates, results in zip(coordinates_list, geo_results):
                cc = results.get('cc')
                country_object = pycountry.countries.get(alpha_2=cc)
                admin1_en = results.get('admin1')
                if admin1_en in ['London', 'Vienna', 'Berlin', 'Hamburg', 'Bremen']:
                    city = results.get('name')  # hier Admin1 eintragen
                else:
                    city = results.get('name')
                self.__coordinates_places_map[coordinates] = {'name': city,
                                                              'admin1': f'{admin1_en} {country_object.flag}',
                                                              'admin2': results.get('admin2'),
                                                              'cc': cc,
                                                              'country_name': f'{sprich_deutsch.territories.get(country_object.alpha_2,
                                                                                                                country_object.name)} '
                                                                              f'{country_object.flag}',
                                                              'official_name': results.get('official_name')}

            tf = TimezoneFinder(in_memory=True)

            print('Getting timezones for journeys...')
            self.__journeys: list[Journey] = []
            userid_set: set[int] = set()
            id_list: set[int] = set()
            number_skipped_journeys = 0
            for journey_data in self.__whole_travel_data:
                status_id = journey_data['id']

                if status_id in id_list:
                    number_skipped_journeys += 1
                else:
                    journey_object = Journey(journey_data, tf, self.__timezone_per_stop,
                                                      self.__coordinates_places_map)
                    if self.__start_date <= journey_object.departure_planned.replace(tzinfo=None) <= self.__end_date:
                        self.__journeys.append(journey_object)
                        self.__timezone_per_stop[journey_object.destination_stop] = journey_object.timezone_destination
                        self.__timezone_per_stop[journey_object.origin_stop] = journey_object.timezone_origin
                        id_list.add(status_id)
                        userid_set.add(journey_object.user_id)
                    else:
                        number_skipped_journeys += 1

            try:
                with open('._data_saver.json', 'w', encoding='utf-8') as f:
                    json.dump(
                        {'coordinates_data': self.__coordinates_dict, 'coordinates_timezone': self.__timezone_per_stop},
                        f, ensure_ascii=False, indent=4)
                    print(f'Created {'data_saver.json'}')
            except Exception as e:
                print(f'Could not create {'data_saver.json'}: {e}')

            self.__journeys.sort()

            self.__user_id_dict: dict[int, User] = dict()
            for userid in userid_set:
                self.__user_id_dict[userid] = User(userid, self.__delay_standard)
            for journey in self.__journeys:
                user = self.__user_id_dict[journey.user_id]
                user.add_journey(journey)
            self.__user_id_dict = dict(sorted(self.__user_id_dict.items(), key=lambda x: x[1], reverse=True))

            self.__user_name_dict: dict[str, User] = dict()
            for u in self.__user_id_dict.values():
                for j in u.journeys:
                    self.__user_name_dict[j.user_name] = u

            raw_travel_data = len(self.__whole_travel_data)
            assigned_travel_data = sum(len(u.journeys) for u in self.__user_id_dict.values())
            if assigned_travel_data == raw_travel_data and number_skipped_journeys == 0:
                print(f'Successfully imported {assigned_travel_data} journeys\n')
            elif number_skipped_journeys > 0 and assigned_travel_data + number_skipped_journeys == raw_travel_data:
                print(
                    f'Successfully imported {assigned_travel_data} journeys ({number_skipped_journeys} excess journeys were deleted)\n')
            else:
                raise Exception(
                    f'Something went wrong: there should be {raw_travel_data} journeys, there are just {assigned_travel_data}\n')
            print(f'\n{50 * '-'}\n')

            try:
                with open('._data_saver.json', 'w', encoding='utf-8') as f:
                    json.dump(
                        {'coordinates_data': self.__coordinates_dict, 'coordinates_timezone': self.__timezone_per_stop},
                        f, ensure_ascii=False, indent=4)
                    print(f'Created {'data_saver.json'}\n')
            except Exception as e:
                print(f'Could not create {'data_saver.json'}: {e}\n')

            for user in self.__user_id_dict.items():
                user[1].get_import_length(self.__user_meta)

        else:
            raise ImportError('Keine User wurden Importiert')




    def __user_input_to_object(self, user_name_list: str | list[str] | None= None, min_requirement: int = 1) -> list[User] | None:
        if len(self.__user_id_dict) < min_requirement:
            return None
        elif user_name_list is None:
            return [user[1] for user in self.__user_id_dict.items()]
        elif isinstance(user_name_list, str) :
            user_name_list = [user_name_list]

        user_list = []
        for user_name in list(user_name_list):
            user_element = self.__user_name_dict.get(user_name)
            if user_element is None:
                print(f'User {user_name} nicht gefunden!')
            else:
                user_list.append(user_element)
        user_list = list(set(user_list))

        if len(user_list) < min_requirement:
            print('Es wurden zu wenige User ausgewählt!')
            return None

        return user_list


    def create_user_excel(self, user_name_list: str | list[str]= None) -> None:
        user_list = self.__user_input_to_object(user_name_list, 1)
        if user_list is None:
            return
        for user in user_list:
            user.create_excel()


    def create_shared_excel(self, user_name_list: str | list[str]= None) -> None:
        user_list = self.__user_input_to_object(user_name_list, 2)
        if user_list is None:
            return

        visits_dataframe: pd.DataFrame = pd.DataFrame()
        visits_dataframe_filterable :pd.DataFrame = pd.DataFrame()

        dicts = [user.stations_with_via for user in user_list]
        station_counter = Counter(operator for d in dicts for operator in d.keys())
        common_stations = []
        for station, count in station_counter.items():
            if count >= min(max(2, len(user_list) - 1 if len(user_list) < 5 else len(user_list) - 2), 5):
                common_stations.append(station)

        if common_stations:
            visits_data = []
            visits_data_filterable = []
            common_stations = sorted(common_stations)
            for station in common_stations:
                visits: list[tuple] = []
                for user in user_list:
                    visits.append((user.name, user.stations_with_via.get(station,0)))
                visits = sorted(visits, key=lambda x: (x[1]), reverse=True)
                current_station = {'Haltestelle': station}
                for index, (user, number_with_via) in enumerate(visits):
                    user_obj = self.__user_name_dict[user]
                    number_without_via = user_obj.visited_stations.get(station,0)
                    visits_data.append({'Haltestelle':(station if index==0 else ''),
                                        'User':user,
                                        'Anzahl ohne Via' : number_without_via,
                                        'Anzahl mit Via':number_with_via,
                                        'Pro Monat ohne Via': round(number_without_via/user_obj.exported_days * 30, 3),
                                        'Pro Monat mit Via': round(number_with_via/user_obj.exported_days * 30, 3),
                                        'Umstiegquote': (f'{number_without_via/number_with_via * 100:.2f}%') if number_with_via > 0 else ''})
                    current_station[f'Anzahl ohne Via ({user})'] = number_without_via
                    current_station[f'Anzahl mit Via ({user})'] = number_with_via
                    current_station[f'Umstiegquote (%) ({user})'] = (round(number_without_via/number_with_via * 100,2)) if number_with_via > 0 else ''
                visits_data.append({'Haltestelle': '',
                                    'User': '',
                                    'Anzahl ohne Via': '',
                                    'Anzahl mit Via': '',
                                    'Pro Monat ohne Via': '',
                                    'Pro Monat mit Via': '',
                                    'Umstiegquote': ''})
                visits_data_filterable.append(current_station)

            visits_dataframe = pd.DataFrame(visits_data)
            visits_dataframe_filterable = pd.DataFrame(visits_data_filterable)

        stats_data = []

        for user in user_list:
            total_hours, total_minutes = divmod(user.total_journey_time / 60, 60)
            per_day_hours, per_day_minutes = divmod((user.total_journey_time / user.exported_days) / 60, 60)
            per_journey_hours, per_journey_minutes = divmod((user.total_journey_time / user.number_of_journeys) / 60,
                                                            60)
            stats_data.append({'User': user.name,
                          'Fahrten': len(user.journeys),
                          'Tage': user.exported_days ,
                          'Fahrten pro Tag':round(len(user.journeys) / user.exported_days,3),
                          'Linien': len(user.distance_operator_line_sorted),
                          'Linien pro Tag':round(len(user.distance_operator_line_sorted) / user.exported_days,3),
                          'Linien pro Fahrt': round(len(user.distance_operator_line_sorted) / user.number_of_journeys,3),
                          'Betreiber': len(user.distance_operator_sorted),
                          'Betreiber pro Tag':round(len(user.distance_operator_sorted) / user.exported_days,3),
                          'Betreiber pro Fahrt':round(len(user.distance_operator_sorted) / user.number_of_journeys,3),
                          'Haltestellen': round(user.number_of_visited_stations,3),
                          'Haltestellen pro Tag':round(user.number_of_visited_stations / user.exported_days,3),
                          'Haltestellen pro Fahrt':round(user.number_of_visited_stations / user.number_of_journeys,3),
                          'Haltestellen mit Via': round(user.number_of_visited_stations_with_via, 3),
                          'Haltestellen pro Tag (via)': round(user.number_of_visited_stations_with_via / user.exported_days, 3),
                          'Haltestellen pro Fahrt (via)': round(user.number_of_visited_stations_with_via / user.number_of_journeys, 3),
                          'Abfahrtsverspätung (min)': round(sum(j.departure_delay for j in user.journeys) / 60,0),
                          'Abfahrt pro Tag':round((sum(j.departure_delay for j in user.journeys) / 60) / user.exported_days ,3),
                          'Abfahrt pro Fahrt':round((sum(j.departure_delay for j in user.journeys) / 60) / user.number_of_journeys ,3),
                          'Ankunftsverspätung (min)': round(sum(j.arrival_delay for j in user.journeys) / 60,0),
                          'Ankunft pro Tag':round((sum(j.arrival_delay for j in user.journeys) / 60) / user.exported_days ,3),
                          'Ankunft pro Fahrt':round((sum(j.arrival_delay for j in user.journeys) / 60) / user.number_of_journeys ,3),
                          'Zeit (min)': round(int(user.total_journey_time / 60),3),
                          'min pro Tag':round(user.total_journey_time / 60 / user.exported_days ,3),
                          'min pro Fahrt':round(user.total_journey_time / user.number_of_journeys / 60,3),
                          'Zeit (hh:mm)': f"{total_hours:02.0f}:{int(round(total_minutes, 0)):02}",
                          'hh:mm pro Tag':f"{per_day_hours:02.0f}:{round(per_day_minutes, 0):02.0f}",
                          'hh:mm pro Fahrt':f"{per_journey_hours:02.0f}:{round(per_journey_minutes, 0):02.0f}",
                          'Distanz (m)': user.total_distance,
                          'm pro Tag':round(user.total_distance / user.exported_days ,3),
                          'm pro Fahrt':round(user.total_distance / user.number_of_journeys,3),
                          'Distanz (km)': round(user.total_distance / 1000,2),
                          'km pro Tag':round(user.total_distance / 1000 / user.exported_days,2),
                          'km pro Fahrt':round(user.total_distance / 1000 / user.number_of_journeys,2),
                          'Punkte': sum(j.journey_points for j in user.journeys),
                          'Punkte pro Tag':round(sum(j.journey_points for j in user.journeys) / user.exported_days,3),
                          'Punkte pro Fahrt':round(sum(j.journey_points for j in user.journeys) / user.number_of_journeys,3),
                          'Echtzeitquote (%)': round(user.realtime_availability * 100,2),
                          f'Mehr als {round(self.__delay_standard / 60)} Minuten verspätet (%)': round(user.delay_rate_standard * 100,2),
                          '\u2205 Zwischenhalte':round(mean(j.number_of_inter_stations for j in user.journeys),3)})

        stats_dataframe = pd.DataFrame(stats_data)


        type_set: set[str] = set()
        for user in user_list:
            for train_type, train_distance in user.distance_type_sorted.items():
                type_set.add(train_type)
        type_list = sorted(type_set)
        type_data = []
        for type in type_list:
            user_list.sort(reverse=True, key = lambda x: (x.distance_type_sorted.get(type, {}).get('distance', 0)))
            for index, user in enumerate(user_list):
                user_type_dict = user.distance_type_sorted.get(type, {})
                h, m = divmod(user_type_dict.get('time',0)/60, 60)
                type_data.append({'Kategorie':(type if index==0 else ''),
                                  'User':user.name,
                                  'Distanz (km)': round(user_type_dict.get('distance', 0)/1000, 2),
                                  'Distanz pro Tag (km)': round(user_type_dict.get('distance', 0)/user.exported_days/1000, 2),
                                  'Zeit (hh:mm)': f"{h:02.0f}:{int(round(m, 0)):02}",
                                  'Zeit pro Tag (min)': round(user_type_dict.get('time', 0)/user.exported_days/60, 2),
                                  'Fahrten': user_type_dict.get('sum',0),
                                  'Fahrten pro Woche': round(user_type_dict.get('sum', 0)/user.exported_days*7, 2)})
            type_data.append({'Kategorie': '',
                              'User': '',
                              'Distanz (km)': '',
                              'Distanz pro Tag (km)': '',
                              'Zeit (hh:mm)': '',
                              'Zeit pro Tag (min)': '',
                              'Fahrten': '',
                              'Fahrten pro Woche': ''})

        type_dataframe = pd.DataFrame(type_data)

        dicts = [user.distance_operator_sorted for user in user_list]
        operator_counter = Counter(operator for d in dicts for operator in d.keys())
        common_operators = []
        for operator, count in operator_counter.items():
            if count >= max(2, min(len(user_list) - 1 if len(user_list) < 4 else len(user_list) - 2, 4)):
                common_operators.append(operator)

        operator_data = []
        if common_operators:
            if 'Unknown Operator' in common_operators:
                common_operators.remove('Unknown Operator')
            common_operators = sorted(common_operators)
            operator_data = []
            for operator in common_operators:
                user_list.sort(reverse=True, key=lambda x: (x.distance_operator_sorted.get(operator, {}).get('distance', 0)))
                for index, user in enumerate(user_list):
                    user_operator_dict = user.distance_operator_sorted.get(operator, {})
                    h, m = divmod(user_operator_dict.get('time', 0) / 60, 60)
                    operator_data.append({'Betreiber': (operator if index == 0 else ''),
                                      'User': user.name,
                                      'Distanz (km)': round(user_operator_dict.get('distance', 0) / 1000, 2),
                                      'Distanz pro Tag (km)': round(
                                          user_operator_dict.get('distance', 0) / user.exported_days / 1000, 2),
                                      'Zeit (hh:mm)': f"{h:02.0f}:{int(round(m, 0)):02}",
                                      'Zeit pro Tag (min)': round(
                                          user_operator_dict.get('time', 0) / user.exported_days / 60, 2),
                                      'Geschwindigkeit (km/h)': (round(user_operator_dict.get('distance', 0)/user_operator_dict.get('time', 0)*3.6, 2)
                                                                 if user_operator_dict.get('time', 0) != 0 else ''),
                                      'Fahrten': user_operator_dict.get('sum', 0),
                                      'Fahrten pro Woche': round(
                                          user_operator_dict.get('sum', 0) / user.exported_days * 7, 2)})
                operator_data.append({'Betreiber': '',
                                  'User': '',
                                  'Distanz (km)': '',
                                  'Distanz pro Tag (km)': '',
                                  'Zeit (hh:mm)': '',
                                  'Zeit pro Tag (min)': '',
                                  'Fahrten': '',
                                  'Fahrten pro Woche': ''})

        operator_dataframe = pd.DataFrame(operator_data)

        shared_trip_data = []
        all_countrys = [] #Vorleistung für Länderauswertung
        all_states = []
        trip_id_dict = {}
        for j in self.__journeys:
            if self.__user_id_dict.get(j.user_id) in user_list:
                if not trip_id_dict.get(j.trip_id):
                    trip_id_dict[j.trip_id] = []
                trip_id_dict[j.trip_id].append((self.__user_id_dict.get(j.user_id), j))
            all_countrys.append(j.country_name_dest)
            all_countrys.append(j.country_name_origin)
            all_states.append(j.admin1_origin)
            all_states.append(j.admin1_dest)


        for trip_id, list_user_journey in trip_id_dict.items():
            if len(list_user_journey) > 1 and trip_id != 'N/A':
                list_user_journey.sort(key=lambda x: x[0] ,reverse=True)
                for index, (user, journey) in enumerate(list_user_journey):
                    shared_trip_data.append({'Kategorie':(journey.train_type if index == 0 else ''),
                                             'Linie': (journey.line_name if index == 0 else ''),
                                             'User': user.name,
                                             'Starthaltestelle': journey.origin_stop,
                                             'Abfahrtszeit (ist)': journey.departure_real.strftime('%d.%m.%Y %H:%M'),
                                             'Zielhaltestelle':journey.destination_stop,
                                             'Ankunftszeit (ist)': journey.arrival_real.strftime('%d.%m.%Y %H:%M'),
                                             'Entfernung (km)': round(journey.journey_distance / 1000, 3),
                                             'Punkte': journey.journey_points,
                                             'Link': f'=HYPERLINK("https://traewelling.de/status/{journey.status_id}",'
                                                     f'"https://traewelling.de/status/{journey.status_id}")',
                                             '_color': ('color:black' if not journey.realtime_availability else
                                                        ('color:red' if journey.delayed_by_standard(self.__delay_standard) else 'color:green')),})
                shared_trip_data.append({'Kategorie': '', 'Linie': '','User': '','Starthaltestelle': '','Abfahrtszeit (ist)': '',
                                         'Zielhaltestelle': '','Ankunftszeit (ist)': '','Entfernung (km)': '','Punkte': '',
                                         'Link': '','_color': '' })

        shared_trip_dataframe = pd.DataFrame(shared_trip_data)

        countrys_data = []
        all_countrys = sorted(filter(None, set(all_countrys)))

        for country in all_countrys:
            user_list.sort(reverse=True,
                           key=lambda x: (x.sorted_countries.get(country, 0)))
            for index, user in enumerate(user_list):
                countrys_data.append({'Land': (country if index == 0 else ''),
                                      'User': user.name,
                                      'Besuche': user.sorted_countries.get(country, 0)})
            countrys_data.append({'Land': '',
                                  'User': '',
                                  'Besuche': ''})

        countrys_dataframe = pd.DataFrame(countrys_data)

        states_data = []
        all_states = sorted(filter(None, set(all_states)))

        for state in all_states:
            user_list.sort(reverse=True,
                           key=lambda x: (x.sorted_admin1.get(state, 0)))
            for index, user in enumerate(user_list):
                states_data.append({'Bundesland': (state if index == 0 else ''),
                                      'User': user.name,
                                      'Besuche': user.sorted_admin1.get(state, 0)})
            states_data.append({'Bundesland': '',
                                  'User': '',
                                  'Besuche': ''})

        states_dataframe = pd.DataFrame(states_data)


        user_list.sort(key=lambda x: x.name)
        names = f'{user_list[0].name}'
        too_long_test = False
        for number, name in enumerate([user.name for user in user_list[1:]]):
            if len(names) + len(name) < 150:
                names += f'_{name}'
            elif not too_long_test:
                too_long_test = True
                names += f'_and_{len(user_list) - number}_more'

        try:
            folder_name = 'finished exports/shared'
            output_dir = Path(folder_name)
            output_dir.mkdir(parents=True, exist_ok=True)
            with pd.ExcelWriter(output_dir / f"{names}'s_shared_data.xlsx", engine='openpyxl') as writer:
                stats_dataframe.to_excel(writer, sheet_name='Statistik', index=False)
                if not visits_dataframe.empty:
                    visits_dataframe.to_excel(writer, sheet_name='Haltestellen', index=False)
                    visits_dataframe_filterable.to_excel(writer, sheet_name='Haltestellen (filterbar)', index=False)
                type_dataframe.to_excel(writer, sheet_name='Kategorie', index=False)
                if not operator_dataframe.empty:
                    operator_dataframe.to_excel(writer, sheet_name='Betreiber', index=False)
                countrys_dataframe.to_excel(writer, sheet_name='Länder', index=False)
                states_dataframe.to_excel(writer, sheet_name='Regionen_Bundesländer', index=False)
                if not shared_trip_dataframe.empty:
                    styled_trip_dataframe = shared_trip_dataframe.style.apply(lambda row:
                                                                              [row['_color'] if col in [
                                                                                  'Ankunftszeit (ist)',
                                                                                  'Abfahrtszeit (ist)']
                                                                               else '' for col in row.index], axis=1)
                    styled_trip_dataframe = styled_trip_dataframe.hide(['_color'], axis='columns')
                    styled_trip_dataframe.hide(['_color'], axis='columns').to_excel(writer, sheet_name='Gemeinsame Fahrten', index=False)

                    # --- AB HIER: Breite anpassen | Achtung: KI generiert ---
                    for sheet_name in writer.sheets:
                        worksheet = writer.sheets[sheet_name]
                        for col in worksheet.columns:
                            max_length = 0
                            column = col[0].column_letter  # Den Buchstaben der Spalte (A, B, C...) holen

                            for cell in col:
                                try:
                                    if len(str(cell.value)) > max_length:
                                        max_length = len(str(cell.value))
                                except:
                                    pass

                            # Ein kleiner Puffer (ca. 2 Einheiten), damit es nicht zu eng ist
                            adjusted_width = (max_length + 2)
                            worksheet.column_dimensions[column].width = adjusted_width



                print(f"Created {folder_name}/{names}'s_shared_data.xlsx")

        except Exception as e:
            print(f"{names}'s_shared_data.xlsx could not be created:\n\n{e}")

        print(f'\n{50 * "-"}\n')


    def create_gis_single_csv(self, user_name_list: str | list[str]= None) -> None:
        user_list = self.__user_input_to_object(user_name_list, 1)
        if user_list is None:
            return
        header_row = ['y', 'x', 'Stopname', 'Username']
        journey_rows = []
        for j in self.__journeys:
            if self.__user_name_dict.get(j.user_name) in user_list:
                journey_rows.extend([[j.destination_coordinates[0], j.destination_coordinates[1], j.destination_stop, j.user_name],
                                     [j.origin_coordinates[0], j.origin_coordinates[1], j.origin_stop, j.user_name]])
        output_dir = Path('finished exports')
        output_dir.mkdir(parents=True, exist_ok=True)
        with open(output_dir / "gis_single_export.csv", 'w', newline='', encoding='utf-8') as datei:
            schreiben = csv.writer(datei)
            schreiben.writerow(header_row)
            schreiben.writerows(journey_rows)
        print(f'Created {datei.name}')
        print(f'\n{50 * "-"}\n')

    def create_gis_number_csv(self, user_name_list: str | list[str]= None) -> None:
        user_list = self.__user_input_to_object(user_name_list, 1)
        if user_list is None:
            return
        header_row = ['y', 'x', 'Stopname', 'Username', 'Number']
        journey_rows = []
        journey_rows_dict = {}
        for j in self.__journeys:
            if self.__user_name_dict.get(j.user_name) in user_list:
                destination_data = (j.destination_coordinates[0], j.destination_coordinates[1], j.destination_stop, j.user_name)
                origin_data = (j.origin_coordinates[0], j.origin_coordinates[1], j.origin_stop, j.user_name)

                journey_rows_dict[destination_data] = journey_rows_dict.get(destination_data,0) + 1
                journey_rows_dict[origin_data] = journey_rows_dict.get(origin_data,0) + 1

        for (y, x, stop, user), count in journey_rows_dict.items():
            journey_rows.append([y, x, stop, user, count])

        output_dir = Path('finished exports')
        output_dir.mkdir(parents=True, exist_ok=True)
        with open(output_dir / "gis_number_export.csv", 'w', newline='', encoding='utf-8') as datei:
            schreiben = csv.writer(datei)
            schreiben.writerow(header_row)
            schreiben.writerows(journey_rows)
        print(f'Created {datei.name}')
        print(f'\n{50 * "-"}\n')

def traewelling_analysis(users_for_api_get: list[str] | str | None = None, delay_standard: int = 300,
                         start_date : str = '2000-01-01', end_date : str = '3000-01-01',
                         search_for_json : bool = True, do_users_analysis: bool = True, do_shared_analysis: bool = True,
                         do_geo_analysis: bool = False, search_for_correct_usernames: bool = False,
                         update_already_downloaded_checkins: bool = False, update_all_downloaded_users=False)->None:
    start_time = time()
    try:
        traewelling = Traewelling(users_for_api_get = users_for_api_get, delay_standard=delay_standard,
                                  start_date=start_date, end_date=end_date, search_for_json=search_for_json,
                                  search_for_correct_usernames=search_for_correct_usernames,
                                  update_past_downloaded_checkins=update_already_downloaded_checkins,
                                  update_all_downloaded_users=update_all_downloaded_users)
        if do_users_analysis:
            traewelling.create_user_excel()
        if do_shared_analysis:
            traewelling.create_shared_excel()
        if do_geo_analysis:
            traewelling.create_gis_number_csv()
            traewelling.create_gis_single_csv()
    except KeyboardInterrupt:
        print("\n\033[93mNotbremse gezogen! (KeyboardInterrupt)\033[0m")
    except Exception as e:
        print(f"\n\033[91mStörung im Betriebsablauf: {type(e).__name__} {e}\033[0m")
        print(f"\033[91m{traceback.format_exc()}\033[0m")
    finally:
        end_time = time()
        traewel_time = end_time - start_time
        if traewel_time > 200:
            print(f"\033[91mThank You, for träwelling {traewel_time / 60} minutes with Deutsche Bahn!\033[0m")
        else:
            print(f"\033[91mThank You, for träwelling {traewel_time} seconds with Deutsche Bahn!\033[0m")
        print(f'\n{50 * "-"}\n')


if __name__ == '__main__':
    traewelling_analysis(users_for_api_get= ['Hier die', 'verschiedenen User eintragen'], do_users_analysis=True, do_shared_analysis=True, search_for_json=True,
                         do_geo_analysis=True, search_for_correct_usernames=False, update_already_downloaded_checkins= False,
                         update_all_downloaded_users = True)#, start_date='2026-03-14', end_date='2026-03-31'



