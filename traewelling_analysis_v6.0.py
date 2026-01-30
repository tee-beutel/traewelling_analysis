
#traewelling_analysis_v6.0.py

import subprocess
import sys

# 1. Check/Install timezonefinder --Achtung KI-generiert--
try:
    import timezonefinder
except ImportError:
    print("Modul 'timezonefinder' nicht gefunden. Installation wird gestartet...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "timezonefinder"])
    print("timezonefinder erfolgreich installiert.")

# 2. Check/Install tzdata für zoneinfo
try:
    from zoneinfo import ZoneInfo, ZoneInfoNotFoundError
    try:
        # Teste, ob eine Standard-Zeitzone geladen werden kann
        ZoneInfo("Europe/Berlin")
    except ZoneInfoNotFoundError:
        print("Zeitzonendaten fehlen (tzdata). Installation wird gestartet...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "tzdata"])
        print("tzdata erfolgreich installiert.")
except ImportError:
    # Falls Python < 3.9 verwendet wird, existiert zoneinfo noch nicht
    print("Fehler: 'zoneinfo' ist erst ab Python 3.9 verfügbar. Bitte Python aktualisieren.")
    sys.exit(1)

import json
from pathlib import Path
from timezonefinder import TimezoneFinder
from zoneinfo import ZoneInfo
from time import time
import csv
import pandas as pd
from datetime import datetime
import re



class Journey:
    def __init__(self, data, tf: TimezoneFinder, timezone_info: dict):
        train_type_to_readable_dict = {'regional':'Regionalzug', 'bus':'Bus', 'ferry':'Fähre', 'national':'Regionalzug', 'regionalExp':'Fernverkehr',
                                       'nationalExpress':'Fernverkehr', 'suburban': 'S-Bahn', 'subway':'U-Bahn', 'taxi': 'Taxi', 'tram':'Tram',
                                       'plane':'Igitt ein Flugzeug'}
        self.__raw_data = data
        self.__status_id = data['status']['id']
        self.__trip_id = data['trip']['id']
        self.__user_name = data['status']['userDetails'].get('username', 'Username not found')
        self.__user_id = data['status']['userDetails'].get('id', 'N/A')

        self.__origin_stop = (f'{data['status']['train']['origin']['name']} ({data['status']['train']['origin']['rilIdentifier']})' if
                              data['status']['train']['origin'].get('rilIdentifier') else data['status']['train']['origin']['name'])
        self.__destination_stop = (f'{data['status']['train']['destination']['name']} ({data['status']['train']['destination']['rilIdentifier']})' if
                              data['status']['train']['destination'].get('rilIdentifier') else data['status']['train']['destination']['name'])
        self.__origin_coordinates =(data['trip']['origin'].get('latitude', 'N/A'), data['trip']['origin'].get('longitude', 'N/A'))
        self.__destination_coordinates = (data['trip']['destination'].get('latitude', 'N/A'), data['trip']['destination'].get('longitude', 'N/A'))

        self.__via_stations_data = data['trip'].get('stopovers', [])
        self.__via_stations = list(f'{d['name']} ({d['rilIdentifier']})' if d.get('rilIdentifier') else d['name']
                                   for d in self.__via_stations_data)
        number_origin = self.__via_stations.index(self.__origin_stop)
        number_destination = self.__via_stations.index(self.__destination_stop)
        self.__via_stations = self.__via_stations[number_origin+1:number_destination]


        self.__timezone_origin = timezone_info.get(self.__origin_stop, None)
        self.__timezone_destination = timezone_info.get(self.__destination_stop, None)
        if self.__timezone_origin is None:
            self.__timezone_origin = tf.timezone_at(lng=self.__origin_coordinates[1], lat=self.__origin_coordinates[0])
        if self.__timezone_destination is None:
            self.__timezone_destination = tf.timezone_at(lng=self.__destination_coordinates[1], lat=self.__destination_coordinates[0])


        self.__operator_info = data['status']['train']['operator']
        if self.__operator_info:
            self.__operator_name = self.__operator_info.get('name', 'Unknown Operator')
            self.__operator_id = self.__operator_info.get('id', None)
        else:
            self.__operator_name = 'Unknown Operator'
            self.__operator_id = None

        self.__line_name_raw = data['status']['train'].get('lineName', 'Unknown Line')
        self.__line_name = re.sub(r'\s*\(\d+\)', '', self.__line_name_raw).strip()

        self.__journey_distance = data['status']['train'].get('distance', 0)
        self.__line_number = data['status']['train']['number']


        given_train_type = train_type_to_readable_dict.get(data['status']['train']['category'], data['status']['train']['category'])
        if re.match(r"^(REX|RE|RB|Stoptrein|FEX)", self.__line_name) and given_train_type != 'Bus':
            self.__train_type = 'Regionalzug'
        elif re.match(r"^(RS)", self.__line_name):
            self.__train_type = 'S-Bahn'
        else:
            self.__train_type = given_train_type

        self.__journey_points = data['status']['train']['points']
        self.__journey_number = data['status']['train']['journeyNumber']

        self.__planned_departure = datetime.fromisoformat(data['status']['train']['origin']['departurePlanned']).astimezone(ZoneInfo(self.__timezone_origin))
        self.__planned_arrival = datetime.fromisoformat(data['status']['train']['destination']['arrivalPlanned']).astimezone(ZoneInfo(self.__timezone_destination))
        manualDeparture = data['status']['train']['manualDeparture']
        manualArrival = data['status']['train']['manualArrival']
        realtimeDeparture = data['status']['train']['origin']['departureReal']
        realtimeArrival = data['status']['train']['destination']['arrivalReal']

        self.__realtime_availability = False
        if manualDeparture is not None:
            self.__realtime_availability = True
            self.__real_departure = datetime.fromisoformat(manualDeparture).astimezone(ZoneInfo(self.__timezone_origin))
        elif realtimeDeparture is not None:
            self.__realtime_availability = True
            self.__real_departure = datetime.fromisoformat(realtimeDeparture).astimezone(ZoneInfo(self.__timezone_origin))
        else:
            self.__real_departure = self.__planned_departure

        if manualArrival is not None:
            self.__realtime_availability = True
            self.__real_arrival = datetime.fromisoformat(manualArrival).astimezone(ZoneInfo(self.__timezone_destination))
        elif realtimeArrival is not None:
            self.__realtime_availability = True
            self.__real_arrival = datetime.fromisoformat(realtimeArrival).astimezone(ZoneInfo(self.__timezone_destination))
        else:
            self.__real_arrival = self.__planned_arrival


        self.__departure_delay = (self.__real_departure - self.__planned_departure).total_seconds()
        self.__arrival_delay = (self.__real_arrival - self.__planned_arrival).total_seconds()



        self.__vehicle_number = 'N/A'
        self.__tags = data['status'].get('tags', None)
        if self.__tags is not None:
            for tag in self.__tags:
                if tag.get('key', None) == 'trwl:vehicle_number':
                    self.__vehicle_number = tag.get('value', 'N/A')

        self.__journey_time_planned = (self.__planned_arrival - self.__planned_departure).total_seconds()
        self.__journey_time_real = (self.__real_arrival - self.__real_departure).total_seconds()
        self.__journey_time_delta = (self.__journey_time_real - self.__journey_time_planned)


    def __str__(self):
        string_base =(f'Fahrt {self.__line_name} ({self.__journey_number}) von {self.__origin_stop} nach {self.__destination_stop} '
                     f'am {self.__planned_departure.strftime("%d.%m.%Y")} ' )
        match self.arrival_delay:
            case d if d>0:
                return string_base + f'mit {self.__arrival_delay / 60:.2f} Minuten Verspätung, Fzg: {self.__vehicle_number}'
            case d if d<0:
                return string_base + f'mit {-self.__arrival_delay / 60:.2f} Minuten Verfrühung, Fzg: {self.__vehicle_number}'
            case _:
                return string_base + f'war pünktlich, Fzg: {self.__vehicle_number}'

    def __lt__(self, other):
        if isinstance(other, Journey):
            return self.__real_arrival < other.__real_arrival
        elif isinstance(other, datetime):
            return self.__real_departure < other
        return None

    def delayed_by_standard(self, standard: int)->bool:
        return self.__arrival_delay >= standard


    @property
    def user_id(self):
        return self.__user_id
    @property
    def trip_id(self):
        return self.__trip_id
    @property
    def status_id(self):
        return self.__status_id
    @property
    def user_name(self):
        return self.__user_name
    @property
    def operator_name(self):
        return self.__operator_name
    @property
    def line_name(self):
        return self.__line_name
    @property
    def journey_distance(self):
        return self.__journey_distance
    @property
    def train_type(self):
        return self.__train_type
    @property
    def arrival_delay(self):
        return self.__arrival_delay
    @property
    def departure_delay(self):
        return self.__departure_delay
    @property
    def arrival_planned(self):
        return self.__planned_arrival
    @property
    def departure_planned(self):
        return self.__planned_departure
    @property
    def departure_real(self):
        return self.__real_departure
    @property
    def arrival_real(self):
        return self.__real_arrival
    @property
    def realtime_availability(self):
        return self.__realtime_availability
    @property
    def vehicle_number(self):
        return self.__vehicle_number
    @property
    def origin_stop(self):
        return self.__origin_stop
    @property
    def destination_stop(self):
        return self.__destination_stop
    @property
    def journey_number(self):
        return self.__journey_number
    @property
    def journey_time_real(self):
        return self.__journey_time_real
    @property
    def journey_time_delta(self):
        return self.__journey_time_delta
    @property
    def journey_points(self):
        return self.__journey_points
    @property
    def origin_coordinates(self):
        return self.__origin_coordinates
    @property
    def destination_coordinates(self):
        return self.__destination_coordinates
    @property
    def timezone_origin(self):
        return self.__timezone_origin
    @property
    def timezone_destination(self):
        return self.__timezone_destination
    @property
    def via_stations(self):
        return self.__via_stations



class User:
    def __init__(self, id: int, delay_standard : int)->None:
        self.__delay_standard = delay_standard
        self.__id : int = id
        self.__name: str | None = None
        self.__journeys : list[Journey] = []
        self.__number_of_journeys: int = 0
        self.__exported_days:int = 0
        self.__total_distance: int = 0
        self.__total_journey_time: int = 0
        self.__distance_type_sorted: dict|None= None
        self.__distance_operator_line_sorted = None
        self.__distance_operator_sorted = None
        self.__average_delay: float | None = None
        self.__realtime_availability: float | None = None
        self.__max_delay_journey: Journey | None = None
        self.__max_ahead_journey: Journey | None = None
        self.__cumulative_delay: int | None = None
        self.__delay_rate_standard: float | None = None
        self.__visited_stations: dict[str,int] = dict()
        self.__stations_with_via: dict[str, int] = dict()
        self.__used_vehicles: dict[tuple,int] = dict()
        self.__type_delay_sorted:dict[str,list]|None = None
        self.__type_number_sorted: dict[str, int] | None = None
        self.__type_delay_by_standard_dict:dict[str,int]|None = None
        self.__number_of_visited_stations: int = 0
        self.__number_of_visited_stations_with_via: int = 0


    def add_journey(self, journey: Journey)->None:
        self.__journeys.append(journey)

    def __repr__(self) -> str:
        return self.__name

    def __str__(self) -> str:
        return self.__name

    def __eq__(self, other)->bool:
        if isinstance(other,Journey):
            return self.__name == other.user_name
        elif isinstance(other,User):
            return self.__id == other.__id
        elif isinstance(other, str):
            return self.__name == other
        else:
            print(f'Kein Vergleich von {self.__name} mit {type(other).__name__} möglich')
            return False

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

    def get_import_length(self) -> None:
        self.__name = self.__journeys[-1].user_name

        date_list = []
        for journey in self.__journeys:
            date_list.extend([journey.arrival_real, journey.departure_real])
        first_day = min(date_list)
        last_day = max(date_list)
        self.__exported_days = (last_day.date ()- first_day.date()).days + 1
        print(f'In {self.__exported_days} Tagen hat {self.__name} {len(self.__journeys)} Fahrten gemacht. '
              f'Das sind {len(self.__journeys) / self.__exported_days:.2f} Fahrten pro Tag!\n'
              f'Es wurden Fahrten von {first_day.strftime("%d.%m.%Y")} bis '
              f'{last_day.strftime("%d.%m.%Y")} berücksichtigt')
        print(f'\n{50 * "-"}\n')
        self.__number_of_journeys = len(self.__journeys)
        self.__total_distance = sum(journey.journey_distance for journey in self.__journeys)
        self.__total_journey_time = sum(journey.journey_time_real for journey in self.__journeys)


    def user_distance_time_analysis_execute(self):
        distance_type: dict = {}
        distance_operator: dict = {}
        distance_operator_line: dict = {}


        for journey in self.__journeys:
            journey_distance = journey.journey_distance
            journey_time = journey.journey_time_real
            journey_arrival_delay = max(journey.arrival_delay,0)
            journey_realtime_availability = journey.realtime_availability
            journey_delayed_by_standard = journey.delayed_by_standard(self.__delay_standard)
            journey_points = journey.journey_points
            train_type = journey.train_type
            operator_name = journey.operator_name
            line_tuple = (operator_name, journey.line_name)


            distance_type[train_type] = {'distance': journey_distance + distance_type.get(train_type, {}).get('distance', 0),
                                         'time': journey_time + distance_type.get(train_type, {}).get('time', 0),
                                         'arrival_delay': journey_arrival_delay + distance_type.get(train_type, {}).get('arrival_delay', 0),
                                         'realtime_availability': journey_realtime_availability + distance_type.get(train_type, {}).get('realtime_availability', 0),
                                         'sum': 1 + distance_type.get(train_type, {}).get('sum', 0),
                                         'delayed_by_standard': journey_delayed_by_standard + distance_type.get(train_type, {}).get('delayed_by_standard', 0),
                                         'points': journey_points + distance_type.get(train_type, {}).get('points', 0)}

            distance_operator_line[line_tuple] = {'distance': journey_distance + distance_operator_line.get(line_tuple, {}).get('distance', 0),
                                                  'time': journey_time + distance_operator_line.get(line_tuple, {}).get('time', 0),
                                                  'arrival_delay': journey_arrival_delay + distance_operator_line.get(line_tuple, {}).get('arrival_delay', 0),
                                                  'realtime_availability': journey_realtime_availability + distance_operator_line.get(line_tuple, {}).get('realtime_availability', 0),
                                                  'sum': 1 + distance_operator_line.get(line_tuple, {}).get('sum', 0),
                                                  'delayed_by_standard': journey_delayed_by_standard + distance_operator_line.get(line_tuple, {}).get('delayed_by_standard', 0),
                                                  'points': journey_points + distance_operator_line.get(line_tuple, {}).get('points', 0)}

            distance_operator[operator_name] = {'distance': journey_distance + distance_operator.get(operator_name, {}).get('distance', 0),
                                                'time': journey_time + distance_operator.get(operator_name, {}).get('time', 0),
                                                'arrival_delay': journey_arrival_delay + distance_operator.get(operator_name, {}).get('arrival_delay', 0),
                                                'realtime_availability': journey_realtime_availability + distance_operator.get(operator_name, {}).get('realtime_availability', 0),
                                                'sum': 1 + distance_operator.get(operator_name, {}).get('sum', 0),
                                                'delayed_by_standard': journey_delayed_by_standard + distance_operator.get(operator_name, {}).get('delayed_by_standard', 0),
                                                'points': journey_points + distance_operator.get(operator_name, {}).get('points', 0)}




        self.__distance_type_sorted = dict(sorted(distance_type.items(), key=lambda x: x[1].get('distance', 0), reverse=True))
        self.__distance_operator_line_sorted = dict(sorted(distance_operator_line.items(), key=lambda x: x[1].get('distance', 0), reverse=True))
        self.__distance_operator_sorted = dict(sorted(distance_operator.items(), key=lambda x: x[1].get('distance', 0), reverse=True))

    def __user_distance_analysis_output(self):
        print(f'Distanzanalyse für {self.__name}:\n')
        print(f'Gesamtdistanz in {self.__exported_days} Tagen: {self.__total_distance / 1000:.2f} Kilometer')
        print(f'Durchschnittlich {(self.__total_distance / self.__exported_days) / 1000:.2f} Kilometer pro Tag gemacht.', '\n')

        print('Anteile der Verschiedenen Verkehrsmittel am Mix (Distanz/Zeit):')
        for train_type, distance_time_per_type in self.__distance_type_sorted.items():
            print(f'{train_type:.<16}: {distance_time_per_type.get('distance') / self.__total_distance * 100:<5.2f}% '
                  f'insgesamt {distance_time_per_type.get('distance') / 1000:<10.2f}km und {distance_time_per_type.get('time') / self.__total_journey_time * 100:<5.2f}% also'
                  f' {distance_time_per_type.get('time')/60:>7.0f} Minuten')
        print(f'\n{50 * "-"}\n')

        current_place = 0
        print(f'Das sind die {len(self.__distance_operator_line_sorted)} Linen von {self.__name}:')
        for (operator_name, line_name), line_distance in self.__distance_operator_line_sorted.items():
            current_place += 1
            print(
                f'{current_place:>3}. Linie {line_name:.<18} von {operator_name:.<60} {line_distance.get('distance',0) / 1000:<7.2f} '
                f'Kilometer und {line_distance.get('time',0) / 60:>7.0f} Minuten')
        print(f'\n{50 * "-"}\n')

        current_place = 0
        print(f'Das sind die {len(self.__distance_operator_sorted)} Betreiber von {self.__name}:')
        for operator_name, operator_data in self.__distance_operator_sorted.items():
            current_place += 1
            print(
                f'{current_place:>3}. Betreiber {operator_name:.<60}hat {operator_data.get('distance',0) / 1000:<9.2f} Kilometer, das sind '
                f'{operator_data.get('distance',0) / self.__total_distance * 100:>5.2f}% und und {operator_data.get('time',0) / 60:>7.0f} Minuten'
                f'{((operator_data.get('distance',0) / 1000) / (operator_data.get('time',0) / 3600) if operator_data.get('time',0) else 0):>7.2f}km/h')
        print(f'\n{50 * "-"}\n')

    def distance_analysis_per_user(self):
        self.user_distance_time_analysis_execute()
        self.__user_distance_analysis_output()

    def delay_analysis_execute(self) -> None:
        self.__cumulative_delay = sum(j.arrival_delay for j in self.__journeys if j.arrival_delay > 0)
        self.__average_delay = self.__cumulative_delay / len(self.__journeys)
        available_vs_not = [0, 0]
        for j in self.__journeys:
            if j.realtime_availability:
                available_vs_not[0] += 1
            else:
                available_vs_not[1] += 1
        self.__realtime_availability = available_vs_not[0] / sum(available_vs_not)

        type_delay_dict: dict[str, list] = dict()
        type_number_dict: dict[str, int] = dict()
        self.__type_delay_by_standard_dict: dict[str, int] = dict()
        current_max_delay_j = self.__journeys[0]
        current_max_ahead_j = self.__journeys[0]
        delayed_by_standard = 0
        for j in self.__journeys[1:]:
            if j.arrival_delay > current_max_delay_j.arrival_delay:
                current_max_delay_j = j
            if j.arrival_delay < current_max_ahead_j.arrival_delay:
                current_max_ahead_j = j
            if j.delayed_by_standard(self.__delay_standard):
                delayed_by_standard += 1

            if j.train_type in type_delay_dict:
                type_delay_dict[j.train_type] += [j.arrival_delay]
                type_number_dict[j.train_type] += 1

            else:
                type_delay_dict[j.train_type] = [j.arrival_delay]
                type_number_dict[j.train_type] = 1
            if j.train_type in self.__type_delay_by_standard_dict:
                if j.delayed_by_standard(self.__delay_standard):
                    self.__type_delay_by_standard_dict[j.train_type] += 1
            else:
                if j.delayed_by_standard(self.__delay_standard):
                    self.__type_delay_by_standard_dict[j.train_type] = 1


        self.__max_delay_journey = current_max_delay_j
        self.__max_ahead_journey = current_max_ahead_j
        self.__delay_rate_standard = delayed_by_standard / len(self.__journeys)
        self.__type_delay_sorted = dict(sorted(type_delay_dict.items(), key=lambda x: sum(d for d in x[1] if d > 0), reverse=True))
        self.__type_number_sorted = dict(sorted(type_number_dict.items(), key=lambda x: x[1], reverse=True))

    def __delay_analysis_output(self) -> None:
        print(f'\nVerspätungsstatistik für {self.__name}')
        print(f'Für {self.__realtime_availability * 100:.2f}% deiner Fahrten gibt es Echtzeitdaten')
        print(f'{self.__delay_rate_standard*100 :.2f}% aller Fahren sind mehr als fünf Minuten zu spät')
        print(f'Durchschnittliche Verspätung {self.__average_delay:.2f} Sekunden.')
        print(f'Insgesamt in {self.__exported_days} Tagen: {self.__cumulative_delay/3600:.3f} Stunden.')
        print(f'Am meisten verspätet war die {self.__max_delay_journey}')
        print(f'Am meisten verfrüht war die {self.__max_ahead_journey}')
        print('\nVerspätung nach Zugtyp')
        for train_type in self.__type_delay_sorted.keys():
            print(
                f'{train_type:.<16}insgesamt {sum(d for d in self.__type_delay_sorted[train_type] if d > 0) / 60:<5.0f} Minuten '
                f'{self.__type_delay_by_standard_dict.get(train_type,0) / self.__type_number_sorted[train_type]*100:<5.2f}% mehr als 5 Minuten '
                f'(bei {self.__type_number_sorted[train_type]:<4} Fahrten)')
        print(f'\n{50 * "-"}\n')

    def delay_analysis(self) -> None:
        self.delay_analysis_execute()
        self.__delay_analysis_output()

    def visited_station_execution(self):
        for j in self.__journeys:
            self.__visited_stations[j.origin_stop] = self.__visited_stations.get(j.origin_stop, 0) + 1
            self.__visited_stations[j.destination_stop] = self.__visited_stations.get(j.destination_stop, 0) + 1
        self.__number_of_visited_stations = len(self.__visited_stations)
        self.__visited_stations = dict(sorted(self.__visited_stations.items(), key=lambda x: x[1], reverse=True))

    def __visited_station_output(self) -> None:
        print(f'\n{self.__name} war an folgenden {self.__number_of_visited_stations} Haltestellen:')
        current_place, place_previous = 0, 0
        visited_previous = list(self.__visited_stations.values())[0] + 1

        for stop, times_visited in self.__visited_stations.items():
            current_place += 1
            if times_visited != visited_previous:
                place_previous = current_place
            visited_previous = times_visited
            print(f"{place_previous:>3}. {stop:.<55} wurde {times_visited:>3}-mal besucht")
        print(f'\n{50 * "-"}\n')

    def __visited_station_with_via_execution(self) -> None:
        self.__stations_with_via = self.visited_stations
        for j in self.__journeys:
            for name in j.via_stations:
                self.__stations_with_via[name] = self.__stations_with_via.get(name, 0) + 1
        self.__number_of_visited_stations_with_via = len(self.__stations_with_via)
        self.__stations_with_via = dict(sorted(self.__stations_with_via.items(), key=lambda x: x[0], reverse=False))
        self.__stations_with_via = dict(sorted(self.__stations_with_via.items(), key=lambda x: x[1], reverse=True))

    def __visited_station_with_via_output(self) -> None:
        print(f'\n{self.__name} war an folgenden {self.__number_of_visited_stations_with_via} Haltestellen (Zwichenhaltestellen inkusive):')
        current_place, place_previous = 0, 0
        visited_previous = list(self.__stations_with_via.values())[0] + 1

        for stop, times_visited in self.__stations_with_via.items():
            current_place += 1
            if times_visited != visited_previous:
                place_previous = current_place
            visited_previous = times_visited
            print(f"{place_previous:>3}. {stop:.<55} wurde {times_visited:>3}-mal besucht")
        print(f'\n{50 * "-"}\n')

    def vehicle_execution(self):
        for j in self.__journeys:
            vehicle_ident = (j.vehicle_number,j.operator_name)
            self.__used_vehicles[vehicle_ident] = self.__used_vehicles.get(vehicle_ident, 0) + 1
        self.__used_vehicles = dict(sorted(self.__used_vehicles.items(), key=lambda x: x[1], reverse=True))

    def __vehicle_output(self) -> None:
        print(f'\n{self.__name} hat folgende {len(self.__used_vehicles)} Fahrzeuge genutzt:')
        current_place, place_previous = 0, 0
        used_previous = list(self.__used_vehicles.values())[0] + 1
        for (vehicle, operator), number_used in self.__used_vehicles.items():
            current_place += 1
            if number_used != used_previous:
                place_previous = current_place
            used_previous = number_used
            if vehicle == 'N/A':
                print(f'{place_previous:>3}.{number_used:>3} Fahrten haben keine Fahrzeugnummer')
            else:
                print(f'{place_previous:>3}.Fahrzeug {vehicle:.<30}. {number_used:>3}-mal genutzt')
        print(f'\n{50 * "-"}\n')

    def visited_station(self):
        self.visited_station_execution()
        self.__visited_station_output()

    def visited_station_with_via(self):
        self.__visited_station_with_via_execution()
        self.__visited_station_with_via_output()

    def vehicle_analysis(self) -> None:
        self.vehicle_execution()
        self.__vehicle_output()

    def create_journeys_csv(self) -> None:
        header_row = ['Kategorie','Linie','Fahrtnummer','Betreiber','Abfahrtshaltestelle','Abfahrt geplant','Abfahrt real','Abw ab',
                      'Zielstation','Ankunft geplant','Ankunft real','Abw an','Reisezeit (min)','Reisezeit (hh:mm)',
                      'Delta','Entfernung (m)', 'Entfernung (km)', 'Punkte']
        journey_rows: list[list] = []
        for j in self.__journeys:
            j_row = [j.train_type, j.line_name, j.journey_number, j.operator_name, j.origin_stop, j.departure_planned.strftime('%d.%m.%Y %H:%M')]
            if j.realtime_availability:
                j_row.append(j.departure_real.strftime('%d.%m.%Y %H:%M'))
            else:
                j_row.append('N/A')
            j_row.extend([int(round(j.departure_delay/60,0)), j.destination_stop, j.arrival_planned.strftime('%d.%m.%Y %H:%M')])
            if j.realtime_availability:
                j_row.append(j.arrival_real.strftime('%d.%m.%Y %H:%M'))
            else:
                j_row.append('N/A')
            #journey_hours = int(j.journey_time_real/3600)
            #journey_minutes = int((j.journey_time_real-journey_hours*3600)/60)
            journey_hours, journey_minutes = divmod(max(j.journey_time_real/60,0), 60)
            j_row.extend([int(round(j.arrival_delay/60,0)),int(round(j.journey_time_real/60,0)),f"{int(journey_hours):02}:{int(journey_minutes):02}",
                          int(round(j.journey_time_delta/60,0)), j.journey_distance,
                          f'{j.journey_distance/1000:.3f} km', j.journey_points, j.realtime_availability])
            journey_rows.append(j_row)

        #time_per_day = self.__total_journey_time/self.__exported_days
        #total_hours = int(self.__total_journey_time/3600)
        #total_minutes = int(((self.__total_journey_time-total_hours*3600)/60))
        total_hours, total_minutes = divmod(self.__total_journey_time/60, 60)


        total_row = ['Gesamtwert:', f'{len(self.distance_operator_line_sorted)} verschiedene Linien',
                     f'{len(self.__journeys)} Fahrten',
                     f'{len(self.distance_operator_sorted)} verschidene Betreiber',
                     f'{self.number_of_visited_stations} einmalige Haltestellen',
                     f'{self.__exported_days} Tage exportiert',
                     '',
                     f'{sum(j.departure_delay for j in self.__journeys) / 60:.0f} Minuten',
                     '', '', '',
                     f'{sum(j.arrival_delay for j in self.__journeys) / 60:.0f} Minuten',
                     f'{int(self.__total_journey_time / 60)} Minuten',
                     f"{total_hours:02.0f}:{int(round(total_minutes, 0)):02}",
                     f'{sum(j.journey_time_delta / 60 for j in self.__journeys):.0f} Minuten',
                     f'{self.__total_distance} Meter', f'{self.__total_distance / 1000:.2f} Kilometer',
                     f'{sum(j.journey_points for j in self.__journeys)} Punkte']

        #per_day_hours = int(time_per_day / 3600)
        #per_day_minutes = int(((time_per_day - per_day_hours * 3600) / 60))
        per_day_hours, per_day_minutes = divmod((self.__total_journey_time/self.__exported_days)/60, 60)

        average_per_day_row = ['Durchschnitt pro Tag:',
                               f'{len(self.distance_operator_line_sorted) / self.__exported_days:.3f} verschiedene Linien',
                               f'{len(self.__journeys) / self.__exported_days:.3f} Fahrten',
                               f'{len(self.distance_operator_sorted) / self.__exported_days:.3f} verschidene Betreiber',
                               f'{self.number_of_visited_stations / self.__exported_days:.3f} einmalige Haltestellen',
                               '', '',
                               f'{(sum(j.departure_delay for j in self.__journeys) / 60) / self.__exported_days :.3f} Minuten',
                               '', '', '',
                               f'{(sum(j.arrival_delay for j in self.__journeys) / 60) / self.__exported_days :.3f} Minuten',
                               f'{self.__total_journey_time / 60 / self.__exported_days :.3f} Minuten',
                               f"{per_day_hours:02.0f}:{round(per_day_minutes, 0):02.0f}",
                               f'{sum(j.journey_time_delta / 60 for j in self.__journeys) / self.__exported_days:.3f} Minuten',
                               f'{self.__total_distance / self.__exported_days :.3f} Meter',
                               f'{self.__total_distance / 1000 / self.__exported_days:.2f} Kilometer',
                               f'{sum(j.journey_points for j in self.__journeys) / self.__exported_days :.3f} Punkte']

        #time_per_journey = self.__total_journey_time / self.__number_of_journeys
        #per_journey_hours = int(time_per_journey / 3600)
        #per_journey_minutes = int(((time_per_journey - per_journey_hours * 3600) / 60))
        per_journey_hours, per_journey_minutes = divmod((self.__total_journey_time / self.__number_of_journeys)/60, 60)
        average_per_journey_row = ['Durchschnitt pro Fahrt:',
                               f'{len(self.distance_operator_line_sorted) / self.__number_of_journeys:.3f} verschiedene Linien',
                               f'{len(self.__journeys) / self.__number_of_journeys:.3f} Fahrten',
                               f'{len(self.distance_operator_sorted) / self.__number_of_journeys:.3f} verschidene Betreiber',
                               f'{self.number_of_visited_stations / self.__number_of_journeys:.3f} einmalige Haltestellen',
                               '', '',
                               f'{(sum(j.departure_delay for j in self.__journeys) / 60) / self.__number_of_journeys :.3f} Minuten',
                               '', '', '',
                               f'{(sum(j.arrival_delay for j in self.__journeys) / 60) / self.__number_of_journeys :.3f} Minuten',
                               f'{self.__total_journey_time / self.__number_of_journeys / 60:.3f} Minuten',
                               f"{per_journey_hours:02.0f}:{round(per_journey_minutes, 0):02.0f}",
                               f'{sum(j.journey_time_delta / 60 for j in self.__journeys) / self.__number_of_journeys:.3f} Minuten',
                               f'{self.__total_distance / self.__number_of_journeys :.3f} Meter',
                               f'{self.__total_distance / 1000 / self.__number_of_journeys:.2f} Kilometer',
                               f'{sum(j.journey_points for j in self.__journeys) / self.__number_of_journeys :.3f} Punkte']

        journey_rows.extend([[], total_row, average_per_day_row,average_per_journey_row])



        with open(f"{self.__name}'s_journeys.csv", 'w', newline='', encoding='utf-8-sig') as datei:
            schreiben = csv.writer(datei, delimiter=';')
            schreiben.writerow(header_row)
            schreiben.writerows(journey_rows)
        print(f'Created {datei.name}')
        print(f'\n{50 * "-"}\n')


    def create_excel(self) -> None:
        j_data: list[dict] = []
        for j in self.__journeys:
            journey_hours, journey_minutes = divmod(max(j.journey_time_real / 60, 0), 60)
            j_data.append({'Kategorie': j.train_type,
                          'Linie': j.line_name,
                          'Fahrtnummer': j.journey_number,
                          'Betreiber': j.operator_name,
                          'Fahrzeugnummer': j.vehicle_number,
                          'Abfahrtshaltestelle': j.origin_stop,
                          'Abfahrt geplant': j.departure_planned.strftime('%d.%m.%Y %H:%M'),
                          'Abfahrt real': j.departure_real.strftime('%d.%m.%Y %H:%M') if j.realtime_availability else 'N/A',
                          'Abw ab': int(round(j.departure_delay / 60, 0)),
                          'Zielstation': j.destination_stop,
                          'Ankunft geplant': j.arrival_planned.strftime('%d.%m.%Y %H:%M'),
                          'Ankunft real': j.arrival_real.strftime('%d.%m.%Y %H:%M') if j.realtime_availability else 'N/A',
                          'Abw an': int(round(j.arrival_delay / 60, 0)),
                          'Reisezeit (min)': int(round(j.journey_time_real / 60, 0)),
                          'Reisezeit (hh:mm)': f"{int(journey_hours):02}:{int(journey_minutes):02}",
                          'Delta': int(round(j.journey_time_delta / 60, 0)),
                          'Entfernung (m)': j.journey_distance,
                          'Entfernung (km)': round(j.journey_distance / 1000, 3),
                          'Punkte': j.journey_points,
                          'Link': f'=HYPERLINK("https://traewelling.de/status/{j.status_id}", "https://traewelling.de/status/{j.status_id}")'})

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
                                f'{sum(j.journey_points for j in self.__journeys) / self.__number_of_journeys :.3f} Punkte']}

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
                              'Punkte': type_info.get('points', 0)})

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
                                  'Punkte': operator_info.get('points', 0)})

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
                              'Punkte': line_info.get('points', 0)})

        line_dataframe = pd.DataFrame(line_data)

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
                              'Besuche': times_visited})

        stop_with_via_dataframe = pd.DataFrame(stop_with_via_data)

        vehicle_data = []
        for (vehicle, operator), number_used in self.used_vehicles.items():
            vehicle_data.append({f'Fahrzeug ({len(self.used_vehicles)})': vehicle,
                                 'Betreiber': operator,
                                 'Fahrten': number_used})
        vehicle_dataframe = pd.DataFrame(vehicle_data)


        try:
            with pd.ExcelWriter(f"{self.__name}'s_data.xlsx",engine='openpyxl') as writer:
                journeys_dataframe.to_excel(writer, sheet_name='Fahrtenliste', index=False)
                stats_dataframe.to_excel(writer, sheet_name='Statistiken', index=False)
                type_dataframe.to_excel(writer, sheet_name='Kategorie', index=False)
                operator_dataframe.to_excel(writer, sheet_name='Betreiber', index=False)
                line_dataframe.to_excel(writer, sheet_name='Linie', index=False)
                stop_dataframe.to_excel(writer, sheet_name='Haltestelle (ohne Via)', index=False)
                stop_with_via_dataframe.to_excel(writer, sheet_name='Haltestelle (mit Via)', index=False)
                vehicle_dataframe.to_excel(writer, sheet_name='Fahrzeuge', index=False)

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


                print(f"Created {self.__name}'s_data.xlsx")

        except Exception as e:
            print(f"{self.__name}'s_data.xlsx could not be created:\n\n{e}")

        print(f'\n{50 * "-"}\n')



    @property
    def name(self) -> str:
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
            self.__visited_station_with_via_execution()
        return self.__number_of_visited_stations_with_via
    @property
    def stations_with_via(self) -> dict:
        if not self.__stations_with_via:
            self.__visited_station_with_via_execution()
        return self.__stations_with_via



class Traewelling:
    def __init__(self, delay_standard: int = 300):
        self.__delay_standard = delay_standard

        print("Importing träwelling data...")

        path = Path.cwd()
        file_names = []
        for file in path.glob('*.json'):
            file_names.append(file)

        if not file_names:
            raise FileNotFoundError("No .json files found")

        self.__whole_travel_data = []

        for file_name in file_names:
            with open(file_name, encoding='utf-8') as f:
                travel = json.load(f)
            travel_data = travel.get('data',{})
            self.__whole_travel_data += travel_data
            print(f'Imported {len(travel_data)} travel data from {file_name}')

        tf = TimezoneFinder(in_memory= True)

        print('Getting timezones for journeys...')
        self.__journeys : list[Journey] = []
        userid_set : set[int] = set()
        id_list : set[int] = set()
        number_skipped_journeys = 0
        self.__timezone_per_stop = {}
        for journey_data in self.__whole_travel_data:

            status_id = journey_data['status']['id']
            if status_id in id_list:
                number_skipped_journeys += 1
            else:
                journey_object = Journey(journey_data, tf, self.__timezone_per_stop)
                self.__journeys.append(journey_object)
                self.__timezone_per_stop[journey_object.destination_stop] = journey_object.timezone_destination
                id_list.add(status_id)
                userid_set.add(journey_object.user_id)



        self.__journeys.sort()

        self.__user_id_dict : dict[int,User] = dict()
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
            print(f'Successfully imported {assigned_travel_data} journeys ({number_skipped_journeys} excess journeys were deleted)\n')
        else:
            raise Exception(f'Something went wrong: there should be {raw_travel_data} journeys, there are just {assigned_travel_data}\n')
        print(f'\n{50*'-'}\n')

        for user in self.__user_id_dict.items():
            user[1].get_import_length()

    def __str__(self)->str:
        print_string = ''
        for journey in self.__journeys:
            print_string += f'{journey}\n'
        return print_string

    def user_journey_list(self) -> str:
        print_string = ''
        for user in self.__user_name_dict.items():
            print_string += f'{user[1]}\n'
        return print_string

    def distance_analysis(self, user_name_list: str | list[str] | None= None) -> None:
        user_list = self.__user_input_to_object(user_name_list, 1)
        if user_list is None:
            return
        for user in user_list:
            user.distance_analysis_per_user()

    def distance_type_comparison_absolut(self, user_name_list: str | list[str] | None= None) -> None:
        user_list = self.__user_input_to_object(user_name_list, 2)
        if user_list is None:
            print('\nAbsoluter Vergleich der Distanzen benötigt mindestens zwei User')
            return
        type_set: set[str] = set()
        for user in user_list:
            if user.distance_type_sorted is None:
                user.user_distance_analysis_execute()
            for train_type , train_distance in user.distance_type_sorted.items():
                type_set.add(train_type)
        type_list = sorted(type_set)
        names = f'{user_list[0].name}, ' + ', '.join([user.name for user in user_list[1:]])
        print(f'\nAbsoluter Vergleich der Distanzen von {names} für Verkehrsmittel:')
        for type in type_list:
            user_type_distance : dict[str,int] = dict()
            for user in user_list:
                user_type_distance[user.name] = user.distance_type_sorted.get(type,{}).get('distance', 0)
            user_type_distance_sorted = sorted(user_type_distance.items(), key=lambda x: (x[1]), reverse=True)
            current_place = 0
            print(f'\n{type}:')
            for user, distance in user_type_distance_sorted:
                current_place += 1
                print(f'{current_place:>3}. User {user:<20} {distance / 1000:<7.2f} Kilometer')
        print(f'\n{50*"-"}\n')

    def distance_type_comparison_relative(self, user_name_list: str | list[str] | None= None) -> None:
        user_list = self.__user_input_to_object(user_name_list, 2)
        if user_list is None:
            print('\nRelativer Vergleich der Distanzen benötigt mindestens zwei User')
            return
        type_set: set[str] = set()
        for user in user_list:
            for train_type , train_distance in user.distance_type_sorted.items():
                type_set.add(train_type)
        type_list = sorted(type_set)
        names = f'{user_list[0].name}, ' + ', '.join([user.name for user in user_list[1:]])
        print(f'\nRelativer Vergleich der Distanzen von {names} für Verkehrsmittel:')
        for type in type_list:
            user_type_distance : dict[str,float] = dict()
            for user in user_list:
                user_type_distance[user.name] = (user.distance_type_sorted.get(type,{}).get('distance', 0) / user.exported_days)
            user_type_distance_sorted = sorted(user_type_distance.items(), key=lambda x: (x[1]), reverse=True)
            current_place = 0
            print(f'\nFür {type}:')
            for user, dic in user_type_distance_sorted:
                current_place += 1
                print(f'{current_place:>3}. User {user:<20} {dic / 1000:<7.4f} Kilometer pro Tag')

    def operator_user_comparison(self, user_name_list: str | list[str] | None= None) -> None:
        user_list = self.__user_input_to_object(user_name_list, 2)
        if user_list is None:
            print('\nVergleich der Distanzen der Betreiber benötigt mindestens zwei User')
            return
        dicts = [user.distance_operator_sorted for user in user_list]
        common_operators = set(dicts[0].keys()).intersection(*(d.keys() for d in dicts[1:]))
        names = f'{user_list[0].name}, ' + ', '.join([user.name for user in user_list[1:]])
        if not common_operators:
            print(f'{names} haben keine gemeinsamen Betreiber')
            return
        if 'Unknown Operator' in common_operators:
            common_operators.remove('Unknown Operator')
        print(f'Hier die gemeinsamen Betreiber von {names}:\n')
        common_operators = sorted(common_operators)
        for operator in common_operators:
            print(f'\nBetreiber {operator}:')
            operators_distances: list[tuple] = []
            for user in user_list:
                operators_distances.append((user.name, user.distance_operator_sorted.get(operator)))
            operators_distances = sorted(operators_distances, key=lambda x: (x[1].get('distance')), reverse=True)
            string_operator = ''
            for user_name, dic in operators_distances:
                string_operator += f'{user_name}: {dic.get('distance', 0)/1000:.2f} km und {dic.get('time', 0)/60:.0f}\n'
            print(string_operator)
        print(f'\n{50 * "-"}\n')


    def delay_analysis(self, user_name_list: str | list[str]= None) -> None:
        user_list = self.__user_input_to_object(user_name_list, 1)
        if user_list is None:
            return
        for user in user_list:
            user.delay_analysis()

    def station_analysis(self, user_name_list: str | list[str]= None) -> None:
        user_list = self.__user_input_to_object(user_name_list, 1)
        if user_list is None:
            return
        for user in user_list:
            user.visited_station()

    def station_analysis_with_via(self, user_name_list: str | list[str]= None) -> None:
        user_list = self.__user_input_to_object(user_name_list, 1)
        if user_list is None:
            return
        for user in user_list:
            user.visited_station_with_via()

    def shared_station_analysis(self, user_name_list: str | list[str]= None) -> None:
        user_list = self.__user_input_to_object(user_name_list, 2)
        if user_list is None:
            return
        dicts = [user.visited_stations for user in user_list]
        common_stations = set(dicts[0].keys()).intersection(*(d.keys() for d in dicts[1:]))
        names = f'{user_list[0].name}, ' + ', '.join([user.name for user in user_list[1:]])
        if not common_stations:
            print(f'{names} haben keine gemeinsamen Stationen')
            return
        print(f'Hier die gemeinsamen Stationen von {names}:\n')
        common_stations = sorted(common_stations)
        for station in common_stations:
            print(f'\nStation {station}:')
            visits: list[tuple]= []
            for user in user_list:
                visits.append((user.name, user.visited_stations.get(station)))
            visits = sorted(visits, key=lambda x: (x[1]), reverse=True)
            string_visits = ''
            for user_name, number in visits:
                string_visits += f'{user_name}: {number}  '
            print(string_visits)
        print(f'\n{50*"-"}\n')

    def vehicle_analysis(self, user_name_list: str | list[str]= None) -> None:
        user_list = self.__user_input_to_object(user_name_list, 1)
        if user_list is None:
            return
        for user in user_list:
            user.vehicle_analysis()

    def shared_trips_analysis(self, user_name_list: str | list[str]= None) -> None:
        user_list = self.__user_input_to_object(user_name_list, 1)
        if user_list is None:
            return
        trip_id_dict_unfiltered = {}
        trip_id_dict_filtered = {}
        trip_dict_object = {}
        for j in self.__journeys:
            if j.user_name in user_list:
                trip_id_dict_unfiltered[j.trip_id] = trip_id_dict_unfiltered.get(j.trip_id, []) + [j.user_id]
                trip_dict_object.update({(j.trip_id, j.user_name): j})


        for trip_id, user_ids in trip_id_dict_unfiltered.items():
            if len(user_ids) >= 2:
                user_names = []
                for user_id in user_ids:
                    user_names.append(self.__user_id_dict.get(user_id))
                user_names.sort(reverse=True)
                trip_id_dict_filtered[trip_id] = user_names

        for trip_id, user_names in trip_id_dict_filtered.items():
            for user_name in user_names:
                print(f'{user_name}: {trip_dict_object.get((trip_id, user_name.name), 'N/A')}')
            print('\n')



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

    def do_whole_calculations(self, user_name_list = None) -> None:
        self.distance_analysis(user_name_list)
        self.delay_analysis(user_name_list)
        self.station_analysis(user_name_list)
        if len(self.__user_id_dict) > 1:
            self.distance_type_comparison_relative(user_name_list)
            self.distance_type_comparison_absolut(user_name_list)
            self.operator_user_comparison(user_name_list)
            self.shared_station_analysis(user_name_list)
        print("\033[91mThank You, for träwelling with Deutsche Bahn!\033[0m")
        print(f'\n{50 * "-"}\n')

    def create_journeys_csv(self, user_name_list: str | list[str]= None) -> None:
        user_list = self.__user_input_to_object(user_name_list, 1)
        if user_list is None:
            return
        for user in user_list:
            user.create_journeys_csv()

    def create_user_excel(self, user_name_list: str | list[str]= None) -> None:
        user_list = self.__user_input_to_object(user_name_list, 1)
        if user_list is None:
            return
        for user in user_list:
            user.create_excel()

    def create_gis_single_csv(self, user_name_list: str | list[str]= None) -> None:
        user_list = self.__user_input_to_object(user_name_list, 1)
        if user_list is None:
            return
        header_row = ['y', 'x', 'Stopname', 'Username']
        journey_rows = []
        for j in self.__journeys:
            if j.user_name in user_list:
                journey_rows.extend([[j.destination_coordinates[0], j.destination_coordinates[1], j.destination_stop, j.user_name],
                                     [j.origin_coordinates[0], j.origin_coordinates[1], j.origin_stop, j.user_name]])
        with open(f"gis_single_export.csv", 'w', newline='', encoding='utf-8') as datei:
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
            if j.user_name in user_list:
                destination_data = (j.destination_coordinates[0], j.destination_coordinates[1], j.destination_stop, j.user_name)
                origin_data = (j.origin_coordinates[0], j.origin_coordinates[1], j.origin_stop, j.user_name)

                journey_rows_dict[destination_data] = journey_rows_dict.get(destination_data,0) + 1
                journey_rows_dict[origin_data] = journey_rows_dict.get(origin_data,0) + 1

        for (y, x, stop, user), count in journey_rows_dict.items():
            journey_rows.append([y, x, stop, user, count])


        with open(f"gis_number_export.csv", 'w', newline='', encoding='utf-8') as datei:
            schreiben = csv.writer(datei)
            schreiben.writerow(header_row)
            schreiben.writerows(journey_rows)
        print(f'Created {datei.name}')
        print(f'\n{50 * "-"}\n')




if __name__ == '__main__':
    start_time = time()
    traewelling = Traewelling()
    traewelling.create_user_excel()
    end_time = time()
    print(f'Traewelling took {end_time - start_time} seconds')




