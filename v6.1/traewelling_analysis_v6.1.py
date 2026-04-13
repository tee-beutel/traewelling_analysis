
#traewelling_analysis_v6.1.py

import json
from pathlib import Path
from timezonefinder import TimezoneFinder
from zoneinfo import ZoneInfo
from time import time
import csv
import pandas as pd
from datetime import datetime
import re
from collections import Counter



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

        o, d = data['status']['train']['origin'], data['status']['train']['destination']
        self.__origin_stop = (f'{o['name']} ({o['rilIdentifier']})' if o.get('rilIdentifier') else (o['name']
                                if o['name'] != 'Braunschweig Hbf ZOB' else 'Braunschweig Hbf (HBS)'))
        self.__destination_stop = (f'{d['name']} ({d['rilIdentifier']})' if d.get('rilIdentifier') else (d['name']
                                if d['name'] != 'Braunschweig Hbf ZOB' else 'Braunschweig Hbf (HBS)'))

        self.__origin_coordinates =(data['trip']['origin'].get('latitude', 'N/A'), data['trip']['origin'].get('longitude', 'N/A'))
        self.__destination_coordinates = (data['trip']['destination'].get('latitude', 'N/A'), data['trip']['destination'].get('longitude', 'N/A'))

        self.__via_stations_data = data['trip'].get('stopovers', [])
        self.__via_stations = list(f'{d['name']} ({d['rilIdentifier']})' if d.get('rilIdentifier') else (d['name'] if d['name'] != 'Braunschweig Hbf ZOB' else 'Braunschweig Hbf (HBS)')
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

        self.__journey_points = data['status']['train']['points']
        self.__journey_number = data['status']['train']['journeyNumber']

        self.__vehicle_number = 'N/A'
        self.__ticket_used = 'N/A'
        self.__wagon_class = 'N/A'
        self.__locomotive_class = 'N/A'
        self.__tags = data['status'].get('tags', None)
        if self.__tags is not None:
            for tag in self.__tags:
                if tag.get('key', None) == 'trwl:vehicle_number':
                    self.__vehicle_number = tag.get('value', 'N/A')
                if tag.get('key', None) == 'trwl:journey_number':
                    self.__journey_number = tag.get('value', 'N/A')
                if tag.get('key', None) == 'trwl:locomotive_class':
                    self.__locomotive_class = tag.get('value', 'N/A')
                if tag.get('key', None) == 'trwl:ticket':
                    self.__ticket_used = tag.get('value', 'N/A')
                if tag.get('key', None) == 'trwl:wagon_class':
                    self.__wagon_class = tag.get('value', 'N/A')

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
        if isinstance(other,User):
            return self.__id == other.__id
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
                                         'points': journey_points + distance_type.get(train_type, {}).get('points', 0),
                                         'all_delay_arr': distance_type.get(train_type, {}).get('all_delay_arr', []) + [journey_arrival_delay]}

            distance_operator_line[line_tuple] = {'distance': journey_distance + distance_operator_line.get(line_tuple, {}).get('distance', 0),
                                                  'time': journey_time + distance_operator_line.get(line_tuple, {}).get('time', 0),
                                                  'arrival_delay': journey_arrival_delay + distance_operator_line.get(line_tuple, {}).get('arrival_delay', 0),
                                                  'realtime_availability': journey_realtime_availability + distance_operator_line.get(line_tuple, {}).get('realtime_availability', 0),
                                                  'sum': 1 + distance_operator_line.get(line_tuple, {}).get('sum', 0),
                                                  'delayed_by_standard': journey_delayed_by_standard + distance_operator_line.get(line_tuple, {}).get('delayed_by_standard', 0),
                                                  'points': journey_points + distance_operator_line.get(line_tuple, {}).get('points', 0),
                                                  'all_delay_arr': distance_operator_line.get(train_type, {}).get('all_delay_arr', []) + [journey_arrival_delay]}

            distance_operator[operator_name] = {'distance': journey_distance + distance_operator.get(operator_name, {}).get('distance', 0),
                                                'time': journey_time + distance_operator.get(operator_name, {}).get('time', 0),
                                                'arrival_delay': journey_arrival_delay + distance_operator.get(operator_name, {}).get('arrival_delay', 0),
                                                'realtime_availability': journey_realtime_availability + distance_operator.get(operator_name, {}).get('realtime_availability', 0),
                                                'sum': 1 + distance_operator.get(operator_name, {}).get('sum', 0),
                                                'delayed_by_standard': journey_delayed_by_standard + distance_operator.get(operator_name, {}).get('delayed_by_standard', 0),
                                                'points': journey_points + distance_operator.get(operator_name, {}).get('points', 0),
                                                'all_delay_arr': distance_operator.get(train_type, {}).get('all_delay_arr', []) + [journey_arrival_delay]}




        self.__distance_type_sorted = dict(sorted(distance_type.items(), key=lambda x: x[1].get('distance', 0), reverse=True))
        self.__distance_operator_line_sorted = dict(sorted(distance_operator_line.items(), key=lambda x: x[1].get('distance', 0), reverse=True))
        self.__distance_operator_sorted = dict(sorted(distance_operator.items(), key=lambda x: x[1].get('distance', 0), reverse=True))


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



    def visited_station_execution(self):
        for j in self.__journeys:
            self.__visited_stations[j.origin_stop] = self.__visited_stations.get(j.origin_stop, 0) + 1
            self.__visited_stations[j.destination_stop] = self.__visited_stations.get(j.destination_stop, 0) + 1
        self.__number_of_visited_stations = len(self.__visited_stations)
        self.__visited_stations = dict(sorted(self.__visited_stations.items(), key=lambda x: x[1], reverse=True))


    def __visited_station_with_via_execution(self) -> None:
        self.__stations_with_via = self.visited_stations.copy()
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
                                f'{sum(j.journey_points for j in self.__journeys) / self.__number_of_journeys :.3f} Punkte'],
                      'Echtzeitquote': [f'{self.realtime_avaliability * 100:.2f}%',f'',f''],
                      f'Mehr als {round(self.__delay_standard/60)} Minuten verspätet': [f'{self.delay_rate_standard*100 :.2f}%',
                                        f'',
                                        f'']}

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
    @property
    def realtime_avaliability(self) -> float:
        if not self.__realtime_availability:
            self.delay_analysis_execute()
        return self.__realtime_availability
    @property
    def delay_rate_standard(self) -> float:
        if not self.__delay_rate_standard:
            self.delay_analysis_execute()
        return self.__delay_rate_standard



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
            print('Wenn du noch weitere User hinzufügst, werden hier spannende Auswertungen erstellt')
            print(f'\n{50 * "-"}\n')
            return

        visits_dataframe: pd.DataFrame = pd.DataFrame()
        visits_dataframe_filterable :pd.DataFrame = pd.DataFrame()

        dicts = [user.stations_with_via for user in user_list]
        station_counter = Counter(operator for d in dicts for operator in d.keys())
        common_stations = []
        for station, count in station_counter.items():
            if count >= max(2, len(user_list) - 1 if len(user_list) < 5 else len(user_list) - 2):
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
        for station, count in operator_counter.items():
            if count >= max(2, len(user_list) - 1 if len(user_list) < 5 else len(user_list) - 2):
                common_operators.append(station)

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

        trip_id_dict = {}
        for j in self.__journeys:
            if not trip_id_dict.get(j.trip_id):
                trip_id_dict[j.trip_id] = []
            trip_id_dict[j.trip_id].append((self.__user_id_dict.get(j.user_id), j))

        for trip_id, list_user_journey in trip_id_dict.items():
            if len(list_user_journey) > 1:
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


        user_list.sort(key=lambda x: x.name)
        names = f'{user_list[0].name}_' + '_'.join([user.name for user in user_list[1:]])
        try:
            with pd.ExcelWriter(f"{names}'s_shared_data.xlsx", engine='openpyxl') as writer:
                if not visits_dataframe.empty:
                    visits_dataframe.to_excel(writer, sheet_name='Haltestellen', index=False)
                    visits_dataframe_filterable.to_excel(writer, sheet_name='Haltestellen (filterbar)', index=False)
                type_dataframe.to_excel(writer, sheet_name='Kategorie', index=False)
                if not operator_dataframe.empty:
                    operator_dataframe.to_excel(writer, sheet_name='Betreiber', index=False)
                if not shared_trip_dataframe.empty:
                    styled_trip_dataframe = shared_trip_dataframe.style.apply(lambda row:
                                                                              [row['_color'] if col in [
                                                                                  'Ankunftszeit (ist)',
                                                                                  'Abfahrtszeit (ist)']
                                                                               else '' for col in row.index], axis=1)
                    styled_trip_dataframe.hide(['_color'], axis='columns')
                    styled_trip_dataframe.to_excel(writer, sheet_name='Gemeinsame Fahrten', index=False)


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

                print(f"Created {names}'s_shared_data.xlsx")

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
            if self.__user_name_dict.get(j.user_name) in user_list:
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
    traewelling.create_shared_excel()
    traewelling.create_gis_number_csv()
    traewelling.create_gis_single_csv()
    end_time = time()
    print(f"\033[91mThank You, for träwelling {end_time - start_time} seconds with Deutsche Bahn!\033[0m")
    print(f'\n{50 * "-"}\n')




