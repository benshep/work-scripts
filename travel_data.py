import re
from datetime import datetime, timedelta

import numpy as np
import requests

from math import radians, sin, cos, sqrt, atan2, copysign

import airportsdata
import pandas
import serpapi

from work_folders import docs_folder
from work_tools import read_excel
from google_routing_credentials import api_key
from serp_credentials import serp_api_key

iata_pattern = re.compile(r'\b[A-Z]{3}\b(?:/[A-Z]{3}\b)+')
airports = airportsdata.load('IATA')
excel_file = docs_folder / 'Sustainability' / 'Travel' / '2026-27 ASTeC spending.xlsx'
uk_stations = pandas.read_csv(
    'https://raw.githubusercontent.com/davwheat/uk-railway-stations/refs/heads/main/stations.csv')
serp = serpapi.Client(api_key=serp_api_key)
known_stations = {
    'PARIS ST LAZARE': (48.876944, 2.324444),
    'PAZ': (48.876944, 2.324444),
    'TROUVILLE/DEAUVILLE': (49.36, 0.084167),
    'DUEVILLE': (49.36, 0.084167),
    'TDV': (49.36, 0.084167),
    'QQS': (51.53, -0.125278),  # St Pancras
    'ST PANCRAS': (51.53, -0.125278),
    'XPG': (48.881111, 2.355278),  # Gare du Nord
}
# from https://www.gov.uk/government/collections/government-conversion-factors-for-company-reporting
hotel_conversions = {
    # hotel (kgCO₂e/night)
    'London': 11.5,
    'United Kingdom': 10.4,
    'Australia': 35,
    'Belgium': 12.2,
    'Brazil': 8.7,
    'Canada': 7.4,
    'Chile': 27.6,
    'China': 53.5,
    'Colombia': 14.7,
    'Costa Rica': 4.7,
    'Egypt': 44.2,
    'France': 6.7,
    'Germany': 13.2,
    'Hong Kong, China': 51.5,
    'India': 58.9,
    'Indonesia': 62.7,
    'Italy': 14.3,
    'Japan': 39,
    'Jordan': 68.9,
    'Korea': 55.8,
    'Malaysia': 61.5,
    'Maldives': 152.2,
    'Mexico': 19.3,
    'Netherlands': 14.8,
    'Oman': 90.3,
    'Philippines': 54.3,
    'Portugal': 19,
    'Qatar': 86.2,
    'Russian Federation': 24.2,
    'Saudi Arabia': 106.4,
    'Singapore': 24.5,
    'South Africa': 51.4,
    'Spain': 7,
    'Switzerland': 6.6,
    'Thailand': 43.4,
    'Turkey': 32.1,
    'United Arab Emirates': 63.8,
    'United States': 16.1,
    'Vietnam': 38.5,
    None: 32.1  # median value
}
# 2026 factors, assuming economy class throughout: values in kgCO₂e/passenger.km
conversions = {
    # air
    'Domestic': 0.22928,
    'Short': 0.12576,
    'Long': 0.11704,
    'International': 0.10916,
    # rail
    'National Rail': 0.03092,  # intended for UK trains
    'International Rail': 0.01135  # i.e. Eurostar
}
routing_url = 'https://routes.googleapis.com/directions/v2:computeRoutes'
routing_headers = {
    "Content-Type": "application/json",
    "X-Goog-Api-Key": api_key,
    "X-Goog-FieldMask": "routes.distanceMeters",
}


def extract_iata_codes(text: str) -> list[str]:
    """Return the 3-letter IATA codes contained within a comment field."""
    matches = iata_pattern.findall(text)  # e.g. ['MAN/OSL/BER']
    codes = []
    for m in matches:
        codes.extend(m.split('/'))  # split each group into individual codes
    return codes or line_to_iata_pairs(text)


def extract_station_codes(text: str) -> list[str]:
    """Return the SNCF station IDs from the station names contained within a comment field."""
    matches = iata_pattern.findall(text)  # e.g. ['MAN/OSL/BER']
    codes = []
    for m in matches:
        codes.extend(m.split('/'))  # split each group into individual codes
    return codes or line_to_iata_pairs(text)


def extract_country_name(row) -> tuple[str | None, str | None]:
    """Return a country name contained within a comment field."""
    # format e.g. 16/05/2026 JONES/MATTHEW Hotel Hotel  France MERCURE TROUVILLE-SUR-MER;14360; - EXPEDIA BOOKING
    text = row['Comment']
    for country in hotel_conversions.keys():
        if country and country in text:
            description = text.split(';')[0]
            pos = description.find(country)
            hotel_name = description[pos + len(country) + 1:]
            return country, hotel_name
    return None, None


def haversine(lat1, lon1, lat2, lon2):
    r = 6371.0088  # mean Earth radius in km
    phi1, phi2 = radians(lat1), radians(lat2)
    dphi = radians(lat2 - lat1)
    dlambda = radians(lon2 - lon1)
    a = sin(dphi / 2) ** 2 + cos(phi1) * cos(phi2) * sin(dlambda / 2) ** 2
    return 2 * r * atan2(sqrt(a), sqrt(1 - a))


def total_flight_distance(row) -> tuple[float, str, float]:
    """Return the total distance in km between the given airports, the haul type (domestic, short, long, international,
    represented as D/S/L/I for each leg), and the emissions in kgCO₂e."""
    codes = row['airport_codes']
    total_distance = 0.0
    emissions = 0.0
    hauls = ''
    for a, b in zip(codes, codes[1:]):
        ap_a, ap_b = airports.get(a), airports.get(b)
        if not ap_a or not ap_b:
            missing = a if not ap_a else b
            raise ValueError(f"Unknown airport code: {missing}")
        countries = {ap_a['country'], ap_b['country']}
        time_zones = {ap_a['tz'].split('/')[0], ap_b['tz'].split('/')[0]}  # e.g. Europe
        # approximation: in reality, some North African countries classed as short-haul too!
        haul = 'Domestic' if countries == {'GB'} else \
            'Short' if 'GB' in countries and 'Europe' in time_zones else \
                'Long' if 'GB' in countries else \
                    'International'
        hauls += haul[0]
        dist_km = haversine(ap_a['lat'], ap_a['lon'], ap_b['lat'], ap_b['lon'])
        total_distance += dist_km
        emissions += conversions[haul] * dist_km
    emissions = copysign(emissions, row['Cost'])  # in case of refunds!
    return total_distance, hauls, emissions


def total_train_distance(row):
    """Return the total distance in km between the given stations."""
    stations = row['stations']
    # print(stations)
    coords = []
    for station in stations:
        if pt := known_stations.get(station, None):
            coords.append((*pt, station))
            continue
        matched = uk_stations[uk_stations['stationName'].str.fullmatch(station, case=False)].squeeze()
        if len(matched):
            coords.append((matched['lat'], matched['long'], matched['stationName']))
        else:
            print('No match for', station)
    total_distance = 0.0
    total_emissions = 0.0
    if len(coords) > 1:
        for a, b in zip(coords, coords[1:]):
            payload = {
                "origin": {'location': {'latLng': {'latitude': a[0], 'longitude': a[1]}}},
                "destination": {'location': {'latLng': {'latitude': b[0], 'longitude': b[1]}}},
                'travelMode': 'TRANSIT',
                'transitPreferences': {'allowedTravelModes': ['TRAIN', 'LIGHT_RAIL', 'RAIL']}
            }
            response = requests.post(routing_url, json=payload, headers=routing_headers).json()
            print(a[2], 'to', b[2], end=' ')
            try:
                distance = response['routes'][0]['distanceMeters'] / 1000
                print(round(distance), 'via Google')
            except (IndexError, KeyError):  # fall back to crow-flight distance + 10%
                distance = haversine(*a[:2], *b[:2]) * 1.1
                print(round(distance), 'as the crow flies')
            total_distance += distance
            rail_type = 'International Rail' if 'International' in row['Comment'] else 'National Rail'
            total_emissions += conversions[rail_type] * distance
    return total_distance, total_emissions


def guess_iata_code(segment: str, airport: bool = True) -> str | None:
    """Guess an IATA code for a free-text city/airport name segment."""
    text = segment.strip().upper()
    if not text or any(ch.isdigit() for ch in text) or text == 'CONFERENCE ATTENDANCE':
        return None

    if not airport:
        return text

    # Known override: "Manchester" is always the UK airport, not the
    # US one that also matches by name.
    if text == "MANCHESTER":
        return "MAN"

    text_words = set(re.findall(r'[A-Z]+', text))
    best_score = 0.0

    for code, a in airports.items():
        candidate = f"{a['city']} {a['name']}".upper()
        candidate_words = set(re.findall(r'[A-Z]+', candidate))
        if not candidate_words:
            continue
        overlap = text_words & candidate_words
        score = len(overlap) / len(text_words)  # fraction of OUR words matched
        if score > best_score:
            best_score, best_code = score, code

    return best_code if best_score >= 0.5 else None


def line_to_iata_pairs(line: str) -> list[str]:
    """Search a comment line for full names of either airports or train stations."""
    m = re.search(r'(Rail|Flight)([- ])(International Rail|Rail|Air)\2', line)
    if not m:
        print(line)
        return []
    sep = m.group(2)  # dash or space
    rest = line[m.end():]
    segments = [s.strip() for s in rest.split(sep)]
    if sep == ' ':
        uk = rest.find(
            'United Kingdom ')  # e.g. LONDON EUSTON United Kingdom WARRINGTON BANK QUAY/LONDON EUSTON/WARRINGTON BANK QUAY
        if uk > 0:
            segments = rest[uk + 15:].split('/')
            dest = rest[:uk - 1]
            if dest not in segments:
                segments.append(dest)
    # print(segments)
    codes = [c for seg in segments if (c := guess_iata_code(seg, m.group(1) == 'Flight'))]
    return codes


# it's trickier for rail bookings!
# All bookings seem to contain either "Rail-Rail" or "Rail Rail"
# After that, we get the stations in full, like
# 16/05/2026-OWEN/HYWEL -Rail-Rail-PARIS ST LAZARE-TROUVILLE/DEAUVILLE-CONFERENCE ATTENDANCE
# 23/05/2026-COWIE/LOUISE-Rail-Rail-LONDON EUSTON-MACCLESFIELD-CONFERENCE ATTENDANCE
# 15/05/2026 ANGAL KALININ/DEEPA Rail Rail LONDON EUSTON United Kingdom WARRINGTON BANK QUAY/LONDON EUSTON/WARRINGTON BANK QUAY
# or short codes (IATA), not in the EU stations list!
# 16/05/2026 BAINBRIDGE/ALEXANDER  International Rail Rail PARIS GARE DU NORD France QQS/XPG/QQS
# https://raw.githubusercontent.com/jpatokal/openflights/master/data/airports-extended.dat


def get_flights_data():
    """Read a spreadsheet of spending (gleaned from OBI) and extract data on flight bookings,
    looking up the total distance travelled and calculating emissions for each booking.
    Add just the rows containing useful flight data to a new sheet in the workbook."""
    travel = get_clarity_rows()
    flights = travel[travel['Comment'].str.contains('Flight')]
    flights['airport_codes'] = flights['Comment'].apply(extract_iata_codes)
    flights[['distance_km', 'haul', 'emissions_kgco2e']] = flights.apply(total_flight_distance, axis=1,
                                                                         result_type='expand')
    print(flights)
    save_data(flights, 'Flights')


def save_data(data, sheet_name: str):
    """Write data to spreadsheet file."""
    with pandas.ExcelWriter(excel_file, mode='a',  # append
                            if_sheet_exists='replace',  # replace existing sheets
                            engine='openpyxl',  # can't write multiple sheets using default engine
                            ) as writer:
        data.to_excel(writer, sheet_name=sheet_name)


def get_rail_data():
    """Read a spreadsheet of spending (gleaned from OBI) and extract data on rail bookings,
    looking up the total distance travelled and calculating emissions for each booking.
    Add just the rows containing useful rail data to a new sheet in the workbook."""
    travel = get_clarity_rows()
    rail = travel[travel['Comment'].str.contains('Rail[- ]Rail')]
    rail['stations'] = rail['Comment'].apply(extract_iata_codes)
    # print(*rail['stations'], sep='\n')
    rail[['distance_km', 'emissions_kgco2e']] = rail.apply(total_train_distance, axis=1, result_type='expand')
    print(rail)
    save_data(rail, 'Rail')


hotel_rates = {}
def estimate_nights(row) -> int:
    """Use the hotel name to estimate the price and therefore number of nights stayed."""
    global hotel_rates
    rate = 150  # VERY APPROXIMATE GUESS
    if name := row['hotel_name']:
        if name in hotel_rates:
            rate = hotel_rates[name]
        else:
            check_in = datetime.now() + timedelta(days=30)  # simulate booking 1 month in advance
            check_out = check_in + timedelta(days=1)  # 1 night
            results = serp.search({'engine': 'google_hotels',
                                   'q': name, 'adults': 1,
                                   'check_in_date': check_in.strftime('%Y-%m-%d'),
                                   'check_out_date': check_out.strftime('%Y-%m-%d'),
                                   'currency': 'GBP', 'gl': 'gb', 'hl': 'en'})
            try:
                rate = results['rate_per_night']['extracted_lowest']
                print(name, rate)
            except KeyError:
                pass
    nights = np.ceil(row['Cost'] / rate)
    max_nights = 7  # seems an OK limit for work trips
    if nights > max_nights:  # i.e. we paid more than the estimated nightly rate
        print(name, nights, 'nights estimated: reducing to', max_nights)
        nights = max_nights
    return nights


def get_hotel_data():
    """Read a spreadsheet of spending (gleaned from OBI) and extract data on hotel bookings,
    calculating emissions for each booking.
    Add just the rows containing useful hotel data to a new sheet in the workbook."""
    travel = get_clarity_rows()
    hotel = travel[travel['Comment'].str.contains('Hotel[- ]Hotel')]
    hotel[['country', 'hotel_name']] = hotel.apply(extract_country_name, axis=1, result_type='expand')
    hotel['nights'] = hotel.apply(estimate_nights, axis=1)
    hotel['emissions_kgco2e'] = hotel['country'].map(hotel_conversions) * hotel['nights']
    print(hotel)
    save_data(hotel, 'Hotel')


def get_clarity_rows():
    """Return the spreadsheet rows where the supplier is Clarity."""
    data = read_excel(excel_file)
    clarity = data[data['Employee/Supplier Name'] == 'Clarity Travel Ltd t/a Clarity']
    return clarity


if __name__ == '__main__':
    get_hotel_data()
