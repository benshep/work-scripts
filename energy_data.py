from datetime import datetime

from rayleigh_connect import EasyAPI
from rayleigh_api_key import api_key

rayleigh = EasyAPI(api_key)


def find_sensors(names: list[str]) -> dict[str, tuple[str, str]]:
    """Returns the gateway and sensor IDs for the given sensor names."""
    return_dict = {}
    for gateway in rayleigh.fetch_gateways():
        for sensor in rayleigh.fetch_sensors(gateway['id']):
            if sensor['name'] in names:
                return_dict[sensor['name']] = (gateway['id'], sensor['id'])
    return return_dict


def list_sensors():
    """Lists sensor names defined for each gateway."""
    for gateway in rayleigh.fetch_gateways():
        for sensor in rayleigh.fetch_sensors(gateway['id']):
            if sensor['name'] not in (None, 'Frequency'):  # lots of sensors called Frequency
                print(sensor['name'])


def archive_data():
    meters = [
        ('Q3025301880080025', 'e5'),
        ('Q3025301880080025', 'e6')
    ]
    gateways = rayleigh.fetch_gateways()
    # print(*gateways, sep='\n')
    for gateway in gateways:
        print(gateway['id'], gateway['name'])
        sensors = rayleigh.fetch_sensors(gateway['id'])
        for sensor in sensors:
            if (gateway['id'], sensor['id']) in meters:
                data = rayleigh.fetch_data(
                    gateway['id'],
                    sensor['id'] + '.kwh',  # want the energy data
                    datetime(2025, 10, 23))
                print(data)


# Data returned has arbitrary timestamps. Need to use something like numpy.interp1d
# to peg it to half-hourly intervals. Then drop into an XLSX file with structure:
# datetime          gateway_id1     gateway_id2
#                   sensor_id1      sensor_id2
#                   sensor_name1    sensor_name2
# 2025-10-24 00:00  123.456         456.789
# 2025-10-24 00:30  123.556         456.889
# ...


if __name__ == '__main__':
    # print(find_sensors(['PD2 - SW4E - PD934 CLARA Rack Room 2 - Magnet Rack 1',
    #                     'PD2 - SW4D - PD935 CLARA Rack Room 2 - Magnet Rack 2']))
    # list_sensors()
    archive_data()