# Copyright 2021 UXEON SP. Z O.O.
# Version: v1.0.0

import requests

from datetime import datetime


class APIException(Exception):
    pass


def dt_to_rfc3339(dt: datetime) -> str:
    return dt.isoformat().split('.')[0] + 'Z'


class EasyAPI:
    """Application Programming Interface for applications displaying and processing data.
    https://docs.rayleighconnect.net/api/easy/"""
    __base_url: str = 'https://api.rayleighconnect.net/easy/v1/'

    def __init__(self, token: str):
        self.__headers: dict[str, str] = {
            'authorization': 'Bearer ' + token,
            'accept': 'application/json'
        }

    def __request(self, url: str, params: dict[str, str] | None = None):
        request_url = self.__base_url + url
        r = requests.get(request_url, params=params, headers=self.__headers)
        if r.status_code == 404: return []
        if r.status_code != 200: raise APIException(r.status_code)
        return r.json()

    def fetch_gateways(self) -> list:
        """A gateway is a device that is collecting and transmitting the data to the system."""
        return self.__request('gateways')

    def fetch_sensors(self, gateway_id: str) -> list:
        """A sensor is single measuring point. One physical device can provide multiple measuring points."""
        return self.__request(f'gateways/{gateway_id}')

    def fetch_data(self, gateway_id: str, sensor_id: str, start: datetime, end: datetime | None = None) -> list:
        """Fetch raw archival data."""
        payload = {'from': dt_to_rfc3339(start)}
        if end is not None:
            payload['to'] = dt_to_rfc3339(end)
        return self.__request(f'gateways/{gateway_id}/{sensor_id}/data', payload)
