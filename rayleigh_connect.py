# Copyright 2021 UXEON SP. Z O.O.
# Version: v1.0.0
import aiohttp

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
        self.token = token
        self.__headers: dict[str, str] = {
            'authorization': 'Bearer ' + token,
            'accept': 'application/json'
        }

    async def __request(self, url: str, params: dict[str, str] | None = None):
        request_url = self.__base_url + url

        async with aiohttp.ClientSession() as session:
            async with session.get(request_url, params=params, headers=self.__headers) as response:
                if response.status == 404: return []
                if response.status != 200: raise APIException(response.status, response.reason)
                response_json = await response.json()

        return response_json

    async def fetch_gateways(self):
        """A gateway is a device that is collecting and transmitting the data to the system."""
        return await self.__request('gateways')

    async def fetch_sensors(self, gateway_id: str):
        """A sensor is single measuring point. One physical device can provide multiple measuring points."""
        return await self.__request(f'gateways/{gateway_id}')

    async def fetch_data(self, gateway_id: str, sensor_id: str, start: datetime, end: datetime | None = None):
        """Fetch raw archival data."""
        payload = {'from': dt_to_rfc3339(start)}
        if end is not None:
            payload['to'] = dt_to_rfc3339(end)
        return await self.__request(f'gateways/{gateway_id}/{sensor_id}/data', payload)

