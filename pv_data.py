import asyncio
import os
import re
from datetime import datetime, timedelta
from urllib.parse import urlencode

import aiohttp

from work_folders import docs_folder

archiver_url = 'http://claraserv2.dl.ac.uk:17668/retrieval/data/getData.json?'


def read_logs():
    os.chdir(os.path.join(docs_folder, 'CLARA', 'logbooks'))
    for file in os.listdir('.'):
        file_date = datetime.strptime(file[:6], '%y%m%d')
        # print(file_date)
        summary = open(file).read().split('<hr />')[0]
        for line in summary.splitlines():
            power_values = re.findall(r'([\d\.]+)\s?MW', line)
            cavities = re.findall(r'L0[1-4]|Linacs?[\s\-]?0?[1-4]|Gun|HRRG|TDC', line, re.IGNORECASE)
            if power_values:  # and not cavities:
                print(file_date.strftime('%d/%m/%Y'),
                      re.sub(r'<.*?>', '', line.replace('&nbsp;', ' ').strip()),
                      ','.join(cavities),
                      *power_values,
                      sep='\t')


async def pv_avg_server(client_session: aiohttp.ClientSession,
                        pv: str,
                        start: datetime,
                        end: datetime,
                        bin_period: timedelta = timedelta(hours=0.5)) -> dict[datetime, float]:
    """Return the average (mean) value of a PV over a bin_period of time.
    Let the server handle the averaging."""
    url = archiver_url + urlencode({
        'pv': f'mean_{int(bin_period.total_seconds())}({pv})',
        'from': (start - bin_period).strftime('%Y-%m-%dT%H:%M:%SZ'),  # '2025-01-01T00:00:00.000Z',
        'to': end.strftime('%Y-%m-%dT%H:%M:%SZ'),  # '2025-01-01T00:30:00.000Z',
        'fetchLatestMetadata': False
    })
    print(url)
    async with client_session.get(url, timeout=None) as response:
        json = await response.json()
    data = json[0]['data']
    # Return a dict with datetime: average_value, to make it easier to line up multiple datasets
    average = {}
    for data_point in data:
        if (val := data_point['val']) < 1e100:  # don't record silly values!
            average[datetime.fromtimestamp(data_point['secs']) - bin_period / 2] = val
    print(pv, len(data))
    return average


async def pv_avg(client_session: aiohttp.ClientSession, pv: str, start: datetime,
                 period: timedelta = timedelta(hours=0.5)) -> float | None:
    """Return the average (mean) value of a PV over a bin_period of time."""
    url = archiver_url + urlencode({
        'pv': pv,
        'from': start.strftime('%Y-%m-%dT%H:%M:%SZ'),  # '2025-01-01T00:00:00.000Z',
        'to': (start + period).strftime('%Y-%m-%dT%H:%M:%SZ'),  # '2025-01-01T00:30:00.000Z',
        'fetchLatestMetadata': False
    })
    # try:
    async with client_session.get(url, timeout=None) as response:
        json = await response.json()
        await asyncio.sleep(0.1)
    # except asyncio.TimeoutError:
    #     print('Timeout error', pv, start)
    #     average = None
    # else:
    data = json[0]['data']
    average = sum(data_point['val'] for data_point in data) / len(data) if data else None

    # Display a counter
    i = client_session.__getattribute__('pv_counter')
    print(i, client_session.__getattribute__("pv_total"), sep=' / ', end='\r')
    client_session.__setattr__('pv_counter', i + 1)

    return average


async def pvs_avg(client_session: aiohttp.ClientSession,
                  pvs: list[str], start: datetime,
                  period: timedelta = timedelta(hours=0.5)) -> list[datetime | float | None]:
    """Return the average (mean) value of each of a set of PVs over a bin_period of time."""
    average_values = await asyncio.gather(*[pv_avg(client_session, pv, start, period) for pv in pvs])
    return [start] + average_values


async def pvs_avg_time_series(client_session: aiohttp.ClientSession,
                              pvs: list[str], start: datetime, period: timedelta = timedelta(hours=0.5),
                              periods: int = 1) -> list[list[datetime | float | None]]:
    return await asyncio.gather(*[pvs_avg(client_session, pvs, start + i * period, period) for i in range(periods)])


async def cav_power_to_csv():
    session = aiohttp.ClientSession(connector=aiohttp.TCPConnector(limit=10))
    pv_list = (['CLA-GUNS-LRF-CTRL-01:PID:ad1:ch6:Power:Wnd:Hi']
               + [f'CLA-L0{i + 1}-LRF-CTRL-01:PID:ad1:ch3:Power:Wnd:Avg' for i in range(4)])

    # Do 2 weeks at a time: the server response is reasonably quick for this amount of data
    start = datetime(2024, 4, 1, 0, 0, 0)
    period = timedelta(minutes=30)
    with open(f'pv_data_{start.strftime("%Y-%m-%d")}.csv', 'w') as f:
        # f.write(','.join(['Date/time'] + pv_list) + '\n')
        for _ in range(27):
            print(start)
            end = start + timedelta(days=14)
            series = await asyncio.gather(*[pv_avg_server(session, pv, start, end, bin_period=period)
                                            for pv in pv_list])
            print('')
            while start < end:
                f.write(
                    ','.join([start.strftime('%d/%m/%Y %H:%M')]
                             + [f'{x[start]:.4g}' if start in x else '' for x in series])
                    + '\n'
                )
                start += period
                # print(*p, sep='\t')
    await session.close()


if __name__ == '__main__':
    # read_logs()
    asyncio.run(cav_power_to_csv())
