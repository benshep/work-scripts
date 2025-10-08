import asyncio
import os
import re
from datetime import datetime, timedelta
from urllib.parse import urlencode

import aiohttp

from folders import docs_folder

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


async def pv_avg(client_session: aiohttp.ClientSession, pv: str, start: datetime, period: timedelta = timedelta(hours=0.5)) -> float | None:
    """Return the average (mean) value of a PV over a period of time."""

    query = urlencode({
        'pv': pv,
        'from': start.strftime('%Y-%m-%dT%H:%M:%SZ'),  # '2025-01-01T00:00:00.000Z',
        'to': (start + period).strftime('%Y-%m-%dT%H:%M:%SZ'),  # '2025-01-01T00:30:00.000Z',
        'fetchLatestMetadata': True
    })
    url = f'http://claraserv2.dl.ac.uk:17668/retrieval/data/getData.json?{query}'
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
    """Return the average (mean) value of each of a set of PVs over a period of time."""
    average_values = await asyncio.gather(*[pv_avg(client_session, pv, start, period) for pv in pvs])
    return [start] + average_values


async def pvs_avg_time_series(client_session: aiohttp.ClientSession,
                              pvs: list[str], start: datetime, period: timedelta = timedelta(hours=0.5),
                              periods: int = 1) -> list[list[datetime | float | None]]:
    return await asyncio.gather(*[pvs_avg(client_session, pvs, start + i * period, period) for i in range(periods)])


async def get_data():
    session = aiohttp.ClientSession(connector=aiohttp.TCPConnector(limit=10))
    periods = 48 * 30
    session.__setattr__('pv_total', periods * 5)
    session.__setattr__('pv_counter', 0)

    start = datetime(2024, 4, 1, 0, 0, 0)
    series = await pvs_avg_time_series(session,
                                       ['CLA-GUNS-LRF-CTRL-01:PID:ad1:ch6:Power:Wnd:Hi'] + \
                                       [f'CLA-L0{i + 1}-LRF-CTRL-01:PID:ad1:ch3:Power:Wnd:Avg' for i in range(4)],
                                       start,
                                       periods=periods)
    await session.close()
    print('')
    with open(f'pv_data_{start.strftime("%Y-%m-%d")}.csv', 'w') as f:
        for p in series:
            f.write(p[0].strftime('%d/%m/%Y %H:%M') + ','
                    + ','.join([f'{x:.4g}' if x else '' for x in p[1:]])
                    + '\n')
            # print(*p, sep='\t')


if __name__ == '__main__':
    # read_logs()
    asyncio.run(get_data())
