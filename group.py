import os
import tempfile
from datetime import date, timedelta, datetime
from itertools import accumulate
from pathlib import Path
from platform import node
from urllib.parse import urlencode

from dateutil.relativedelta import relativedelta
from pushbullet import Pushbullet

import oracle
import outlook
import staff
from mars_group import members
from otl import fy_start
from pushbullet_api_key import api_key  # local file, keep secret!
from work_folders import downloads_folder, docs_folder


def run_otl_calculator(force_this_week: bool = False) -> tuple[str, str] | None:
    """Iterate through staff, listing the hours to upload for new OTL cards required."""
    # staff.verbose = True
    cards_to_book = 0
    links_filename = os.path.join(downloads_folder, 'otl_upload_links.html')
    with open(links_filename, 'w') as links_file:
        folder = os.path.split(__file__)[0]
        css_path = Path(os.path.join(folder, 'redwood.css')).as_uri()
        links_file.write(f'''<!doctype html><html lang="en">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width,initial-scale=1" />
    <title>OTL bookings</title>
    <link rel="stylesheet" href="{css_path}" id="css" type="text/css">
</head>
<body>
    <main class="app">
''')
        for member in members:
            # if member.known_as in ('Nasiq',):
            if True:
                print('\n', member.name)
                member.update_off_days(force_reload=False)
                start = date.today()
                start -= timedelta(days=start.weekday())  # Monday of current week
                start -= timedelta(days=7*4)  # check last few weeks
                # usually wait until Thu to do current week - can force override
                while start + timedelta(days=0 if force_this_week else 3) <= date.today():
                # for _ in range(7):
                    member.prev_bookings = {}  # reset in case this has already run
                    hours_booked = member.hours_for_week(start)
                    hours_needed = sum(member.hours_needed(start + timedelta(days=day)) for day in range(5))
                    print(f'{start.strftime("%d/%m/%Y")}: {hours_booked=:.2f}, {hours_needed=:.2f}')
                    end = start + timedelta(days=6)
                    if hours_needed - hours_booked > 0.01:
                        cards_to_book += 1
                        params = {'calledFromAddTimeCard': 'true', 'pAsofdate': start.strftime('%Y-%m-%d')}
                        if member.known_as != 'Ben':
                            params |= {'pPersonId': member.person_id, 'userContext': 'LINE_MANAGER'}
                        url = oracle.apps[('home',)] + 'time/timecards/landing-page?' + urlencode(params)
                        links_file.write(f'''
        <section class="card">
            <header class="card-header">
                <h1>Time Card</h1>
            </header>
            <div class="meta">
                <div class="item">
                    <div class="label">Person</div>
                    <div class="value">{member.name}</div>
                </div>
                <div class="item">
                    <div class="label">Person Number</div>
                    <div class="value">{member.person_number}</div>
                </div>
                <div class="item">
                    <div class="label">Time Card Period</div>
                    <div class="value">{start.strftime("%d/%m/%Y")} - {end.strftime("%d/%m/%Y")}</div>
                </div>
                <div class="item">
                    <div class="label">Scheduled Hours</div>
                    <div class="value">{hours_needed:.1f}</div>
                </div>
                <div class="item">
                    <div class="label">Upload Link</div>
                    <div class="value"><a href="{url}">Submit Timecard</a></div>
                </div>
            </div>
''')
                        links_file.write(member.otl_upload_page(start))
                        links_file.write('        </section>')
                    start += timedelta(days=7)
        links_file.write('    </main>\n</body>\n</html>\n')
    if cards_to_book:
        return f'{cards_to_book=}', links_filename
    return None


def leave_cross_check():
    """Iterate through staff, and check Oracle vs Outlook leave bookings."""
    toast = ''
    _, output_filename = tempfile.mkstemp(prefix='leave_cross_check', suffix='.txt')
    with open(output_filename, 'w', encoding='utf-8-sig') as output_file:
        for member in members:
            # if member.known_as in ('Ben',):
            if True:
                not_in_oracle, not_in_outlook, output = member.leave_cross_check()
                if not_in_oracle:
                    toast += f'{member.known_as}: {not_in_oracle=}, {not_in_outlook=}\n'
                    output_file.write(output)
    return (toast, output_filename) if toast else toast


def show_leave_dates():
    """Iterate through staff, and show the leave dates for each one, not including bank holidays."""
    for member in members:
        print('\n' + member.known_as)
        member.update_off_days(force_reload=True)
        print(*sorted(day for day in member.off_days - staff.site_holidays.keys()), sep='\n')


def get_checkins() -> list[str]:
    """Read messages back regarding check-ins. For instance: Joe chat in office"""
    pushbullet = Pushbullet(api_key)
    pushes = pushbullet.get_pushes(modified_after=(datetime.now() - timedelta(days=7)).timestamp())
    return [datetime.fromtimestamp(push['modified']).strftime('%Y-%m-%d %H:%M: ') + push.get('body', '')
            for push in pushes if 'title' not in push]  # most have titles: looking for one without (sent from phone)]


def check_in() -> str | bool:
    """Pick a member of staff to check in with."""
    if node() != 'DDAST0025':  # only run on desktop
        return False
    now = datetime.now()
    # when am I free? start with 0900-1700
    print('Checking my free time')
    my_free_times = outlook.find_free_times()
    if not my_free_times:
        print('No free times available for me')
        return False

    print('Getting previous checkins')
    checkins = get_checkins()
    os.chdir(os.path.join(docs_folder, 'Group Leader', 'check_in'))
    files = [filename for filename in os.listdir() if filename.endswith('.txt')]
    for filename in files:
        name = filename[:-4]
        for checkin in checkins:
            if name not in checkin:
                continue
            line = checkin + '\n'
            if line not in open(filename).read():
                print(checkin)
                open(filename, 'a').write(line)
                timestamp = datetime.strptime(checkin[:16], '%Y-%m-%d %H:%M').timestamp()
                os.utime(filename, (timestamp, timestamp))

    files = sorted(files, key=os.path.getmtime)
    # modified X.Y days ago: round down to nearest X
    ages = list(accumulate((now - datetime.fromtimestamp(os.path.getmtime(filename))).days for filename in files))
    # Select basically at random, but pick the same one on a given day
    # Weight older entries higher, and zero weighting to any that are less than one day old
    selected = now.toordinal() % ages[-1]
    index = next(i for i, age in enumerate(ages) if age > selected)
    # When going through list, go back to oldest, then onwards to newest
    files = files[index::-1] + files[index + 1:]
    for filename in files:
        name = filename[:-4]
        email = open(filename).read().splitlines()[0]
        print(f'Checking free time for {name}')
        their_free_times = outlook.find_free_times(email)
        free_overlap = my_free_times & their_free_times
        if not free_overlap:
            print(f'No free time overlap with {name}')
            continue
        break
    else:
        print('No free time overlap with anyone today')
        return False
    return f'Check in with {name} at {min(free_overlap).strftime("%H:%M")}'


def list_ftes():
    """List staff and committed time against projects."""
    for member in sorted(members, key=lambda person: person.formal_name()):
        for entry in member.booking_plan.entries:
            print(
                member.formal_name(),
                entry.code,
                entry.start_date.strftime('%d/%m/%Y'),
                entry.end_date.strftime('%d/%m/%Y'),
                entry.annual_fte,
                '',
                *[entry.annual_fte if entry.start_date <= fy_start + relativedelta(months=+i) <= entry.end_date else 0
                  for i in range(12)],
            sep='\t')



if __name__ == '__main__':
    print(run_otl_calculator(force_this_week=True))
    # print(leave_cross_check())
    # print(check_in())
    # list_ftes()