import os
import tempfile
from datetime import date, timedelta, datetime
from itertools import accumulate
from math import ceil
from urllib.parse import urlencode

from pushbullet import Pushbullet

import oracle
import outlook
import staff
from mars_group import members
from work_folders import downloads_folder, docs_folder
from work_tools import floor_date
from pushbullet_api_key import api_key  # local file, keep secret!


def run_otl_calculator(force_this_week: bool = False) -> tuple[str, str] | None:
    """Iterate through staff, listing the hours to upload for new OTL cards required."""
    # staff.verbose = True
    cards_to_book = 0
    links_filename = os.path.join(downloads_folder, 'otl_upload_links.html')
    with open(links_filename, 'w') as links_file:
        for member in members:
            # if member.known_as in ('Nasiq',):
            if True:
                links_file.write(f'<h2>{member.known_as}</h2>\n')
                print('\n', member.name)
                member.update_off_days(force_reload=False)
                start = date.today()
                start -= timedelta(days=start.weekday())  # Monday of current week
                start -= timedelta(days=7*4)  # check last few weeks
                # usually wait until Thu to do current week - can force override
                while start + timedelta(days=0 if force_this_week else 3) <= date.today():
                # for _ in range(7):
                    hours_booked = member.hours_for_week(start)
                    hours_needed = sum(member.hours_needed(start + timedelta(days=day)) for day in range(5))
                    print(f'{start.strftime("%d/%m/%Y")}: {hours_booked=:.2f}, {hours_needed=:.2f}')
                    if hours_needed - hours_booked > 0.01:
                        cards_to_book += 1
                        params = {'calledFromAddTimeCard': 'true', 'pAsofdate': start.strftime('%Y-%m-%d')}
                        if member.known_as != 'Ben':
                            params |= {'pPersonId': member.person_id, 'userContext': 'LINE_MANAGER'}
                        url = oracle.apps[('home',)] + 'time/timecards/landing-page?' + urlencode(params)
                        links_file.write(f'<p><a href="{url}">{start.strftime("%d/%m/%Y")}</a></p>\n')
                        lines = list(member.bulk_upload_lines(start, manual_mode=True))
                        print(*lines, sep='\n')
                        for line in lines:
                            links_file.write(f'<p>{line}</p>\n')
                    start += timedelta(days=7)
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
    return [datetime.fromtimestamp(pushes[0]['modified']).strftime('%Y-%m-%d %H:%M: ') + push.get('body', '')
            for push in pushes if 'title' not in push]  # most have titles: looking for one without (sent from phone)]


def check_in():
    """Pick a member of staff to check in with."""
    now = datetime.now()
    # when am I free? start with 0900-1700
    my_free_times = outlook.find_free_times()
    if not my_free_times:
        print('No free times available for me')
        return (now + timedelta(days=1)).replace(hour=9)  # do it 9-10am tomorrow

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

    files = sorted(files, key=os.path.getmtime)
    # modified X.Y days ago: round down to nearest X
    ages = list(accumulate((now - datetime.fromtimestamp(os.path.getmtime(filename))).days for filename in files))
    selected = now.toordinal() % ages[-1]
    index = next(i for i, age in enumerate(ages) if age > selected)
    # go back to oldest, then onwards to newest
    files = files[index::-1] + list(reversed(files[-1:index:-1]))
    for filename in files:
        name = filename[:-4]
        email = open(filename).read().splitlines()[0]
        their_free_times = outlook.find_free_times(email)
        free_overlap = my_free_times & their_free_times
        if not free_overlap:
            print(f'No free time overlap with {name}')
            index += 1
            continue
        break
    else:
        print('No free time overlap with anyone today')
        return (now + timedelta(days=1)).replace(hour=9)  # do it 9-10am tomorrow
    return f'Check in with {name} at {min(free_overlap).strftime("%H:%M")}'



if __name__ == '__main__':
    # run_otl_calculator(force_this_week=True)
    # print(leave_cross_check())
    print(check_in())