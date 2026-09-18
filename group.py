import os
import tempfile
from datetime import date, timedelta, datetime
from itertools import accumulate
from pathlib import Path
from time import sleep
from urllib.parse import urlencode
from selenium.webdriver.common.by import By

from dateutil.relativedelta import relativedelta
from pushbullet import Pushbullet

import oracle
import outlook
import staff
from mars_group import members
from otl import fy_start
from pushbullet_api_key import api_key  # local file, keep secret!
from work_folders import downloads_folder, docs_folder


def run_otl_calculator(weeks_ahead: int = 0, **kwargs) -> tuple[str, str] | None:
    """Iterate through staff, listing the hours to upload for new OTL cards required."""
    # staff.verbose = True
    cards_to_book = 0
    links_filename = downloads_folder / 'otl_upload_links.html'
    with open(links_filename, 'w', encoding='utf-8') as links_file:
        folder = Path(__file__).parent
        css_filename = folder / 'redwood.css'
        css = css_filename.read_text()
        # language=HTML
        links_file.write(f'''<!doctype html><html lang="en">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width,initial-scale=1" />
    <link rel="icon" href="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABHNCSVQICAgIfAhkiAAAAAlwSFlzAAAAdgAAAHYBTnsmCAAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZ5vuPBoAAALNSURBVDiNlZFfSFt3FMc/v3uvofGaaHTxb9qmtrR1raz4onvZy2CMjQ0fdPShqIzChoLt7GiL2A4GQwqFgTjcwyb4IAx0uD/Owf48zK2l1QepTgubJmptOzXJbkzMvbnJzW8Pw04sVPaFwzkczjkcPl/YpWgifeqPtUzyz6j07u4bhvQtPsomwn/JevZIyJebXyTHCJBv177gWrlyRT/c+d6WK7bu7AxlnivX1vo+9pT0D6S8tybTiZAUSNq8D0a/VnDES0AV4HOmQ7o+P49c2vQCvp3ILUY87tlZ1MnZfAS+PI8oklmaAIRZ0zTgWOLdva89S9ltIMNE0d8jr2tafTDkMpPrOV0vRlXz9t3OSUdZMlJp2x3jNmjx/htqykyXGUkLf0kBjx7HORTwEV6OciRY8iQvP4gRqPSxHkmoHrfuQaqCyj4UgMk7Ya73TrC0EuPSB19hxE0uXhsDoLNnDCNu0v3Rd8zMrnGj72emZ1b/cyGSdK5KIXv/DwPLBNtm+Gildk4BGP9xgZaOYe4vbtLY8hlG3OS1s58C8MpbA4zft+j+JUvbuM35b0xu3rHY8VhEks5VO5vtte0sB9wutrctvB43RtykqNDN4IzFSkKluUaj3A3RJAz9brNmZDa+by0q+5fB3TCXP/yW0GqM9sujGFsm57u+IGxI5qIa7zdonPYLfv1tgePlggtnXBimLK37JP6q2nGxp7WsNL+u9vkKDgeKOVldyqFAMceCfpadAvQ8QUOVAkiKnR9YjZZTVqJjpAXzkYymJLaYymUO4C/0k9pSOFhRgRGTVAeqSKXIyV3wDlYVPql3+trC0bzBkyHnITn0p2hnZOlcjP6ELfG4BHbBG9ScgI1NyU9hW7oUdUjsZ1njiPl5ZQFvN9doVORLIknB0JzNrWV7auGCr37fA0gp3vwy/Y4iZZeVkcGklV1PWM7YvXZfJ8A/J+09d4eRGfAAAAAASUVORK5CYII=">
    <title>OTL bookings</title>
    <style>
{css}
    </style>
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
                # wait until Thu to do current week
                while start + timedelta(days=3) <= date.today() + weeks_ahead * timedelta(days=7):
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
                        card_body, copy_text = member.otl_upload_page(start)
                        # language=HTML
                        links_file.write(f'''
        <section class="card" id="card{cards_to_book:02d}">
            <a onclick="copyText(this.getElementsByTagName('h1')[0], '{copy_text}')" ondragstart="copyText(this.getElementsByTagName('h1')[0], '{copy_text}')" href="{url}">
                <header class="card-header">
                    <h1>Time Card</h1>
                </header>
            </a>
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
                    <div class="value"><button onclick="copyText(this, '{copy_text}')">📋 Copy grid</button></div>
                </div>
                <div class="item">
                    <div class="value"><button onclick="document.getElementById('card{cards_to_book:02d}').remove()">☑️ Done</button></div>
                </div>
            </div>
''')
                        links_file.write(card_body)
                        links_file.write('        </section>')
                    start += timedelta(days=7)
        # language=javascript
        copy_code = '''
        function copyText(button, s) {
            const originalText = button.textContent;
            navigator.clipboard.writeText(s.replaceAll(',', '\\t\\t\\t').replaceAll(';', '\\n')).then(() => {
                button.textContent = "✔️ Copied";
                setTimeout(() => {
                    button.textContent = originalText;
                }, 2000);
            });
            }
        function copyColumn(button, s) {
            const originalText = button.textContent;
                navigator.clipboard.writeText(s.replaceAll(';', '\\n')).then(() => {
                button.textContent = "✔️ Copied";
                setTimeout(() => {
                    button.textContent = originalText;
                }, 2000);
            });
            }
'''
        # language=HTML
        links_file.write(f'''
    </main>
    <script>{copy_code}</script>
</body>
</html>
                         ''')
    if cards_to_book:
        return f'{cards_to_book=}', links_filename
    return None


def leave_cross_check(**kwargs):
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


def check_in(**kwargs) -> str | bool:
    """Pick a member of staff to check in with."""
    now = datetime.now()
    # when am I free? start with 0900-1700
    print('Checking my free time')
    my_free_times = outlook.find_free_times()
    if not my_free_times:
        print('No free times available for me')
        return True

    print('Getting previous checkins')
    checkins = get_checkins()

    files: list[Path] = list((docs_folder / 'Group Leader' / 'check_in').glob('*.txt'))
    for filename in files:
        name = filename.stem
        for checkin in checkins:
            if name not in checkin:
                continue
            line = checkin + '\n'
            if line not in filename.read_text():
                print(checkin)
                open(filename, 'a').write(line)
                timestamp = datetime.strptime(checkin[:16], '%Y-%m-%d %H:%M').timestamp()
                os.utime(filename, (timestamp, timestamp))

    files = sorted(files, key=os.path.getmtime)
    # modified X.Y days ago: round down to nearest X
    ages = list(accumulate((now - datetime.fromtimestamp(filename.stat().st_mtime)).days for filename in files))
    # Select basically at random, but pick the same one on a given day
    # Weight older entries higher, and zero weighting to any that are less than one day old
    selected = now.toordinal() % ages[-1]
    index = next(i for i, age in enumerate(ages) if age > selected)
    # When going through list, go back to oldest, then onwards to newest
    files = files[index::-1] + files[index + 1:]
    for filename in files:
        name = filename.stem
        email = filename.read_text().splitlines()[0]
        print(f'Checking free time for {name}')
        their_free_times = outlook.find_free_times(email)
        free_overlap = my_free_times & their_free_times
        if not free_overlap:
            print(f'No free time overlap with {name}')
            continue
        break
    else:
        print('No free time overlap with anyone today')
        return True
    return f'Check in with {name} at {datetimes_to_ranges(free_overlap)}'


def datetimes_to_ranges(datetimes: set[datetime]) -> str:
    """Convert a set of datetimes to a list of ranges in a string."""
    if not datetimes:
        return ""

    sorted_times = sorted(datetimes)
    print(sorted_times)
    half_hour = timedelta(minutes=30)

    ranges = []
    start = sorted_times[0]
    end = sorted_times[0]

    for dt in sorted_times[1:]:
        if dt - end == half_hour:
            # Extends the current range
            end = dt
        else:
            # Gap found — save current range and start a new one
            ranges.append((start, end))
            start = end = dt

    ranges.append((start, end))

    comma_separated_list = ", ".join(
        f"{start.strftime('%H:%M')}-{(end + half_hour).strftime('%H:%M')}"
        for start, end in ranges)
    return ' or '.join(comma_separated_list.rsplit(', ', 1))  # a, b, c -> a, b or c


def list_ftes():
    """List staff and committed time against projects."""
    for member in sorted(members, key=lambda person: person.formal_name()):
        for entry in member.booking_plan.entries:
            print(
                member.formal_name(surname_first=False),
                entry.code.name,
                entry.start_date.strftime('%d/%m/%Y'),
                entry.end_date.strftime('%d/%m/%Y'),
                entry.annual_fte,
                '',
                *[entry.annual_fte if entry.start_date <= fy_start + relativedelta(months=+i) <= entry.end_date else 0
                  for i in range(12)],
            sep='\t')


def goal_page_urls():
    """List this year's objectives for each staff member."""
    web = oracle.go_to_oracle_page('home', show_window=True)
    for member in members:
        url = f'{oracle.fusion_url}goals/goal-center?pPersonId={member.person_id}'
        web.get(url)
        sleep(10)
        goal_titles = web.find_elements(By.CLASS_NAME, 'oj-link')
        print(member.name, len(goal_titles) - 1, 'objectives', url)
        for goal_title in goal_titles[1:]:  # first is 'Skip to main content'
            print('', goal_title.text)


if __name__ == '__main__':
    staff.verbose = True
    print(run_otl_calculator(weeks_ahead=1))
    # print(leave_cross_check())
    # print(check_in())
    # list_ftes()
    # goal_page_urls()