import contextlib
import os
from datetime import datetime, timedelta
from collections import Counter
import pandas
import outlook
from pushbullet import Pushbullet  # to show notifications
from pushbullet_api_key import api_key  # local file, keep secret!
from tools import read_excel

folder = os.path.join(os.environ['UserProfile'], 'Science and Technology Facilities Council',
                      'DL Activity Coordination Working Group - Occupancy Reports')
file_suffix = '-Daresbury Access.xlsx'
group_name = 'ASTeC Magnetics and Radiation Sources'
# group_name = 'ASTeC Accelerator Physics'


def read_old_access_data():
    os.chdir(folder)
    files = os.listdir()
    days_on_site = Counter()
    for filename in files:
        if not filename.endswith(file_suffix):
            continue
        date = filename[:8]
        min_date = '20200902'
        if date < min_date:  # different format before this
            continue
        date_value = datetime.strptime(date, '%Y%m%d')
        year, week, weekday = date_value.isocalendar()
        if weekday == 1:  # Monday - show last week's data
            last_monday = date_value - timedelta(days=7)
            prev_year, prev_week, _ = last_monday.isocalendar()
            print(prev_year, prev_week,
                  [(surname, days) for (year, week, surname), days in days_on_site.items()
                   if year == prev_year and week == prev_week])
        for person in list(read_excel(filename, sheet_name='Sheet1', dtype='string').itertuples()):
            with contextlib.suppress(TypeError):
                name = person._1
                if name == 'King, Matthew MP' or person._4 == group_name:
                    surname = name.split(', ')[0]
                    days_on_site[(year, week, surname)] += 1
    all_names = sorted({surname for _, _, surname in days_on_site.keys()})
    print('\t', *all_names, sep='\t')
    all_weeks = sorted({(year, week) for year, week, _ in days_on_site.keys()})
    for year, week in all_weeks:
        print(year, week, *(days_on_site[(year, week, surname)] for surname in all_names), sep='\t')


def check_prev_day(date):
    """Check access data of the previous working day."""
    group_members = pandas.Series(['Bainbridge, Alex AR', 'Coku, Mexhisan M', 'Dunning, David DJ', 'Hinton, Alex AG',
                                   'King, Matthew MP',
                                   'Marinov, Kiril KB',
                                   'Pollard, Amelia AE',
                                   'Shepherd, Ben BJA', 'Thompson, Neil NR'])
    # find the latest file
    os.chdir(folder)
    filename = datetime.strftime(date, '%Y%m%d') + file_suffix
    # for each group member: here yesterday? if not - check calendar
    if not os.path.exists(filename):
        return False
    access_list = read_excel(filename, sheet_name='Sheet1', dtype='string').iloc[:, 0]  # first column
    not_on_site = group_members[~group_members.isin(access_list)]
    no_away_events = []
    for name in not_on_site:
        surname, rest = name.split(', ')
        first = rest.split(' ')[0]
        email = f'{first}.{surname}@stfc.ac.uk'.lower()
        print(email)
        # days_away = outlook.get_away_dates(-3, -2, user=email, look_for=outlook.is_not_in_office)
        away = any(event.BusyStatus in (3, 4) for event in
                   outlook.get_appointments_in_range(date, date + timedelta(days=1), user=email))
        if not away:
            no_away_events.append(f'{first} {surname}')
    return no_away_events


def check_prev_week():
    """Check access data of the previous working week."""
    toast = ''
    for days_ago in range(7, 0, -1):
        date = datetime.now() - timedelta(days=days_ago)
        if date.weekday() > 4:  # Sat or Sun
            continue
        if names := check_prev_day(date):
            toast += date.strftime('%#d %b %Y: ') + ', '.join(names) + '\n'
        else:
            return False  # postpone the task
    if toast:
        print(toast)
        Pushbullet(api_key).push_note('Away last week (no calendar event)', toast)


if __name__ == '__main__':
    # read_old_access_data()
    print(check_prev_week())
