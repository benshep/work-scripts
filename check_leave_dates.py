from datetime import datetime
from oracle import go_to_oracle_page
import outlook
from pushbullet import Pushbullet  # to show notifications
from pushbullet_api_key import api_key  # local file, keep secret!


def list_missing(date_set):
    """Convert a set of missing dates to a comma-separated string."""
    return ', '.join(date.strftime("%#d %b") for date in sorted(list(date_set)))  # e.g. 4 Aug

# def nice_range_text():
#     from datetime import datetime
#     import itertools
#
#     dates = [datetime.strptime(s + ' 2023', "%d %b %Y") for s in '30 May, 31 May, 1 Jun, 2 Jun, 7 Jun'.split(', ')]
#
#     # group consecutive elements using itertools.groupby()
#     groups = []
#     for k, g in itertools.groupby(enumerate(dates), lambda x: x[0] - x[1].toordinal()):
#         groups.append(list(map(lambda x: x[1], g)))
#
#     # create a list of tuples with the first and last element of each group
#     def range_text(date_range):
#         first = date_range[0]
#         last = date_range[-1]
#         if first.year == last.year:
#             if first.month == last.month:
#                 if first.day == last.day:
#                     return first.strftime("%#d %b")
#                 return first.strftime('%#d') + '-' + last.strftime("%#d %b")
#             return first.strftime('%#d %b') + ' - ' + last.strftime("%#d %b")
#         return first.strftime('%#d %b %Y') + ' - ' + last.strftime("%#d %b %Y")
#
#     res = [range_text(group) for group in groups]
#
#     # print result
#     print("Grouped list is : " + str(res))


def check_leave_dates():
    """Compare annual leave dates in an Outlook calendar and in the list submitted to Oracle."""
    oracle_off_dates = get_oracle_off_dates()
    outlook_off_dates = outlook.get_away_dates(min(oracle_off_dates), 0, look_for=outlook.is_annual_leave)

    if not_in_outlook := oracle_off_dates - outlook_off_dates:
        toast = f'Missing from Outlook: {list_missing(not_in_outlook)}\n'
    else:
        toast = ''
    if not_in_oracle := outlook_off_dates - oracle_off_dates:
        toast += f'Missing from Oracle: {list_missing(not_in_oracle)}'
    if toast:
        print(toast)
        Pushbullet(api_key).push_note('📅 Check leave dates', toast)


def get_oracle_off_dates():
    web = go_to_oracle_page(('RCUK Self-Service Employee', 'Attendance Management'))
    cells = web.find_elements_by_class_name('x1w')
    start_dates = cells[::8]
    end_dates = cells[1::8]
    absence_type = cells[2::8]
    off_dates = outlook.to_set(outlook.get_date_list(from_dmy(start.text), from_dmy(end.text))
                               if 'Leave' in ab_type.text else []
                               for start, end, ab_type in zip(start_dates, end_dates, absence_type))
    web.quit()
    return off_dates


def from_dmy(text):
    """Convert a date in the format 31-Oct-2022 to a datetime."""
    return datetime.strptime(text, '%d-%b-%Y').date()


if __name__ == '__main__':
    check_leave_dates()
