import pandas
from oracle import go_to_oracle_page
from outlook import get_appointments_in_range
from pushbullet import Pushbullet  # to show notifications
from pushbullet_api_key import api_key  # local file, keep secret!


def get_date_list(start, end, **kwargs):
    """Return a list of business dates (i.e. Mon-Fri) in a given date range, inclusive."""
    return pandas.bdate_range(start, end, **kwargs).to_pydatetime().tolist()


def to_set(date_lists):
    """Convert a list of date lists to a flat set."""
    return set(sum(date_lists, []))


def list_missing(date_set):
    """Convert a set of missing dates to a comma-separated string."""
    return ', '.join(date.strftime("%#d %b") for date in sorted(list(date_set)))  # e.g. 4 Aug


def check_leave_dates():
    """Compare annual leave dates in an Outlook calendar and in the list submitted to Oracle."""
    oracle_off_dates = get_oracle_off_dates()

    al_events = filter(is_annual_leave, get_appointments_in_range(min(oracle_off_dates), 0))
    outlook_off_dates = to_set(get_date_list(event.Start.date(), event.End.date(), closed='left') for event in al_events)

    not_in_outlook = oracle_off_dates - outlook_off_dates
    toast = ''
    if not_in_outlook:
        toast = f'Missing from Outlook: {list_missing(not_in_outlook)}\n'
    not_in_oracle = outlook_off_dates - oracle_off_dates
    if not_in_oracle:
        toast += f'Missing from Oracle: {list_missing(not_in_oracle)}'
    if toast:
        print(toast)
        Pushbullet(api_key).push_note('📅 Check leave dates', toast)


def get_oracle_off_dates():
    try:
        web = go_to_oracle_page(('RCUK Self-Service Employee', 'Attendance Management'))
        cells = web.driver.find_elements_by_class_name('x1w')
    finally:
        web.driver.quit()
    start_dates = cells[::8]
    end_dates = cells[1::8]
    return to_set(get_date_list(start.text, end.text) for start, end in zip(start_dates, end_dates))


def is_annual_leave(event):
    """Check whether a given Outlook event is an annual leave booking."""
    return all([event.AllDayEvent, event.BusyStatus == 3, event.Subject == 'Annual Leave'])


if __name__ == '__main__':
    check_leave_dates()
