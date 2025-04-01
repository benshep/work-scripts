from typing import Any, Iterator
from datetime import date
import pandas
from oracle import go_to_oracle_page
from outlook import get_appointments_in_range, is_annual_leave
from pushbullet import Pushbullet  # to show notifications
from pushbullet_api_key import api_key  # local file, keep secret!
from selenium.webdriver.common.by import By


def get_date_list(start: Any, end: Any, **kwargs: Any) -> list[date]:
    """Return a list of business dates (i.e. Mon-Fri) in a given date range, inclusive."""
    return pandas.bdate_range(start, end, **kwargs).to_pydatetime().tolist()


def to_set(date_lists: Iterator[list[date]]) -> set[date]:
    """Convert a list of date lists to a flat set."""
    return set(sum(date_lists, []))


def list_missing(date_set: set[date]) -> str:
    """Convert a set of missing dates to a comma-separated string."""
    return ', '.join(date.strftime("%#d %b") for date in sorted(list(date_set)))  # e.g. 4 Aug


def check_leave_dates() -> str:
    """Compare annual leave dates in an Outlook calendar and in the list submitted to Oracle."""
    oracle_off_dates = get_oracle_off_dates()

    al_events = filter(is_annual_leave, get_appointments_in_range(min(oracle_off_dates), 0))
    outlook_off_dates = to_set(get_date_list(event.Start.date(), event.End.date(), closed='left') for event in al_events)

    toast = ''
    if not_in_outlook := oracle_off_dates - outlook_off_dates:
        toast = f'Missing from Outlook: {list_missing(not_in_outlook)}\n'
    if not_in_oracle := outlook_off_dates - oracle_off_dates:
        toast += f'Missing from Oracle: {list_missing(not_in_oracle)}'
    return toast


def get_oracle_off_dates() -> set[date]:
    web = go_to_oracle_page('RCUK Self-Service Employee', 'Attendance Management')
    cells = web.find_elements(By.CLASS_NAME, 'x1w')
    start_dates = cells[::8]
    end_dates = cells[1::8]
    off_dates = to_set(get_date_list(start.text, end.text)
                       for start, end in zip(start_dates, end_dates))
    web.quit()
    return off_dates


if __name__ == '__main__':
    check_leave_dates()
