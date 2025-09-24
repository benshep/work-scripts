import os
import shutil
import tempfile
from collections import Counter
from datetime import date, timedelta, datetime
from math import isclose
from typing import Generator
from functools import cache

import pandas

import oracle
import otl
import outlook
from array_round import fair_round
from check_leave_dates import get_off_dates
from folders import docs_folder

site_holidays = outlook.get_dl_ral_holidays(otl.fy)
try:
    site_holidays |= outlook.get_dl_ral_holidays(otl.fy + 1)
except FileNotFoundError:  # not published yet!
    pass


@cache  # so we don't have to keep reading the spreadsheet
def get_obi_data() -> pandas.DataFrame:
    """Read budget info from spreadsheet."""
    data = excel_to_dataframe(os.path.join(docs_folder, 'Budget', 'OBI Staff Bookings.xlsx'))
    if otl.fy == 2025:  # import data from old system, used at beginning of FY25/26
        old_data_file = os.path.join(docs_folder, 'Budget', 'MaRS bookings FY25 pre-Fusion.xlsx')
        old_data = excel_to_dataframe(old_data_file, sheet_name='Sheet2')
        data = pandas.concat([data, old_data])
    return data


def excel_to_dataframe(excel_filename: str, **kwargs) -> pandas.DataFrame:
    """Read Excel data into a pandas dataframe, using a temporary file to circumvent permission issues."""
    # copy to a temporary file to get around permission errors due to workbook being open
    handle, temp_filename = tempfile.mkstemp(suffix='.xlsx')
    os.close(handle)
    shutil.copy2(excel_filename, temp_filename)
    data = pandas.read_excel(temp_filename, dtype={'Task Number': str}, **kwargs)
    os.remove(temp_filename)
    return data


def keep_in_bounds(daily_hours):
    return max(0, min(daily_hours, otl.hours_per_day))


class GroupMember:
    def __init__(self, name: str, person_number: int, person_id: int, assignment_id: int, email: str = '',
                 known_as: str = '', booking_plan: otl.BookingPlan = None):
        self.name = name
        self.person_number = person_number
        self.person_id = person_id
        self.assignment_id = assignment_id
        self.email = email or name.replace(' ', '.').lower() + '@stfc.ac.uk'
        self.known_as = known_as or name.split(' ')[0]
        print(self.known_as)
        self.booking_plan = booking_plan or otl.BookingPlan([])
        self.off_days = site_holidays
        self.new_bookings = Counter()

    def get_oracle_leave_dates(self, test_mode: bool = False, table_format: bool = False) -> set[date] | list[list]:
        """Get leave dates in Oracle for a staff member."""
        web = oracle.go_to_oracle_page('absences', show_window=test_mode)
        web.get(oracle.apps[('absences',)] + f'?pPersonId={self.person_id}')
        off_dates = get_off_dates(web, fetch_all=True, page_count=1, table_format=table_format)
        print(off_dates)
        return off_dates

    def update_off_days(self, force_reload: bool = False):
        cache_file = os.path.join(docs_folder, 'Group Leader', 'off_days_cache', f'{self.known_as}.txt')
        last_week = datetime.now() - timedelta(days=7)
        if force_reload or not os.path.exists(cache_file) or datetime.fromtimestamp(os.path.getmtime(cache_file)) < last_week:
            self.off_days |= outlook.get_away_dates(otl.fy_start, otl.fy_end,
                                                    user=self.email, look_for=outlook.is_annual_leave)
            self.off_days |= self.get_oracle_leave_dates()
            open(cache_file, 'w').write('\n'.join([d.strftime('%d/%m/%Y') for d in sorted(list(self.off_days))]))
        else:
            print(f'Loading off days from cache for {self.known_as}')
            self.off_days = set(datetime.strptime(l, '%d/%m/%Y').date() for l in open(cache_file, 'r').read().splitlines())


    def working_days_in_period(self, start: date, end: date) -> int:
        """Return the number of working days in a period between start and end, given a set of off days."""
        dates = [start + timedelta(day) for day in range((end - start).days + 1)]
        dates = [date for date in dates if date.weekday() < 5 and date not in self.off_days]
        return len(dates)

    def daily_hours(self, entry: otl.Entry, when: date) -> float:
        """Check back through this year's OBI data, and return a number of daily hours to book going forward,
        based on entries so far."""
        if entry.end_date < when or entry.start_date > when:  # already ended, or not started yet
            return 0.0
        obi_data = get_obi_data()
        my_bookings = obi_data[
            (obi_data['Employee/Supplier Name'] == self.name) &
            (obi_data['Expenditure Type'] == 'Labour')
        ]
        code_bookings = my_bookings[
            (my_bookings['Project Number'] == entry.code.project) &
            (my_bookings['Task Number'] == entry.code.task)]
        # self.new_bookings stores bookings added this time
        hours_logged = self.new_bookings[entry.code] + sum(code_bookings['Quantity'])
        hours_needed = otl.hours_per_day * otl.days_per_fte * entry.annual_fte - hours_logged
        rest_of_year_day_count = self.working_days_in_period(when, entry.end_date)
        if rest_of_year_day_count == 0:  # no more days in period
            return 0.0
        hours_per_day = keep_in_bounds(hours_needed / rest_of_year_day_count)
        # print(f'{entry.code} {entry.annual_fte=}, {hours_logged=:.2f} {rest_of_year_day_count=} {hours_per_day=:.2f}')
        return hours_per_day

    def daily_bookings(self, when: date) -> list[tuple[otl.Code, float]]:
        """Return a list of codes and hours to book on a given date."""
        if when in self.off_days:
            return [(otl.unproductive_code, otl.hours_per_day)]

        current_projects = sorted(
            [entry for entry in self.booking_plan.entries if entry.start_date <= when <= entry.end_date],
            key = lambda entry: entry.priority)  # top priority (0) first
        current_projects[-1].priority = otl.Priority.BALANCING  # in case there are no low-priority projects
        total_low_priority = sum(entry.annual_fte
                                 for entry in current_projects
                                 if entry.priority == otl.Priority.BALANCING)
        codes = []
        hours = []
        hours_left = otl.hours_per_day  # keep track of remaining hours for balancing projects
        for entry in current_projects:
            codes.append(entry.code)
            if entry.priority != otl.Priority.BALANCING:  # higher-priority project
                # assign correct number of hours to ensure overall booking over the year is correct
                daily_hours = self.daily_hours(entry, when)
                hours_left -= daily_hours
            else:  # lower-priority project
                # apportion balancing hours depending on share of expected booking
                daily_hours = keep_in_bounds(hours_left * entry.annual_fte / total_low_priority)
            print(entry.code, daily_hours)
            hours.append(daily_hours)
        hours = fair_round(hours, 0.01)  # Oracle rounds to nearest 0.01 and will complain if sum != 7.4
        # Keep track of new bookings so that amounts don't change through the week
        for code, hour in zip(codes, hours):
            self.new_bookings[code] += hour
        assert isclose(sum(hours), otl.hours_per_day)
        return list(zip(codes, hours))

    def bulk_upload_lines(self, week_beginning: date) -> Generator[str]:
        """Generate a series of entries for the bulk upload CSV file, with OTL hours for the given week."""
        # Coerce back to the beginning of the week (Monday)
        week_beginning -= timedelta(days=week_beginning.weekday())
        for day in range(5):
            use_date = week_beginning + timedelta(days=day)
            for code, hours in self.daily_bookings(use_date):
                if hours == 0:
                    continue
                yield ','.join([
                    str(self.person_number),
                    code.project, code.task, 'Labour',
                    use_date.strftime('%d/%m/%Y'), f'{hours:.02f}',
                    f'{self.known_as} timecard submitted through bulk upload {date.today().strftime("%d/%m/%Y")}',
                    'CREATE', '', '', ''
                ])


def check_total_ftes(members: list[GroupMember]):
    print(*[f'{member.known_as}: {member.booking_plan.total_fte()}' for member in members], sep='\n')
