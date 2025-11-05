import os
import shutil
import tempfile
from collections import Counter
from datetime import date, timedelta, datetime
from functools import cache
from math import isclose, prod
from typing import Generator

import pandas

import otl
import outlook
from array_round import fair_round
from folders import docs_folder
from work_tools import read_excel

site_holidays = outlook.get_dl_ral_holidays(otl.fy)
try:
    site_holidays |= outlook.get_dl_ral_holidays(otl.fy + 1)
except FileNotFoundError:  # not published yet!
    pass
site_holidays = {day: otl.hours_per_day for day in site_holidays}

web = None
verbose = False


def report(*args, **kwargs):
    if verbose:
        print(*args, **kwargs)


@cache  # so we don't have to keep reading the spreadsheet
def get_obi_data() -> pandas.DataFrame:
    """Read budget info from spreadsheet."""
    data = read_excel(os.path.join(docs_folder, 'Budget', 'OBI Staff Bookings.xlsx'))
    if otl.fy == 2025:  # import data from old system, used at beginning of FY25/26
        old_data_file = os.path.join(docs_folder, 'Budget', 'MaRS bookings FY25 pre-Fusion.xlsx')
        old_data = read_excel(old_data_file, sheet_name='Sheet2')
        data = pandas.concat([data, old_data])
    return data


@cache
def get_absence_data() -> pandas.DataFrame:
    """Read absence info from spreadsheet."""
    return read_excel(os.path.join(docs_folder, 'Budget', 'OBI Absence Report.xlsx'))


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
        # print(self.known_as)
        self.booking_plan = booking_plan or otl.BookingPlan([])
        self.off_days = site_holidays.copy()  # need an independent copy of it, since we'll be making changes!
        self.new_bookings = Counter()
        self.prev_bookings = {}

    def get_oracle_leave_dates(self) -> dict[date, float]:
        """Get leave dates using Oracle data for a staff member."""
        all_absences = get_absence_data()
        my_absences = all_absences[all_absences['Name'] == self.name]
        off_dates = {}
        for i, row in my_absences.iterrows():
            start_date = row['Date Start'].date()
            end_date = row['Date End'].date()
            hours = row['Duration'] * 1 if row['UOM'] == 'H' else otl.hours_per_day
            date_list = outlook.get_date_list(start_date, end_date)
            # Assumption: listed hours are spread equally over listed days
            # with any 'left-over' hours going into the last day
            # e.g. 8.4 hours over 2 days would be 7.4 hours on day 1, 1 hour on day 2
            for day in date_list:
                hours_this_day = min(otl.hours_per_day, hours)
                off_dates[day] = hours_this_day
                hours -= hours_this_day

        report('Oracle:', *[day.strftime('%d/%m/%Y') for day in sorted(off_dates)])
        return off_dates

    def leave_cross_check(self):
        """Perform a cross-check between leave days recorded in Outlook and Oracle."""
        start = max(otl.fy_start, date(2025, 6, 2))  # don't go back before Fusion start date
        end = date.today()  # don't look in the future
        outlook_days = outlook.get_away_dates(start, end, user=self.email, look_for=outlook.is_annual_leave)
        print('Outlook:', *[day.strftime('%d/%m/%Y') for day in sorted(outlook_days)])
        oracle_days = {day for day in self.get_oracle_leave_dates().keys() if start <= day <= end}
        not_in_outlook = oracle_days - outlook_days
        not_in_oracle = outlook_days - oracle_days
        if not_in_outlook:
            print('Not in Outlook:', *sorted(not_in_outlook))
        if not_in_oracle:
            print('Not in Oracle:', *sorted(not_in_oracle))

    def update_off_days(self, force_reload: bool = False) -> None:
        """Load off days from cached file, or if that's more than a week old, reload from Outlook and Oracle.
        :param force_reload: Ignore any cached information.
        """
        cache_file = os.path.join(docs_folder, 'Group Leader', 'off_days_cache', f'{self.name}.txt')
        last_week = datetime.now() - timedelta(days=7)
        cache_exists = os.path.exists(cache_file)
        if force_reload or not cache_exists or datetime.fromtimestamp(os.path.getmtime(cache_file)) < last_week:
            print(f'Fetching Outlook off days for {self.known_as}')
            outlook_days = outlook.get_away_dates(otl.fy_start, otl.fy_end,
                                                  user=self.email, look_for=outlook.is_annual_leave)
            self.off_days |= {day: otl.hours_per_day for day in outlook_days}
            print(f'Fetching Oracle off days for {self.known_as}')
            self.off_days |= self.get_oracle_leave_dates()
            with open(cache_file, 'w') as f:
                f.write('\n'.join(
                    [f"{d.strftime('%d/%m/%Y')}\t{hrs:.2f}" for d, hrs in sorted(self.off_days.items())]))
        else:
            print(f'Loading off days from cache for {self.known_as}')
            self.off_days = {}
            for l in open(cache_file, 'r').read().splitlines():
                day, hrs = l.split('\t')
                self.off_days[datetime.strptime(day, '%d/%m/%Y').date()] = float(hrs)

    def working_days_in_period(self, start: date, end: date) -> int:
        """Return the number of working days in a bin_period between start and end, given a set of off days."""
        dates = [start + timedelta(day) for day in range((end - start).days + 1)]
        dates = [day for day in dates
                 if day.weekday() < 5  # Mon-Fri
                 and (day, otl.hours_per_day) not in self.off_days.items()]  # not a whole day off!
        return len(dates)

    def daily_hours(self, entry: otl.Entry, when: date) -> float:
        """Check back through this year's OBI data, and return a number of daily hours to book going forward,
        based on entries so far."""
        if entry.end_date < when or entry.start_date > when:  # already ended, or not started yet
            return 0.0
        my_bookings = self.get_my_bookings()
        code_bookings = my_bookings[
            (my_bookings['Project Number'] == entry.code.project)
            & (my_bookings['Task Number'] == entry.code.task)
            # fix when booking to same code over multiple periods (e.g. UKXFEL CDOA vs continuation)
            & (my_bookings['Item Date'] >= pandas.to_datetime(entry.code.start))
            ]
        # self.new_bookings stores bookings added this time
        hours_logged = self.new_bookings[entry.code] + sum(code_bookings['Quantity'])
        hours_needed = otl.hours_per_day * otl.days_per_fte * entry.annual_fte - hours_logged
        report(entry.code, hours_logged, hours_needed)
        rest_of_year_day_count = self.working_days_in_period(when, entry.end_date)
        if rest_of_year_day_count == 0:  # no more days in bin_period
            return 0.0
        hours_per_day = max(0.0, hours_needed / rest_of_year_day_count)
        # print(f'{entry.code} {entry.annual_fte=}, {hours_logged=:.2f} {rest_of_year_day_count=} {hours_per_day=:.2f}')
        return hours_per_day

    def get_my_bookings(self) -> pandas.DataFrame:
        """Return my bookings stored in the OBI data spreadsheet."""
        obi_data = get_obi_data()
        return obi_data[obi_data['Employee/Supplier Name'] == self.name]

    def hours_for_week(self, week_beginning: date) -> float:
        """Return my total number of booked hours in the week for the given date."""
        # Coerce back to the beginning of the week (Monday)
        week_beginning -= timedelta(days=week_beginning.weekday())
        my_bookings = self.get_my_bookings()
        weekly_bookings = my_bookings[(my_bookings['Item Date'] >= pandas.to_datetime(week_beginning)) &
                                      (my_bookings['Item Date'] < pandas.to_datetime(
                                          week_beginning + timedelta(days=7)))]
        return sum(weekly_bookings['Quantity'])

    def daily_bookings(self, when: date) -> dict[otl.Code, float]:
        """Return a list of codes and hours to book on a given date."""
        off_hours = self.off_days.get(when, 0)
        unproductive_bookings = {otl.unproductive_code: off_hours} if off_hours else {}
        working_hours = otl.hours_per_day - off_hours  # in case of a half-day off etc
        if not working_hours:
            return unproductive_bookings

        current_projects = sorted(
            [entry for entry in self.booking_plan.entries if entry.start_date <= when <= entry.end_date],
            key=lambda entry: entry.priority)  # top priority (0) first
        current_projects[-1].priority = otl.Priority.BALANCING  # in case there are no low-priority projects
        total_low_priority = sum(entry.annual_fte
                                 for entry in current_projects
                                 if entry.priority == otl.Priority.BALANCING)
        hours_left = working_hours  # keep track of remaining hours for balancing projects
        for entry in current_projects:
            if entry.priority != otl.Priority.BALANCING:  # higher-priority project
                # assign correct number of hours to ensure overall booking over the year is correct
                entry.hours = self.daily_hours(entry, when)
                hours_left -= entry.hours
            else:  # lower-priority project
                # apportion balancing hours depending on share of expected booking
                entry.hours = keep_in_bounds(hours_left * entry.annual_fte / total_low_priority)
            report(entry.code, entry.hours)

        # At this point, we might only have priority projects on this list
        # The total is not necessarily exactly 7.4 hours - it might be more or less
        # How to fix this? Let's try weighting by time left on each project
        # A project with a longer remaining duration should have a lower weight
        hours = [entry.hours for entry in current_projects]
        if not isclose(sum(hours), working_hours, abs_tol=0.5):  # fine if within half an hour
            # Weight by remaining time: shortest gets a weight of 1, twice as long gets 0.5, etc
            # and also by priority: 0 gets weight 1, 1 gets 1/2, 2 gets 1/3
            hours = [hrs / prod([self.working_days_in_period(when, entry.end_date),
                                 entry.priority.value + 1])
                     for hrs, entry in zip(hours, current_projects)]
            report(*hours, sep='\n')
        # Finally scale up or down to match correct number of hours
        scale_factor = sum(hours) / working_hours
        hours = [hrs / scale_factor for hrs in hours]
        hours = fair_round(hours, 0.01)  # Oracle rounds to nearest 0.01 and will complain if sum != 7.4
        report(*hours, sep='\n')
        assert isclose(sum(hours), working_hours)

        # Add up any codes that are identical
        projects_counter = Counter()
        for project, hrs in zip(current_projects, hours):
            if hrs > 0:  # leave out zero bookings
                projects_counter[project.code] += hrs
                # Keep track of new bookings so that amounts don't change through the week
                self.new_bookings[project.code] += hrs

        return {**unproductive_bookings, **projects_counter}

    def bulk_upload_lines(self, week_beginning: date, manual_mode: bool = False) -> Generator[str]:
        """Generate a series of entries for the bulk upload CSV file, with OTL hours for the given week.
        If manual_mode is True, outputs more user-readable entries."""
        # Coerce back to the beginning of the week (Monday)
        week_beginning -= timedelta(days=week_beginning.weekday())
        for day in range(5):
            use_date = week_beginning + timedelta(days=day)
            bookings = self.daily_bookings(use_date)
            if manual_mode:
                yield use_date.strftime('%a')  # Mon
                if dicts_close(bookings, self.prev_bookings):
                    continue
            self.prev_bookings = bookings
            for code, hours in bookings.items():
                yield '\t'.join([str(code), f'{hours:.02f}']) if manual_mode else \
                    ','.join([
                        str(self.person_number),
                        code.project, code.task, 'Labour',
                        use_date.strftime('%d/%m/%Y'), f'{hours:.02f}',
                        f'{self.known_as} timecard submitted through bulk upload {date.today().strftime("%d/%m/%Y")}',
                        'CREATE', '', '', ''
                    ])


def dicts_close(dict1: dict, dict2: dict) -> bool:
    """Return True if dict1 and dict2 have the same keys and all their values are close."""
    return dict1.keys() == dict2.keys() and all([isclose(dict1[key], dict2[key]) for key in dict1.keys()])


def check_total_ftes(members: list[GroupMember]):
    print(*[f'{member.known_as}: {member.booking_plan.total_fte()}' for member in members], sep='\n')
