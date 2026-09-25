import random
from datetime import date, datetime
from enum import IntEnum
from math import isclose

import cvxpy
import numpy as np
from dateutil.relativedelta import relativedelta

months = 12
hours_per_day = 7.4
days_per_fte = 215
hours_per_fte = hours_per_day * days_per_fte

today = datetime.now()
fy = today.year - (today.month < 4)  # last calendar year if before April
fy_start = date(fy, 4, 1)
fy_end = date(fy + 1, 3, 31)

unproductive = 'Unproductive - Straight Time'
straight_time = 'Labour - Straight Time'


# Classes used for OTL bookings


class Priority(IntEnum):
    """Priority level of a given project. Externally-funded projects take top priority."""
    EXTERNAL = 0
    """Externally-funded project. Top priority - aim to book exactly as expected."""
    AGREED = 1
    """Internally-funded project with a committed amount of time."""
    BALANCING = 2
    """Internally-funded project that can be used to balance bookings."""

    def weight(self):
        """Weighting for cvxpy-based resource levelling."""
        return 10 ** (3 - self.value)


class Profile(IntEnum):
    """Profile type of a project."""
    FLAT = 0
    """Level of effort is expected to be flat through the project."""
    FRONT_LOADED = 1
    """Most of the effort at the start."""
    BACK_LOADED = 2
    """Most of the effort at the end."""
    BELL = 3
    """Low at the start and end, high in the middle."""

    def shape(self, n):
        """Return the shape of the project given how many months it runs for."""
        match self.value:
            case self.FRONT_LOADED:
                profile = np.linspace(n, 1, n)
            case self.BACK_LOADED:
                profile = np.linspace(1, n, n)
            case self.BELL:
                profile = np.hanning(n) + 0.1
            case _:
                profile = np.ones(n)
        return profile / profile.sum()  # sum to 1.0


def table_cell(key: str, display_value: str) -> str:
    """A representation of a table cell that can be pasted into Fusion."""
    return f'{{"data":"{key}","valueItem":{{"key":"{key}","data":{{"Code":"{key}","DisplayValue":"{display_value}","UnitOfMeasure":"HR"}}}}}}'


class Code:
    """A project-task pair."""

    def __init__(self, project: str, task: str = '01',
                 name: str = '', fusion_name: str = '',
                 start: date = fy_start, end: date = fy_end,
                 priority: Priority = Priority.EXTERNAL,
                 project_key: str = '', task_key: str = '',
                 hours_type: str = straight_time):
        """
        A booking code.
        :param project: The project code, e.g. STGA00001.
        :param task: The task number. Defaults to 01 if not supplied.
        :param name: The name of the task, as used in the budget sheet. Defaults to blank.
        :param fusion_name: The name of the task, as used in Oracle Fusion. Defaults to same as name.
        :param start: The project's start date. Use the start of the current FY if not provided or starts earlier.
        :param end: The project's end date. Use the end of the current FY if not provided or finishes later.
        :param priority: The project's priority level.
        :param project_key: The project key number, intended for copy-pasting within Fusion.
        :param task_key: The project key number, intended for copy-pasting within Fusion.
        :param hours_type: The project's hours type. Defaults to 'Labour - Straight Time'.
        """
        self.project = project
        self.task = task
        self.name = name
        self.fusion_name = fusion_name or name
        self.start = max(start, fy_start)
        self.end = min(end, fy_end)
        self.priority = priority
        self.project_key = project_key
        self.task_key = task_key
        self.hours_type = hours_type

    def __repr__(self):
        return f'{self.project} {self.task}'

    def project_cell(self):
        """A representation of the project code that can be pasted into Fusion."""
        if not self.project_key:
            return ''
        return table_cell(self.project_key, f"{self.project} - {self.fusion_name}")

    def task_cell(self):
        """A representation of the task code that can be pasted into Fusion."""
        if not self.task_key:
            return ''
        return table_cell(self.task_key, f"{self.task} - {self.project}")

    def hours_cell(self):
        """A representation of the hours type that can be pasted into Fusion."""
        key = {
            straight_time: '300000012241990',
            unproductive: '',
        }[self.hours_type]
        return table_cell(key, self.hours_type)


# What to book various types of leave to? These are specific ASTeC codes
annual_leave, special_paid_leave, parental_leave, sick_leave = [
    Code('STRA00009', f'01.{i + 1:02d}',  # 01.01, 01.02, 01.03, 01.04
         fusion_name='ASTeC Non-Productive Time', hours_type=unproductive,
         project_key='300000105495142')
    for i in range(4)]
unpaid = Code('(no booking)', 'N/A',
              hours_type=unproductive)  # TODO: need to deal with when we move to automated bookings
no_booking = Code('(no booking)', 'N/A', hours_type=unproductive)
unproductive_code = {  # list of possible Absence Types in Fusion
    'Annual Leave': annual_leave,
    'Bank Holiday': annual_leave,
    'Privilege Day': annual_leave,
    'Privilege Days': annual_leave,
    # ‘No booking’ days and half-days relate to the ‘compensating leave’ accrued through the
    # additional minutes worked each week throughout the year; these should be left blank on OTL
    'Compensating Leave': no_booking,
    'Career Break': special_paid_leave,
    'Phased Return - Leave': special_paid_leave,
    'Sabbatical Leave': special_paid_leave,
    'Special Leave - Paid': special_paid_leave,
    'Special Leave - Unpaid': unpaid,
    'Unpaid Leave': unpaid,
    'Unpaid Parental Leave': unpaid,
    'Maternity Leave': parental_leave,
    'Adoption Leave': parental_leave,
    'Paternity/Maternity Support - Adoption': parental_leave,
    'Paternity/Maternity Support - Birth': parental_leave,
    'Sickness Absence': sick_leave,
    'Half Day Sickness Absence': sick_leave,
    'Industrial Action': unpaid,
}


class BadDataError(Exception):
    pass


class LevellingFailed(Exception):
    pass


class Entry:
    """A booking plan entry for a staff member.
    Override the default start and end by providing those parameters."""

    def __init__(self, code: Code, annual_fte: float | None = None,
                 start_date: date | None = None, end_date: date | None = None,
                 priority: Priority | None = None, profile: Profile = Profile.FLAT):
        self.code = code
        self.annual_fte = annual_fte
        self.start_date = start_date or code.start
        self.end_date = end_date or code.end
        self.priority = priority or code.priority
        self.hours = 0
        self.monthly_fte = []
        self.profile = profile

    def __repr__(self):
        return f'{self.code}: {self.annual_fte or 0:.2f}, from {self.start_date.strftime("%d/%m/%Y")}-{self.end_date.strftime("%d/%m/%Y")}, priority {self.priority.name}'

    def __str__(self):
        return f'{self.code}\t{self.annual_fte or 0:.2f}\t{self.start_date.strftime("%d/%m/%Y")}\t{self.end_date.strftime("%d/%m/%Y")}\t{self.priority.value}'

    def is_active_on(self, when: date | int) -> bool:
        """Return True if the entry is active on the given date, False otherwise.
        Supply a date or an integer month (0-11)."""
        when_date = when if isinstance(when, date) else fy_start + relativedelta(months=when)
        return self.start_date <= when_date <= self.end_date


def month_delta(start: date, end: date) -> int:
    """Return the number of months between start and end date."""
    return end.year * 12 + end.month - start.year * 12 - start.month


class BookingPlan:
    """Booking plan for a staff member."""

    def __init__(self, entries: list[Entry]):
        self.entries = entries
        # Any low-priority (balancing) entries? If not, make the last one low-priority
        if not [entry for entry in self.entries if entry.priority == Priority.BALANCING]:
            # print(f'No low-priority projects found. Setting {entries[-1].code} to low priority')
            entries[-1].priority = Priority.BALANCING
        # If any entries don't have an annual FTE amount, assign them one to make the total up to 1
        blank_entries = [entry for entry in self.entries if entry.annual_fte is None]
        if blank_entries:
            fte_share = (1 - self.total_fte()) / len(blank_entries)
            if fte_share < -1e-4:  # ignore rounding errors
                raise BadDataError(f'Total FTE for plan ({self.total_fte():.2f}) > 1.0')
            for entry in blank_entries:
                # print(f'Setting {entry.code} FTE share to {fte_share:.2f}')
                entry.annual_fte = fte_share
        elif not isclose(self.total_fte(), 1, abs_tol=1e-3):
            raise BadDataError(f'Total FTE for plan ({self.total_fte():.4f}) != 1.0, no blank entries')
        # Naive allocation: refine it later
        for entry in self.entries:
            month_count = month_delta(entry.start_date, entry.end_date) + 1
            shape = entry.profile.shape(month_count) * entry.annual_fte
            shape = np.pad(shape, (
                max(0, month_delta(fy_start, entry.start_date)),
                max(0, month_delta(entry.end_date, fy_end))
            ))
            entry.monthly_fte = shape

    def total_fte(self):
        """Calculate the total FTE across all booking codes."""
        return sum(entry.annual_fte or 0 for entry in self.entries)

    def monthly_total(self):
        """Calculate the monthly total FTE across all booking codes."""
        return np.sum(np.array([entry.monthly_fte for entry in self.entries]), axis=0)

    def monthly_resource_levelling(self):
        """Set a monthly plan, ensuring all months add up to 100% effort and all projects have the correct allocation
        for the year, respecting time boundaries and priorities."""
        for m in range(months):
            # Any month exceeding 100% total capacity gets reduced starting with the lowest-priority projects.
            # Any month under 100% total capacity gets increased starting with the lowest-priority projects.
            total = sum([entry.monthly_fte[m] for entry in self.entries])
            surplus = total - 1.0 / months  # positive surplus: allocation too large this month
            # print(i, total, surplus)
            for project in sorted([entry for entry in self.entries if entry.is_active_on(m)],
                                  key=lambda entry: entry.priority, reverse=True):
                if isclose(surplus, 0):
                    break
                new_amount = max(0, min(project.monthly_fte[m] - surplus, 1))
                # print(project.code, project.monthly_fraction[i], new_amount)
                # Increased allocation: change is positive
                # Decreased allocation: change is negative
                change = new_amount - project.monthly_fte[m]
                project.monthly_fte[m] = new_amount
                surplus += change

    def convex_levelling(self):
        """Set a monthly plan, ensuring all months add up to 100% effort and all projects have the correct allocation
        for the year, respecting time boundaries and priorities. Use cvxpy to do it with MATHS."""
        n_projects = len(self.entries)
        allocation_matrix = cvxpy.Variable(shape=(n_projects, months), nonneg=True)
        shortfall = cvxpy.Variable(n_projects, nonneg=True)

        # Set constraints. Monthly capacity:
        constraints = [cvxpy.sum(allocation_matrix[:, m]) <= 1.0 / months for m in range(months)]
        # Zero outside project dates
        for i, entry in enumerate(self.entries):
            for m in range(months):
                if not entry.is_active_on(m):
                    constraints.append(allocation_matrix[i, m] == 0)
            # Required effort + shortfall
            constraints.append(cvxpy.sum(allocation_matrix[i, :]) + shortfall[i] == entry.annual_fte)

        target = np.array([entry.monthly_fte for entry in self.entries])

        # Objective terms. Shortfall penalties:
        priority_weights = np.array([entry.priority.weight() for entry in self.entries])
        shortfall_term = cvxpy.sum(cvxpy.multiply(priority_weights, shortfall))
        # Even allocation term
        evenness_term = cvxpy.sum_squares(allocation_matrix - target)
        # Month-to-month smoothing
        smoothness_term = 0
        for m in range(months - 1):
            smoothness_term += cvxpy.sum_squares(allocation_matrix[:, m + 1] - allocation_matrix[:, m])
        # Fairly arbitrary weighting for terms
        objective = cvxpy.Minimize(
            100 * shortfall_term
            + 10 * evenness_term
            + 1 * smoothness_term
        )
        problem = cvxpy.Problem(objective, constraints)
        problem.solve()
        if problem.status != 'optimal':
            raise LevellingFailed(f'Levelling failed, status {problem.status}')
        for entry, row in zip(self.entries, allocation_matrix.value):
            entry.monthly_fte = row


def to_percent(x: float) -> str:
    """0.00% representation of number."""
    digits = 0 if abs(x) >= 0.1 else 1
    return (f'{x * 100:.{digits}f}%' if x >= 0.005 else '').ljust(5)


def levelling_test():
    words = open('1000-most-common-english-words.txt').read().splitlines()
    n_projects = random.randint(5, 10)
    remaining_fte = 1.0
    projects = []
    min_start = None
    max_end = None
    max_name_length = 0
    for i in range(n_projects):
        fte = random.random() * remaining_fte if i < n_projects - 1 else remaining_fte
        dates = sorted([random.randint(0, 11), random.randint(0, 11)])
        start = fy_start + relativedelta(months=dates[0])
        end = fy_start + relativedelta(months=dates[1] + 1, days=-1)
        min_start = start if min_start is None else min(min_start, start)
        max_end = end if max_end is None else max(max_end, end)
        if i == n_projects - 1:
            if min_start > fy_start:
                start = fy_start
            if max_end < fy_end:
                end = fy_end
        name = ' '.join(random.choice(words).title() for _ in range(3))
        max_name_length = max(max_name_length, len(name))
        project = Entry(Code(name), fte, start_date=start, end_date=end, priority=Priority(random.randint(0, 2)))
        remaining_fte -= fte
        projects.append(project)

    projects = [
        Entry(Code('External-Funded Thing'), 0.3, start_date=date(2026, 8, 1), end_date=date(2026, 12, 31),
              priority=Priority.EXTERNAL, profile=Profile.FRONT_LOADED),
        Entry(Code('Important ASTeC Thing'), 0.3, start_date=date(2026, 10, 1), end_date=fy_end,
              priority=Priority.AGREED, profile=Profile.BACK_LOADED),
        Entry(Code('ASTeC Core Task'), 0.4, start_date=fy_start, end_date=fy_end, priority=Priority.BALANCING),
    ]
    max_name_length = max([len(entry.code.project) for entry in projects])

    bp = BookingPlan(projects)
    for entry in bp.entries:
        entry.code.project = entry.code.project.ljust(max_name_length)
        print(str(entry), *[to_percent(effort * 12) for effort in entry.monthly_fte], f'{sum(entry.monthly_fte):.02f}',
              sep='\t')
        l = len(str(entry))
    print(' ' * l, '', '', *[to_percent(x * 12) for x in bp.monthly_total()], sep='\t')
    print('')
    bp.convex_levelling()
    for entry in bp.entries:
        print(str(entry), *[to_percent(effort * 12) for effort in entry.monthly_fte], f'{sum(entry.monthly_fte):.02f}',
              sep='\t')
    print(' ' * l, '', '', *[to_percent(x * 12) for x in bp.monthly_total()], sep='\t')


if __name__ == '__main__':
    # shape = Profile.FRONT_LOADED.shape(6)
    # print(shape, sum(shape))
    levelling_test()
