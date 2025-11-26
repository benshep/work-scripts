import os
from datetime import date, datetime
from enum import IntEnum
from math import isclose

from work_folders import budget_folder
from work_tools import read_excel

hours_per_day = 7.4
days_per_fte = 215

today = datetime.now()
fy = today.year - (today.month < 4)  # last calendar year if before April
fy_start = date(fy, 4, 1)
fy_end = date(fy + 1, 3, 31)

obi_data = read_excel(os.path.join(budget_folder, 'OBI Finance Report.xlsx'), dtype={'Task Number': str})


# Classes used for OTL bookings


class Priority(IntEnum):
    """Priority level of a given project. Externally-funded projects take top priority."""
    EXTERNAL = 0
    """Externally-funded project. Top priority - aim to book exactly as expected."""
    AGREED = 1
    """Internally-funded project with a committed amount of time."""
    BALANCING = 2
    """Internally-funded project that can be used to balance bookings."""


class Code:
    """A project-task pair."""

    def __init__(self, project: str, task: str = '01',
                 start: date = fy_start, end: date = fy_end,
                 priority: Priority = Priority.EXTERNAL):
        """
        A booking code.
        :param project: The project code, e.g. STGA00001.
        :param task: The task number. Defaults to 01 if not supplied.
        :param start: The project's start date. Defaults to the start of the current FY.
        :param end: The project's end date. Defaults to the end of the current FY.
        :param priority: The project's priority level.
        """
        self.project = project
        self.task = task
        self.start = start
        self.end = end
        self.priority = priority

    def __repr__(self):
        return f'{self.project} {self.task}'


# What to book various types of leave to? These are specific ASTeC codes
annual_leave = Code('STRA00009', '01.01')
special_paid_leave = Code('STRA00009', '01.02')
parental_leave = Code('STRA00009', '01.03')
sick_leave = Code('STRA00009', '01.04')
unpaid = Code('(no booking)', 'N/A')  # TODO: need to deal with when we move to automated bookings
no_booking = Code('(no booking)', 'N/A')
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


class Entry:
    """A booking plan entry for a staff member.
    Override the default start and end by providing those parameters."""

    def __init__(self, code: Code, annual_fte: float | None = None,
                 start_date: date | None = None, end_date: date | None = None,
                 priority: Priority | None = None):
        self.code = code
        self.annual_fte = annual_fte
        self.start_date = start_date or code.start
        self.end_date = end_date or code.end
        self.priority = priority or code.priority
        self.hours = 0

    def __repr__(self):
        return f'{self.code}: {self.annual_fte or 0:.2f}, from {self.start_date.strftime("%d/%m/%Y")}-{self.end_date.strftime("%d/%m/%Y")}, priority {self.priority.name}'


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
        if not blank_entries:
            if not isclose(self.total_fte(), 1):
                raise BadDataError(f'Total FTE for plan ({self.total_fte():.2f}) != 1.0, no blank entries')
            return
        fte_share = (1 - self.total_fte()) / len(blank_entries)
        if fte_share < -1e-4:  # ignore rounding errors
            raise BadDataError(f'Total FTE for plan ({self.total_fte():.2f}) > 1.0')
        for entry in blank_entries:
            # print(f'Setting {entry.code} FTE share to {fte_share:.2f}')
            entry.annual_fte = fte_share

    def total_fte(self):
        """Calculate the total FTE across all booking codes."""
        return sum(entry.annual_fte or 0 for entry in self.entries)


if __name__ == '__main__':
    bp = BookingPlan([Entry(Code('x', 'y'), 0.5), Entry(Code('x', 'z'))])
    print(*bp.entries, sep='\n')
