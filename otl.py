from math import isclose
import os
import pandas
from datetime import date, timedelta, datetime
from enum import IntEnum

from folders import docs_folder
from tools import read_excel

hours_per_day = 7.4
days_per_fte = 215

today = datetime.now()
fy = today.year - (today.month < 4)  # last calendar year if before April
fy_start = date(fy, 4, 1)
fy_end = date(fy + 1, 3, 31)

obi_data = read_excel(os.path.join(docs_folder, 'Budget', 'OBI Finance Report.xlsx'), dtype={'Task Number': str})

# Classes used for OTL bookings

class Priority(IntEnum):
    EXTERNAL = 0
    """Externally-funded project. Top priority - aim to book exactly as expected."""
    AGREED = 1
    """Internally-funded project with a committed amount of time."""
    BALANCING = 2
    """Internally-funded project that can be used to balance bookings."""


class Code:
    """Class to represent a project-task pair."""
    def __init__(self, project: str, task: str = '01',
                 start: date = fy_start, end: date = fy_end,
                 priority: Priority = Priority.EXTERNAL):
        self.project = project
        self.task = task
        self.start = start
        self.end = end
        self.priority = priority

    def __repr__(self):
        return f'{self.project} {self.task}'


unproductive_code = Code('STRA00009', '01.01')  # for ASTeC


class BadDataError(Exception):
    pass


class Entry:
    """Class to represent a booking plan entry for a staff member.
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
    """Class to represent a booking plan for a staff member."""
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
