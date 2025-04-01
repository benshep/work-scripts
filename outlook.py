# import pandas
import os
from typing import Protocol, Any

import win32com.client
import pywintypes
import re
from datetime import date, datetime, timedelta
from enum import Enum
import pythoncom
from icalendar import Calendar
from folders import hr_info_folder


class OlResponseStatus(Enum):
    """Indicates the response to a meeting request."""
    # https://learn.microsoft.com/en-us/office/vba/api/outlook.olresponsestatus

    accepted = 3
    """Meeting accepted."""
    declined = 4
    """Meeting declined."""
    none = 0
    """The appointment is a simple appointment and does not require a response."""
    not_responded = 5
    """Recipient has not responded."""
    organized = 1
    """The AppointmentItem is on the Organizer's calendar or the recipient is the Organizer of the meeting."""
    tentative = 2
    """Meeting tentatively accepted."""


class Recipient(Protocol):
    """Represents a user or resource in Outlook, generally a mail or mobile message addressee."""
    # https://learn.microsoft.com/en-us/office/vba/api/outlook.recipient

    Name: str
    """The display name for the Recipient."""
    Address: str
    """The email address of the Recipient."""
    MeetingResponseStatus: OlResponseStatus
    """The overall status of the response to the meeting request for the recipient."""
    AddressEntry: object
    Application: object
    AutoResponse: str
    """The text of an automatic response for a Recipient."""
    Class: object
    DisplayType: object
    EntryID: str
    """The unique Entry ID of the object."""
    Index: int
    """The position of the object within the collection."""
    Parent: object
    PropertyAccessor: object
    Resolved: bool
    """Indicates True if the recipient has been validated against the Address Book."""
    Sendable: bool
    """Indicates whether a meeting request can be sent to the Recipient."""
    Session: object
    TrackingStatus: object
    TrackingStatusTime: object
    Type: int

    def Delete(self) -> None:
        """Deletes an object from the collection."""
        pass

    def FreeBusy(self, start: pywintypes.TimeType, min_per_char: int, complete_format: bool = False) -> str:
        """Returns free/busy information for the recipient.
        The default is to return a string representing one month of free/busy information compatible with the
        Microsoft Schedule+ Automation format (that is, the string contains one character for each MinPerChar minute,
        up to one month of information from the specified Start date).
        If the optional argument complete_format is omitted or False, then "free" is indicated by the character 0
        and all other states by the character 1.
        If CompleteFormat is True, then the same length string is returned as defined above,
        but the characters now correspond to the OlBusyStatus constants."""

    def Resolve(self) -> bool:
        """Attempts to resolve a Recipient object against the Address Book.
        Returns true if the object was resolved; otherwise, false."""

class OlBusyStatus(Enum):
    busy = 2
    """The user is busy."""
    free = 0
    """The user is available."""
    out_of_office = 3
    """The user is out of office."""
    tentative = 1
    """The user has a tentative appointment scheduled."""
    working_elsewhere = 4
    """The user is working in a location away from the office."""


class AppointmentItem(Protocol):
    """Represents a meeting, a one-time appointment, or a recurring appointment or meeting in the Calendar folder."""
    # https://learn.microsoft.com/en-us/office/vba/api/outlook.appointmentitem

    Subject: str
    """The subject for the Outlook item."""
    Body: str
    """The clear-text body of the Outlook item."""
    AllDayEvent: bool
    """Returns True if the appointment is an all-day event (as opposed to a specified time)."""
    BusyStatus: OlBusyStatus
    """The busy status of the user for the appointment."""
    Start: pywintypes.TimeType
    """The starting date and time for the Outlook item."""
    StartUTC: pywintypes.TimeType
    """The start date and time of the appointment expressed in the Coordinated Universal Time (UTC) standard."""
    End: pywintypes.TimeType
    """The end date and time for the Outlook item."""
    EndUTC: pywintypes.TimeType
    """The end date and time of the appointment expressed in the Coordinated Universal Time (UTC) standard."""
    Duration: int
    """Duration in minutes."""
    Recipients: list[Recipient]
    """All the recipients for the Outlook item."""
    RequiredAttendees: str
    """A semicolon-delimited String of required attendee names for the meeting appointment."""
    OptionalAttendees: str
    """A semicolon-delimited String of optional attendee names for the meeting appointment."""


date_spec = datetime | date | float | int
"""A date/datetime or relative days from today."""


def get_appointments_in_range(start: date_spec = 0.0,
                              end: date_spec = 30.0,
                              user: str = 'me') -> list[AppointmentItem]:
    """Get a list of appointments in the given range. start and end are relative days from today, or datetimes."""
    today = date.today()
    from_date = today + timedelta(days=start) if isinstance(start, (int, float)) else start - timedelta(days=1)
    to_date = today + timedelta(days=end) if isinstance(end, (int, float)) else end
    date_filter = f"[Start] >= '{datetime_text(from_date)}' AND [Start] <= '{datetime_text(to_date)}'"
    return get_appointments(date_filter, user=user)


def get_appointments(restriction: str, sort_order: str = 'Start', user: str = 'me') -> list[AppointmentItem]:
    """Get a list of calendar appointments with the given Outlook filter."""
    calendar = get_calendar(user)
    appointments = calendar.Items
    appointments.Sort(f"[{sort_order}]")
    appointments.IncludeRecurrences = True
    return appointments.Restrict(restriction)


def get_calendar(user: str = 'me'):
    """Return the calendar folder for a given user. If none supplied, default to my calendar."""
    namespace = get_outlook().GetNamespace('MAPI')
    if user == 'me':
        return namespace.GetDefaultFolder(9)
    recipient = namespace.CreateRecipient(user)
    if not recipient.Resolve():
        raise RuntimeError(f'User "{user}" not found.')
    return namespace.GetSharedDefaultFolder(recipient, 9)


def get_outlook():
    """Return a reference to the Outlook application."""
    pythoncom.CoInitialize()  # try to combat the "CoInitialize has not been called" error
    return win32com.client.Dispatch('Outlook.Application')


def get_meeting_time(event: AppointmentItem, get_end: bool = False) -> datetime:
    """Return the start or end time of the meeting (in UTC) as a Python datetime."""
    meeting_time = event.EndUTC if get_end else event.StartUTC
    return datetime.fromtimestamp(meeting_time.timestamp())


def happening_now(event: AppointmentItem, hours_ahead: float = 0.5) -> bool:
    """Return True if an event is currently happening or is due to start in the next half-hour."""
    try:
        # print(appointmentItem.Start)
        start_time = get_meeting_time(event)
        end_time = get_meeting_time(event, get_end=True)
        buffer = timedelta(hours=hours_ahead)
        return start_time - buffer <= datetime.now() <= end_time + buffer
    except (OSError, pywintypes.com_error):  # appointments with weird dates!
        return False


def get_current_events(user: str = 'me', hours_ahead: float = 0.5) -> list[AppointmentItem]:
    """Return a list of current events from Outlook, sorted by subject."""
    current_events = filter(lambda event: happening_now(event, hours_ahead),
                            get_appointments_in_range(-7, 1 + hours_ahead / 24, user=user))
    # Sort by start time then subject, so we have a predictable order for events to be returned
    return sorted(current_events, key=lambda event: f'{event.StartUTC} {event.Subject}')


def datetime_text(advance_time: datetime) -> str:
    return advance_time.strftime('%#d/%#m/%Y %H:%M')


def is_my_all_day_event(event: AppointmentItem) -> bool:
    """Check whether a given Outlook event is the user's own all day event."""
    return event.AllDayEvent and len(event.Recipients) <= 1  # not sent by someone else


def is_out_of_office(event: AppointmentItem) -> bool:
    """Check whether a given Outlook event is an out of office booking."""
    return is_my_all_day_event(event) and event.BusyStatus == OlBusyStatus.out_of_office


def is_annual_leave(event: AppointmentItem) -> bool:
    """Check whether a given Outlook event is an annual leave booking."""
    return is_out_of_office(event) and any([event.Subject.lower().endswith('annual leave'), event.Subject == 'Off',
                                            re.search(r'\bAL$', event.Subject)])  # e.g. "ARB AL" but not "INTERNAL"


def is_wfh(event: AppointmentItem) -> bool:
    """Check whether a given Outlook event is a work from home day."""
    return all([is_my_all_day_event(event),
                event.BusyStatus == OlBusyStatus.working_elsewhere,
                re.search(r'\bWork(ing)? from home$', event.Subject) or re.search(r'\bWFH$', event.Subject)])


def get_away_dates(start: date_spec = -30, end: date_spec = 90,
                   user: str = 'me', look_for: callable = is_out_of_office) -> set[AppointmentItem]:
    """Return a set of the days in the given range that the user is away from the office.
    The look_for parameter can be:
     - is_my_all_day_event: look for any all day event with only the user as an attendee
     - is_out_of_office (default): as above, but set to out of office
     - is_annual_leave: as above, but subject is "Annual Leave" or "AL"
     - is_wfh: set to out of office, and subject is "Work(ing) from home" or "WFH"
     """
    events = filter(look_for, get_appointments_in_range(start, end, user=user))
    # Need to subtract a day here since the end time of an all-day event is 00:00 on the next day
    try:
        away_list = [get_date_list(event.Start.date(), event.End.date() - timedelta(days=1)) for event in events]
        return to_set(away_list)
    except pywintypes.com_error:
        print(f"Warning: couldn't fetch away dates for {user}")
        return set()


def get_date_list(start: datetime,
                  end: datetime | None = None,
                  count: int = 0) -> list[datetime]:
    """Return a list of business dates (i.e. Mon-Fri) in a given date range, inclusive.
    Specify count instead of end to return a fixed number of dates."""
    if end is None and count == 0:  # neither optional argument specified - assume one day
        count = 1
    date_list = []
    while True:
        weekday = start.weekday()
        if weekday < 5:  # Mon-Fri
            date_list.append(start)
            if 0 < count == len(date_list):
                break
        start += timedelta(days=1)
        if end and start > end:
            break
    return date_list


def to_set(date_lists: list[list[date | datetime]]) -> set[date | datetime]:
    """Convert a list of date lists to a flat set."""
    return set(sum(date_lists, []))


def list_meetings():
    event_list = get_appointments_in_range(start=date(2022, 4, 1), end=date(2023, 3, 31))
    for event in event_list:
        print(event.Duration, event.AllDayEvent, event.Subject, OlBusyStatus(event.BusyStatus), sep='\t')
    print(len(event_list))


def get_dl_ral_holidays() -> set[date]:
    """Return a list of holiday dates for DL/RAL. Fetches ICS file from STFC HR info folder (synced via OneDrive)."""
    filename = os.path.join(hr_info_folder, f'DL_RAL_Site_Holidays_{datetime.now().year}.ics')
    calendar = Calendar.from_ical(open(filename, encoding='utf-8').read())
    return {event.decoded('dtstart') for event in calendar.walk('VEVENT')}


if __name__ == '__main__':
    # away_dates = sorted(
    #     list(get_away_dates(datetime.date(2024, 2, 12), 0, look_for=is_annual_leave)))
    # print(len(away_dates))
    # print(*away_dates, sep='\n')
    print(list_meetings())
