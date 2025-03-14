# import pandas
import os
import win32com.client
import pywintypes
import re
import datetime
import pythoncom
from icalendar import Calendar


def get_appointments_in_range(start=0.0, end=30.0, user='me'):
    """Get a list of appointments in the given range. start and end are relative days from today, or datetimes."""
    today = datetime.date.today()
    from_date = today + datetime.timedelta(days=start) if isinstance(start, (int, float)) else start - datetime.timedelta(days=1)
    to_date = today + datetime.timedelta(days=end) if isinstance(end, (int, float)) else end
    date_filter = f"[Start] >= '{datetime_text(from_date)}' AND [Start] <= '{datetime_text(to_date)}'"
    return get_appointments(date_filter, user=user)


def get_appointments(restriction, sort_order='Start', user='me'):
    """Get a list of calendar appointments with the given Outlook filter."""
    calendar = get_calendar(user)
    appointments = calendar.Items
    appointments.Sort(f"[{sort_order}]")
    appointments.IncludeRecurrences = True
    return appointments.Restrict(restriction)


def get_calendar(user='me'):
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


def get_meeting_time(event, get_end=False):
    """Return the start or end time of the meeting (in UTC) as a Python datetime."""
    meeting_time = event.EndUTC if get_end else event.StartUTC
    return datetime.datetime.fromtimestamp(meeting_time.timestamp())


def happening_now(event, hours_ahead=0.5):
    """Return True if an event is currently happening or is due to start in the next half-hour."""
    try:
        # print(appointmentItem.Start)
        start_time = get_meeting_time(event)
        end_time = get_meeting_time(event, get_end=True)
        buffer = datetime.timedelta(hours=hours_ahead)
        return start_time - buffer <= datetime.datetime.now() <= end_time + buffer
    except (OSError, pywintypes.com_error):  # appointments with weird dates!
        return False


def get_current_events(user='me', hours_ahead=0.5):
    """Return a list of current events from Outlook, sorted by subject."""
    current_events = filter(lambda event: happening_now(event, hours_ahead),
                            get_appointments_in_range(-7 if hours_ahead < 1 else 0, 1 + hours_ahead / 24, user=user))
    # Sort by start time then subject, so we have a predictable order for events to be returned
    return sorted(current_events, key=lambda event: f'{event.StartUTC} {event.Subject}')


def datetime_text(advance_time):
    return advance_time.strftime('%#d/%#m/%Y %H:%M')


def is_my_all_day_event(event):
    """Check whether a given Outlook event is the user's own all day event."""
    return event.AllDayEvent and len(event.Recipients) <= 1  # not sent by someone else


def is_out_of_office(event):
    """Check whether a given Outlook event is an out of office booking."""
    return is_my_all_day_event(event) and event.BusyStatus == 3  # out of office


def is_annual_leave(event):
    """Check whether a given Outlook event is an annual leave booking."""
    return is_out_of_office(event) and any([event.Subject.lower().endswith('annual leave'), event.Subject == 'Off',
                                           re.search(r'\bAL$', event.Subject)])  # e.g. "ARB AL" but not "INTERNAL"


def is_wfh(event):
    """Check whether a given Outlook event is a work from home day."""
    return all([is_my_all_day_event(event),
                event.BusyStatus == 4,  # working elsewhere
                re.search(r'\bWork(ing)? from home$', event.Subject) or re.search(r'\bWFH$', event.Subject)])


def get_away_dates(start=-30, end=90, user='me', look_for=is_out_of_office):
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
        away_list = [get_date_list(event.Start.date(), event.End.date() - datetime.timedelta(days=1)) for event in events]
        return to_set(away_list)
    except pywintypes.com_error:
        print(f"Warning: couldn't fetch away dates for {user}")
        return set()


def get_date_list(start, end=None, count=0):
    """Return a list of business dates (i.e. Mon-Fri) in a given date range, inclusive.
    Specify count instead of end to return a fixed number of dates."""
    if end is not None:  # end date specified, not count
        count = (end - start).days + 1
    elif count == 0:  # neither optional argument specified - assume one day
        count = 1
    date_list = []
    while len(date_list) < count:
        date_list.append(start)
        weekday = start.weekday()
        start += datetime.timedelta((7 - weekday) if weekday > 3 else 1)  # add 3 on Fri, 2 on Sat, 1 for all others
    return date_list


def to_set(date_lists):
    """Convert a list of date lists to a flat set."""
    return set(sum(date_lists, []))


def list_meetings():
    event_list = get_appointments_in_range(start=datetime.date(2022, 4, 1), end=datetime.date(2023, 3, 31))
    for event in event_list:
        print(event.Duration, event.AllDayEvent, event.Subject, event.BusyStatus, sep='\t')
    print(len(event_list))


def get_dl_ral_holidays():
    """Return a list of holiday dates for DL/RAL. Assumes ICS file is already downloaded to Documents/Other folder."""
    filename = os.path.join(os.environ['UserProfile'], 'STFC', 'Documents', 'Other', 'DL_RAL_Site_Holidays_2024.ics')
    calendar = Calendar.from_ical(open(filename, encoding='utf-8').read())
    return {event.decoded('dtstart') for event in calendar.walk('VEVENT')}


if __name__ == '__main__':
    away_dates = sorted(
        list(get_away_dates(datetime.date(2023, 1, 1), datetime.date(2023, 12, 31),
                            user='nasiq.ziyan@stfc.ac.uk')))
    print(len(away_dates))
    print(*away_dates, sep='\n')
