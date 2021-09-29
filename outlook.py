import pandas
import win32com.client
import pywintypes
from datetime import datetime, timedelta


def get_appointments_in_range(start=0, end=30):
    """Get a list of appointments in the given range. start and end are relative days from today, or datetimes."""
    from_date = datetime.today() + timedelta(days=start) if isinstance(start, int) else start - timedelta(days=1)
    to_date = datetime.today() + timedelta(days=end) if isinstance(end, int) else end
    return get_appointments(f"[Start] >= '{datetime_text(from_date)}' AND [Start] <= '{datetime_text(to_date)}'")


def get_appointments(restriction, sort_order='Start'):
    """Get a list of calendar appointments with the given Outlook filter."""
    outlook = win32com.client.Dispatch('Outlook.Application')
    appointments = outlook.GetNamespace('MAPI').GetDefaultFolder(9).Items
    appointments.Sort(f"[{sort_order}]")
    appointments.IncludeRecurrences = True
    return appointments.Restrict(restriction)


def get_meeting_time(event, get_end=False):
    """Return the start or end time of the meeting (in UTC) as a Python datetime."""
    meeting_time = event.EndUTC if get_end else event.StartUTC
    return datetime.fromtimestamp(meeting_time.timestamp())


def happening_now(event):
    """Return True if an event is currently happening or is due to start in the next half-hour."""
    try:
        # print(appointmentItem.Start)
        start_time = get_meeting_time(event)
        end_time = get_meeting_time(event, get_end=True)
        return start_time - timedelta(hours=0.5) <= datetime.now() <= end_time
    except (OSError, pywintypes.com_error):  # appointments with weird dates!
        return False


def get_current_events():
    """Return a list of current events from Outlook, sorted by subject."""
    current_events = filter(happening_now, get_appointments_in_range(-7, 1))
    # Sort by subject so we have a predictable order for events to be returned
    current_events = sorted(current_events, key=lambda event: event.Subject)
    return current_events


def datetime_text(advance_time):
    return advance_time.strftime('%#d/%#m/%Y %H:%M')


def is_annual_leave(event):
    """Check whether a given Outlook event is an annual leave booking."""
    return all([event.AllDayEvent, event.BusyStatus == 3, event.Subject == 'Annual Leave'])


def get_outlook_leave_dates(start=-30, end=90):
    al_events = filter(is_annual_leave, get_appointments_in_range(start, end))
    return [get_date_list(event.Start.date(), event.End.date(), closed='left') for event in al_events]


def get_date_list(start, end, **kwargs):
    """Return a list of business dates (i.e. Mon-Fri) in a given date range, inclusive."""
    return pandas.bdate_range(start, end, **kwargs).to_pydatetime().tolist()