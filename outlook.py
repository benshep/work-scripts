import win32com.client
import pywintypes
from datetime import datetime, timedelta


def get_appointments_in_range(start=0, end=30):
    """Get a set of appointments in the given range. start and end are relative days from today, or datetimes."""
    start_is_int = isinstance(start, int)
    from_date = datetime.today() + timedelta(days=start) if start_is_int else start - timedelta(days=1)
    end_is_int = isinstance(end, int)
    to_date = datetime.today() + timedelta(days=end) if end_is_int else end
    outlook = win32com.client.Dispatch('Outlook.Application')
    calendar = outlook.GetNamespace('MAPI').GetDefaultFolder(9)
    appointments = calendar.Items
    appointments.Sort("[Start]")
    appointments.IncludeRecurrences = True
    d_m_y = "%#d/%#m/%Y"  # no leading zeros
    from_date_text = from_date.strftime(d_m_y)
    to_date_text = to_date.strftime(d_m_y)
    restriction = f"[Start] >= '{from_date_text}' AND [Start] <= '{to_date_text}'"
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


