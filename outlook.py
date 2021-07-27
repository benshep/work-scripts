import win32com.client
from datetime import datetime, timedelta


def get_appointments_in_range(start=0, end=30):
    """Get a set of appointments in the given range. start and end are relative days from today."""
    today = datetime.today() - timedelta(days=start)
    tomorrow = today + timedelta(days=end)
    appointments = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI').GetDefaultFolder(9).Items
    appointments.Sort("[Start]")
    appointments.IncludeRecurrences = True
    d_m_y = "%#d/%#m/%Y"  # no leading zeros
    restriction = f"[Start] >= '{today.strftime(d_m_y)}' AND [Start] <= '{tomorrow.strftime(d_m_y)}'"
    return appointments.Restrict(restriction)
