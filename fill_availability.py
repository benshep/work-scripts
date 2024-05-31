import os
import pandas
from outlook import get_appointments_in_range


def fill_availability():
    weeks = 8
    available_dates = get_available_dates(weeks)

    user_profile = os.environ['UserProfile']
    stfc = 'Science and Technology Facilities Council'
    rota = 'CLARA_Shifts_rota'
    file = 'Ben.xlsx'
    spreadsheet_xlsx = os.path.join(user_profile, f'OneDrive - {stfc}', rota, file)
    if not os.path.exists(spreadsheet_xlsx):  # alternative sync location (laptop)
        spreadsheet_xlsx = os.path.join(user_profile, stfc, f'CLARA Workflow - {rota}', file)
    mod_time = pandas.to_datetime(os.path.getmtime(spreadsheet_xlsx), unit='s')
    old_end_date = mod_time + pandas.to_timedelta(weeks, 'W') - pandas.to_timedelta(1, 'd')
    column_name = 'available'
    old_free_dates = pandas.read_excel(spreadsheet_xlsx)[column_name]  # returns a Series
    new_data = pandas.DataFrame(sorted(list(available_dates)), columns=[column_name])
    today = pandas.to_datetime('today').floor('d')
    with pandas.ExcelWriter(spreadsheet_xlsx, mode='a', if_sheet_exists='replace') as writer:
        new_data.to_excel(writer, sheet_name='Ben')
        pandas.DataFrame([today]).to_excel(writer, sheet_name='Updated')

    print(f"Old list has {len(old_free_dates)} entries, file modified {format_date(mod_time)}")
    old_free_dates = old_free_dates[old_free_dates >= today]
    print(f'Ignoring dates in old list before today, {format_date(today)} - reduced to {len(old_free_dates)} entries')
    new_free_dates = new_data[column_name]
    print(f"New list has {len(new_free_dates)} entries")
    new_free_dates = new_free_dates[new_free_dates < old_end_date]
    print(f'Ignoring dates in new list after {format_date(old_end_date)} - reduced to {len(new_free_dates)} entries')
    old_free_dates = set(old_free_dates)
    new_free_dates = set(new_free_dates)
    changed_to_yes = sorted(list(new_free_dates - old_free_dates))
    changed_to_no = sorted(list(old_free_dates - new_free_dates))
    toast = ''
    if changed_to_yes:
        toast += '👍 ' + ', '.join(format_date(date) for date in changed_to_yes) + '\n'
    if changed_to_no:
        toast += '👎 ' + ', '.join(format_date(date) for date in changed_to_no)

    return toast


def format_date(date, include_time=False):
    """Return a string with the date formatted as (e.g.) 4 Aug, with an optional HH:MM time component."""
    return date.strftime("%#d %b" + (' %H:%M' if include_time else ''))


def get_available_dates(weeks):
    """Get a list of available dates from Outlook."""
    available_dates = bdate_set('today', periods=5 * weeks)
    print('Fixed events:')
    for event in get_appointments_in_range(0, 7 * weeks):
        recurrence = event.GetRecurrencePattern()
        weekly = event.IsRecurring and recurrence.RecurrenceType == 1 and recurrence.Interval == 1
        busy = event.BusyStatus == 2
        away = event.BusyStatus == 3
        people = len(event.RequiredAttendees.split('; '))
        hours = event.Duration / 60
        days = hours / 24
        if fixed_event := (away or (busy and not weekly)) and (people < 2 or people > 5 or hours >= 2):
            all_day = event.AllDayEvent
            true_values = ', '.join(var for var, value in locals().items() if value is True and var != 'fixed_event')
            duration = f"{days=}" if all_day else f"{hours=:.1f}"
            start = format_date(event.Start, not all_day)
            print(start, event.Subject, true_values, f"{people=}", duration)
            available_dates -= bdate_set(event.Start.date(), event.End.date())
    return available_dates


def bdate_set(*args, **kwargs):
    """Return a set of business dates returned from pandas.bdate_range."""
    return set(pandas.bdate_range(*args, **kwargs).tolist())


if __name__ == '__main__':
    fill_availability()
