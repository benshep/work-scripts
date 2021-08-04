import os
import pandas
from outlook import get_appointments_in_range
from pushbullet import Pushbullet  # to show notifications
from pushbullet_api_key import api_key  # local file, keep secret!

weeks = 8
d_mon = "%#d %b"  # e.g. 4 Aug

available_dates = set(pandas.bdate_range('today', periods=5*weeks).tolist())
print('Fixed events:')
for event in get_appointments_in_range(0, 7 * weeks):
    recurrence = event.GetRecurrencePattern()
    weekly = event.IsRecurring and recurrence.RecurrenceType == 1 and recurrence.Interval == 1
    busy = event.BusyStatus == 2
    away = event.BusyStatus == 3
    people = len(event.RequiredAttendees.split('; '))
    hours = event.Duration / 60
    days = hours / 24
    if (away or (busy and not weekly)) and (people < 2 or people > 5 or hours >= 2):
        all_day = event.AllDayEvent
        print(event.Subject, event.Start.strftime(d_mon + (' %H:%M' if not all_day else '')),
              ', '.join(key for key, value in locals().items() if value is True),
              f"{people=}", f"{days=}" if all_day else f"{hours=:.1f}")
        available_dates -= set(pandas.bdate_range(event.Start.date(), event.End.date()).tolist())

print('\nLooking for changes')
spreadsheet_xlsx = os.path.join(os.environ['UserProfile'], 'OneDrive - Science and Technology Facilities Council',
                                'CLARA_Shifts_rota', 'Ben.xlsx')
mod_time = pandas.to_datetime(os.path.getmtime(spreadsheet_xlsx), unit='s')
old_end_date = mod_time + pandas.to_timedelta(weeks, 'W')
column_name = 'available'
old_free_dates = pandas.read_excel(spreadsheet_xlsx)[column_name]  # returns a Series
new_data = pandas.DataFrame(sorted(list(available_dates)), columns=[column_name])
today = pandas.to_datetime('today').floor('d')
with pandas.ExcelWriter(spreadsheet_xlsx, mode='a', if_sheet_exists='replace') as writer:
    new_data.to_excel(writer, sheet_name='Ben')
    pandas.DataFrame([today]).to_excel(writer, sheet_name='Updated')

print(f"Old list has {len(old_free_dates)} entries, file modified {mod_time.strftime(d_mon)}")
old_free_dates = old_free_dates[old_free_dates >= today]
print(f'Ignoring dates in old list before today, {today.strftime(d_mon)} - reduced to {len(old_free_dates)} entries')
new_free_dates = new_data[column_name]
print(f"New list has {len(new_free_dates)} entries")
new_free_dates = new_free_dates[new_free_dates < old_end_date]
print(f'Ignoring dates in new list after {old_end_date.strftime(d_mon)} - reduced to {len(new_free_dates)} entries')
old_free_dates = set(old_free_dates)
new_free_dates = set(new_free_dates)
changed_to_yes = new_free_dates - old_free_dates
changed_to_no = old_free_dates - new_free_dates
toast = ''
if changed_to_yes:
    toast += '👍 ' + ', '.join(date.strftime(d_mon) for date in changed_to_yes) + '\n'
if changed_to_no:
    toast += '👎 ' + ', '.join(date.strftime(d_mon) for date in changed_to_no)

if toast:
    print(toast)
    Pushbullet(api_key).push_note('📅 Availability', toast)
