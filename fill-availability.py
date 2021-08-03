import os
import pandas
from datetime import datetime, timedelta
from outlook import get_appointments_in_range

today = datetime.today()
available_dates = set(pandas.bdate_range(today, today + timedelta(days=60)).tolist())
for appt in get_appointments_in_range(0, 60):
    # print(appt.Subject, appt.IsRecurring, appt.GetRecurrencePattern().RecurrenceType)
    recurrence = appt.GetRecurrencePattern()
    weekly = appt.IsRecurring and recurrence.RecurrenceType == 1 and recurrence.Interval == 1
    busy = appt.BusyStatus == 2
    oof = appt.BusyStatus == 3
    people = len(appt.RequiredAttendees.split('; '))
    hours = appt.Duration / 60
    fixed = (oof or (busy and not weekly)) and (people < 2 or people > 4 or hours >= 2)
    if fixed:
        print(appt.Subject, appt.Start, busy, oof, weekly, people, hours, fixed)
        # remove dates in range (Start, End) from available_dates
        available_dates -= set(pandas.bdate_range(appt.Start.date(), appt.End.date()).tolist())

df = pandas.DataFrame(available_dates)
spreadsheet_xlsx = os.path.join(os.environ['UserProfile'], 'OneDrive - Science and Technology Facilities Council',
                                'CLARA_Shifts_rota', 'Ben.xlsx')
with pandas.ExcelWriter(spreadsheet_xlsx, mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name='Ben')
