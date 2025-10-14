import os
import pandas
from datetime import datetime
import outlook
from folders import user_profile
from tools import read_excel


def check_on_site_support():
    sheet = get_spreadsheet_data()
    me = 'BS'
    # get row labels
    row_labels = [row[0] for row in sheet.values]
    months = sorted({label for label in row_labels if isinstance(label, datetime)})  # remove duplicates
    spreadsheet_support_days = {me: set(), 'AB': set(), 'AH': set()}
    for month in months:
        # which row do we need to look at to find the date numbers?
        row = row_labels.index(month)
        # now find the column with the first day of the month
        col = list(sheet.values[row]).index(1)
        # further down, the months are repeated with rows for initials - find the corresponding one
        row = row_labels.index(month, row + 1)
        for i, initials in enumerate(sheet.values[row][col:]):
            if initials in spreadsheet_support_days and (date := datetime(month.year, month.month, i + 1)) >= datetime.now():
                spreadsheet_support_days[initials].add(date)
    print('From spreadsheet')
    for initials, support_days in spreadsheet_support_days.items():
        print(initials, format_date_list(support_days))

    # final month of this calendar is the first cell of the last row
    final_month = months[-1]
    # find the last day of this month
    last_day = final_month + pandas.DateOffset(months=1) + pandas.DateOffset(days=-1)
    # get events from my calendar
    events = outlook.get_appointments_in_range(0, last_day)
    support_subject = '🧲 CLARA magnets on-site support'
    outlook_support_days = {outlook.get_meeting_time(event) for event in events if event.Subject == support_subject}
    print('From Outlook', format_date_list(outlook_support_days))

    # any days marked for me missing from my calendar?
    not_in_outlook = sorted(spreadsheet_support_days[me] - outlook_support_days)
    create_support_events(not_in_outlook, support_subject)
    toast = ['📅 Added to Outlook: ' + format_date_list(not_in_outlook)] if not_in_outlook else []

    # any days marked in calendar missing from spreadsheet?
    not_in_sheet = outlook_support_days - spreadsheet_support_days[me]
    for missing_day in not_in_sheet:
        for event in events:
            if outlook.get_meeting_time(event) == missing_day and event.Subject == support_subject:
                event.Delete()
    if not_in_sheet:
        toast.append('🗑️ Removed from Outlook: ' + format_date_list(not_in_sheet))

    user_dict = {me: 'me', 'AB': 'alex.bainbridge@stfc.ac.uk', 'AH': 'alex.hinton@stfc.ac.uk'}
    for initials, support_days in spreadsheet_support_days.items():
        user = user_dict[initials]
        wfh_days = outlook.get_away_dates(0, last_day, user=user, look_for=outlook.is_wfh)
        leave_days = outlook.get_away_dates(0, last_day, user=user)
        off_days = wfh_days | leave_days
        print(f'{initials} off days', format_date_list(off_days))
        if clashes := off_days & support_days:
            toast.append(f'🏠 {initials} out of office clashes: ' + format_date_list(clashes))

    return toast


def create_support_events(not_in_outlook, support_subject):
    """Go through a list of missing days and create a support event for each one in Outlook."""
    if not not_in_outlook:
        return
    outlook_app = outlook.get_outlook()
    for missing_day in not_in_outlook:
        event = outlook_app.CreateItem(1)  # AppointmentItem
        event.Subject = support_subject
        event.Start = missing_day.strftime('%Y-%m-%d')
        event.AllDayEvent = True
        event.Duration = 24 * 60
        event.Save()  # shows as free by default


def get_spreadsheet_data():
    """Read in data from CLARA support calendar."""
    spreadsheet_filename = os.path.join(user_profile, r"Science and Technology Facilities Council",
                                        'ASTeC-RTF Safety - General', "CLARA NWH Support Calendar.xlsx")
    return read_excel(spreadsheet_filename, sheet_name='Magnets')


def format_date_list(date_list):
    """Return a comma-separated formatted string list of the dates in date_list."""
    return ', '.join(day.strftime('%#d %b') for day in sorted(date_list))


if __name__ == '__main__':
    check_on_site_support()
