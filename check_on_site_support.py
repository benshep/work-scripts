import os
import pandas
from datetime import datetime
import outlook
from pushbullet import Pushbullet  # to show notifications
from pushbullet_api_key import api_key  # local file, keep secret!


def check_on_site_support():
    # read in data from CLARA support calendar
    user_profile = os.environ['UserProfile']
    spreadsheet_filename = os.path.join(user_profile, r"Science and Technology Facilities Council",
                                        'ASTeC-RTF Safety - General', "CLARA NWH Support Calendar.xlsx")
    sheet = pandas.read_excel(spreadsheet_filename, sheet_name='Magnets')
    # get row labels
    row_labels = [row[0] for row in sheet.values]
    months = [label for label in row_labels if isinstance(label, datetime)]
    spreadsheet_support_days = set()
    for month in months:
        month_text = month.strftime('%b-%y')  # e.g. Oct-21
        # which row do we need to look at to find the date numbers?
        row = next(i for i, label in enumerate(row_labels) if isinstance(label, str) and label.startswith(month_text))
        # now find the column with the first day of the month
        col = list(sheet.values[row]).index(1)
        # find my name in the list
        row = row_labels.index(month)
        for i, initials in enumerate(sheet.values[row][col:]):
            if initials == 'BS':
                spreadsheet_support_days.add(datetime(month.year, month.month, i + 1))
    print('From spreadsheet', spreadsheet_support_days)

    # final month of this calendar is the first cell of the last row
    final_month = months[-1]
    # find the last day of this month
    last_day = final_month + pandas.DateOffset(months=1) + pandas.DateOffset(days=-1)
    # get events from my calendar
    events = outlook.get_appointments_in_range(0, last_day)
    support_subject = '🧲 CLARA magnets on-site support'
    outlook_support_days = {outlook.get_meeting_time(event) for event in events if event.Subject == support_subject}
    wfh_subject = '🏠 Working from home'
    wfh_days = {outlook.get_meeting_time(event) for event in events if event.Subject == wfh_subject}
    print('From Outlook', outlook_support_days)

    # any days marked for me missing from my calendar?
    not_in_outlook = sorted(list(spreadsheet_support_days - outlook_support_days))
    print('Not in Outlook', not_in_outlook)

    for missing_day in not_in_outlook:
        outlook_app = outlook.get_outlook()
        event = outlook_app.CreateItem(1)  # AppointmentItem
        event.Subject = support_subject
        event.Start = missing_day.strftime('%Y-%m-%d')
        event.AllDayEvent = True
        event.Duration = 24 * 60
        event.Save()  # shows as free by default
    toast = ['📅 Added to Outlook: ' + format_date_list(not_in_outlook)] if not_in_outlook else []

    # any days marked in calendar missing from spreadsheet?
    not_in_sheet = sorted(list(outlook_support_days - spreadsheet_support_days))
    print('Not in spreadsheet', not_in_sheet)
    for missing_day in not_in_sheet:
        for event in events:
            if outlook.get_meeting_time(event) == missing_day and event.Subject == support_subject:
                event.Delete()
    if not_in_sheet:
        toast.append('🗑️ Removed from Outlook: ' + format_date_list(not_in_sheet))

    wfh_clashes = sorted(list(wfh_days & spreadsheet_support_days))
    if wfh_clashes:
        toast.append('🏠 WFH clashes: ' + format_date_list(wfh_clashes))
    print('WFH clashes', wfh_clashes)

    if toast:
        print(toast)
        Pushbullet(api_key).push_note(support_subject, '\n'.join(toast))


def format_date_list(date_list):
    """Return a comma-separated formatted string list of the dates in date_list."""
    return ', '.join(day.strftime('%#d %b') for day in date_list)


if __name__ == '__main__':
    check_on_site_support()
