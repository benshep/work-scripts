import win32com.client as win32
import pandas
from datetime import datetime
from pushbullet import Pushbullet  # to show notifications
from pushbullet_api_key import api_key  # local file, keep secret!
from oracle import go_to_oracle_page


def get_staff_table(web):
    """Return the list of staff from the People in Hierarchy page."""
    # expand all rows
    i = 0
    while True:
        i += 1
        table = web.find_elements(classname='x1o', tag='table')[0]
        rows = table.find_elements_by_tag_name('tr')
        imgs = table.find_elements_by_tag_name('img')
        for img in imgs:
            if img.get_attribute('title') == 'Select to expand':
                img.click()
                break
        else:  # no more expand buttons to click
            break
        if i >= len(rows):
            break
    return len(rows) - 5  # table has a few extra rows!


def get_name(web):
    """Return the first and last name from the recent timecards page."""
    name_translate = {'Alexander': 'Alex'}
    recent = 'Recent Timecards: '
    name = web.find_elements(recent)[0].text
    name = name[len(recent):]
    surname, first, number = name.split(', ')  # e.g. 'Bainbridge, Doctor Alexander Robert, 155807'
    title, first, *middles = first.split(' ')  # middles will be empty if no middle name
    first = name_translate.get(first, first)
    return first, surname


def annual_leave_check():
    """Check remaining annual leave days for each staff member."""
    today = pandas.to_datetime('today')
    last_day = f'{today.year}-12-31'
    # assume full week of public holidays / privilege days at end of December
    days_left_in_year = len(pandas.bdate_range('today', last_day)) - 5
    check_staff_list(['RCUK Self-Service Manager', 'Attendance Management'], check_al_page,
                     'Return to People in Hierarchy', f'🏖️ Group Leave ({days_left_in_year=})')


def otl_check():
    """Check last-submitted OTL timecard for each staff member."""
    check_staff_list('STFC OTL Supervisor', check_otl_page, 'Return to Hierarchy', '👔 Group Timecards')


def check_staff_list(page, check_function, return_button_text, toast_title):
    """Go to a specific page for each staff member and perform a function."""
    web = go_to_oracle_page(page)
    row_count = get_staff_table(web)

    toast = []
    try:
        for i in range(1, row_count):  # first one is mine - ignore this
            web.click(return_button_text)
            web.click(id=f'N3:Y:{i}')  # link from 'Action' column at far right
            toast.append(check_function(return_button_text, web))
    finally:
        web.driver.quit()
    toast = '\n'.join(filter(None, toast))  # remove None - any left?
    if toast:
        print(toast_title)
        print(toast)
        Pushbullet(api_key).push_note(toast_title, toast)


def check_al_page(return_button_text, web):
    """On an individual annual leave balance page, check the remaining balance, and return toast text."""
    if web.exists('The selected action is not available'):  # not available for honorary scientists
        web.go_back()
        return None
    web.click('Entitlement Balances')
    web.click('Show Accrual Balances')
    spans = web.find_elements(tag='span', classname='x2')
    name = spans[0].text
    surname, first = name.split(', ')
    remaining_days = float(spans[-1].text)  # can have half-days too
    # Only show if more than 10 days remaining (max carry over)
    return f'{first} {surname}: {remaining_days:g} days' if remaining_days > 10 else None


def check_otl_page(return_button_text, web):
    """On an individual OTL timecards submitted page,
    check the last date, return toast text and send a reminder if necessary."""

    # show latest first
    web.click('Period Starting')  # sort ascending
    web.click('Period Starting')  # sort descending
    first, surname = get_name(web)
    outlook = win32.Dispatch('outlook.application')
    last_card_date = web.find_elements(id='Hxctcarecentlist:Hxctcaperiodstarts:0')[0].text
    print(f'{first} {surname}: {last_card_date}')
    if last_card_date == return_button_text:  # some staff (e.g. Hon Sci) don't have timecards
        return
    weeks = last_card_age(last_card_date)
    toast = None
    if weeks >= 1:
        toast = f'{first} {surname}: last card {last_card_date}'
    if weeks >= 2:
        send_reminder(outlook, first, surname, last_card_date, weeks)
    return toast


def send_reminder(outlook, first, surname, last_card_date, weeks):
    mail = outlook.CreateItem(0)
    mail.To = f'{first}.{surname}@stfc.ac.uk'
    mail.Subject = 'OTL Reminder'
    mail.Body = f"""{first},

This is your automated reminder to bring your OTL timecards up to date. Your last card was dated {last_card_date}, and so you're {weeks} weeks behind.

Thanks very much.

Ben Shepherd (via a bot)"""
    # mail.Send()
    mail.Display()


def last_card_age(last_card_date):
    delta = datetime.today() - datetime.strptime(last_card_date, '%d-%b-%Y')  # e.g. 12-Jul-2021
    return delta.days // 7


if __name__ == '__main__':
    annual_leave_check()
