import win32com.client as win32
from datetime import datetime
from pushbullet import Pushbullet  # to show notifications
from pushbullet_api_key import api_key  # local file, keep secret!
from oracle import go_to_oracle_page


def get_staff_table(web):
    """Return the list of staff from the OTL Supervisor page."""
    # expand all rows
    i = 0
    while True:
        i += 1
        tables = web.find_elements(classname='x1o', tag='table')
        rows = tables[0].find_elements_by_tag_name('tr')
        if i >= len(rows):
            break
        web.click(id=f'N3:hgi:{i}')
    return len(rows) - 5  # table has a few extra rows!


def get_name(web):
    """Return the first and last name from the recent timecards page."""
    name_translate = {'Alexander': 'Alex'}
    recent = 'Recent Timecards: '
    name = web.find_elements(recent)[0].text
    name = name[len(recent):]
    surname, first, number = name.split(', ')  # e.g. 'Bainbridge, Doctor Alexander Robert, 155807'
    title, first, *middles = first.split(' ')  # middles will be empty if no middle name
    first = name_translate[first] if first in name_translate else first
    return first, surname


def oracle_otl_check():
    web = go_to_oracle_page('STFC OTL Supervisor')

    print('Fetching list of staff')
    row_count = get_staff_table(web)

    print('Finding recent timecards for each staff member')
    toast = ''
    outlook = win32.Dispatch('outlook.application')
    return_button_text = 'Return to Hierarchy'
    for i in range(1, row_count):  # first one is mine - ignore this
        web.click(return_button_text)
        # time.sleep(1)
        web.click(id=f'N3:Y:{i}')  # link from 'Action' column at far right
        # show latest first (only need to do once)
        if i == 1:
            web.click('Period Starting')  # sort ascending
            web.click('Period Starting')  # sort descending
        first, surname = get_name(web)
        last_card_date = web.find_elements(id='Hxctcarecentlist:Hxctcaperiodstarts:0')[0].text
        print(f'{first} {surname}: {last_card_date}')
        if last_card_date == return_button_text:  # some staff (e.g. Hon Sci) don't have timecards
            continue
        weeks = last_card_age(last_card_date)
        if weeks >= 1:
            toast += f'{first} {surname}: last card {last_card_date}\n'
        if weeks >= 2:
            send_reminder(outlook, first, surname, last_card_date, weeks)
    web.close_current_tab()
    if toast:
        print('\nLate timecards:')
        print(toast)
        Pushbullet(api_key).push_note('👔 Group Timecards', toast)


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
    oracle_otl_check()
