import win32com.client as win32
from datetime import datetime
from pushbullet import Pushbullet  # to show notifications
from pushbullet_api_key import api_key  # local file, keep secret!
from oracle import go_to_oracle_page

name_translate = {'Alexander': 'Alex'}

web = go_to_oracle_page('STFC OTL Supervisor')

# expand all rows
print('Fetching list of staff')
i = 0
while True:
    i += 1
    table = web.find_elements(classname='x1o', tag='table')[0]
    rows = table.find_elements_by_tag_name('tr')
    if i >= len(rows):
        break
    web.click(id=f'N3:hgi:{i}')

print('Finding recent timecards for each staff member')
toast = ''
outlook = win32.Dispatch('outlook.application')
recent = 'Recent Timecards: '
for i in range(1, len(rows) - 5):  # table has a few extra rows!
    # time.sleep(1)
    web.click(id=f'N3:Y:{i}')
    name = web.find_elements(recent)[0].text
    name = name[len(recent):]
    surname, first, number = name.split(', ')  # e.g. 'Bainbridge, Doctor Alexander Robert, 155807'
    title, first = first.split(' ', 1)
    try:
        first, middles = first.split(' ', 1)
    except ValueError:  # some people have no middle name!
        pass
    first = name_translate[first] if first in name_translate else first
    # show latest first (only need to do once?)
    if i == 1:
        web.click('Period Starting')  # sort ascending
        web.click('Period Starting')  # sort descending

    last_card_date = web.find_elements(id='Hxctcarecentlist:Hxctcaperiodstarts:0')[0].text

    print(f'{first} {surname}: {last_card_date}')
    if last_card_date != 'Return to Hierarchy':  # some staff (e.g. Hon Sci) don't have timecards
        delta = datetime.today() - datetime.strptime(last_card_date, '%d-%b-%Y')  # e.g. 12-Jul-2021
        weeks = delta.days // 7
        if weeks >= 1:
            toast += f'{first} {surname}: last card {last_card_date}\n'
        if weeks >= 2:
            mail = outlook.CreateItem(0)
            mail.To = f'{first}.{surname}@stfc.ac.uk'
            mail.Subject = 'OTL Reminder'
            mail.Body = f"""{first},

This is your automated reminder to bring your OTL timecards up to date. Your last card was dated {last_card_date}, and so you're {weeks} weeks behind.

Thanks very much.

Ben Shepherd (via a bot)"""
            mail.Send()

    web.click('Return to Hierarchy')

web.close_current_tab()

if toast:
    print('\nLate timecards:')
    print(toast)
    Pushbullet(api_key).push_note('👔 Group Timecards', toast)