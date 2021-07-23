import time
from datetime import datetime
from webbot import Browser
from pushbullet import Pushbullet  # to show notifications
from pushbullet_api_key import api_key  # local file, keep secret!
from oracle_credentials import username, password  # local file, keep secret!

web = Browser(showWindow=False)
# log in to Oracle
web.go_to('https://ebs.ssc.rcuk.ac.uk/OA_HTML/AppsLogin')
web.type(username, 'username')
web.type(password, 'password')
web.click('Login')

time.sleep(2)
# web.refresh()  # sometimes fails to load home page
web.click('STFC OTL Supervisor')

# expand all rows
i = 0
while True:
    i += 1
    table = web.find_elements(classname='x1o', tag='table')[0]
    rows = table.find_elements_by_tag_name('tr')
    if i >= len(rows):
        break
    web.click(id=f'N3:hgi:{i}')

# go to recent timecards for each one
toast = ''
for i in range(1, len(rows) - 5):  # table has a few extra rows!
    # time.sleep(1)
    web.click(id=f'N3:Y:{i}')
    recent = 'Recent Timecards: '
    name = web.find_elements(recent)[0].text
    name = name[len(recent):]
    surname, first, number = name.split(', ')  # e.g. 'Bainbridge, Doctor Alexander Robert, 155807'
    title, first = first.split(' ', 1)
    try:
        first, middles = first.split(' ', 1)
    except ValueError:  # some people have no middle name!
        pass
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
            toast += f'{first} {surname}: last card {last_card_date}'

    web.click('Return to Hierarchy')

web.close_current_tab()

if toast:
    print(toast)
    Pushbullet(api_key).push_note('📅 OTL Timecards', toast)
