import os
import time

import win32com.client
import pandas
from webbot import Browser
from oracle_credentials import username, password  # local file, keep secret!
from datetime import datetime, timedelta
from pushbullet import Pushbullet  # to show notifications
from pushbullet_api_key import api_key  # local file, keep secret!

user_profile = os.path.join(os.environ['UserProfile'], 'Documents')

# get standard bookings
filename = os.path.join(user_profile, r'Budgets\MaRS Staff Booking 2021.xlsx')
booking_plan = pandas.read_excel(filename, header=[3, 4], index_col=0)
ftes = booking_plan.loc['Ben Shepherd']
ftes = ftes.iloc[:-1]  # remove total column
ftes = ftes.dropna()  # get rid of projects with zero hours
hours = ftes * 7.4  # hours per day for each project

# find out of office dates
today = datetime.today()
in_three_months = today + timedelta(days=90)
appointments = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI').GetDefaultFolder(9).Items
appointments.Sort("[Start]")
appointments.IncludeRecurrences = True
d_m_y = "%#d/%#m/%Y"  # no leading zeros
yyyy_mm_dd = '%Y-%m-%d'
restriction = f"[Start] >= '{today.strftime(d_m_y)}' AND [End] <= '{in_three_months.strftime(d_m_y)}'"
days_away = []
for appointment in appointments.Restrict(restriction):
    if appointment.BusyStatus == 3:  # out of office
        days = appointment.End - appointment.Start
        days_away += [appointment.Start + timedelta(days=i) for i in range(days.days)]
days_away = [date.replace(tzinfo=None) for date in days_away]  # ignore time zone info

web = Browser(showWindow=True)
web.go_to('https://ebs.ssc.rcuk.ac.uk/OA_HTML/AppsLogin')
web.type(username, 'username')
web.type(password, 'password')
web.click('Login')

time.sleep(2)
web.click('STFC OTL Timecards')
web.click('Time')


def type_into(element, text):
    """Type this text into a given element. After finding the elements, it seems faster than webbot's 'type' method."""
    element.click()
    web.press(web.Key.CONTROL + 'a')  # select all
    web.type(text)
    time.sleep(0.5)

toast = ''
while True:
    web.click('Create Timecard')
    web.refresh()
    web.type('Angal-Kalinin, Doctor Deepa (Deepa)', id='A150N1display')  # approver
    select = web.find_elements(id='N66')[0]
    options = select.find_elements_by_tag_name('option')
    for option in options:
        card_label = option.text
        if not card_label.endswith('~'):  # not entered yet
            option.click()
            wb_date = datetime.strptime(card_label.split(' - ')[0], '%B %d, %Y')  # e.g. August 16, 2021

            # enter hours
            for row, ((project, task), daily_hours) in enumerate(hours.items()):
                web.click('Add Another Row')
                time.sleep(2)
                hours_boxes = web.find_elements(classname='x1v')  # 7x6 of these
                project_boxes = web.find_elements(xpath='//*[@class="x8" and @title="Project"]')
                task_boxes = web.find_elements(xpath='//*[@class="x8" and @title="Task"]')
                type_boxes = web.find_elements(xpath='//*[@class="x8" and @title="Type"]')
                type_into(project_boxes[row], project.strip())
                type_into(task_boxes[row], task.strip())
                type_into(type_boxes[row], 'Labour - (Straight Time)')
                for day in range(5):
                    i = row * 7 + day
                    hrs = '0' if wb_date + timedelta(days=day) in days_away else f'{daily_hours:.2f}'
                    type_into(hours_boxes[i], hrs)
            # do a row for leave and holidays
            row = len(hours)
            type_into(project_boxes[row], 'STRA00009')
            type_into(task_boxes[row], '01.01')
            type_into(type_boxes[row], 'Labour - (Straight Time)')
            days_off = 0
            for day in range(5):
                i = row * 7 + day
                if wb_date + timedelta(days=day) in days_away:
                    days_off += 1
                    hrs = '7.4'
                else:
                    hrs = '0'
                type_into(hours_boxes[i], hrs)

            web.click(id='review')  # Continue button
            web.click(id='HxcSubmit')  # Submit button
            toast += f'Submitted timecard for {card_label}'
            if days_off:
                toast += f' (with {days_off} days off)'
            toast += '\n'
            break
    else:  # no unentered timecards found
        break

if toast:
    print(toast)
    Pushbullet(api_key).push_note('📅 OTL Timecards', toast)
