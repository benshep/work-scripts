import os
import time
import pandas
from oracle import type_into, go_to_oracle_page
from datetime import datetime, timedelta
from pushbullet import Pushbullet  # to show notifications
from pushbullet_api_key import api_key  # local file, keep secret!
from outlook import get_appointments_in_range

# get standard bookings
fy = pandas.Timestamp(datetime.today()).to_period('Q-MAR').qyear  # e.g. 2022 for FY21/22
filename = os.path.join(os.environ['UserProfile'], 'Documents', 'Budgets', f'MaRS Staff Booking {fy - 1}.xlsx')
booking_plan = pandas.read_excel(filename, header=[3, 4], index_col=0)
ftes = booking_plan.loc['Ben Shepherd']
ftes = ftes.iloc[:-1]  # remove total column
ftes = ftes.dropna()  # get rid of projects with zero hours
hours = ftes * 7.4  # hours per day for each project

# find out of office dates in the next 3 months
days_away = []
for appointment in get_appointments_in_range(0, 90):
    if appointment.AllDayEvent and appointment.BusyStatus == 3:  # out of office
        days = appointment.End - appointment.Start
        days_away += [appointment.Start + timedelta(days=i) for i in range(days.days)]
days_away = [date.replace(tzinfo=None) for date in days_away]  # ignore time zone info

web = go_to_oracle_page(('STFC OTL Timecards', 'Time'))

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
            if pandas.Timestamp(wb_date + timedelta(days=5)).to_period('Q-MAR').qyear != fy:
                raise NotImplementedError('Going into new financial year not supported yet')
            on_holiday = [wb_date + timedelta(days=day) in days_away for day in range(5)]  # list of True/False for on holiday that day

            # enter hours
            if not all(on_holiday):
                for row, ((project, task), daily_hours) in enumerate(hours.items()):
                    web.click('Add Another Row')
                    time.sleep(2)
                    hours_boxes = web.find_elements(classname='x1v')  # 7x6 of these
                    project_boxes = web.find_elements(xpath='//*[@class="x8" and @title="Project"]')
                    task_boxes = web.find_elements(xpath='//*[@class="x8" and @title="Task"]')
                    type_boxes = web.find_elements(xpath='//*[@class="x8" and @title="Type"]')
                    type_into(web, project_boxes[row], project.strip())
                    type_into(web, task_boxes[row], task.strip())
                    type_into(web, type_boxes[row], 'Labour - (Straight Time)')
                    for day in range(5):
                        i = row * 7 + day
                        hrs = '0' if on_holiday[day] else f'{daily_hours:.2f}'
                        type_into(web, hours_boxes[i], hrs)
            else:
                hours = []
                hours_boxes = web.find_elements(classname='x1v')  # 7x6 of these
                project_boxes = web.find_elements(xpath='//*[@class="x8" and @title="Project"]')
                task_boxes = web.find_elements(xpath='//*[@class="x8" and @title="Task"]')
                type_boxes = web.find_elements(xpath='//*[@class="x8" and @title="Type"]')
            if any(on_holiday):
                # do a row for leave and holidays
                row = len(hours)
                type_into(web, project_boxes[row], 'STRA00009')
                type_into(web, task_boxes[row], '01.01')
                type_into(web, type_boxes[row], 'Labour - (Straight Time)')
                [type_into(web, hours_boxes[row * 7 + day], '7.4' if on_holiday[day] else '0') for day in range(5)]

            web.click(id='review')  # Continue button
            web.click(id='HxcSubmit')  # Submit button
            toast += f'Submitted timecard for {card_label}'
            if days_off := sum(on_holiday):
                toast += f' (with {days_off} days off)'
            toast += '\n'
            break
    else:  # no unentered timecards found
        break

if toast:
    print(toast)
    Pushbullet(api_key).push_note('📅 OTL Timecards', toast)