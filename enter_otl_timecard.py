import os
import time
import pandas

import outlook
from oracle import type_into, go_to_oracle_page
from datetime import datetime, timedelta
from pushbullet import Pushbullet  # to show notifications
from pushbullet_api_key import api_key  # local file, keep secret!


def enter_otl_timecard(show_window=False):
    # get standard bookings
    hours = get_project_hours_new()

    # find out of office dates in the next 3 months (and previous 1)
    days_away = outlook.get_away_dates(-30, 90)

    web = go_to_oracle_page(('STFC OTL Timecards', 'Time'), show_window=show_window)
    toast = ''
    try:
        while True:
            print('Going to timesheet page')
            web.click('Create Timecard')
            web.refresh()
            select = web.find_elements(id='N66')[0]
            options = select.find_elements_by_tag_name('option')
            for option in options:
                card_label = option.text
                if card_label.endswith('~'):  # already entered
                    continue
                print('Creating timecard for', card_label)
                option.click()
                web.type('Angal-Kalinin, Doctor Deepa (Deepa)', id='A150N1display')  # approver
                wb_date = datetime.strptime(card_label.split(' - ')[0], '%B %d, %Y').date()  # e.g. August 16, 2021
                on_holiday = [wb_date + timedelta(days=day) in days_away for day in range(5)]  # list of True/False for on holiday that day
                print(f'{on_holiday=}')

                # enter hours
                if not all(on_holiday):
                    for row, ((project_task, _), daily_hours) in enumerate(hours.items()):
                        project, task = project_task.split(' ')
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
    finally:
        web.driver.quit()

    if toast:
        print(toast)
        Pushbullet(api_key).push_note('📅 OTL Timecards', toast)


def get_project_hours(fy):
    """Fetch the standard hours worked on each project from a spreadsheet."""
    filename = os.path.join(os.environ['UserProfile'], 'Documents', 'Budgets', f'MaRS Staff Booking {fy - 1}.xlsx')
    booking_plan = pandas.read_excel(filename, header=[3, 4], index_col=0)
    ftes = booking_plan.loc['Ben Shepherd']
    ftes = ftes.iloc[:-1]  # remove total column
    ftes = ftes.dropna()  # get rid of projects with zero hours
    return ftes * 7.4


def get_project_hours_new():
    """Fetch the standard hours worked on each project from a spreadsheet."""
    filename = os.path.join(os.environ['UserProfile'], 'Documents', 'Group Leader', 'MaRS staff and projects.xlsx')
    booking_plan = pandas.read_excel(filename, sheet_name='Book', header=0, index_col=[2, 3], skiprows=1)
    print(booking_plan)
    ftes = booking_plan['Ben']
    ftes = ftes.iloc[:-2]  # remove "not on code" and total rows
    ftes = ftes.dropna()  # get rid of projects with zero hours
    return ftes * 7.4


if __name__ == '__main__':
    enter_otl_timecard(show_window=True)
