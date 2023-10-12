import os
import time

import win32com.client as win32
import pandas
from datetime import datetime, timedelta
from subprocess import check_output
from pushbullet import Pushbullet  # to show notifications
from pushbullet_api_key import api_key  # local file, keep secret!
from oracle import go_to_oracle_page, type_into
import outlook


def enter_otl_timecard(show_window=False):
    """Go to the time card entry page for each staff member and fill it in based on their formula."""
    # get standard booking formula
    all_hours = get_project_hours()

    # loop through each staff member
    for name, hours in all_hours.items():
        # find out of office dates in the next 3 months (and previous 1)
        email = name.replace(' ', '.').lower() + '@stfc.ac.uk'
        days_away = outlook.get_away_dates(-30, 90, user=email)

        if name == 'Ben Shepherd':  # have to do my own this way
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
                        on_holiday = [wb_date + timedelta(days=day) in days_away for day in
                                      range(5)]  # list of True/False for on holiday that day
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


def get_project_hours():
    """Fetch the standard hours worked on each project from a spreadsheet."""
    filename = os.path.join(os.environ['UserProfile'], 'Documents', 'Group Leader', 'MaRS staff and projects.xlsx')
    booking_plan = pandas.read_excel(filename, sheet_name='Book', header=0, index_col=[2, 3], skiprows=1)
    print(booking_plan)
    hours = {}
    for name in booking_plan.columns[2:]:  # first two are for project code and name
        if name == 'Total':  # got to the end of the names
            break
        ftes = booking_plan[name]
        ftes = ftes.iloc[:-2]  # remove "not on code" and total rows
        ftes = ftes.dropna()  # get rid of projects with zero hours
        hours[name] = ftes * 7.4  # turn percentages to daily hours
    return hours


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


def otl_submit():
    """Submit this week's OTL timecard for each staff member."""
    # don't bother before Thursday (to give people time to book the end of the week off)
    if datetime.now().weekday() < 3:
        print('Too early in the week')
        return False
    if b'LogonUI.exe' not in check_output('TASKLIST'):  # workstation not locked
        print('Workstation not locked')
        return False  # for run_tasks, so we know it should retry the task but not report an error
    # get standard booking formula
    all_hours = get_project_hours()

    def submit_staff_timecard(return_button_text, web, doing_my_cards=False):
        """On an individual OTL timecards submitted page, submit a timecard for the current week if necessary."""

        # show latest first
        web.click('Period Starting')  # sort ascending
        web.click('Period Starting')  # sort descending
        if doing_my_cards:
            name = 'Ben Shepherd'
            email = 'me'
        else:
            first, surname = get_name(web)
            name = f'{first} {surname}'
            # find out of office dates in the next 3 months (and previous 1)
            email = name.replace(' ', '.').lower() + '@stfc.ac.uk'
        if name not in all_hours.keys():  # e.g. for Hon Sci
            return ''
        hours = all_hours[name]
        days_away = outlook.get_away_dates(-30, 90, user=email)
        web.refresh()
        cards_done = 0
        total_days_away = 0
        while True:  # loop to do all outstanding timecards - break out when done
            last_card_date = web.find_elements(id='Hxctcarecentlist:Hxctcaperiodstarts:0')[0].text
            print(f'{name}: {last_card_date=}')
            if last_card_date == return_button_text:  # some staff (e.g. Hon Sci) don't have timecards
                break
            weeks = last_card_age(last_card_date)
            if weeks <= 0:
                print('Up to date')
                break
            web.click('Create Timecard')
            select = web.find_elements(id='N66' if doing_my_cards else 'N89')[0]
            options = select.find_elements_by_tag_name('option')
            # Find the first non-entered one (reverse the order since newer ones are at the top)
            next_option = next(option for option in reversed(options)
                               if not option.text.endswith('~') and ' - ' in option.text)
            card_date_text = next_option.text
            next_option.click()
            print('Creating timecard for', card_date_text)
            if doing_my_cards:
                web.type('Angal-Kalinin, Doctor Deepa (Deepa)', id='A150N1display')  # approver
            wb_date = datetime.strptime(card_date_text.split(' - ')[0], '%B %d, %Y').date()  # e.g. August 16, 2021
            # list of True/False for on holiday that day
            on_holiday = [wb_date + timedelta(days=day) in days_away for day in range(5)]
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
            cards_done += 1
            total_days_away += sum(on_holiday)
            print('Submitted timecard for', card_date_text)
            web.click('Return to Recent Timecards')
        if cards_done > 0:
            return f'{name}: {cards_done=}' + (f', {total_days_away=}' if total_days_away else '') + '\n'
        else:
            return ''

    # everyone else's first
    toast = check_staff_list('STFC OTL Supervisor', submit_staff_timecard, 'Return to Hierarchy')
    # now do mine
    web = go_to_oracle_page(('STFC OTL Timecards', 'Time', 'Recent Timecards'))
    toast += submit_staff_timecard('Home', web, doing_my_cards=True)
    print(toast)
    Pushbullet(api_key).push_note('👔 Group Timecards', toast)


def check_staff_list(page, check_function, return_button_text, toast_title=''):
    """Go to a specific page for each staff member and perform a function.
    Don't specify a toast title if you don't want a toast displayed."""
    web = go_to_oracle_page(page, show_window=True)
    row_count = get_staff_table(web)

    toast = []
    try:
        for i in range(1, row_count):  # first one is mine - ignore this
            web.click(return_button_text)
            web.click(id=f'N3:Y:{i}')  # link from 'Action' column at far right
            toast.append(check_function(return_button_text, web))
    finally:
        web.driver.quit()
    toast = '\n'.join(filter(None, toast))
    if toast_title and toast:
        print(toast_title)
        print(toast)
        Pushbullet(api_key).push_note(toast_title, toast)
    return toast


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
    if weeks >= 3:
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
    delta = datetime.now() - datetime.strptime(last_card_date, '%d-%b-%Y')  # e.g. 12-Jul-2021
    return delta.days // 7


if __name__ == '__main__':
    otl_submit()
