import contextlib
import os
import time

import selenium.common.exceptions
import pandas
from datetime import datetime, timedelta
from subprocess import check_output
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as condition
from selenium.webdriver.common.by import By

from pushbullet import Pushbullet  # to show notifications
from pushbullet_api_key import api_key  # local file, keep secret!
from oracle import go_to_oracle_page
import outlook


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
        # ensure all requests have completed each time
        wait_until_page_ready(web)
        i += 1
        table = web.find_element_by_class_name('x1o')
        rows = table.find_elements_by_tag_name('tr')
        imgs = table.find_elements_by_tag_name('img')
        for img in imgs:
            if img.get_attribute('title') == 'Select to expand':
                img.click()
                time.sleep(0.1)
                break
        else:  # no more expand buttons to click
            break
        if i >= len(rows):
            break
    return len(rows) - 5  # table has a few extra rows!


def wait_until_page_ready(web, timeout=2):
    """Wait until a page has finished loading or a background request is completed."""
    wait = WebDriverWait(web, timeout)
    wait.until(lambda driver: driver.execute_script('return document.readyState') == 'complete')


def get_name(web):
    """Return the first and last name from the recent timecards page."""
    name_translate = {'Alexander': 'Alex'}
    name = web.find_element_by_class_name('x1f').text.split(': ')[1]  # Recent Timecards heading
    surname, first, number = name.split(', ')  # e.g. 'Bainbridge, Doctor Alexander Robert, 155807'
    title, first, *middles = first.split(' ')  # middles will be empty if no middle name
    first = name_translate.get(first, first)
    return first, surname


def annual_leave_check(test_mode=False):
    """Check remaining annual leave days for each staff member."""
    today = pandas.to_datetime('today')
    last_day = f'{today.year}-12-31'
    # assume full week of public holidays / privilege days at end of December
    days_left_in_year = len(pandas.bdate_range('today', last_day)) - 5
    iterate_staff(['RCUK Self-Service Manager', 'Attendance Management'], check_al_page,
                  f'🏖️ Group Leave ({days_left_in_year=})', show_window=test_mode)


def otl_submit(test_mode=False):
    """Submit this week's OTL timecard for each staff member."""
    # don't bother before Thursday (to give people time to book the end of the week off)
    if not test_mode:
        if datetime.now().weekday() < 3:
            print('Too early in the week')
            return False
        if b'LogonUI.exe' not in check_output('TASKLIST'):  # workstation not locked
            print('Workstation not locked')
            return False  # for run_tasks, so we know it should retry the task but not report an error
    # get standard booking formula
    all_hours = get_project_hours()

    # everyone else's first
    def submit_card(web):
        submit_staff_timecard(web, all_hours)

    toast = iterate_staff('STFC OTL Supervisor', submit_card, 'HxcHieReturnButton', show_window=test_mode)
    # now do mine
    web = go_to_oracle_page(('STFC OTL Timecards', 'Time', 'Recent Timecards'), show_window=test_mode)
    toast += submit_staff_timecard(web, all_hours, doing_my_cards=True)
    print(toast)
    Pushbullet(api_key).push_note('👔 Group Timecards', toast)


def iterate_staff(page, check_function, toast_title='', show_window=False):
    """Go to a specific page for each staff member and perform a function.
    Don't specify a toast title if you don't want a toast displayed."""
    web = go_to_oracle_page(page, show_window=show_window)
    row_count = get_staff_table(web)
    print(f'{row_count=}')

    toast = []
    for i in range(1, row_count):  # first one is mine - ignore this
        web.find_element_by_id(f'N3:Y:{i}').click()  # link from 'Action' column at far right
        toast.append(check_function(web))
    web.quit()
    toast = '\n'.join(filter(None, toast))
    if toast_title and toast:
        print(toast_title)
        print(toast)
        Pushbullet(api_key).push_note(toast_title, toast)
    return toast


def submit_staff_timecard(web, all_hours, doing_my_cards=False):
    """On an individual OTL timecards submitted page, submit a timecard for the current week if necessary."""

    # show latest first
    click_when_ready(web, 'HxcPeriodStarts')  # sort ascending
    click_when_ready(web, 'HxcPeriodStarts')  # sort descending
    wait_until_page_ready(web)
    if doing_my_cards:
        name = 'Ben Shepherd'
        email = 'me'
    else:
        first, surname = get_name(web)
        name = f'{first} {surname}'
        # find out of office dates in the next 3 months (and previous 1)
        email = name.replace(' ', '.').lower() + '@stfc.ac.uk'
    do_timecards = name in all_hours.keys()
    if do_timecards:
        hours = all_hours[name]
        days_away = outlook.get_away_dates(-30, 90, user=email)
    cards_done = 0
    total_days_away = 0
    while do_timecards:  # loop to do all outstanding timecards - break out when done
        first_card_in_list = web.find_elements_by_id('Hxctcarecentlist:Hxctcaperiodstarts:0')
        if not first_card_in_list:  # some staff (e.g. Hon Sci) don't have timecards
            print('No timecards in list')
            break
        last_card_date = first_card_in_list[0].text
        print(f'{name}: {last_card_date=}')
        weeks = last_card_age(last_card_date)
        if weeks <= 0:
            print('Up to date')
            break
        web.find_element_by_id('Hxccreatetcbutton').click()  # Create Timecard
        select = web.find_element_by_id('N66' if doing_my_cards else 'N89')
        options = select.find_elements_by_tag_name('option')
        # Find the first non-entered one (reverse the order since newer ones are at the top)
        next_option = next(option for option in reversed(options)
                           if not option.text.endswith('~') and ' - ' in option.text)
        card_date_text = next_option.text
        next_option.click()
        print('Creating timecard for', card_date_text)
        if doing_my_cards:
            web.find_element_by_id('A150N1display').send_keys('Angal-Kalinin, Doctor Deepa (Deepa)')  # approver
        wb_date = datetime.strptime(card_date_text.split(' - ')[0], '%B %d, %Y').date()  # e.g. August 16, 2021
        # list of True/False for on holiday that day
        on_holiday = [wb_date + timedelta(days=day) in days_away for day in range(5)]
        print(f'{on_holiday=}')

        # enter hours
        if not all(on_holiday):
            for row, ((project_task, _), daily_hours) in enumerate(hours.items()):
                project, task = project_task.split(' ')
                web.find_element_by_xpath('//button[contains(text(), "Add Another Row")]').click()
                wait_until_page_ready(web)
                hours_boxes = web.find_elements_by_class_name('x1v')  # 7x6 of these
                boxes = get_boxes(web)
                fill_boxes(boxes, project, task)
                for day in range(5):
                    hrs = '0' if on_holiday[day] else f'{daily_hours:.2f}'
                    hours_boxes[row * 7 + day].send_keys(hrs)
                    wait_until_page_ready(web)
        else:
            hours = []
            hours_boxes = web.find_elements_by_class_name('x1v')  # 7x6 of these
            boxes = get_boxes(web)
        if any(on_holiday):
            # do a row for leave and holidays
            row = len(hours)
            fill_boxes(boxes,'STRA00009', '01.01')
            [hours_boxes[row * 7 + day].send_keys('7.4' if on_holiday[day] else '0') for day in range(5)]

        click_when_ready(web, 'review')  # Continue button
        click_when_ready(web, 'HxcSubmit')  # Submit button
        cards_done += 1
        total_days_away += sum(on_holiday)
        print('Submitted timecard for', card_date_text)
        click_when_ready(web, 'Return to Recent Timecards', By.LINK_TEXT)
    click_when_ready(web, 'HxcHieReturnButton')
    if cards_done > 0:
        return f'{name}: {cards_done=}' + (f', {total_days_away=}' if total_days_away else '') + '\n'
    else:
        return ''


def click_when_ready(web, element_id, by=By.ID):
    """Wait for an element to appear, then click it."""
    WebDriverWait(web, 2).until(condition.presence_of_element_located((by, element_id))).click()


def fill_boxes(boxes, project, task):
    """Fill in project, task, type boxes."""
    boxes['Project'].send_keys(project.strip())
    boxes['Task'].send_keys(task.strip())
    boxes['Type'].send_keys('Labour - (Straight Time)')


def get_boxes(web):
    """Find boxes to fill in."""
    return {title: web.find_elements_by_xpath(f'//*[@class="x8" and @title="{title}"]')
            for title in ('Project', 'Task', 'Type')}


def check_al_page(web):
    """On an individual annual leave balance page, check the remaining balance, and return toast text."""
    if web.find_elements_by_class_name('x5y'):  # error - not available for honorary scientists
        web.back()
        return None
    click_when_ready(web, 'Entitlement Balances', By.LINK_TEXT)
    with contextlib.suppress(selenium.common.exceptions.TimeoutException):  # balances are expanded on the second look
        click_when_ready(web, 'Show Accrual Balances', By.LINK_TEXT)
    name = web.find_element_by_id('EmpName').text
    surname, first = name.split(', ')
    remaining_days = float(web.find_elements_by_class_name('x2')[-1].text)  # can have half-days too
    web.find_element_by_id('Return').click()
    # Only show if more than 10 days remaining (max carry over)
    return f'{first} {surname}: {remaining_days:g} days' if remaining_days > 10 else None


def last_card_age(last_card_date):
    delta = datetime.now() - datetime.strptime(last_card_date, '%d-%b-%Y')  # e.g. 12-Jul-2021
    return delta.days // 7


if __name__ == '__main__':
    otl_submit(test_mode=True)
