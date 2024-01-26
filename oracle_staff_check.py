import contextlib
import os
import time
from platform import node

import selenium.common.exceptions
import pandas
from datetime import datetime, timedelta
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as condition
from selenium.webdriver.common.by import By

from oracle import go_to_oracle_page
import outlook

end_of_year_wb = None  # datetime.date(2023, 12, 25)  # special case - for entering Christmas timecards early


def get_project_hours():
    """Fetch the standard hours worked on each project from a spreadsheet."""
    filename = os.path.join(os.environ['UserProfile'], 'Documents', 'Group Leader', 'MaRS staff and projects.xlsx')
    booking_plan = pandas.read_excel(filename, sheet_name='Book', header=0, index_col=2, skiprows=1)
    booking_plan = booking_plan[booking_plan.index.notnull()]  # remove any rows without a project code
    print(booking_plan)
    hours = {}
    assert all(booking_plan.columns[:3] == ['Unnamed: 0', 'Official name', 'Project name'])
    for name in booking_plan.columns[3:]:  # first two are for project code and name
        if name == 'Total':  # got to the end of the names
            break
        ftes = booking_plan[name]
        # ftes = ftes.iloc[:-2]  # remove "not on code" and total rows
        ftes = ftes.dropna()  # get rid of projects with zero hours
        total_effort = round(sum(ftes), 2)
        if total_effort != 1.0:
            raise ValueError(f"{name}'s total FTE adds up to {total_effort:.2f} - should be 1.00")
        daily_hours = ftes * 7.4  # turn percentages to daily hours
        hours[name] = daily_hours  # turn percentages to daily hours

    return hours


def get_staff_table(web):
    """Return the list of staff from the People in Hierarchy page."""
    # expand all rows
    i = 0
    while True:
        # ensure all requests have completed each time
        wait_until_page_ready(web)
        i += 1
        table = web.find_element(By.CLASS_NAME, 'x1o')
        rows = table.find_elements(By.TAG_NAME, 'tr')
        imgs = table.find_elements(By.TAG_NAME, 'img')
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
    name = web.find_element(By.CLASS_NAME, 'x1f').text.split(': ')[1]  # Recent Timecards heading
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
    return iterate_staff(('RCUK Self-Service Manager', 'Attendance Management'), check_al_page,
                         show_window=test_mode) + f'\n{days_left_in_year=}'


def otl_submit(test_mode=False):
    """Submit this week's OTL timecard for each staff member."""
    # don't bother before Thursday (to give people time to book the end of the week off)
    if not test_mode:
        if node() == 'DLAST0023':
            print('Not running on laptop')
            return False
        if datetime.now().weekday() < 3:
            print('Too early in the week')
            return False
        # if b'LogonUI.exe' not in check_output('TASKLIST'):  # workstation not locked
        #     print('Workstation not locked')
        #     return False  # for run_tasks, so we know it should retry the task but not report an error
    # get standard booking formula
    all_hours = get_project_hours()

    # everyone else's first
    def submit_card(web):
        submit_staff_timecard(web, all_hours)

    toast = iterate_staff(('STFC OTL Supervisor',), submit_card, show_window=test_mode)
    # now do mine
    web = go_to_oracle_page(('STFC OTL Timecards', 'Time', 'Recent Timecards'), show_window=test_mode)
    toast += submit_staff_timecard(web, all_hours, doing_my_cards=True)
    web.quit()
    return toast


def iterate_staff(page, check_function, toast_title='', show_window=False):
    """Go to a specific page for each staff member and perform a function.
    Don't specify a toast title if you don't want a toast displayed."""
    web = go_to_oracle_page(page, show_window=show_window)
    row_count = get_staff_table(web)
    print(f'{row_count=}')

    toast = []
    for i in range(1, row_count):  # first one is mine - ignore this
        web.find_element(By.ID, f'N3:Y:{i}').click()  # link from 'Action' column at far right
        toast.append(check_function(web))
    web.quit()
    return '\n'.join(filter(None, toast))


def submit_staff_timecard(web, all_hours, doing_my_cards=False):
    """On an individual OTL timecards submitted page, submit a timecard for the current week if necessary."""

    if doing_my_cards:
        name = 'Ben Shepherd'
        email = 'me'
    else:
        first, surname = get_name(web)
        name = f'{first} {surname}'
        email = name.replace(' ', '.').lower() + '@stfc.ac.uk'
    print(name)

    # show latest first
    header_id = 'HxcPeriodStarts'
    web.find_element(By.ID, header_id).click()  # sort ascending
    time.sleep(2)
    header = web.find_element(By.ID, header_id)
    try:
        sort_arrow = header.find_element(By.TAG_NAME, 'img')
        if sort_arrow.get_attribute('title') != 'Sorted in descending order':
            header.click()  # sort descending
        do_timecards = name in all_hours.keys()
    except selenium.common.exceptions.NoSuchElementException:
        # some staff (e.g. Hon Sci) don't have timecards, so no sort up/down button
        print('No timecards in list')
        do_timecards = False

    wait_until_page_ready(web)
    if do_timecards:
        hours = all_hours[name]
        # find out of office dates in the next 3 months (and previous 1)
        days_away = outlook.get_away_dates(-30, 90, user=email)
        if days_away is False:  # error fetching dates
            print(f'Warning: no away dates fetched. Suggest manual check for {name}')

    cards_done = 0
    total_days_away = 0
    while do_timecards:  # loop to do all outstanding timecards - break out when done
        first_card_in_list = web.find_elements(By.ID, 'Hxctcarecentlist:Hxctcaperiodstarts:0')
        last_card_date = first_card_in_list[0].text
        print(f'{last_card_date=}')
        weeks = last_card_age(last_card_date)
        if weeks <= 0 and not end_of_year_wb:
            print('Up to date')
            break
        web.find_element(By.ID, 'Hxccreatetcbutton').click()  # Create Timecard
        select = web.find_element(By.ID, 'N66' if doing_my_cards else 'N89')
        options = select.find_elements(By.TAG_NAME, 'option')
        # Find the first non-entered one (reverse the order since newer ones are at the top)
        next_option = next(opt for opt in reversed(options) if not opt.text.endswith('~') and ' - ' in opt.text)
        card_date_text = next_option.text
        wb_date = datetime.strptime(card_date_text.split(' - ')[0], '%B %d, %Y').date()  # e.g. August 16, 2021
        if not doing_my_cards and datetime.now().date() < wb_date != end_of_year_wb:
            continue  # don't do future cards for other staff
        next_option.click()
        print('Creating timecard for', card_date_text)
        if doing_my_cards:
            web.find_element(By.ID, 'A150N1display').send_keys('Angal-Kalinin, Doctor Deepa (Deepa)')  # approver
        # list of True/False for on holiday that day
        on_holiday = [wb_date + timedelta(days=day) in days_away for day in range(5)]
        print(f'{on_holiday=}')

        # enter hours
        if not all(on_holiday) and end_of_year_wb != wb_date:
            for row, (project_task, daily_hours) in enumerate(hours.items()):
                project, task = project_task.split(' ')
                web.find_element(By.XPATH, '//button[contains(text(), "Add Another Row")]').click()
                wait_until_page_ready(web)
                hours_boxes = web.find_elements(By.CLASS_NAME, 'x1v')  # 7x6 of these
                boxes = get_boxes(web)
                fill_boxes(boxes, row, project, task)
                for day in range(5):
                    hrs = '0' if on_holiday[day] else f'{daily_hours:.2f}'
                    hours_boxes[row * 7 + day].send_keys(hrs)
                    wait_until_page_ready(web)
        else:
            hours = []
            hours_boxes = web.find_elements(By.CLASS_NAME, 'x1v')  # 7x6 of these
            boxes = get_boxes(web)
        row = len(hours)
        if end_of_year_wb == wb_date:  # end of year is a bit special
            fill_boxes(boxes, row, 'STRA00009', '01.01')
            [hours_boxes[row * 7 + i].send_keys(str(duration)) for i, duration in enumerate([7.4, 7.4, 7.4, 3.7, 0])]
        elif any(on_holiday):
            # do a row for leave and holidays
            fill_boxes(boxes, row, 'STRA00009', '01.01')
            [hours_boxes[row * 7 + day].send_keys('7.4' if on_holiday[day] else '0') for day in range(5)]

        web.find_element(By.ID, 'review').click()  # Continue button
        web.find_element(By.ID, 'HxcSubmit').click()  # Submit button
        cards_done += 1
        total_days_away += sum(on_holiday)
        print('Submitted timecard for', card_date_text)
        web.find_element(By.LINK_TEXT, 'Return to Recent Timecards').click()
    if not doing_my_cards:
        web.find_element(By.ID, 'HxcHieReturnButton').click()
    if cards_done > 0:
        return f'{name}: {cards_done=}' + (f', {total_days_away=}' if total_days_away else '') + '\n'
    else:
        return ''


def fill_boxes(boxes, row, project, task):
    """Fill in project, task, type boxes."""
    boxes['Project'][row].send_keys(project.strip())
    boxes['Task'][row].send_keys(task.strip())
    boxes['Type'][row].send_keys('Labour - (Straight Time)')


def get_boxes(web):
    """Find boxes to fill in."""
    return {title: web.find_elements(By.XPATH, f'//*[@class="x8" and @title="{title}"]')
            for title in ('Project', 'Task', 'Type')}


def check_al_page(web):
    """On an individual annual leave balance page, check the remaining balance, and return toast text."""
    if web.find_elements(By.CLASS_NAME, 'x5y'):  # error - not available for honorary scientists
        web.back()
        return None
    web.find_element(By.LINK_TEXT, 'Entitlement Balances').click()
    # balances are expanded on the second look, so the 'Show' button won't be present
    with contextlib.suppress(selenium.common.exceptions.NoSuchElementException):
        web.find_element(By.LINK_TEXT, 'Show Accrual Balances').click()
    name = web.find_element(By.ID, 'EmpName').text
    surname, first = name.split(', ')
    remaining_days = float(web.find_elements(By.CLASS_NAME, 'x2')[-1].text)  # can have half-days too
    web.find_element(By.ID, 'Return').click()
    # Only show if more than 10 days remaining (max carry over)
    return f'{first} {surname}: {remaining_days:g} days' if remaining_days > 10 else None


def last_card_age(last_card_date):
    delta = datetime.now() - datetime.strptime(last_card_date, '%d-%b-%Y')  # e.g. 12-Jul-2021
    return delta.days // 7


if __name__ == '__main__':
    print(annual_leave_check(test_mode=True))
