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
from check_leave_dates import get_off_dates
import outlook

# At Christmas, we get guidance for filling in OTLs. This gives hours to book to the leave code for each day
end_of_year_time_off = {
    datetime(2024, 12, 25): 7.4,
    datetime(2024, 12, 26): 7.4,
    datetime(2024, 12, 27): 7.4,
    datetime(2024, 12, 30): 3.7,
    datetime(2024, 12, 31): 0,
    datetime(2025, 1, 1): 7.4
}


def get_project_hours():
    """Fetch the standard hours worked on each project from a spreadsheet."""
    filename = os.path.join(os.environ['UserProfile'], 'STFC', 'Documents',
                            'Group Leader', 'MaRS staff and projects.xlsx')
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
        # get rid of projects with zero hours
        ftes = ftes.dropna().drop(ftes[ftes == 0].index)
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
    name = web.find_element(By.CLASS_NAME, 'x1f').text.split(': ')[1]  # Recent Timecards heading
    return translate_name(name)


def translate_name(name):
    name_translate = {'Alexander': 'Alex'}
    surname, first, *number = name.split(', ')  # e.g. 'Bainbridge, Doctor Alexander Robert, 155807'
    if ' ' in first:
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


def otl_submit(test_mode=False, weeks_in_advance=0):
    """Submit this week's OTL timecard for each staff member."""
    # don't bother before Thursday (to give people time to book the end of the week off)
    if not test_mode:
        if node() == 'DLAST0023':
            print('Not running on laptop')
            return False
        now = datetime.now()
        weekday = now.weekday()
        if weekday < 3:
            print('Too early in the week')
            return now + timedelta(days=3 - weekday)
        # if b'LogonUI.exe' not in check_output('TASKLIST'):  # workstation not locked
        #     print('Workstation not locked')
        #     return False  # for run_tasks, so we know it should retry the task but not report an error
    # get standard booking formula
    all_hours = get_project_hours()

    # get Oracle holiday dates
    all_off_dates = get_staff_leave_dates(test_mode)

    # everyone else's first
    def submit_card(web):
        return submit_staff_timecard(web, all_hours, all_off_dates, weeks_in_advance=weeks_in_advance)

    toast = iterate_staff(('STFC OTL Supervisor',), submit_card, show_window=test_mode)
    # now do mine
    web = go_to_oracle_page(('STFC OTL Timecards', 'Time', 'Recent Timecards'), show_window=test_mode)
    try:
        toast += '\n' + submit_staff_timecard(web, all_hours, weeks_in_advance=weeks_in_advance)
    finally:
        web.quit()
    return toast


def get_all_off_dates(web):
    return get_off_dates(web, fetch_all=True, me=False)


def get_staff_leave_dates(test_mode=False):
    """Get leave dates in Oracle for each staff member."""
    return iterate_staff(('RCUK Self-Service Manager', 'Attendance Management'),
                         get_all_off_dates, show_window=test_mode)


def iterate_staff(page, check_function, show_window=False):
    """Go to a specific page for each staff member and perform a function.
    If the function returns a string, concatenate them all together and return a toast.
    Otherwise, return a dict of {name: value}."""
    web = go_to_oracle_page(page, show_window=show_window)
    print(page)
    try:
        row_count = get_staff_table(web)
        print(f'{row_count=}')

        toast = []
        return_dict = {}
        for i in range(1, row_count):  # first one is mine - ignore this
            # id changes after you've clicked 'Action' link once!
            tag_id = 'N25' if i > 1 and page == ('STFC OTL Supervisor',) else 'N26'
            name = ' '.join(translate_name(web.find_element(By.ID, f'N3:{tag_id}:{i}').text))
            print(name)
            web.find_element(By.ID, f'N3:Y:{i}').click()  # link from 'Action' column at far right
            with contextlib.suppress(selenium.common.exceptions.NoSuchElementException):
                if web.find_element(By.CLASS_NAME, 'x5y').text == 'Error':  # got an error page, not the expected page
                    web.find_element(By.CLASS_NAME, 'x7n').click()  # click 'OK'
                    continue
            result = check_function(web)
            if isinstance(result, str):
                toast.append(result)
            elif result:  # make sure we're not returning a None
                return_dict[name] = result
    finally:
        web.quit()
    return return_dict or '\n'.join(filter(None, toast))


def submit_staff_timecard(web, all_hours, all_absences=None, weeks_in_advance=0):
    """On an individual OTL timecards submitted page, submit a timecard for the current week if necessary.
    Specify weeks_in_advance to do the given number of extra cards after this current week."""

    doing_my_cards = not all_absences
    if doing_my_cards:
        name = 'Ben Shepherd'
        email = 'me'
        all_absences = {name: set()}
    else:
        first, surname = get_name(web)
        name = f'{first} {surname}'
        email = name.replace(' ', '.').lower() + '@stfc.ac.uk'

    # show latest first
    header = web.find_element(By.ID, 'HxcPeriodStarts')
    header.click()  # sort ascending
    time.sleep(2)
    header = web.find_element(By.ID, 'HxcPeriodStarts')  # try to prevent element going stale
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
        outlook_off_dates = outlook.get_away_dates(-30, 90, user=email)
        if outlook_off_dates is False:  # error fetching dates
            print(f'Warning: no away dates fetched. Suggest manual check for {name}')
            outlook_off_dates = set()
        days_away = outlook_off_dates | outlook.get_dl_ral_holidays() | all_absences[name]

    cards_done = 0
    total_days_away = 0
    while do_timecards:  # loop to do all outstanding timecards - break out when done
        first_card_in_list = web.find_elements(By.ID, 'Hxctcarecentlist:Hxctcaperiodstarts:0')
        last_card_date = first_card_in_list[0].text
        print(f'{last_card_date=}')
        weeks = last_card_age(last_card_date)
        if weeks <= -weeks_in_advance:
            print('Up to date')
            break
        web.find_element(By.ID, 'Hxccreatetcbutton').click()  # Create Timecard
        select = web.find_element(By.ID, 'N66' if doing_my_cards else 'N89')
        options = select.find_elements(By.TAG_NAME, 'option')
        # Find the first non-entered one (reverse the order since newer ones are at the top)
        next_option = next(opt for opt in reversed(options) if not opt.text.endswith('~') and ' - ' in opt.text)
        card_date_text = next_option.text
        wb_date = datetime.strptime(card_date_text.split(' - ')[0], '%B %d, %Y').date()  # e.g. August 16, 2021
        # if not doing_my_cards and datetime.now().date() < wb_date != end_of_year_wb:
        #     continue  # don't do future cards for other staff
        next_option.click()
        print('Creating timecard for', card_date_text)
        if doing_my_cards:
            web.find_element(By.ID, 'A150N1display').send_keys('Angal-Kalinin, Doctor Deepa (Deepa)')  # approver
        # list of True/False for on holiday that day
        on_leave = [wb_date + timedelta(days=day) in days_away for day in range(5)]
        print(f'{on_leave=}')

        # enter hours
        if not all(on_leave):
            # Compensate for rounding errors using 'cascade rounding' method
            # https://stackoverflow.com/questions/13483430/how-to-make-rounded-percentages-add-up-to-100#answer-13483486
            # Needed because Oracle only accepts 2 decimal places, and STFC need exactly 7.4 hours per day
            running_total = 0
            total_rounded = 0
            for row, (project_task, daily_hours) in enumerate(hours.items()):
                running_total += daily_hours
                previous_total_rounded = round(running_total, 2)
                rounded_hours = previous_total_rounded - total_rounded
                total_rounded = previous_total_rounded
                project, task = project_task.split(' ')
                web.find_element(By.XPATH, '//button[contains(text(), "Add Another Row")]').click()
                wait_until_page_ready(web)
                hours_boxes = web.find_elements(By.CLASS_NAME, 'x1v')  # 7x6 of these
                boxes = get_boxes(web)
                fill_boxes(boxes, row, project, task)
                for day in range(5):
                    hrs = 0 if on_leave[day] or wb_date + timedelta(days=day) in end_of_year_time_off else rounded_hours
                    hours_boxes[row * 7 + day].send_keys(f'{hrs:.2f}')
                    wait_until_page_ready(web)
                row = len(hours)
        else:
            row = 0
            hours_boxes = web.find_elements(By.CLASS_NAME, 'x1v')  # 7x6 of these
            boxes = get_boxes(web)
        # do a row for leave and holidays
        fill_boxes(boxes, row, 'STRA00009', '01.01')
        for day in range(5):
            hours_boxes[row * 7 + day].send_keys(
                ':.2f'.format(end_of_year_time_off.get(wb_date + timedelta(days=day), 7.4 if on_leave[day] else 0)))


        web.find_element(By.ID, 'review').click()  # Continue button
        web.find_element(By.ID, 'HxcSubmit').click()  # Submit button
        cards_done += 1
        total_days_away += sum(on_leave)
        print('Submitted timecard for', card_date_text)
        time.sleep(2)
        web.find_element(By.LINK_TEXT, 'Return to Recent Timecards').click()
    if not doing_my_cards:
        print('Return to hierarchy')
        web.find_element(By.ID, 'HxcHieReturnButton').click()  # Return to Hierarchy button
    if cards_done > 0:
        return f'{name}: {cards_done=}' + (f', {total_days_away=}' if total_days_away else '')
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
    """Return weeks since supplied Monday, starting on Mondays.
    Returns 0 for this week, 1 for last week, -1 for next week."""
    now = datetime.now()
    this_monday = now - timedelta(days=now.weekday())
    delta = this_monday - datetime.strptime(last_card_date, '%d-%b-%Y')  # e.g. 12-Jul-2021
    return round(delta.days / 7)


if __name__ == '__main__':
    print(get_staff_leave_dates(test_mode=True))
    # print(last_card_age('25-Mar-2024'))
