import contextlib
import os
import shutil
import tempfile
import time
import pickle
from typing import Any, Callable, TypedDict

import selenium.common.exceptions
import pandas
from datetime import datetime, timedelta, date

from pandas.io.common import file_exists
from selenium.webdriver.firefox.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By

from folders import user_profile
from oracle import go_to_oracle_page
from check_leave_dates import get_off_dates
import outlook

# At Christmas, we get guidance for filling in OTLs. This gives hours to book to the leave code for each day
end_of_year_time_off = {
    date(2024, 12, 25): 7.4,
    date(2024, 12, 26): 7.4,
    date(2024, 12, 27): 7.4,
    date(2024, 12, 30): 3.7,
    date(2024, 12, 31): 0,
    date(2025, 1, 1): 7.4
}

supervisor_page = 'STFC OTL Supervisor'


def get_project_hours() -> tuple[dict[str, pandas.Series], set[pandas.Timestamp]]:
    """Fetch the standard hours worked on each project from a spreadsheet."""
    excel_filename = os.path.join(user_profile, 'STFC', 'Documents', 'Group Leader', 'MaRS staff and projects.xlsx')
    # copy to a temporary file to get around permission errors due to workbook being open
    handle, temp_filename = tempfile.mkstemp(suffix='.xlsx')
    os.close(handle)
    shutil.copy2(excel_filename, temp_filename)
    booking_plan = pandas.read_excel(temp_filename, sheet_name='Book', header=0, index_col=2, skiprows=1)
    os.remove(temp_filename)
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

    return hours, set(booking_plan['Change date'].dropna())  # a list of dates when the booking formula changes


def get_staff_table(web: WebDriver) -> int:
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


def wait_until_page_ready(web: WebDriver, timeout: float = 2.0) -> None:
    """Wait until a page has finished loading or a background request is completed."""
    wait = WebDriverWait(web, timeout)
    wait.until(lambda driver: driver.execute_script('return document.readyState') == 'complete')


def get_name(web: WebDriver) -> tuple[str, str]:
    """Return the first and last name from the recent timecards page."""
    name = web.find_element(By.CLASS_NAME, 'x1f').text.split(': ')[1]  # Recent Timecards heading
    return translate_name(name)


def translate_name(name: str) -> tuple[str, str]:
    name_translate = {'Alexander': 'Alex'}
    surname, first, *number = name.split(', ')  # e.g. 'Bainbridge, Doctor Alexander Robert, 155807'
    if ' ' in first:
        title, first, *middles = first.split(' ')  # middles will be empty if no middle name
    first = name_translate.get(first, first)
    return first, surname


def annual_leave_check(test_mode: bool = False) -> str:
    """Check remaining annual leave days for each staff member."""
    today = pandas.to_datetime('today')
    last_day = f'{today.year}-12-31'
    # assume full week of public holidays / privilege days at end of December
    days_left_in_year = len(pandas.bdate_range('today', last_day)) - 5
    return iterate_staff(check_al_page, 'RCUK Self-Service Manager', 'Attendance Management',
                         show_window=test_mode) + f'\n{days_left_in_year=}'


def otl_submit(test_mode: bool = False,
               weeks_in_advance: int = 0,
               staff_names: list[str] | None = None) -> Any:
    """Submit this week's OTL timecard for each staff member."""
    # don't bother before Thursday (to give people time to book the end of the week off)
    now = datetime.now()
    if not test_mode:
        weekday = now.weekday()
        if weekday < 3:
            print('Too early in the week')
            return now + timedelta(days=3 - weekday)
    # get standard booking formula
    all_hours, change_dates = get_project_hours()

    # get Oracle holiday dates
    off_dates_file = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'off_dates.db')
    if file_exists(off_dates_file) and (now - datetime.fromtimestamp(os.path.getmtime(off_dates_file))).days < 1:
        all_off_dates = pickle.load(open(off_dates_file, 'rb'))
    else:
        all_off_dates = get_staff_leave_dates(test_mode, staff_names=staff_names)
        # noinspection PyTypeChecker
        pickle.dump(all_off_dates, open(off_dates_file, 'wb'))

    # everyone else's first
    def submit_card(web_driver: WebDriver) -> str:
        return submit_staff_timecard(web_driver, all_hours, change_dates,
                                     all_off_dates, weeks_in_advance=weeks_in_advance)

    toast = iterate_staff(submit_card, supervisor_page,
                          show_window=test_mode, staff_names=staff_names)
    # now do mine
    if staff_names is None or 'me' in staff_names:
        web = go_to_oracle_page('STFC OTL Timecards', 'Time', 'Recent Timecards', show_window=test_mode)
        try:
            toast = '\n'.join([toast, submit_staff_timecard(web, all_hours, change_dates,
                                                            weeks_in_advance=weeks_in_advance)]).strip()
        finally:
            if not test_mode:
                web.quit()
    for change_date in change_dates:
        if now <= change_date < now + timedelta(days=7):
            toast = f'Project code change this week\n{toast}'
    return toast


def get_all_off_dates(web: WebDriver) -> set[date]:
    return get_off_dates(web, fetch_all=True, me=False, page_count=1)


def get_staff_leave_dates(test_mode: bool = False,
                          staff_names: list[str] | None = None) -> dict[str, set[date]]:
    """Get leave dates in Oracle for each staff member."""
    return iterate_staff(get_all_off_dates, 'RCUK Self-Service Manager', 'Attendance Management',
                         show_window=test_mode, staff_names=staff_names)


def iterate_staff(check_function: Callable[[WebDriver], Any],
                  *page: str,
                  show_window: bool = False,
                  staff_names: list[str] | None = None) -> None | dict[str, Any] | str:
    """Go to a specific page for each staff member and perform a function.
    If the function returns a string, concatenate them all together and return a toast.
    Otherwise, return a dict of {name: value}.
    Specify a list of staff names to only perform the function for a subset. Otherwise, it will go through them all."""
    web = go_to_oracle_page(*page, show_window=show_window)
    print(page)
    try:
        row_count = get_staff_table(web)
        print(f'{row_count=}')

        toast = []
        return_dict = {}
        tag_id = 'N26'
        for i in range(1, row_count):  # first one is mine - ignore this
            name = ' '.join(translate_name(web.find_element(By.ID, f'N3:{tag_id}:{i}').text))
            if staff_names and name not in staff_names:
                continue
            print(name)
            web.find_element(By.ID, f'N3:Y:{i}').click()  # link from 'Action' column at far right
            if page == (supervisor_page,):
                tag_id = 'N25'  # id changes after you've clicked 'Action' link once!
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
        if not show_window:
            web.quit()
    return return_dict or '\n'.join(filter(None, toast))


def submit_staff_timecard(web: WebDriver,
                          all_hours: dict[str, pandas.Series],
                          change_dates: set[pandas.Timestamp],
                          all_absences: dict[str, set[date]] | None = None,
                          weeks_in_advance: int = 0) -> str:
    """On an individual OTL timecards submitted page, submit a timecard for the current week if necessary.
    Specify weeks_in_advance to do the given number of extra cards after this current week."""
    this_fy = fy(datetime.now())
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
    period_starts = 'HxcPeriodStarts'
    header = web.find_element(By.ID, period_starts)
    header.click()  # sort ascending
    time.sleep(0.5)
    header = web.find_element(By.ID, period_starts)  # try to prevent element going stale
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
        if not outlook_off_dates:  # error fetching dates
            print(f'Warning: no away dates fetched. Suggest manual check for {name}')
        days_away = outlook_off_dates | outlook.get_dl_ral_holidays() | all_absences[name]
        # print(*sorted(list(days_away)), sep='\n')

    cards_done = 0
    failed_cards = 0
    total_days_away = 0
    while do_timecards:  # loop to do all outstanding timecards - break out when done
        first_card_in_list_start = web.find_elements(By.ID, 'Hxctcarecentlist:Hxctcaperiodstarts:0')
        last_card_date = datetime.strptime(first_card_in_list_start[0].text, '%d-%b-%Y').date()
        first_card_in_list_hours = web.find_elements(By.ID, 'Hxctcarecentlist:Hxctcahoursworked:0')
        last_card_hours = float(first_card_in_list_hours[0].text)
        # recorded_hours = [float(el.text) for el in web.find_elements(By.CLASS_NAME, 'x1u x57')]
        # print(recorded_hours)
        print(f'{last_card_date=}, {last_card_hours=}')
        weeks = last_card_age(last_card_date)
        if weeks <= -weeks_in_advance and last_card_hours >= 37:
            print('Up to date')
            break
        if last_card_hours < 37:  # last card <37 hours, click update link in table
            web.find_element(By.ID, 'Hxctcarecentlist:UpdEnable:0').click()
            print('Updating timecard for', last_card_date)
            card_date_text = last_card_date.strftime('%B %d, %Y')
            wb_date = last_card_date
        else:  # new one instead: Create Timecard
            web.find_element(By.ID, 'Hxccreatetcbutton').click()
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

        def date_in_week(day: int) -> date:
            return wb_date + timedelta(days=day)

        # list of True/False for on holiday that day
        on_leave = [date_in_week(day) in days_away for day in range(5)]
        print(f'{on_leave=}')

        # enter hours
        boxes, hours_boxes, empty_rows = get_boxes(web)
        # figure out which boxes to fill: if updating, some will be filled already
        days_to_fill = [day for day in range(5) if
                        hours_boxes[day].get_attribute('value') == '']  # limitation: only checks first row!
        # don't cross over financial years
        days_to_fill = [day for day in days_to_fill if fy(date_in_week(day)) == this_fy]
        # bookings will change on given dates: fill in up to there
        for change_date in change_dates:
            after_change = [date_in_week(day) >= change_date.date() for day in days_to_fill]
            days_to_fill = [day for day, is_after in zip(days_to_fill, after_change) if is_after == after_change[0]]
        print(f'{days_to_fill=}')
        if len(days_to_fill) == 0:  # didn't find any to fill, must be up to date
            print('Up to date')
            break
        if not all(on_leave) and len(hours) > 0:
            # Compensate for rounding errors using 'cascade rounding' method
            # https://stackoverflow.com/questions/13483430/how-to-make-rounded-percentages-add-up-to-100#answer-13483486
            # Needed because Oracle only accepts 2 decimal places, and STFC need exactly 7.4 hours per day
            running_total = 0
            total_rounded = 0
            boxes, hours_boxes, empty_rows = get_boxes(web)
            for project_task, daily_hours in hours.items():
                print(project_task, daily_hours)
                running_total += daily_hours
                previous_total_rounded = round(running_total, 2)
                rounded_hours = previous_total_rounded - total_rounded
                total_rounded = previous_total_rounded
                project, task = project_task.split(' ')

                if not empty_rows:
                    boxes, hours_boxes, empty_rows = get_boxes(web, add_more_first=True)
                wait_until_page_ready(web)
                row_to_fill = empty_rows.pop(0)
                fill_boxes(boxes, row_to_fill, project, task)
                for day in days_to_fill:
                    hrs = 0 if on_leave[day] or date_in_week(day) in end_of_year_time_off else rounded_hours
                    hours_boxes[row_to_fill * 7 + day].send_keys(f'{hrs:.2f}')
                    wait_until_page_ready(web)

        boxes, hours_boxes, empty_rows = get_boxes(web)
        # do a row for leave and holidays
        row_to_fill = empty_rows.pop(0)
        fill_boxes(boxes, row_to_fill, 'STRA00009', '01.01')
        for day in days_to_fill:
            hrs = end_of_year_time_off.get(date_in_week(day), 7.4 if on_leave[day] else 0)
            hours_boxes[row_to_fill * 7 + day].send_keys(f'{hrs:.2f}')

        wait_until_page_ready(web)
        web.find_element(By.ID, 'review').click()  # Continue button
        wait_until_page_ready(web)
        try:
            web.find_element(By.ID, 'HxcSubmit').click()  # Submit button
        except selenium.common.exceptions.NoSuchElementException:
            print('Submit button not present: suspect an error')
            web.find_element(By.ID, 'Hxctccancelbutton').click()  # Cancel button
        else:
            cards_done += 1
            total_days_away += sum(on_leave)
            print('Submitted timecard for', card_date_text)
            time.sleep(0.5)
            web.find_element(By.LINK_TEXT, 'Return to Recent Timecards').click()
    if not doing_my_cards:
        print('Return to hierarchy')
        web.find_element(By.ID, 'HxcHieReturnButton').click()  # Return to Hierarchy button
    toast = []
    if cards_done:
        toast.append(f'{cards_done=}')
        if total_days_away:
            toast.append(f'{total_days_away=}')
    if failed_cards:
        toast.append(f'{failed_cards=}')
    if toast:
        return f'{name}: ' + ', '.join(toast)
    else:
        return ''


def fy(this_date: date) -> int:
    """Return the financial year, given a date."""
    return this_date.year if this_date.month > 3 else (this_date.year - 1)


Boxes = TypedDict('Boxes', {'Project': list[WebElement],
                            'Task': list[WebElement],
                            'Type': list[WebElement]})
"""A dict pointing to the boxes for project, type and task."""


def fill_boxes(boxes: Boxes, row: int, project: str, task: str) -> None:
    """Fill in project, task, type boxes."""
    boxes['Project'][row].send_keys(project.strip())
    boxes['Task'][row].send_keys(task.strip())
    boxes['Type'][row].send_keys('Labour - (Straight Time)')


def get_boxes(web: WebDriver,
              add_more_first: bool = False) -> tuple[Boxes, list[WebElement], list[int]]:
    """Find boxes to fill in."""
    if add_more_first:
        web.find_element(By.XPATH, '//button[contains(text(), "Add Another Row")]').click()

    wait_until_page_ready(web)
    boxes = {title: web.find_elements(By.XPATH, f'//*[@class="x8" and @title="{title}"]')
             for title in ('Project', 'Task', 'Type')}
    # find all the empty rows
    empty_rows = [i for i, project_box in enumerate(boxes['Project']) if
                       project_box.get_attribute('value') == '']
    if not empty_rows:
        return get_boxes(web, add_more_first=True)
    hours_boxes = web.find_elements(By.CLASS_NAME, 'x1v')  # 7x6 of these
    return boxes, hours_boxes, empty_rows


def check_al_page(web: WebDriver) -> str | None:
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


def last_card_age(last_card_date: date) -> int:
    """Return weeks since supplied Monday, starting on Mondays.
    Returns 0 for this week, 1 for last week, -1 for next week."""
    today = datetime.now().date()
    this_monday = today - timedelta(days=today.weekday())
    delta = this_monday - last_card_date  # e.g. 12-Jul-2021
    return round(delta.days / 7)


if __name__ == '__main__':
    print(otl_submit(test_mode=True, weeks_in_advance=0))
    # get_staff_leave_dates(test_mode=False)
    # print(last_card_age('25-Mar-2024'))
