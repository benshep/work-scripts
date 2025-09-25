from datetime import datetime, date
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.webdriver import WebDriver
from selenium.webdriver.support.ui import WebDriverWait

import otl
from oracle import go_to_oracle_page
import outlook


def list_missing(date_set : set[date]) -> str:
    """Convert a set of missing dates to a comma-separated string."""
    return ', '.join(date.strftime("%#d %b") for date in sorted(list(date_set)))  # e.g. 4 Aug

# def nice_range_text():
#     from datetime import datetime
#     import itertools
#
#     dates = [datetime.strptime(s + ' 2023', "%d %b %Y") for s in '30 May, 31 May, 1 Jun, 2 Jun, 7 Jun'.split(', ')]
#
#     # group consecutive elements using itertools.groupby()
#     groups = []
#     for k, g in itertools.groupby(enumerate(dates), lambda x: x[0] - x[1].toordinal()):
#         groups.append(list(map(lambda x: x[1], g)))
#
#     # create a list of tuples with the first and last element of each group
#     def range_text(date_range):
#         first = date_range[0]
#         last = date_range[-1]
#         if first.year == last.year:
#             if first.month == last.month:
#                 if first.day == last.day:
#                     return first.strftime("%#d %b")
#                 return first.strftime('%#d') + '-' + last.strftime("%#d %b")
#             return first.strftime('%#d %b') + ' - ' + last.strftime("%#d %b")
#         return first.strftime('%#d %b %Y') + ' - ' + last.strftime("%#d %b %Y")
#
#     res = [range_text(group) for group in groups]
#
#     # print result
#     print("Grouped list is : " + str(res))


def check_leave_dates(test_mode: bool = False) -> str:
    """Compare annual leave dates in an Outlook calendar and in the list submitted to Oracle."""
    oracle_off_dates = get_oracle_off_dates(test_mode=test_mode).keys()  # only want date!
    outlook_off_dates = outlook.get_away_dates(min(oracle_off_dates), 0, look_for=outlook.is_annual_leave)

    if not_in_outlook := oracle_off_dates - outlook_off_dates:
        toast = f'Missing from Outlook: {list_missing(not_in_outlook)}\n'
    else:
        toast = ''
    if not_in_oracle := outlook_off_dates - oracle_off_dates:
        toast += f'Missing from Oracle: {list_missing(not_in_oracle)}'
    return toast


def get_oracle_off_dates(page_count: int = 1, test_mode: bool = False,
                         table_format: bool = False) -> dict[date, float] | list[list]:
    web = go_to_oracle_page('absences', show_window=test_mode)
    try:
        off_dates = get_off_dates(web, page_count=page_count, table_format=table_format)
    finally:
        web.quit()
    return off_dates


def wait_until_page_ready(web: WebDriver, timeout: float = 2.0) -> None:
    """Wait until a page has finished loading or a background request is completed."""
    wait = WebDriverWait(web, timeout)
    wait.until(lambda driver: driver.execute_script('return document.readyState') == 'complete')


def get_off_dates(web: WebDriver,
                  fetch_all: bool = False,
                  page_count: int = 1,
                  table_format: bool = False) -> dict[date, float] | list[list]:
    """Get absence dates from an Oracle 'Attendance Management' page.
    If fetch_all is True, looks for all absences, otherwise just ones with "Leave" in the "Absence Type".
    Fetches one page by default: specify page_count for more (or set to 0 for all).
    Returns a set of dates by default; if table_format is True, returns a list of lists as shown on the page."""
    sleep(10)
    pages_done = 0
    return_value = [] if table_format else {}
    while pages_done < page_count or page_count < 1:  # Prev and Next buttons TODO, but for now just one page!
        cells = web.find_elements(By.XPATH, '//div[@data-oj-tabmod="0"]')
        while True:
            # returns empty array [] if no absences, but array of blank text if absences are present
            # need to wait for them to be filled in!
            cells_text = [cell.text for cell in cells]
            # e.g. ['Annual Leave', '', '05/09/2025 - 05/09/2025', '7.4 Hours', 'Completed', ... ]
            print(cells_text)
            if len(cells_text) == 0 or cells_text.count('') / len(cells_text) <= 0.2:  # normal: 1 in 5 is blank
                break
        columns = 5
        absence_types = cells_text[::columns]
        if not absence_types:
            break
        date_ranges = cells_text[2::columns]
        all_hours = cells_text[3::columns]
        approvals = cells_text[4::columns]
        for absence_type, dates, hours, approval in zip(absence_types, date_ranges, all_hours, approvals):
            if (not fetch_all and 'Leave' not in absence_type) or approval == 'Withdrawn':
                continue
            start_text, end_text = dates.split(' - ', maxsplit=1)
            start_date = from_dmy(start_text)
            end_date = from_dmy(end_text)
            date_list = outlook.get_date_list(start_date, end_date)
            hours = float(hours.split(' ')[0])  # e.g. 7.4 Hours
            if table_format:
                return_value.append([start_date, end_date, absence_type, hours, approval])
            else:
                # Assumption: listed hours are spread equally over listed days
                # with any 'left-over' hours going into the last day
                # e.g. 8.4 hours over 2 days would be 7.4 hours on day 1, 1 hour on day 2
                for day in date_list:
                    hours_this_day = min(otl.hours_per_day, hours)
                    return_value |= {day: hours_this_day}
                    hours -= hours_this_day
        # nav_links = web.find_elements(By.CLASS_NAME, 'x48')  # Prev and Next buttons TODO, but for now just one page!
        # for link in nav_links:
        #     if link.text.startswith('Next'):
        #         link.click()  # go to next page
        #         sleep(1)
        #         pages_done += 1
        #         break
        # else:
        break  # out of while loop: no more pages
    # if not me:  # click 'Return to People in Hierarchy'
    #     web.find_element(By.ID, 'Return').click()
    if not table_format:
        print(len(return_value), 'absences, latest:', max(return_value) if return_value else 'N/A')
    return return_value


def from_dmy(text: str) -> date:
    """Convert a date in the format 31/07/2022 to a datetime."""
    return datetime.strptime(text, '%d/%m/%Y').date()


if __name__ == '__main__':
    print(get_oracle_off_dates(table_format=True, test_mode=True))
    # web = go_to_oracle_page('absences', show_window=True)
    # for _ in range(300):
    #     cells = web.find_elements(By.CLASS_NAME, 'oj-typography-body-md')
    #     print([cell.text for cell in cells])
    #     sleep(0.1)
