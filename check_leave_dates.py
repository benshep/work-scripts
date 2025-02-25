from datetime import datetime
from time import sleep

from selenium.webdriver.common.by import By
from oracle import go_to_oracle_page
import outlook


def list_missing(date_set):
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


def check_leave_dates():
    """Compare annual leave dates in an Outlook calendar and in the list submitted to Oracle."""
    oracle_off_dates = get_oracle_off_dates()
    outlook_off_dates = outlook.get_away_dates(min(oracle_off_dates), 0, look_for=outlook.is_annual_leave)

    if not_in_outlook := oracle_off_dates - outlook_off_dates:
        toast = f'Missing from Outlook: {list_missing(not_in_outlook)}\n'
    else:
        toast = ''
    if not_in_oracle := outlook_off_dates - oracle_off_dates:
        toast += f'Missing from Oracle: {list_missing(not_in_oracle)}'
    return toast


def get_oracle_off_dates(page_count=1):
    web = go_to_oracle_page(('RCUK Self-Service Employee', 'Attendance Management'))
    try:
        off_dates = get_off_dates(web, page_count=page_count)
    finally:
        web.quit()
    return off_dates


def get_off_dates(web, fetch_all=False, me=True, page_count=1):
    """Get absence dates from an Oracle 'Attendance Management' page.
    If fetch_all is True, looks for all absences, otherwise just ones with "Leave" in the "Absence Type".
    Fetches one page by default: specify page_count for more (or set to 0 for all)."""
    pages_done = 0
    return_value = set()
    while pages_done < page_count or page_count < 1:
        cells = web.find_elements(By.CLASS_NAME, 'x1w')
        columns = 8 if me else 9  # staff under me get a 'delete' column too
        start_dates = cells[::columns]
        if start_dates[0].text == 'No results found.':
            break
        end_dates = cells[1::columns]
        absence_types = cells[2::columns]
        day_counts = cells[3::columns]
        for start, end, day_count, absence_type_cell in zip(start_dates, end_dates, day_counts, absence_types):
            absence_type = absence_type_cell.text
            if fetch_all or 'Leave' in absence_type:
                date_list = outlook.get_date_list(from_dmy(start.text), count=round(float(day_count.text)))
                # for date in date_list:
                #     print(date, absence_type, sep='\t')
                # end_date = from_dmy(end.text)
                # if date_list[-1] != end_date:
                #     print('\t\tlisted end date', end_date)
                return_value |= set(date_list)
        nav_links = web.find_elements(By.CLASS_NAME, 'x48')  # Prev and Next buttons
        for link in nav_links:
            if link.text.startswith('Next'):
                link.click()  # go to next page
                sleep(1)
                pages_done += 1
                break
        else:
            break  # out of while loop: no more pages
    if not me:  # click 'Return to People in Hierarchy'
        web.find_element(By.ID, 'Return').click()
    print(len(return_value), 'absences, latest:', max(return_value) if return_value else 'N/A')
    return return_value


def from_dmy(text):
    """Convert a date in the format 31-Oct-2022 to a datetime."""
    return datetime.strptime(text, '%d-%b-%Y').date()


if __name__ == '__main__':
    get_oracle_off_dates(page_count=0)
