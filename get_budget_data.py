import os
import pandas
import selenium.common
import polling2
from datetime import datetime
from time import sleep

from oracle import go_to_oracle_page


class InformationFetchFailure(Exception):
    pass


class DataExportFailure(Exception):
    pass


def poll(target, arg=None, ignore_exceptions=(selenium.common.exceptions.NoSuchElementException, ),
         check_success=polling2.is_truthy, **kwargs):
    """Poll with some default values and a single argument."""
    print(f'Polling {target.__name__} for {arg}' + (f', {kwargs}' if kwargs else ''))
    return polling2.poll(target, args=(arg, ), kwargs=kwargs, step=1, max_tries=50, check_success=check_success,
                         ignore_exceptions=ignore_exceptions)


def search_via_prompt_using_poll(web, index, search_term, check_in_dropdown=False):
    """Search using the dropdown box and the pop-up prompt."""
    print(f'Searching for {search_term} in box {index}')
    # Sometimes webbot takes ages to find the elements - base Selenium is much quicker
    # try this a few times - sometimes the dropdown disappears
    for i in range(10):
        dropdowns = poll(web.driver.find_elements_by_class_name, 'promptTextField')
        print(f'Clicking dropdown, attempt {i}')
        dropdowns[index].click()
        try:
            search_buttons = poll(web.driver.find_elements_by_class_name, 'DropDownSearch')
        except polling2.PollingException:
            continue  # no search button - maybe menu disappeared?
        for button in search_buttons:
            if button.is_displayed():
                search_button = button.find_element_by_tag_name('span')
                break
        else:
            continue  # retry clicking dropdown
        if check_in_dropdown:  # maybe we can skip the dialog and just click a menu option
            try:
                menu_items = poll(web.driver.find_elements_by_class_name, 'promptMenuOption')
                for item in menu_items:
                    if item.get_attribute('title') == search_term:
                        item.click()
                        return
            except (polling2.PollingException, selenium.common.exceptions.StaleElementReferenceException):
                continue  # no menu options - maybe menu disappeared?
        print('Clicking Search')
        try:
            search_button.click()
            search_box = poll(web.driver.find_element_by_id, 'choiceListSearchString_D')
            search_box.send_keys(search_term)
        except (polling2.PollingException, selenium.common.exceptions.StaleElementReferenceException):
            continue
        break
    else:
        raise InformationFetchFailure('Failed to click Search button')
    sleep(10)

    def get_search_result(*args):
        poll(web.driver.find_element_by_name, 'searchButton').click()
        return poll(web.driver.find_element_by_class_name, 'searchHighlightedText')  # raises NoSuchElementException if no results
    try:
        poll(get_search_result)
    except polling2.PollingException:
        raise InformationFetchFailure(f'No results for {search_term}')
    move_all = poll(web.driver.find_element_by_id, 'idMoveAllButton')
    move_all.click()
    sleep(3)
    ok_button = poll(web.driver.find_element_by_name, 'OK')
    ok_button.click()
    sleep(10)


def get_budget_data(show_window=False, project_names='all'):
    user_profile = os.environ['UserProfile']
    today = datetime.today()
    fy = today.year - (1 if today.month < 4 else 0)  # last calendar year if before April
    excel_filename = os.path.join(user_profile, 'Documents', 'Budgets', f'Budget summaries {fy}.xlsx')
    # The Index sheet here must have columns Name, Project and Task at least.
    # Budget data will be placed into sheets named after the Name column, overwriting anything already in there.
    csv_filename = os.path.join(user_profile, 'Downloads', 'Detailed Cost by project.csv')
    for line in list(pandas.read_excel(excel_filename, sheet_name='Index', dtype='string').itertuples()):
        if project_names != 'all' and line.Name not in project_names:
            continue
        web = go_to_oracle_page(use_obi=True, show_window=show_window)
        # web.driver.implicitly_wait(10)
        try:
            print(f'Fetching data for {line.Name}')
            get_task_data(web, line.Project, line.Task)
            export_csv_using_poll(csv_filename, web)
            print('Writing data to spreadsheet')
            with pandas.ExcelWriter(excel_filename, mode='a', if_sheet_exists='replace') as writer:
                data = pandas.read_csv(csv_filename, parse_dates=['Date', 'Fiscal Date', 'Fiscal Period'])
                data.to_excel(writer, sheet_name=line.Name)
            print(f'Complete for {line.Name}')
        except InformationFetchFailure as e:
            print(e)
        finally:
            web.driver.quit()


def export_csv_using_poll(csv_filename, web):
    if os.path.exists(csv_filename):
        os.remove(csv_filename)
    print('Exporting data')
    # there are two tables to export, we want the last one
    poll(web.driver.find_elements_by_link_text, 'Export', check_success=lambda result: len(result) > 1)[-1].click()
    poll(web.driver.find_elements_by_link_text, 'Data')[0].click()
    poll(web.driver.find_elements_by_link_text, 'CSV')[0].click()
    poll(os.path.exists, csv_filename)


def get_task_data(web, project_code, task):
    """Get budget data for given project code and task."""
    # try:
    search_via_prompt_using_poll(web, 1, project_code)  # Project Number
    search_via_prompt_using_poll(web, 8, task)  # , check_in_dropdown=True)  # Task Number
    # except selenium.common.exceptions.NoSuchElementException:  # no results for this project/task combination
    #     raise InformationFetchFailure(f'No results for {project_code=}, {task=}')
    print('Getting results')
    poll(web.driver.find_element_by_name, 'gobtn').click()  # Apply
    sleep(30)  # wait for results to be returned
    print('Going to transactions page')
    overflows = poll(web.driver.find_elements_by_class_name, 'obipsTabBarOverflow')
    for element in overflows:
        if element.is_displayed():
            element.click()
            break
    # sleep(10)
    poll(web.find_elements, 'Transactions')[0].click()
    # sleep(10)
    return True


if __name__ == '__main__':
    get_budget_data(True)
