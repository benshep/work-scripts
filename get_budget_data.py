import os

import pandas
import selenium.common
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as condition
import polling2
from datetime import datetime
from time import sleep

from oracle import go_to_oracle_page
from folders import user_profile, docs_folder, downloads_folder

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


def search_via_prompt(web, field_name, search_term, check_in_dropdown=False):
    """Search using the dropdown box and the pop-up prompt."""
    print(f'Searching for {search_term} in box {field_name}')
    prompt_labels = web.find_elements(By.CLASS_NAME, 'promptLabel')
    index = next(i for i, label in enumerate(prompt_labels) if label.text == field_name)
    # try this a few times - sometimes the dropdown disappears
    for i in range(10):
        print(f'Clicking dropdown, attempt {i}')
        web.find_elements(By.CLASS_NAME, 'promptTextField')[index].click()
        try:
            # search_buttons = wait_for_elements(web, By.CLASS_NAME, 'DropDownSearch')
            search_buttons = web.find_elements(By.CLASS_NAME, 'DropDownSearch')
        except selenium.common.exceptions.NoSuchElementException:
            continue  # no search button - maybe menu disappeared?
        for button in search_buttons:
            if button.is_displayed():
                search_button = button.find_element(By.TAG_NAME, 'span')
                break
        else:
            continue  # retry clicking dropdown
        if check_in_dropdown:  # maybe we can skip the dialog and just click a menu option
            try:
                for item in web.find_elements(By.CLASS_NAME, 'promptMenuOption'):
                    if item.get_attribute('title') == search_term:
                        item.click()
                        return
            except selenium.common.exceptions.StaleElementReferenceException:
                continue  # no menu options - maybe menu disappeared?
        print('Clicking Search')
        try:
            search_button.click()
            web.find_element(By.ID, 'choiceListSearchString_D').send_keys(search_term)
        except selenium.common.exceptions.StaleElementReferenceException:
            continue
        break
    else:
        raise InformationFetchFailure('Failed to click Search button')
    sleep(10)

    web.find_element(By.NAME, 'searchButton').click()
    # raises NoSuchElementException if no results
    try:
        web.find_element(By.CLASS_NAME, 'searchHighlightedText')
    except selenium.common.exceptions.NoSuchElementException as e:
        raise InformationFetchFailure(f'No results for {search_term}') from e
    web.find_element(By.ID, 'idMoveAllButton').click()
    web.find_element(By.NAME, 'OK').click()
    sleep(10)


def wait_for_elements(web, by, value):
    WebDriverWait(web, 5).until(condition.presence_of_element_located((by, value)))
    return web.find_elements(by, value)


def get_budget_data(project_names='all', test_mode=False):
    """Fetch data on spending on one or more projects from Oracle."""
    if project_names != 'all' and isinstance(project_names, str):  # just one project name supplied, wrap it in a list
        project_names = [project_names, ]

    today = datetime.now()
    fy = today.year - (today.month < 4)  # last calendar year if before April
    excel_filename = os.path.join(docs_folder, 'Budget', f'Budget summaries {fy}.xlsx')
    # The Index sheet here must have columns Name, Project and Task at least.
    # Budget data will be placed into sheets named after the Name column, overwriting anything already in there.
    csv_filename = os.path.join(downloads_folder, 'Detailed Cost by project.csv')
    for line in list(pandas.read_excel(excel_filename, sheet_name='Index', dtype='string').itertuples()):
        if project_names != 'all' and line.Name not in project_names:
            continue
        web = go_to_oracle_page('obi', show_window=test_mode)
        try:
            print(f'Fetching data for {line.Name}')
            get_task_data(web, line.Project, line.Task)
            export_csv(csv_filename, web)
            print('Writing data to spreadsheet')
            with pandas.ExcelWriter(excel_filename, mode='a', if_sheet_exists='replace') as writer:
                data = pandas.read_csv(csv_filename, parse_dates=['Date', 'Fiscal Date', 'Fiscal Period'])
                data.to_excel(writer, sheet_name=line.Name)
            print(f'Complete for {line.Name}')
        except InformationFetchFailure as e:
            print(e)
        finally:
            web.quit()


def export_csv(csv_filename, web):
    if os.path.exists(csv_filename):
        os.remove(csv_filename)
    print('Exporting data')
    # there are two tables to export, we want the last one
    by, value = By.LINK_TEXT, 'Export'
    wait = WebDriverWait(web, 5)
    wait.until(lambda web: len(web.find_elements(by, value)) > 1)
    web.find_elements(by, value)[-1].click()
    # wait_for_elements(web, By.LINK_TEXT, 'Export')[-1].click()
    web.find_element(By.LINK_TEXT, 'Data').click()
    web.find_element(By.LINK_TEXT, 'CSV').click()
    poll(non_zero_size, csv_filename)


def non_zero_size(filename):
    return os.path.exists(filename) and os.path.getsize(filename) > 0


def get_task_data(web, project_code, task):
    """Get budget data for given project code and task."""
    # try:
    search_via_prompt(web, 'Project Number', project_code)
    search_via_prompt(web, 'Task Number', task)  # , check_in_dropdown=True)
    # except selenium.common.exceptions.NoSuchElementException:  # no results for this project/task combination
    #     raise InformationFetchFailure(f'No results for {project_code=}, {task=}')
    print('Getting results')
    wait = WebDriverWait(web, 5)
    wait.until(condition.presence_of_element_located((By.NAME, 'gobtn'))).click()  # Apply
    # sleep(30)  # wait for results to be returned
    print('Going to transactions page')
    overflows = wait_for_elements(web, By.CLASS_NAME, 'obipsTabBarOverflow')
    for element in overflows:
        if element.is_displayed():
            element.click()
            break
    # sleep(10)
    wait.until(condition.presence_of_element_located((By.XPATH, '//div[@title="Transactions"]'))).click()
    # sleep(10)
    return True


if __name__ == '__main__':
    get_budget_data(project_names='MaRS Group', test_mode=True)
