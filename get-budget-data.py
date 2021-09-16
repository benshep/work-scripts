import os
import pandas
import selenium.common
from datetime import datetime
from time import sleep
from webbot import Browser
from oracle_credentials import username, password  # local file, keep secret!


class InformationFetchFailure(Exception):
    pass


def search_via_prompt(web, index, search_term):
    """Search using the dropdown box and the pop-up prompt."""
    sleep(5)
    print(f'Searching for {search_term} in box {index}')
    # Sometimes webbot takes ages to find the elements - base Selenium is much quicker
    web.driver.find_elements_by_class_name('promptTextField')[index].click()
    web.click('Search...')
    sleep(3)
    web.driver.find_element_by_xpath('//*[@id="choiceListSearchString_D"]').send_keys(search_term)  # search box
    web.driver.find_element_by_name('searchButton').click()
    sleep(2)
    web.driver.find_element_by_class_name('searchHighlightedText')  # raises NoSuchElementException if no results
    web.click(id='idMoveAllButton')
    web.driver.find_element_by_name('OK').click()
    sleep(2)


def get_budget_data():
    user_profile = os.environ['UserProfile']
    today = datetime.today()
    fy = today.year - (1 if today.month < 4 else 0)  # last calendar year if before April
    excel_filename = os.path.join(user_profile, 'Documents', 'Budgets', f'Budget summaries {fy}.xlsx')
    # The Index sheet here must have columns Name, Project and Task at least.
    # Budget data will be placed into sheets named after the Name column, overwriting anything already in there.
    csv_filename = os.path.join(user_profile, 'Downloads', 'Detailed Cost by project.csv')
    for project in pandas.read_excel(excel_filename, sheet_name='Index', dtype='string'):
        try:
            web = obi_login()
            get_task_data(web, project.Project, project.Task)
            export_csv(csv_filename, web)
            with pandas.ExcelWriter(excel_filename, mode='a', if_sheet_exists='replace') as writer:
                pandas.read_csv(csv_filename).to_excel(writer, sheet_name=project.Name)
        except InformationFetchFailure as e:
            print(e)
        finally:
            web.quit()


def export_csv(csv_filename, web):
    if os.path.exists(csv_filename):
        os.remove(csv_filename)
    print('Exporting data')
    web.click('Export', loose_match=False, number=2)  # there are two tables to export, we want the lower one
    web.click('Data', loose_match=False)
    web.click('CSV', loose_match=False)
    for _ in range(30):
        sleep(1)
        if os.path.exists(csv_filename):
            break
    else:
        raise InformationFetchFailure('Data export not successful')


def get_task_data(web, project_code, task):
    """Get budget data for given project code and task."""
    try:
        search_via_prompt(web, 1, project_code)  # Project Number
        search_via_prompt(web, 8, task)  # Task Number
    except selenium.common.exceptions.NoSuchElementException:  # no results for this project/task combination
        raise InformationFetchFailure(f'No results for {project_code=}, {task=}')
    print('Getting results')
    web.driver.find_element_by_name('gobtn').click()  # Apply
    sleep(5)  # wait for results to be returned
    print('Going to transactions page')
    web.driver.find_element_by_xpath('//div[@title="Transactions"]').click()
    sleep(5)
    return True


def obi_login():
    """Login to Oracle Business Intelligence website."""
    print('Starting Chrome')
    web = Browser(showWindow=False)
    web.driver.maximize_window()
    print('Logging in to Oracle')
    web.go_to(
        'https://obi.ssc.rcuk.ac.uk/analytics/saw.dll?dashboard&PortalPath=%2Fshared%2FSTFC%20Shared%2F_portal%2FSTFC%20Projects')
    web.type(username, 'username')
    web.type(password, 'password')
    web.click('Login')
    return web


if __name__ == '__main__':
    get_budget_data()
