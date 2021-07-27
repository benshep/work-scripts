import os
import pandas
import selenium.common
from datetime import datetime
from time import sleep
from webbot import Browser
from oracle_credentials import username, password  # local file, keep secret!


def search_via_prompt(index, search_term):
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


user_profile = os.environ['UserProfile']
csv_filename = os.path.join(os.path.join(user_profile, 'Downloads'), 'Detailed Cost by project.csv')
today = datetime.today()
fy = today.year - (1 if today.month < 4 else 0)  # last calendar year if before April
excel_filename = os.path.join(user_profile, 'Documents', 'Budgets', f'Budget summaries {fy}.xlsx')
# The Index sheet here must have columns Name, Project and Task at least.
# Budget data will be placed into sheets named after the Name column, overwriting anything already in there.
projects = pandas.read_excel(excel_filename, sheet_name='Index', dtype='string')
for name, project_code, task in zip(projects['Name'], projects['Project'], projects['Task']):
    print('Starting Chrome')
    web = Browser()
    print('Logging in to Oracle')
    web.go_to('https://obi.ssc.rcuk.ac.uk/analytics/saw.dll?dashboard&PortalPath=%2Fshared%2FSTFC%20Shared%2F_portal%2FSTFC%20Projects')
    web.type(username, 'username')
    web.type(password, 'password')
    web.click('Login')

    try:
        search_via_prompt(1, project_code)  # Project Number
        search_via_prompt(8, task)  # Task Number
    except selenium.common.exceptions.NoSuchElementException:  # no results for this project/task combination
        web.quit()
        continue
    print('Getting results')
    web.driver.find_element_by_name('gobtn').click()  # Apply
    sleep(5)  # wait for results to be returned
    print('Going to transactions page')
    web.driver.find_element_by_xpath('//div[@title="Transactions"]').click()
    sleep(5)

    try:
        os.remove(csv_filename)
    except FileNotFoundError:
        pass
    print('Exporting data')
    web.click('Export', loose_match=False, number=2)  # there are two tables to export, we want the lower one
    web.click('Data', loose_match=False)
    web.click('CSV', loose_match=False)
    sleep(5)
    csv_data = pandas.read_csv(csv_filename)

    with pandas.ExcelWriter(excel_filename, mode='a', if_sheet_exists='replace') as writer:
        csv_data.to_excel(writer, sheet_name=name)

    web.quit()
