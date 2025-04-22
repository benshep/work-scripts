import os  # for working with files and folders
from time import sleep  # for waiting in the script
from datetime import datetime, date  # for working with dates

# selenium imports: for controlling a web browser instance
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement

import oracle  # my script for working with Oracle from a web browser

# Files are downloaded to C:\Users\username\Downloads (on Windows)
user_profile = os.path.expanduser('~')
downloads_folder = os.path.join(user_profile, 'Downloads')


def get_options(web: WebDriver) -> list[WebElement]:
    """Find the options in the payslip dropdown list."""
    return web.find_element(By.ID, 'AdvicePicker').find_elements(By.TAG_NAME, 'option')


def get_one_slip(web: WebDriver, index: int) -> str:
    """Download one payslip from Oracle."""
    option = get_options(web)[index]
    name = option.text
    # The date format in the dropdown is e.g. 31 Mar 2025
    slip_date = datetime.strptime(name.split(' - ')[0], '%d %b %Y')
    date_text = slip_date.strftime('%b %Y')  # convert to e.g. Mar 2025
    new_filename = slip_date.strftime('%Y-%m.pdf')  # filename is e.g. "2025-03.pdf"
    if os.path.exists(new_filename):
        print(f'Already got payslip for {date_text}')
        return ''
    else:
        print(f'Downloading payslip for {date_text}')
    print(name)
    option.click()  # select the option from the dropdown list
    sleep(2)  # wait a moment
    for _ in range(2):  # Click 'Go' twice, it doesn't work first time!
        web.find_element(By.ID, 'Go').click()
        sleep(5)  # wait again
    # we will already be in the downloads folder
    old_files = set(os.listdir())  # list files before download
    web.find_element(By.ID, 'Export').click()  # start the download
    sleep(2)  # wait for download
    new_file = (set(os.listdir()) - old_files).pop()  # find the file that's been added
    os.rename(new_file, new_filename)  # rename from Oracle's unhelpful names
    return new_filename


def get_payslips(only_latest: bool = True) -> None | str | date:
    """Download all my payslips, or just the latest."""
    web = oracle.go_to_oracle_page('RCUK Self-Service Employee', 'Payslip',
                                   browser=oracle.Browser.edge, manual_login=True)

    print(f'Downloading payslips to {downloads_folder}')
    try:
        payslip_count = len(get_options(web))
        print(f'Found {payslip_count} payslips on Oracle')
        os.chdir(downloads_folder)

        result = '\n'.join(get_one_slip(web, i) for i in ((-1,) if only_latest else range(payslip_count)))
    finally:
        web.quit()  # close the browser even if errors occurred
    print('Finished')

    return result


if __name__ == '__main__':
    print(get_payslips(only_latest=False))
