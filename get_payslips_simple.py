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


def get_radio_options(web: WebDriver) -> list[WebElement]:
    """Find the options in the P60 radio buttons."""
    return web.find_elements(By.NAME, 'YearGroup')


def get_one_slip(web: WebDriver, index: int, p60: bool) -> str:
    """Download one payslip from Oracle."""
    options_func = get_radio_options if p60 else get_options
    option = options_func(web)[index]
    if p60:
        # Need to look for span elements with the year, there will be the same number as input elements
        text_options = web.find_elements(By.CLASS_NAME, 'xo')
        name = text_options[index].text
        # The date format in the P60 radio button starts with the year
        date_text = name[:4]
        new_filename = f'P60 {date_text}.pdf'  # filename is e.g. "2024.pdf"
    else:
        name = option.text
        # The date format in the payslip dropdown is e.g. 31 Mar 2025
        slip_date = datetime.strptime(name.split(' - ')[0], '%d %b %Y')
        date_text = slip_date.strftime('%b %Y')  # convert to e.g. Mar 2025
        new_filename = slip_date.strftime('Payslip %Y-%m.pdf')  # filename is e.g. "Payslip 2025-03.pdf"
    if os.path.exists(new_filename):
        print(f'Already got file for {date_text}')
        return ''
    else:
        print(f'Downloading file for {date_text}')
    print(name)
    option.click()  # select the option from the dropdown list
    if not p60:
        sleep(2)  # wait a moment
        for _ in range(2):  # Click 'Go' twice, it doesn't work first time!
            web.find_element(By.ID, 'Go').click()
            sleep(5)  # wait again
    # we will already be in the downloads folder
    old_files = set(os.listdir())  # list files before download
    web.find_element(By.ID, 'ViewReport' if p60 else 'Export').click()  # start the download
    for _ in range(10):
        sleep(2)  # wait for download
        new_files = set(os.listdir()) - old_files  # only files that weren't there before
        if new_files:
            break
    else:  # didn't break out of loop - no new files
        raise FileNotFoundError('Timed out waiting for download')
    new_file = new_files.pop()  # find the file that's been added
    os.rename(new_file, new_filename)  # rename from Oracle's unhelpful names
    return new_filename


def get_payslips() -> None | str | date:
    """Download all my payslips from Oracle."""
    web = oracle.go_to_oracle_page('RCUK Self-Service Employee',
                                   browser=oracle.Browser.edge, manual_login=True)
    pages = ['Payslip', 'P60 - 2018 Onwards', 'P60 - 2017 and Prior Years']
    result = ''
    for page in pages:
        web.find_element(By.LINK_TEXT, page).click()
        p60 = page.startswith('P60')

        print(f'{page.split(" ")[0]}s downloading to {downloads_folder}')
        try:
            options_func = get_radio_options if p60 else get_options
            payslip_count = len(options_func(web))
            print(f'Found {payslip_count} files to download')
            os.chdir(downloads_folder)

            result = '\n'.join(
                get_one_slip(web, i, p60)
                for i in range(payslip_count)
            )
        finally:
            web.find_element(By.LINK_TEXT, 'Home').click()
    web.quit()  # close the browser even if errors occurred
    print('Finished')

    return result


if __name__ == '__main__':
    print(get_payslips())
