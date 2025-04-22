import os
import time
from datetime import datetime, date

from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement

import oracle
from oracle import go_to_oracle_page

user_profile = os.path.expanduser('~')
downloads_folder = os.path.join(user_profile, 'Downloads')


def get_options(web: WebDriver) -> list[WebElement]:
    return web.find_element(By.ID, 'AdvicePicker').find_elements(By.TAG_NAME, 'option')


def get_one_slip(web: WebDriver, index: int) -> str:
    option = get_options(web)[index]
    name = option.text
    slip_date = datetime.strptime(name.split(' - ')[0], '%d %b %Y')
    date_text = slip_date.strftime('%b %Y')
    new_filename = slip_date.strftime('%Y-%m.pdf')
    if os.path.exists(new_filename):
        print(f'Already got payslip for {date_text}')
        return ''
    else:
        print(f'Downloading payslip for {date_text}')
    print(name)
    option.click()
    time.sleep(2)
    for _ in range(2):  # doesn't work first time!
        web.find_element(By.ID, 'Go').click()
        time.sleep(5)
    old_files = set(os.listdir())
    web.find_element(By.ID, 'Export').click()
    time.sleep(2)
    new_file = (set(os.listdir()) - old_files).pop()
    os.rename(new_file, new_filename)
    return new_filename

def get_payslips(only_latest: bool = True) -> None | str | date:
    """Download all my payslips, or just the latest."""
    web = go_to_oracle_page('RCUK Self-Service Employee', 'Payslip',
                            browser=oracle.Browser.edge, manual_login=True)

    print(f'Downloading payslips to {downloads_folder}')
    try:
        payslip_count = len(get_options(web))
        print(f'Found {payslip_count} payslips on Oracle')
        os.chdir(downloads_folder)

        result = '\n'.join(get_one_slip(web, i) for i in ((-1,) if only_latest else range(payslip_count)))
    finally:
        web.quit()
    print('Finished')

    return result


if __name__ == '__main__':
    print(get_payslips(only_latest=False))
