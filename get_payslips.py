import os
import time
import re
from datetime import datetime, timedelta, date
from calendar import monthrange
from shutil import move

from selenium.webdriver.common.by import By
from pypdf import PdfReader
from govuk_bank_holidays.bank_holidays import BankHolidays
from selenium.webdriver.firefox.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement

from oracle import go_to_oracle_page
from folders import misc_folder, downloads_folder


def get_options(web: WebDriver) -> list[WebElement]:
    return web.find_element(By.ID, 'AdvicePicker').find_elements(By.TAG_NAME, 'option')


def get_one_slip(web: WebDriver, index: int) -> str:
    option = get_options(web)[index]
    name = option.text
    slip_date = datetime.strptime(name.split(' - ')[0], '%d %b %Y')
    date_text = slip_date.strftime('%b %Y')
    new_filename = os.path.join(misc_folder, 'Money', 'Payslips', slip_date.strftime('%y-%m.pdf'))
    if os.path.exists(new_filename):
        print(f'Already got payslip for {date_text}')
        return ''
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
    move(new_file, new_filename)
    payslip_content = PdfReader(new_filename).pages[0].extract_text()
    net_pay = float(re.findall(r'SUMMARY NET PAY (\d+\.\d\d)', payslip_content)[0])
    print(net_pay)
    payments = re.findall('Units Rate Amount\n(.*)\n Amount', payslip_content, re.MULTILINE + re.DOTALL)[0]
    payments = re.sub(r' (\d)', ': £\\1', payments)  # neaten up a bit
    return f'Net pay for {date_text}: £{net_pay:.2f}\n{payments}'


def get_payslips(only_latest: bool = True, test_mode: bool = False) -> None | str | date:
    """Download all my payslips, or just the latest."""
    web = go_to_oracle_page('RCUK Self-Service Employee', 'Payslip', show_window=test_mode)

    try:
        payslip_count = len(get_options(web))
        os.chdir(downloads_folder)

        result = '\n'.join(get_one_slip(web, i) for i in ((-1,) if only_latest else range(payslip_count)))
    finally:
        web.quit()

    if result:
        return result
    # When do we expect to get next one? Paid on second-to-last working day, should see payslip day before
    today = datetime.now().date()
    _, days_in_month = monthrange(today.year, today.month)
    rest_of_month = [today + timedelta(days=d) for d in range(1, days_in_month - today.day + 1)]
    bank_holidays = [bh['date'] for bh in BankHolidays().get_holidays(division=BankHolidays.ENGLAND_AND_WALES)]
    working_days = [day for day in rest_of_month if day.weekday() < 5 and day not in bank_holidays]
    return working_days[-3] if len(working_days) > 2 else today + timedelta(days=1)


if __name__ == '__main__':
    print(get_payslips(test_mode=True))
