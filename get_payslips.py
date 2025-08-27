import os
import time
import re
from collections import Counter
from datetime import datetime, timedelta, date
from calendar import monthrange
from shutil import move
from zipfile import ZipFile

from selenium.webdriver.common.by import By
from pypdf import PdfReader
from govuk_bank_holidays.bank_holidays import BankHolidays
from selenium.webdriver.remote.webdriver import WebDriver

from oracle import go_to_oracle_page
from folders import misc_folder, downloads_folder


def income_pre_tax(year: int) -> float:
    os.chdir(os.path.join(misc_folder, 'Money', 'Payslips'))
    total_payments = Counter()
    for fiscal_month in range(12):
        month = (fiscal_month + 3) % 12 + 1  # convert to calendar month
        payslip_filename = f'{year - 2000}-{month:02d}.pdf'
        payslip_content = PdfReader(payslip_filename).pages[0].extract_text()
        payments = re.findall('Units Rate Amount\n(.*)\n Amount', payslip_content, re.MULTILINE + re.DOTALL)[0]
        for payment in payments.split('\n'):
            label, amount = payment.rsplit(' ', maxsplit=1)
            total_payments[label] += float(amount)
    print(*total_payments.items(), sep='\n')
    return total_payments.total()


def get_payslips(only_latest: bool = True, test_mode: bool = False) -> str | date:
    """Download all my payslips, or just the latest."""
    web = go_to_oracle_page('payslips', show_window=test_mode)
    # On page load, shows latest slip

    # Get list of payslips we already have
    os.chdir(os.path.join(misc_folder, 'Money', 'Payslips'))
    existing_files = os.listdir()
    output = ''

    if only_latest:
        # Latest date
        title = web.find_element(By.CLASS_NAME, 'oj-sp-detail-panel-title').text
        slip_date = datetime.strptime(title, '%A, %d/%m/%Y')  # e.g. Monday, 28/07/2025
        downloaded_file = run_download(web)
        new_filename = slip_date.strftime('%y-%m.pdf')
        if new_filename in existing_files:
            print(f'Already got payslip for {slip_date.strftime("%b %Y")}')
            # When do we expect to get next one? Paid on second-to-last working day, should see payslip day before
            today = datetime.now().date()
            _, days_in_month = monthrange(today.year, today.month)
            rest_of_month = [today + timedelta(days=d) for d in range(1, days_in_month - today.day + 1)]
            bank_holidays = [bh['date'] for bh in BankHolidays().get_holidays(division=BankHolidays.ENGLAND_AND_WALES)]
            working_days = [day for day in rest_of_month if day.weekday() < 5 and day not in bank_holidays]
            return working_days[-3] if len(working_days) > 2 else today + timedelta(days=1)

        move(os.path.join(downloads_folder, downloaded_file), new_filename)
        payments_this_month = payslip_items(new_filename)
        last_month = slip_date - timedelta(days=31)
        last_month_filename = last_month.strftime('%y-%m.pdf')
        payments_last_month = payslip_items(last_month_filename)
        for description, amount in payments_this_month.items():
            prev_amount = payments_last_month.get(description, None)
            symbol = '⭐' if prev_amount is None else \
                '🔺' if amount > prev_amount else \
                '🔻' if amount < prev_amount else '➖'
            output += f'{description}: £{amount:.02f} {symbol}\n'

        # Amounts for latest
        # number_cells = web.find_elements(By.CLASS_NAME, 'oj-sp-scoreboard-metric-card-metric')
        # net_pay = number_cells[1].text.split(' ')[1]  # text is e.g. GBP 1,999.99
        # Summary table i.e. basic pay and any extras
        # This doesn't seem to work well
        # summary_table = web.find_element(By.XPATH, '//div[@aria-label="Summary"]')
        # web.execute_script("arguments[0].scrollIntoView(true);",
        #                    summary_table.find_elements(By.TAG_NAME, 'span')[-1])
        # time.sleep(2)
        # while True:
        #     summary = summary_table.text.strip().splitlines()
        #     if 'Loading' not in summary:
        #         break
        # output = f'Net pay: £{net_pay}\n'
        # for label, value in zip(summary[::2], summary[1::2]):
        #     output += f'{label}: £{value.replace(" GBP", "")}\n'


    else:  # get all
        # maybe doesn't work...
        web.find_element(By.XPATH, '//input[@aria-label="select all"]').click()
        downloaded_file = run_download(web)
        archive = ZipFile(os.path.join(downloads_folder, downloaded_file))
        for filename in archive.namelist():
            # list like ['1_2025-07-28_GBP Payslip.pdf', '2_2025-06-27_GBP Payslip.pdf']
            slip_date = datetime.strptime(filename.split('_')[1], '%Y-%m-%d')
            new_filename = slip_date.strftime('%y-%m.pdf')
            if new_filename not in existing_files:
                archive.extract(filename)
                move(filename, new_filename)
                output += new_filename + '\n'


    web.quit()

    return output


def run_download(web: WebDriver):
    old_files = set(os.listdir(downloads_folder))
    web.find_element(By.CLASS_NAME, 'oj-ux-ico-download').click()  # download button
    time.sleep(5)
    new_file = (set(os.listdir(downloads_folder)) - old_files).pop()
    return new_file


def payslip_items(filename: str) -> dict[str, float]:
    """Read a payslip PDF and return net pay and each payment line in a dict."""
    # read amounts from the payslip
    payslip_content = PdfReader(filename).pages[0].extract_text()
    net_pay = re.findall(r'NET PAY ([\d,]+\.\d\d)', payslip_content)[0]
    return_dict = {'Net pay': float(net_pay.replace(',', ''))}
    # payments and deductions are both preceded by these headers
    headers = 'Description Reference Amount YTD'
    payments = re.findall(f'{headers}\n(.*){headers}\n(.*)\n   Total Payments',
                          payslip_content, re.MULTILINE + re.DOTALL)
    payments = ''.join(payments[0])  # join payments and deductions
    payments = payments.replace('\n', ' ').replace('  ', ' ')
    items = re.findall(r'(.*?) ?([\d,]+\.\d\d) ([\d,]+\.\d\d)', payments)
    for description, amount, _ in items:
        return_dict[description.strip()] = float(amount.replace(',', ''))
    return return_dict

if __name__ == '__main__':
    print(get_payslips(test_mode=True, only_latest=True))
    # print(income_pre_tax(2024))
    # os.chdir(os.path.join(misc_folder, 'Money', 'Payslips'))
    # payments_this_month = payslip_items('25-07.pdf')
    # payments_last_month = payslip_items('25-06.pdf')
    # output = ''
    # for description, amount in payments_this_month.items():
    #     prev_amount = payments_last_month.get(description, None)
    #     symbol = '⭐' if prev_amount is None else \
    #         '🔺' if amount > prev_amount else \
    #             '🔻' if amount < prev_amount else '➖'
    #     output += f'{description}: £{amount:.02f} {symbol}\n'
    # print(output)