import os
import time
import re
from datetime import datetime
from shutil import move
from selenium.webdriver.common.by import By
from pypdf import PdfReader
from oracle import go_to_oracle_page

user_profile = os.environ['UserProfile']


def get_options(web):
    return web.find_element(By.ID, 'AdvicePicker').find_elements(By.TAG_NAME, 'option')


def get_one_slip(web, index):
    option = get_options(web)[index]
    name = option.text
    slip_date = datetime.strptime(name.split(' - ')[0], '%d %b %Y')
    date_text = slip_date.strftime('%b %Y')
    new_filename = os.path.join(user_profile, 'Misc', 'Money', 'Payslips', slip_date.strftime('%y-%m.pdf'))
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
    payments = re.sub(' (\d)', ': £\\1', payments)  # neaten up a bit
    return f'Net pay for {date_text}: £{net_pay:.2f}\n{payments}'


def get_payslips(only_latest=True, test_mode=False):
    """Download all my payslips, or just the latest."""
    web = go_to_oracle_page(('RCUK Self-Service Employee', 'Payslip'), show_window=test_mode)

    try:
        payslip_count = len(get_options(web))
        os.chdir(os.path.join(user_profile, 'Downloads'))

        result = '\n'.join(get_one_slip(web, i) for i in ((-1,) if only_latest else range(payslip_count)))
    finally:
        web.quit()
    return result


if __name__ == '__main__':
    get_payslips(test_mode=True)
