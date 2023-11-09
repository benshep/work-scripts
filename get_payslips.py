import os
import time
import re
from datetime import datetime
from shutil import move
from PyPDF2 import PdfReader
from oracle import go_to_oracle_page
from pushbullet import Pushbullet  # to show notifications
from pushbullet_api_key import api_key  # local file, keep secret!

user_profile = os.environ['UserProfile']


def get_options(web):
    return web.find_element_by_id('AdvicePicker').find_elements_by_tag_name('option')


def get_one_slip(web, index):
    option = get_options(web)[index]
    name = option.text
    slip_date = datetime.strptime(name.split(' - ')[0], '%d %b %Y')
    date_text = slip_date.strftime('%b %Y')
    new_filename = os.path.join(user_profile, 'Misc', 'Money', 'Payslips', slip_date.strftime('%y-%m.pdf'))
    if os.path.exists(new_filename):
        print(f'Already got payslip for {date_text}')
        return
    print(name)
    option.click()
    time.sleep(2)
    for _ in range(2):  # doesn't work first time!
        web.find_element_by_id('Go').click()
        time.sleep(5)
    old_files = set(os.listdir())
    web.find_element_by_id('Export').click()
    time.sleep(2)
    new_file = (set(os.listdir()) - old_files).pop()
    move(new_file, new_filename)
    payslip_content = PdfReader(new_filename).pages[0].extract_text()
    net_pay = float(re.findall(r'SUMMARY NET PAY (\d+\.\d\d)', payslip_content)[0])
    print(net_pay)
    Pushbullet(api_key).push_note('💰 Payslip', f'Net pay for {date_text}: £{net_pay:.2f}')


def get_payslips(only_latest=True, test_mode=False):
    """Download all my payslips, or just the latest."""
    web = go_to_oracle_page(['RCUK Self-Service Employee', 'Payslip'], show_window=test_mode)

    payslip_count = len(get_options(web))
    os.chdir(os.path.join(user_profile, 'Downloads'))

    for i in (-1,) if only_latest else range(payslip_count):
        get_one_slip(web, i)
    web.quit()


if __name__ == '__main__':
    get_payslips(test_mode=True)
