from datetime import date, timedelta
from time import sleep
from urllib.parse import urlencode

from selenium.webdriver.common.by import By

import oracle

from mars_group import members


def run_otl_calculator():
    """Iterate through staff, listing the hours to upload for new OTL cards required."""
    for member in members:
        # if member.known_as in ('Matt', 'Nasiq'):
        print('\n' + member.known_as)
        member.update_off_days(force_reload=False)
        start = date(2025, 9, 22)
        while start + timedelta(days=4) <= date.today():  # wait until Fri to do current week
            hours_booked = member.hours_for_week(start)
            print(f'{start.strftime("%d/%m/%Y")}: {hours_booked=:.2f}')
            if hours_booked < 37:
                print(*member.bulk_upload_lines(start, manual_mode=True), sep='\n')
            start += timedelta(days=7)


def get_leave_balances():
    """Iterate through staff, and show the leave balance for each one."""
    web = oracle.go_to_oracle_page(show_window=True)
    for member in members:
        print('\n' + member.known_as)
        url = (f'https://fa-evzn-saasfaukgovprod1.fa.ocs.oraclecloud.com/fscmUI/redwood/absences/plan-balances?'
               + urlencode(
                    {'pPersonId': member.person_id,
                     'pAssignmentId': member.assignment_id,
                     'pfsUserMode': "'MGR'"
                     }))
        print(url)
        web.get(url)
        sleep(10)
        spans = web.find_elements(By.XPATH, '//span[@class="oj-text-color-primary"]')
        print(*[span.text for span in spans])
    web.quit()


def leave_cross_check():
    """Iterate through staff, and check Oracle vs Outlook leave bookings."""
    for member in members:
        print('\n' + member.known_as)
        member.leave_cross_check()


if __name__ == '__main__':
    leave_cross_check()
