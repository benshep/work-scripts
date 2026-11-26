import re
from base64 import b64decode
from calendar import monthrange
from collections import Counter
from datetime import datetime, timedelta
from urllib.parse import quote

import requests
from govuk_bank_holidays.bank_holidays import BankHolidays
from pypdf import PdfReader

import oracle
from mars_group import members
from work_folders import misc_folder

payslips_folder = misc_folder / 'Money' / 'Payslips'


def income_pre_tax(year: int) -> float:
    """Add up the pre-tax income for a given year."""
    total_payments = Counter()
    for fiscal_month in range(12):
        month = (fiscal_month + 3) % 12 + 1  # convert to calendar month
        payslip_filename = payslips_folder / f'{year - 2000}-{month:02d}.pdf'
        payslip_content = PdfReader(payslip_filename).pages[0].extract_text()
        payments = re.findall('Units Rate Amount\n(.*)\n Amount', payslip_content, re.MULTILINE + re.DOTALL)[0]
        for payment in payments.split('\n'):
            label, amount = payment.rsplit(' ', maxsplit=1)
            total_payments[label] += float(amount)
    print(*total_payments.items(), sep='\n')
    return total_payments.total()


def get_payslips(test_mode: bool = False, **kwargs) -> str | datetime:
    """Download all my payslips, or just the latest."""
    # Get list of payslips we already have
    existing_files = list(payslips_folder.glob('*.pdf'))
    output = ''

    web = oracle.go_to_oracle_page('payslips', show_window=test_mode, wired=True)
    auth = web.auth
    web.quit()
    query = quote(f'PersonId={members[0].person_id}')
    headers = {'Authorization': auth, 'Content-Type': 'application/vnd.oracle.adf.resourcecollection+json'}
    response = requests.get(oracle.rest_url_3 + f'/payslips?onlyData=true&q={query}&totalResults=true', headers=headers)
    payslips = response.json()
    latest_slip, previous_slip = payslips['items'][:2]
    slip_date = datetime.strptime(latest_slip['PaymentDate'], '%Y-%m-%d')
    new_filename = payslips_folder / slip_date.strftime('%y-%m.pdf')
    if new_filename in existing_files:
        print(f'Already got payslip for {slip_date.strftime("%b %Y")}')
        # When do we expect to get next one? Paid on second-to-last working day, should see payslip day before
        today = datetime.now()
        _, days_in_month = monthrange(today.year, today.month)
        rest_of_month = [today + timedelta(days=d) for d in range(1, days_in_month - today.day + 1)]
        bank_holidays = [bh['date'] for bh in BankHolidays().get_holidays(division=BankHolidays.ENGLAND_AND_WALES)]
        working_days = [day for day in rest_of_month if day.weekday() < 5 and day.date() not in bank_holidays]
        return working_days[-3] if len(working_days) > 2 else today + timedelta(days=1)

    slip_id = latest_slip['DocumentsOfRecordId']
    response = requests.get(oracle.rest_url_3 + f'/payslips/{slip_id}/child/PayslipBalances', headers=headers)
    this_payslip = response.json()
    prev_id = previous_slip['DocumentsOfRecordId']
    response = requests.get(oracle.rest_url_3 + f'/payslips/{prev_id}/child/PayslipBalances', headers=headers)
    prev_payslip = response.json()
    for line in this_payslip['items']:
        name = line['ReportingName']
        amount = float(line['Amount'] or 0)
        if amount == 0:
            continue
        prev_amount = next((float(prev_line['Amount'] or 0) for prev_line in prev_payslip['items']
                            if prev_line['BaseBalanceName'] == line['BaseBalanceName']), 0)
        symbol = '⭐' if prev_amount == 0 else '🔺' if amount > prev_amount else '🔻' if amount < prev_amount else '➖'
        output += f'{name}: £{amount:.02f} {symbol}\n'
        print(name, amount, symbol)

    response = requests.post(oracle.rest_url_3 + '/payslips/action/downloadAttachments', json={"dorIDList": [slip_id]},
                             headers={'Authorization': auth, 'Content-Type': 'application/vnd.oracle.adf.action+json'})
    open(new_filename, 'wb').write(b64decode(response.json()['result']['data']))
    return output


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
