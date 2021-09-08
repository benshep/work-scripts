import pandas
from oracle import go_to_oracle_page
from outlook import get_appointments_in_range
from pushbullet import Pushbullet  # to show notifications
from pushbullet_api_key import api_key  # local file, keep secret!

web = go_to_oracle_page(('RCUK Self-Service Employee', 'Attendance Management'))
cells = web.driver.find_elements_by_class_name('x1w')

oracle_off_dates = set(
    sum(
        (
            pandas.bdate_range(start.text, end.text).to_pydatetime().tolist()
            for start, end in zip(cells[::8], cells[1::8])
        ),
        [],
    )
)


outlook_off_dates = set(
    sum(
        (
            pandas.bdate_range(ev.Start.date(), ev.End.date(), closed='left')
            .to_pydatetime()
            .tolist()
            for ev in get_appointments_in_range(min(oracle_off_dates), 0)
            if ev.AllDayEvent
            and ev.BusyStatus == 3
            and ev.Subject == 'Annual Leave'
        ),
        [],
    )
)


not_in_outlook = oracle_off_dates - outlook_off_dates
print('Not in Outlook', not_in_outlook)
not_in_oracle = outlook_off_dates - oracle_off_dates
print('Not in Oracle', not_in_oracle)
dm = "%#d %b"  # e.g. 4 Aug
toast = ('Missing from Outlook: ' + ', '.join(date.strftime(dm) for date in not_in_outlook)) if not_in_outlook else ''
toast += ('\nMissing from Oracle: ' + ', '.join(date.strftime(dm) for date in not_in_oracle)) if not_in_oracle else ''
if toast:
    Pushbullet(api_key).push_note('📅 Check leave dates', toast)