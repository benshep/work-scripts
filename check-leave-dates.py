import pandas
from oracle import go_to_oracle_page
from outlook import get_appointments_in_range
from pushbullet import Pushbullet  # to show notifications
from pushbullet_api_key import api_key  # local file, keep secret!

web = go_to_oracle_page(('RCUK Self-Service Employee', 'Attendance Management'))
cells = web.driver.find_elements_by_class_name('x1w')

start_dates = [element.text for element in cells[::8]]
end_dates = [element.text for element in cells[1::8]]
oracle_off_dates = set(sum([pandas.bdate_range(start, end).to_pydatetime().tolist()
                            for start, end in zip(start_dates, end_dates)], []))

outlook_off_dates = set(sum([pandas.bdate_range(ev.Start.date(), ev.End.date(), closed='left').to_pydatetime().tolist()
                             for ev in get_appointments_in_range(min(oracle_off_dates), 0)
                             if ev.AllDayEvent and ev.BusyStatus == 3 and ev.Subject == 'Annual Leave'], []))

not_in_outlook = oracle_off_dates - outlook_off_dates
print('Not in Outlook', not_in_outlook)
not_in_oracle = outlook_off_dates - oracle_off_dates
print('Not in Oracle', not_in_oracle)
dmy = "%#d/%#m/%Y"
toast = ('Missing from Outlook: ' + ', '.join(date.strftime(dmy) for date in not_in_outlook)) if not_in_outlook else ''
toast += ('\nMissing from Oracle: ' + ', '.join(date.strftime(dmy) for date in not_in_oracle)) if not_in_oracle else ''
if toast:
    Pushbullet(api_key).push_note('📅 Check leave dates', toast)
