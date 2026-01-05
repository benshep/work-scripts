import os
from datetime import date, timedelta
from urllib.parse import urlencode

import oracle
import staff
from mars_group import members
from work_folders import downloads_folder


def run_otl_calculator() -> str | None:
    """Iterate through staff, listing the hours to upload for new OTL cards required."""
    # staff.verbose = True
    cards_to_book = 0
    with open(os.path.join(downloads_folder, 'otl_upload_links.html'), 'w') as links_file:
        for member in members:
            # if member.known_as in ('Nasiq',):
            if True:
                links_file.write(f'<h2>{member.known_as}</h2>\n')
                print('\n', member.name)
                member.update_off_days(force_reload=False)
                start = date.today()
                start -= timedelta(days=start.weekday())  # Monday of current week
                start -= timedelta(days=7*4)  # check last few weeks
                while start + timedelta(days=3) <= date.today():  # wait until Thu to do current week
                # for _ in range(7):
                    hours_booked = member.hours_for_week(start)
                    hours_needed = sum(member.hours_needed(start + timedelta(days=day)) for day in range(5))
                    print(f'{start.strftime("%d/%m/%Y")}: {hours_booked=:.2f}, {hours_needed=:.2f}')
                    if hours_needed - hours_booked > 0.01:
                        cards_to_book += 1
                        params = {'calledFromAddTimeCard': 'true', 'pAsofdate': start.strftime('%Y-%m-%d')}
                        if member.known_as != 'Ben':
                            params |= {'pPersonId': member.person_id, 'userContext': 'LINE_MANAGER'}
                        url = oracle.apps[('home',)] + 'time/timecards/landing-page?' + urlencode(params)
                        links_file.write(f'<p><a href="{url}">{start.strftime("%d/%m/%Y")}</a></p>\n')
                        lines = list(member.bulk_upload_lines(start, manual_mode=True))
                        print(*lines, sep='\n')
                        for line in lines:
                            links_file.write(f'<p>{line}</p>\n')
                    start += timedelta(days=7)
    if cards_to_book:
        return f'{cards_to_book=}'
    return None


def leave_cross_check():
    """Iterate through staff, and check Oracle vs Outlook leave bookings."""
    toast = ''
    for member in members:
        # if member.known_as in ('Ben',):
        if True:
            print('\n' + member.known_as)
            not_in_oracle, not_in_outlook = member.leave_cross_check()
            if not_in_oracle:
                toast += f'{member.known_as}: {not_in_oracle=}, {not_in_outlook=}\n'
    return toast


def show_leave_dates():
    """Iterate through staff, and show the leave dates for each one, not including bank holidays."""
    for member in members:
        print('\n' + member.known_as)
        member.update_off_days(force_reload=True)
        print(*sorted(day for day in member.off_days - staff.site_holidays.keys()), sep='\n')


if __name__ == '__main__':
    run_otl_calculator()
    ######
    # print(leave_cross_check())