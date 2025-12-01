from datetime import date, timedelta

import staff
from mars_group import members


def run_otl_calculator():
    """Iterate through staff, listing the hours to upload for new OTL cards required."""
    # staff.verbose = True
    for member in members:
        # if member.known_as in ('Alex B',):
        if True:
            print('\n' + member.known_as)
            member.update_off_days(force_reload=False)
            start = date.today()
            start -= timedelta(days=start.weekday())  # Monday of current week
            start -= timedelta(days=7*4)  # check last few weeks
            while start + timedelta(days=3) <= date.today():  # wait until Thu to do current week
            # for _ in range(1):
                hours_booked = member.hours_for_week(start)
                print(f'{start.strftime("%d/%m/%Y")}: {hours_booked=:.2f}')
                if hours_booked < 37:
                    print(*member.bulk_upload_lines(start, manual_mode=True), sep='\n')
                start += timedelta(days=7)


def leave_cross_check():
    """Iterate through staff, and check Oracle vs Outlook leave bookings."""
    toast = ''
    for member in members:
        if member.known_as in ('Ben',):
        # if True:
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
    # run_otl_calculator()
    print(leave_cross_check())