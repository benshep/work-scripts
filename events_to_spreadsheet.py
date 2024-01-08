import sys
import os
import outlook
import subprocess
import datetime
from platform import node

sys.path.append(os.path.join(os.environ['UserProfile'], 'Misc', 'Scripts'))
import google_sheets

sheet_id = {'me': '1qjcicqh02LMCdQUJwkvnOYli8TJb9qRdZyvyOwKMFHA',  # 🗺️ Where is Ben?
            'Hywel': '1usvoxxjpPZT0C4rZHKz9mVreDtEc-ahSfHYkpIJZ9FY'}


def dmy(date, time=True):
    """Convert datetime into dd/mm/yyyy format."""
    return date.strftime('%d/%m/%Y' + (' %H:%M' if time else ''))


def events_to_spreadsheet(user='me'):
    """Fetch a list of events from Outlook, and drop it into a spreadsheet."""
    events = outlook.get_appointments_in_range(user=user)
    if sheet_data := [[dmy(event.Start), dmy(event.End), event.Subject,
                       event.Location, event.BusyStatus, event.AllDayEvent] for event in events]:
        print(f'Found {len(sheet_data)} events, starting with {sheet_data[0][2]} at {sheet_data[0][0]}')
        google_sheets.update_cells(sheet_id[user], 'Events', f'A2:F{len(sheet_data) + 1}', sheet_data)


def set_pc_unlocked_flag(user='me'):
    """Check if PC is locked or not, and update a spreadsheet accordingly."""
    # https://stackoverflow.com/questions/34514644/in-python-3-how-can-i-tell-if-windows-is-locked#answer-57258754
    unlocked = 'LogonUI.exe' not in str(subprocess.check_output('TASKLIST'))
    print('Updating online status:', unlocked)
    google_sheets.update_cell(sheet_id[user], '', f'{node()}_unlocked', unlocked)
    if unlocked:
        google_sheets.update_cell(sheet_id[user], '', f'{node()}_unlocked_updated', dmy(datetime.datetime.now()))


if __name__ == '__main__':
    events_to_spreadsheet()
    # set_pc_unlocked_flag()
