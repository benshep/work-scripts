import sys
import os
import outlook
import subprocess
import datetime
from platform import node
sys.path.append(os.path.join(os.environ['UserProfile'], 'Misc', 'Scripts'))
import google_sheets

sheet_id = '1qjcicqh02LMCdQUJwkvnOYli8TJb9qRdZyvyOwKMFHA'  # 🗺️ Where is Ben?


def dmy(date, time=True):
    """Convert datetime into dd/mm/yyyy format."""
    return date.strftime('%d/%m/%Y' + (' %H:%M' if time else ''))


def events_to_spreadsheet():
    """Fetch a list of events from Outlook, and drop it into a spreadsheet."""
    events = outlook.get_appointments_in_range()
    for i, event in enumerate(events):
        print(event.Start, event.End, event.Subject, event.BusyStatus, sep='\t')
        google_sheets.update_cells(sheet_id, 'Events', f'A{i + 2}:F{i + 2}',
                                   [[dmy(event.Start), dmy(event.End), event.Subject, event.Location,
                                     event.BusyStatus, event.AllDayEvent]])


def set_pc_unlocked_flag():
    """Check if PC is locked or not, and update a spreadsheet accordingly."""
    # https://stackoverflow.com/questions/34514644/in-python-3-how-can-i-tell-if-windows-is-locked#answer-57258754
    unlocked = 'LogonUI.exe' not in str(subprocess.check_output('TASKLIST'))
    print('Updating online status:', unlocked)
    google_sheets.update_cell(sheet_id, '', f'{node()}_unlocked', unlocked)
    if unlocked:
        google_sheets.update_cell(sheet_id, '', f'{node()}_unlocked_updated', dmy(datetime.datetime.now()))


if __name__ == '__main__':
    # events_to_spreadsheet()
    set_pc_unlocked_flag()
