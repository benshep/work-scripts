from datetime import datetime, timedelta
import win32com.client
import pywintypes
import subprocess
import re
import os
import sys

# get today's meetings from Outlook calendar
today = datetime.today() - timedelta(days=100)
tomorrow = today + timedelta(days=100)
appointments = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI').GetDefaultFolder(9).Items
appointments.Sort("[Start]")
appointments.IncludeRecurrences = True
date_format = "%d/%m/%Y"
restriction = f"[Start] >= '{today.strftime(date_format)}' AND [End] <= '{tomorrow.strftime(date_format)}'"
closest = None
# Try to capture all possible subdomains
# j is join - w seems to be webinar
# 9-11 digits for the meeting ID
# Following that, URL params with alphanumeric, =&.- (any more required?)
zoom_url = re.compile(r'(https://(?:[\w\-]+.)?zoom.us/[wj]/\d{9,11}(\?[\w=&\.\-]+)?)')

# try to find the closest to now
for appointmentItem in appointments.Restrict(restriction):
    try:
        # print(appointmentItem.Start)
        start = datetime.fromtimestamp(appointmentItem.StartUTC.timestamp())
    except (OSError, pywintypes.com_error):  # appointments with weird dates!
        continue
    dt = int(abs(start - datetime.now()).seconds / 60)  # nearest minute, otherwise we get clashes
    print(f'{start:} ({dt=}) {appointmentItem.Subject}')
    if closest is None or dt < closest or ('second' in sys.argv and dt == closest):
        closest = dt
        print(f'closest: {appointmentItem.Subject}, {closest}')
        # look for a join URL in the body or the subject
        try:
            url = zoom_url.search(f'{appointmentItem.Body}\r\n{appointmentItem.Subject}').group()
            print(f'{url=}')
        except AttributeError:  # not found in string
            pass

appdata_exe = r'C:\Users\bjs54\AppData\Roaming\Zoom\bin\Zoom.exe'
program_files_exe = r'C:\Program Files (x86)\Zoom\bin\Zoom.exe'
# subprocess.call([appdata_exe if os.path.exists(appdata_exe) else program_files_exe, f"--url={url}"])
