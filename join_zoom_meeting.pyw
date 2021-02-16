from datetime import datetime, timedelta
import win32com.client
import pywintypes
import subprocess
import re
import os
import sys

# get today's meetings from Outlook calendar
today = datetime.today()
tomorrow = today + timedelta(days=1)
appointments = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI').GetDefaultFolder(9).Items
appointments.Sort("[Start]")
appointments.IncludeRecurrences = True
d_m_y = "%#d/%#m/%Y"  # no leading zeros
yyyy_mm_dd = '%y-%m-%d'
restriction = f"[Start] >= '{today.strftime(d_m_y)}' AND [End] <= '{tomorrow.strftime(d_m_y)}'"
min_dt = None
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
    if min_dt is None or dt < min_dt or ('second' in sys.argv and dt == min_dt):
        meeting = appointmentItem
        min_dt = dt
        print(f'closest: {meeting.Subject}, {min_dt}')
        # look for a join URL in the body or the subject
        try:
            url = zoom_url.search(f'{meeting.Body}\r\n{meeting.Subject}').group()
            print(f'{url=}')
        except AttributeError:  # not found in string
            pass

user_profile = os.environ['USERPROFILE']
if 'notes' in sys.argv:
    # start a file for notes
    os.chdir(os.path.join(user_profile, 'Documents'))
    folders = next(os.walk('.'))[1]
    subject = meeting.Subject
    try:
        folder = next(folder for folder in folders if folder.lower() in subject.lower())
    except StopIteration:
        try:
            folder = next(folder for folder in folders if folder.lower() in meeting.Body.lower())
        except StopIteration:
            folder = '.'  # couldn't work out what folder based on subject or body: put in root folder
    os.chdir(folder)
    names = []
    name_format = re.compile(r'(\w+), (\w+) \(\w+,\w+,\w+\)')
    for person in meeting.Recipients:
        person_name = person.Name
        match = name_format.match(person_name)
        if match:  # matches "Surname, Firstname (ORG,DEPT,GROUP)"
            names.append(f'{match.group(2)} {match.group(1)}')
        elif '@' in person_name:  # email address
            name, domain = person_name.split('@')
            names.append(name.title().replace('.', ' '))
        else:
            names.append(person_name.title())
    people_list = ', '.join(names)
    text = f'# {subject}\n\n*{start.strftime(d_m_y)}. {people_list}*\n\n'
    illegal_chars = '*?/\\<>:|"'
    for c in illegal_chars:
        subject = subject.replace(c, ' ')
    filename = f'{start.strftime(yyyy_mm_dd)} {subject}.md'
    open(filename, 'a').write(text)
    os.startfile(filename)
else:
    appdata_exe = os.path.join(user_profile, r'AppData\Roaming\Zoom\bin\Zoom.exe')
    program_files_exe = r'C:\Program Files (x86)\Zoom\bin\Zoom.exe'
    subprocess.call([appdata_exe if os.path.exists(appdata_exe) else program_files_exe, f"--url={url}"])
