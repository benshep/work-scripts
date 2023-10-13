#!python3
# -*- coding: utf-8 -*-

"""join_zoom_meeting
Ben Shepherd, 2020
Look through today's events in an Outlook calendar, and follow the Zoom link contained within the subject or body
of the closest one to the current time.
"""

import os
import re
import subprocess
from time import sleep
import outlook

user_profile = os.environ['UserProfile']


def join_zoom_meeting():
    """Look through today's events in an Outlook calendar, and follow the Zoom link contained within the subject or body
    of the closest one to the current time. If there are two, ask the user which one to join."""

    os.system('title Join Zoom meeting')
    # Zoom regex: try to capture all possible subdomains
    # j is join - w seems to be webinar
    # 9-11 digits for the meeting ID
    # Following that, URL params with alphanumeric, =&.- (any more required?)
    zoom_url = re.compile(r'https://(?:[\w\-]+.)?zoom.us/[\w/\?=&\-]+\b')  # simpler regex
    current_events = outlook.get_current_events()
    print('Current events:', [event.Subject for event in current_events])
    joinable_meetings = {}
    for meeting in current_events:
        body_location = f'{meeting.Body}\r\n{meeting.Location}'
        if match := zoom_url.search(body_location):
            url = match.group()
            joinable_meetings[meeting.Subject] = url
    meeting_count = len(joinable_meetings)
    if meeting_count == 1:
        (subject, url), = joinable_meetings.items()
    elif meeting_count > 1:
        subjects = list(joinable_meetings.keys())
        [print(f'{i}. {subject}') for i, subject in enumerate(subjects)]
        try:
            i = min(int(input('Choose meeting to join [0]: ')), meeting_count - 1)
        except ValueError:
            i = 0
        subject = subjects[i]
        url = joinable_meetings[subject]
    else:
        print('No joinable meetings found')
        sleep(10)
        return
    print(f'Joining {subject} ...')
    appdata_exe = os.path.join(user_profile, r'AppData\Roaming\Zoom\bin\Zoom.exe')
    program_files_exe = r'C:\Program Files\Zoom\bin\Zoom.exe'
    subprocess.Popen([appdata_exe if os.path.exists(appdata_exe) else program_files_exe, f"--url={url}"])
    sleep(5)


if __name__ == '__main__':
    join_zoom_meeting()
