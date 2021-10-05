#!python3
# -*- coding: utf-8 -*-

"""join_zoom_meeting
Ben Shepherd, 2020
Look through today's events in an Outlook calendar, and follow the Zoom link contained within the subject or body
of the closest one to the current time.
Command line arguments:
    second      find the second meeting if there are more than one starting at the same time (order is arbitrary)
"""

import os
import re
import subprocess
import sys
import outlook

user_profile = os.environ['UserProfile']


def join_zoom_meeting(skip_one=False):
    """Look through today's events in an Outlook calendar, and follow the Zoom link contained within the subject or body
    of the closest one to the current time."""

    # Zoom regex: try to capture all possible subdomains
    # j is join - w seems to be webinar
    # 9-11 digits for the meeting ID
    # Following that, URL params with alphanumeric, =&.- (any more required?)
    zoom_url = re.compile(r'(https://(?:[\w\-]+.)?zoom.us/[wj]/\d{9,11}(\?[\w=&\.\-]+)?)')
    current_events = outlook.get_current_events()
    print('Current events:', [event.Subject for event in current_events])
    for meeting in current_events:
        match = zoom_url.search(f'{meeting.Body}\r\n{meeting.Location}')
        if match:
            if skip_one:
                skip_one = False
                continue
            url = match.group()
            appdata_exe = os.path.join(user_profile, r'AppData\Roaming\Zoom\bin\Zoom.exe')
            program_files_exe = r'C:\Program Files\Zoom\bin\Zoom.exe'
            subprocess.call([appdata_exe if os.path.exists(appdata_exe) else program_files_exe, f"--url={url}"])
            break


if __name__ == '__main__':
    join_zoom_meeting('second' in sys.argv)
