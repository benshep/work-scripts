import os
import re
import subprocess
from time import sleep

from win32api import GetKeyState

import outlook
from work_folders import user_profile
from start_meeting_notes import start_notes


def join_vc_meeting(force_sync: bool = False):
    """Look through today's events in an Outlook calendar, and follow the Zoom or Teams link contained within the body
    or subject of the closest one to the current time. If there are more than one, ask the user which one to join."""

    os.system('title 🖥️ Join VC meeting')
    # Zoom regex: try to capture all possible subdomains
    # j is join - w seems to be webinar
    # 9-11 digits for the meeting ID
    # Following that, URL params with alphanumeric, =&.- (any more required?)
    zoom_url = re.compile(r'https://(?:[\w\-]+.)?(?:zoom\.us|zoomgov\.com|zoom-x\.de)/[\w/?=&\-]+\b')  # simpler regex
    teams_url = re.compile(r'https://teams\.microsoft\.com/l/meetup-join/.*\b')
    current_events = outlook.get_current_events(min_count=-1 if force_sync else 1)
    # print('Current events:', [event.Subject for event in current_events])
    joinable_meetings = {}
    for meeting in current_events:
        body_location = f'{meeting.Body}\r\n{meeting.Location}'
        if (match := zoom_url.search(body_location)) or teams_url.search(body_location):
            url = match.group()
            joinable_meetings[meeting.Subject] = (url, meeting)
    meeting_count = len(joinable_meetings)
    if meeting_count == 1:
        (subject, (url, meeting)), = joinable_meetings.items()
    elif meeting_count > 1:
        subjects = list(joinable_meetings.keys())
        for i, subject in enumerate(subjects):
            print(f'{i:2d}. {subject}')
        print(f'{meeting_count:2d}. Force sync...')
        i = min(int(input('Choose meeting to join [0]: ') or 0), meeting_count)
        if i < meeting_count:
            subject = subjects[i]
            url, meeting = joinable_meetings[subject]
        else:  # 'Force sync' selected
            join_vc_meeting(True)
    else:
        print('No joinable meetings found')
        if not force_sync:
            print('Syncing Outlook events')
            join_vc_meeting(True)  # try one more time and carry out a sync
        else:
            sleep(10)
        return
    print(f'Joining {subject} ...')
    if 'zoom' in url:
        appdata_exe = os.path.join(user_profile, r'AppData\Roaming\Zoom\bin\Zoom.exe')
        program_files_exe = r'C:\Program Files\Zoom\bin\Zoom.exe'
        subprocess.Popen([appdata_exe if os.path.exists(appdata_exe) else program_files_exe, f"--url={url}"])
    else:  # Teams
        subprocess.Popen(['ms-teams.exe', url])
    print('Starting note file')
    start_notes(meeting)
    sleep(5)


if __name__ == '__main__':
    # Hold down Ctrl on starting the script to force a blocking sync
    # sleep(2)
    # print(win32api.GetKeyState(0x11) < 0)  # Ctrl
    join_vc_meeting(GetKeyState(0x11) < 0)  # Ctrl
