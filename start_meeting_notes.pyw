#!python3
# -*- coding: utf-8 -*-

"""start_meeting_notes
Ben Shepherd, 2021
Look through today's events in an Outlook calendar,
and start a Markdown notes file for the closest one to the current time.
"""


import contextlib
import os
import re
import requests
from icalendar import Calendar
from time import sleep
import outlook

user_profile = os.environ['UserProfile']


def folder_match(name: str, test_against: str):
    """Case-insensitive string comparison."""
    # Add optional 's' to end of keyword - sometimes the folder contains a plural of a name
    return re.search(fr'\b{re.escape(name)}s?\b', test_against, re.IGNORECASE) is not None


def create_note_file():
    """Start a file for notes relating to the given meeting. Find a relevant folder in the user's Documents folder,
     searching first the subject then the body of the meeting for a folder name. Use Other if none found.
     Don't use the Zoom folder (this is often found in the meeting body).
     The file is in Markdown format, with the meeting title, date and attendees filled in at the top."""
    if not (meeting := target_meeting()):
        return  # no current meeting
    go_to_folder(meeting)
    # Hierarchy of responses: (we want to list attendees from the top)
    # olResponseOrganized 	    1 	On the Organizer's calendar or the recipient is the Organizer of the meeting.
    # olResponseAccepted 	    3 	Meeting accepted.
    # olResponseTentative 	    2 	Meeting tentatively accepted.
    # olResponseNone 	        0 	The appointment is a simple appointment and does not require a response.
    # olResponseNotResponded 	5 	Recipient has not responded.
    # olResponseDeclined 	    4 	Meeting declined.
    priority = [1, 3, 2, 0, 5, 4]
    response = {r.Name: r.MeetingResponseStatus for r in meeting.Recipients}
    attendees = filter(None, '; '.join([meeting.RequiredAttendees, meeting.OptionalAttendees]).split('; '))
    attendees = sorted(attendees, key=lambda r: priority.index(response[r]))
    people_list = ', '.join(format_name(person_name, response[person_name]) for person_name in attendees)
    start_time = outlook.get_meeting_time(meeting)
    meeting_date = start_time.strftime("%#d/%#m/%Y")  # no leading zeros
    subject = meeting.Subject.strip()  # remove leading and trailing spaces
    subtitle = f'*{meeting_date}. {people_list}*'
    description = meeting.Body
    description = description.replace('o   ', '  * ')  # second-level lists
    description = description.replace('\r\n\r\n', '\n')  # double-spaced paragraphs
    # snip out Zoom joining instructions
    start = description.find(' <http://zoom.us/> ')  # begins with this
    if start >= 0:
        # last bit is "Skype on a SurfaceHub" link, only one with @lync in it - or sometime SIP: xxxx@zoomcrc.com
        end = max(description.rfind('@lync.zoom.us'), description.rfind('@zoomcrc.com'))
        end = description.find('\r\n', end)  # end of line
        description = (description[:start] + description[end:]).strip()

    if match := re.search(r'https://[\w\.]+/event/\d+', meeting.Body):  # Indico link: look for an agenda
        url = match[0]
        description += ical_to_markdown(url)
        subject = f'[{subject}]({url})'
    text = f'# {subject}\n\n{subtitle}\n\n{description}\n\n'

    bad_chars = str.maketrans({char: ' ' for char in '*?/\\<>:|"'})  # can't use these in filenames
    filename = f'{start_time.strftime("%Y-%m-%d")} {subject.translate(bad_chars)}.md'
    open(filename, 'a', encoding='utf-8').write(text)
    os.startfile(filename)


def target_meeting():
    os.system('title 📓 Start meeting notes')
    current_events = outlook.get_current_events(hours_ahead=12)
    current_events = filter(lambda event: not outlook.is_wfh(event), current_events)
    current_events = filter(lambda event: event.Subject != 'ASTeC/CI Coffee', current_events)

    current_events = list(current_events)
    meeting_count = len(current_events)
    if meeting_count == 1:
        i = 0
    elif meeting_count > 1:
        [print(f'{i}. {event.Start.strftime("%H:%M")} {event.Subject}') for i, event in enumerate(current_events)]
        try:
            i = min(int(input('Choose meeting for note file [0]: ')), meeting_count - 1)
        except ValueError:
            i = 0
    else:
        print('No current meetings found')
        sleep(10)
        return None
    return current_events[i]


def walk(top, max_depth):
    """Pared-down version of os.walk that just lists folders and is limited to a given max depth."""
    # https://stackoverflow.com/questions/35315873/travel-directory-tree-with-limited-recursion-depth
    dirs = [name for name in os.listdir(top) if os.path.isdir(os.path.join(top, name))]
    yield top, dirs
    if max_depth > 1:
        for name in dirs:
            yield from walk(os.path.join(top, name), max_depth - 1)


def find_folder_from_text(text):
    """Choose an appropriate folder based on some text. Traverse the first-level folders, then the second level."""
    for path, folders in walk(os.getcwd(), 2):
        if is_banned(os.path.split(path)[1]):
            continue
        for folder in folders:
            if is_banned(folder):
                continue
            if folder_match(folder, text):
                return os.path.join(path, folder)
    return False


def is_banned(name):
    """Test banned state of folders. Folders are banned for the following reasons:
    - Name is Zoom, Other, or Old work - don't put notes in those
    - Name is lowercase only
    - Name is 3 characters or less"""
    return name in ('Zoom', 'Other', 'Old work') or name.lower() == name or len(name) <= 3


def go_to_folder(meeting):
    """Pick a folder in which to place the meeting notes, looking at the subject and the body."""
    os.chdir(os.path.join(user_profile, 'Documents'))
    folder = find_folder_from_text(meeting.Subject) or find_folder_from_text(meeting.Body) or 'Other'
    os.chdir(folder)
    return folder


def format_name(person_name, response):
    """Convert display name to more readable Firstname Surname format. Works with Surname, Firstname (ORG,DEPT,GROUP)
    or firstname.surname@company.com."""
    # match "Surname, Firstname (ORG,DEPT,GROUP)" - last bit in brackets is optional
    if match := re.match(r'(.+?), ([^ ]+)( \(.+?,.+?,.+?\))?', person_name):
        return_value = f'{match[2]} {match[1]}'  # Firstname Surname
    elif '@' in person_name:  # email address
        name, _ = person_name.split('@')
        return_value = name.title().replace('.', ' ')
    else:
        # title case, but preserve upper case words
        return_value = ' '.join([word if word.isupper() else word.title() for word in person_name.split()])
        return_value = return_value.replace(' - UKRI', '')
    if response == 4:  # declined
        return f'~~{return_value}~~'  # strikethrough
    elif response in (1, 3):  # organised, accepted
        return f'**{return_value}**'  # bold
    else:
        return return_value


def ical_to_markdown(url):
    """Given an Indico event URL, return a Markdown agenda."""
    # url = 'https://indico.desy.de/event/35655/event.ics?scope=contribution'
    if not url.endswith('/'):
        url += '/'
    # Some Indico versions use ?scope=contribution instead - use both!
    event = Calendar.from_ical(requests.get(f'{url}event.ics?detail=contributions&scope=contribution').text)
    prefix = 'Speakers: '
    agenda = ''
    for component in sorted(event.walk('VEVENT'), key=lambda c: c.decoded('dtstart')):
        agenda += f"## [{component.get('summary')}]({component.get('url')})\n"
        # print(component.get("description", ''))
        description = component.get("description", '').splitlines()
        speakers = description[0]
        abstract = description[2] if len(description) > 2 else ''
        if speakers.startswith(prefix):
            agenda += f'*{speakers[len(prefix):]}*\n'
        if abstract:
            agenda += abstract + '\n'
    return agenda


if __name__ == '__main__':
    # for meeting in outlook.get_appointments_in_range(-30, 30):
    #     print(meeting.Subject, go_to_folder(meeting), sep='; ')
    create_note_file()
