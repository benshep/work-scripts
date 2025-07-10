import contextlib
import os
import re
from typing import Iterator, Any
from win32api import GetKeyState

try:
    import pyvda  # virtual desktops
except ImportError:  # not available everywhere
    pyvda = None
import requests
from icalendar import Calendar
from time import sleep
import outlook
from folders import docs_folder, sharepoint_folder


def folder_match(name: str, test_against: str) -> bool:
    """Case-insensitive string comparison."""
    # Add optional 's' to end of keyword - sometimes the folder contains a plural of a name
    return re.search(fr'\b{re.escape(name)}s?\b', test_against, re.IGNORECASE) is not None


def create_note_file() -> None:
    """Prompt for a current meeting and start a note file."""
    if not (meeting := target_meeting()):
        return  # no current meeting
    start_notes(meeting)


def start_notes(meeting: outlook.AppointmentItem) -> None:
    """Start a file for notes relating to the given meeting. Find a relevant folder in the user's Documents folder,
     searching first the subject then the body of the meeting for a folder name. Use Other if none found.
     Don't use the Zoom folder (this is often found in the meeting body).
     The file is in Markdown format, with the meeting title, date and attendees filled in at the top."""
    go_to_folder(meeting)
    # Hierarchy of responses: (we want to list attendees from the top)
    priority = [
        outlook.ResponseStatus.organized,
        outlook.ResponseStatus.accepted,
        outlook.ResponseStatus.tentative,
        outlook.ResponseStatus.none,
        outlook.ResponseStatus.not_responded,
        outlook.ResponseStatus.declined
    ]
    response = {r.Name: outlook.ResponseStatus(r.MeetingResponseStatus) for r in meeting.Recipients}
    attendees = filter(None, '; '.join([meeting.RequiredAttendees, meeting.OptionalAttendees]).split('; '))
    attendees = sorted(attendees, key=lambda r: priority.index(response[r]))
    people_list = ', '.join(format_name(person_name, response[person_name]) for person_name in attendees)
    start_time = outlook.get_meeting_time(meeting)
    meeting_date = start_time.strftime("%#d/%#m/%Y")  # no leading zeros
    subject = meeting.Subject.strip()  # remove leading and trailing spaces
    subtitle = f'*{meeting_date}' + (f'. {people_list}*' if people_list else '*')
    description = clean_body(meeting.Body)

    if match := re.search(r'https://[\w\.]+/event/\d+', meeting.Body):  # Indico link: look for an agenda
        url = match[0]
        agenda = ical_to_markdown(url)
        text = f'# [{subject}]({url})\n\n{subtitle}\n\n{description}\n\n{agenda}\n\n'
    else:
        text = f'# {subject}\n\n{subtitle}\n\n{description}\n\n'

    bad_chars = str.maketrans({char: ' ' for char in '*?/\\<>:|"'})  # can't use these in filenames
    filename = f'{start_time.strftime("%Y-%m-%d")} {subject.translate(bad_chars)}.md'
    open(filename, 'a', encoding='utf-8').write(text)
    if pyvda:  # switch to meetings desktop if possible
        for desktop in pyvda.get_virtual_desktops():
            if desktop.name == '🤝 Meetings':
                desktop.go()
                break
    os.startfile(filename)


def clean_body(description: str) -> str:
    """Strip out Teams and Zoom joining instructions from meeting body, and generally improve formatting."""
    description = description.replace('o   ', '  * ')  # second-level lists
    description = description.replace('\r\n', '\n')  # CRLF -> LF
    description = description.replace('\n\n', '\n')  # double-spaced paragraphs
    # snip out Zoom joining instructions
    start = description.find(' <http://zoom.us')  # begins with this
    if start >= 0:
        # last bit is "Skype on a SurfaceHub" link, only one with @lync in it - or sometimes SIP: xxxx@zoomcrc.com
        end = max(description.rfind('@lync.zoom.us>') + 14,
                  description.rfind('@zoomcrc.com') + 12)
        # end = description.find('\n', end)  # end of line
        if end >= 0:
            description = (description[:start] + description[end:]).strip()
    # snip out Teams joining instructions
    horizontal_line = '_' * 80
    start = description.find(f'{horizontal_line}\nMicrosoft Teams Need help?')
    if start >= 0:
        end = description.rfind(horizontal_line) + 80
        if end >= 0:
            description = (description[:start] + description[end:]).strip()
    # print(description)
    return description


def target_meeting(min_count: int = 1, hours_ahead: float = 12) -> outlook.AppointmentItem:
    os.system('title 📓 Start meeting notes')
    time_format = "%a %d/%m %H:%M" if hours_ahead > 12 else "%H:%M"
    current_events = outlook.get_current_events(hours_ahead=hours_ahead, min_count=min_count)
    unfiltered_count = len(current_events)
    current_events = filter(lambda event: not outlook.is_wfh(event), current_events)
    current_events = filter(lambda event: event.Subject != 'ASTeC/CI Coffee', current_events)

    # put all the declined meetings at the end of the list
    current_events = sorted(list(current_events), key=lambda event: event.Subject.startswith('Declined: '))
    meeting_count = len(current_events)
    print('')
    for i, event in enumerate(current_events):
        print(f'{i:2d}. {event.Start.strftime(time_format)} {event.Subject}')
    print(f'{meeting_count:2d}. Look for more now...')
    print(f'{meeting_count + 1:2d}. Further ahead...')
    i = min(int(input('Choose meeting for note file [0]: ') or 0), meeting_count + 1)
    if i < meeting_count:  # valid meeting
        return current_events[i]
    # one of last two selected: get more now, or look further ahead
    if i == meeting_count:
        return target_meeting(unfiltered_count + 1, hours_ahead)
    else:
        return target_meeting(hours_ahead=hours_ahead + 24 * 7)  # look ahead to next week


def walk(top: str, max_depth: int) -> Iterator[tuple[str, list[str]]]:
    """Pared-down version of os.walk that just lists folders and is limited to a given max depth."""
    # https://stackoverflow.com/questions/35315873/travel-directory-tree-with-limited-recursion-depth
    dirs = [name for name in os.listdir(top) if os.path.isdir(os.path.join(top, name))]
    yield top, dirs
    if max_depth > 1:
        for name in dirs:
            yield from walk(os.path.join(top, name), max_depth - 1)


def find_folder_from_text(top_folder: str, text: str) -> str:
    """Choose an appropriate folder based on some text. Traverse the first-level folders, then the second level."""
    for path, folders in walk(top_folder, 2):
        if is_banned(os.path.split(path)[1]):
            continue
        for folder in folders:
            if is_banned(folder):
                continue
            if folder_match(folder, text):
                return os.path.join(path, folder)
    return ''


def find_subject_folder(folder: str, text: str) -> str:
    """Find a folder which names the specific text provided."""
    subject_match_file = 'meeting_subjects.txt'
    for path, folders, files in os.walk(folder):
        if subject_match_file not in files:
            continue
        subjects = open(os.path.join(path, subject_match_file)).read().splitlines()
        if text in subjects:
            return path
    return ''


def is_banned(name: str) -> bool:
    """Test banned state of folders. Folders are banned for the following reasons:
    - Name is Zoom, Other, or Old work - don't put notes in those
    - Name is lowercase only
    - Name is 3 characters or fewer"""
    return name in ('Zoom', 'Other', 'Old work') or name.lower() == name or len(name) <= 3


def go_to_folder(meeting: outlook.AppointmentItem) -> str:
    """Pick a folder in which to place the meeting notes, looking at the subject and the body."""
    folder = find_subject_folder(docs_folder, meeting.Subject) or \
             find_subject_folder(sharepoint_folder, meeting.Subject) or \
             find_folder_from_text(docs_folder, meeting.Subject) or \
             find_folder_from_text(docs_folder, meeting.Body) or \
             os.path.join(docs_folder, 'Other')
    os.chdir(folder)
    return folder


def format_name(person_name: str, response: outlook.ResponseStatus) -> str:
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
    if response == outlook.ResponseStatus.declined:
        return f'~~{return_value}~~'  # strikethrough
    elif response in (outlook.ResponseStatus.organized, outlook.ResponseStatus.accepted):
        return f'**{return_value}**'  # bold
    else:
        return return_value


def ical_to_markdown(url: str) -> str:
    """Given an Indico event URL, return a Markdown agenda."""
    # url = 'https://indico.desy.de/event/35655/event.ics?scope=contribution'
    if not url.endswith('/'):
        url += '/'
    # Some Indico versions use ?scope=contribution instead - use both!
    try:
        event = Calendar.from_ical(requests.get(f'{url}event.ics?detail=contributions&scope=contribution').text)
    except ValueError:
        return ''
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
