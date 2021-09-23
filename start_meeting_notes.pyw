#!python3
# -*- coding: utf-8 -*-

"""start_meeting_notes
Ben Shepherd, 2021
Look through today's events in an Outlook calendar,
and start a Markdown notes file for  the closest one to the current time.
Command line arguments:
    second      find the second meeting if there are more than one starting at the same time (order is arbitrary)
"""

import os
import re
import sys
from itertools import chain
import outlook

user_profile = os.environ['UserProfile']


def folder_match(name: str, test_against: str):
    """Case-insensitive string comparison. Checks that name is a folder as well."""
    return name.casefold() in test_against.casefold() and os.path.isdir(name) and name != 'Zoom'


def create_note_file(skip_one=False):
    """Start a file for notes relating to the given meeting. Find a relevant folder in the user's Documents folder,
     searching first the subject then the body of the meeting for a folder name. Use Other if none found.
     Don't use the Zoom folder (this is often found in the meeting body).
     The file is in Markdown format, with the meeting title, date and attendees filled in at the top."""

    current_events = outlook.get_current_events()
    print('Current events:', [event.Subject for event in current_events])

    index = 1 if skip_one else 0
    if len(current_events) <= index:
        return
    meeting = current_events[index]

    os.chdir(os.path.join(user_profile, 'Documents'))
    files = os.listdir()
    subject = meeting.Subject
    filter_subject = filter(lambda file: folder_match(file, subject), files)
    filter_body = filter(lambda file: folder_match(file, meeting.Body), files)
    # If we can't work out what folder based on subject or body: put in Other folder
    folder = next(chain(filter_subject, filter_body, ['Other']))
    os.chdir(folder)
    attendees = '; '.join([meeting.RequiredAttendees, meeting.OptionalAttendees])
    print(attendees)
    people_list = ', '.join(format_name(person_name) for person_name in filter(None, attendees.split('; ')))
    start = outlook.get_meeting_time(meeting)
    meeting_date = start.strftime("%#d/%#m/%Y")  # no leading zeros
    text = f'# {subject}\n\n*{meeting_date}. {people_list}*\n\n'
    bad_chars = str.maketrans({char: ' ' for char in '*?/\\<>:|"'})  # can't use these in filenames
    filename = f'{start.strftime("%Y-%m-%d")} {subject.translate(bad_chars)}.md'
    open(filename, 'a', encoding='utf-8').write(text)
    os.startfile(filename)


# match "Surname, Firstname (ORG,DEPT,GROUP)" - last bit in brackets is optional
name_format = re.compile(r'(.+?), ([^ ]+)( \(.+?,.+?,.+?\))?')


def format_name(person_name):
    """Convert display name to more readable Firstname Surname format. Works with Surname, Firstname (ORG,DEPT,GROUP)
    or firstname.surname@company.com."""
    if match := name_format.match(person_name):
        return f'{match.group(2)} {match.group(1)}'  # Firstname Surname
    elif '@' in person_name:  # email address
        name, domain = person_name.split('@')
        return name.title().replace('.', ' ')
    else:
        return person_name.title()


if __name__ == '__main__':
    create_note_file('second' in sys.argv)
