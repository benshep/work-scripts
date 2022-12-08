import os
import re
import win32com.client as win32
from datetime import datetime
from trello_auth import email_to_board


def todos_from_notes():
    """Go through Markdown files in Documents; find any action points that should be added to my to-do list."""
    outlook = win32.Dispatch('outlook.application')
    user_profile = os.environ['UserProfile']
    documents_folder = os.path.join(user_profile, 'Documents')
    os.chdir(documents_folder)
    notes_checked = 'notes_checked'
    last_checked = os.path.getmtime(notes_checked) if os.path.exists(notes_checked) else 0
    for folder, _, file_list in os.walk(documents_folder):
        for file in file_list:
            if not file.lower().endswith('.md'):
                continue
            try:
                meeting_date = datetime.strptime(file[:10], '%Y-%m-%d')
            except ValueError:  # not a filename starting with YYYY-MM-DD
                continue
            filename = os.path.join(folder, file)
            if os.path.getmtime(filename) < last_checked:  # old file
                continue
            # Now read through the file looking for actions
            text = open(filename, encoding='utf-8', errors='replace').read()
            matches = re.finditer(r'(.*)\*\*(.*)\*\*(.*)', text)  # text in bold
            for match in matches:
                bold_text = match.group(2)
                if bold_text.startswith('Action'):
                    action_name = bold_text[6:].strip(' -:.')
                    if len(action_name) <= 3:
                        continue
                    action_name = action_name[0].upper() + action_name[1:]
                    _, folder_name = os.path.split(folder)
                    meeting_title = file[11:-3]  # the bit between the date and the file extension
                    desc = f'From **{meeting_title}** on {meeting_date.strftime("%#d/%#m/%y")}\n{match.string}'
                    print(meeting_date, meeting_title, action_name, sep='; ')
                    mail = outlook.CreateItem(0)
                    mail.To = email_to_board
                    mail.Subject = action_name + \
                                   ('' if folder_name == 'Other' else f' #{folder_name.replace(" ", "_")}')
                    mail.Body = desc
                    mail.Send()
        open(notes_checked, 'w').write('')  # touch file, so we know when it was done last


if __name__ == '__main__':
    todos_from_notes()
