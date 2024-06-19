import os
import re
from datetime import datetime
from trello import TrelloClient
import trello_auth
from folders import docs_folder

def todos_from_notes():
    """Go through Markdown files in Documents; find any action points that should be added to my to-do list."""
    trello_client = TrelloClient(**trello_auth.oauth)
    to_do_board = next(board for board in trello_client.list_boards() if board.name == 'To-Do List')
    new_list = next(task_list for task_list in to_do_board.list_lists() if task_list.name == '💡 New')
    labels = to_do_board.get_labels()
    os.chdir(docs_folder)
    notes_checked = 'notes_checked'
    last_checked = os.path.getmtime(notes_checked) if os.path.exists(notes_checked) else 0
    today = datetime.now()
    toast = ''
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
            print(filename)
            text = open(filename, encoding='utf-8', errors='replace').read()
            matches = re.finditer(r'(.*)\*\*Actions?[ \-.:]+(.{4,})\*\*(.*)', text)  # text in bold
            for match in matches:
                action_name = match[2].strip('. ')  # remove full stops from action name
                action_name = action_name[0].upper() + action_name[1:]  # Sentence case.
                folder_names = folder.lower().split(os.path.sep)[-2:]  # e.g. C:\users\ben\docs\temp -> docs, temp
                meeting_title = file[11:-3]  # the bit between the date and the file extension
                # match[0] is entire match - put it all in for context
                desc = f'From **{meeting_title}** on {meeting_date.strftime("%#d/%#m/%y")}\n{match[0]}'
                print(meeting_date, meeting_title, action_name, sep='; ')

                toast += action_name + '\n'
                card = new_list.add_card(name=action_name, desc=desc)
                for label in labels:
                    if label.name.lower() in folder_names:
                        print(label.name)
                        card.add_label(label)

                # try to figure out if there's a deadline date
                after_text = match.group(3)
                # matches "By 1/2", "deadline 4 Nov", "by 23 November 2042", etc
                deadline_regex = r'(?:[Bb]y|[Dd]eadline) (\d\d?)[/ ](\d\d?|[A-Z][a-z]{2,8})(?:[/ ](?:20)?(\d\d))?'
                if date_match := re.search(deadline_regex, after_text):
                    card.add_label(next(label for label in labels if label.name == 'deadline 📆'))
                    day = date_match[1]
                    month = date_match[2]
                    year = date_match[3] or today.year
                    year = int(year)
                    if year < 100:
                        year += 2000
                    date_format = '%d/%m/%Y' if month.isnumeric() else '%d/%b/%Y'
                    date = datetime.strptime(f'{day}/{month[:3]}/{year}', date_format)
                    if date < today:  # try to deal with dates that go into next year
                        date = datetime.strptime(f'{day}/{month[:3]}/{year + 1}', date_format)
                    date = date.replace(hour=8)
                    print(date)
                    card.set_due(date)
                    card.set_reminder(3 * 24 * 60)  # 3 days before due date

    open(notes_checked, 'w').write('')  # touch file, so we know when it was done last
    return toast


if __name__ == '__main__':
    print(todos_from_notes())
