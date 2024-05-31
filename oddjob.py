import os
from random import choice
from trello import TrelloClient
import trello_auth


def random_task():
    """Pick a task at random from my Trello to-do list."""
    os.system('title 🎲 Random Task')
    trello_client = TrelloClient(**trello_auth.oauth)
    to_do_board = next(board for board in trello_client.list_boards() if board.name == 'To-Do List')
    lists = to_do_board.open_lists()
    # strip emoji and space from name (won't be displayed nicely in a command prompt window)
    for i, task_list in enumerate(lists):
        print(i, task_list.name.encode('ascii', 'ignore').decode('ascii').strip())
    try:
        i = min(int(input('Which list for a random task [0]: ')), len(lists) - 1)
    except ValueError:
        i = 0
    task = choice(lists[i].list_cards())
    os.startfile(task.url)


if __name__ == '__main__':
    random_task()
