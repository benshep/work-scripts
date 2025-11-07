import os
from time import sleep

from winsound import Beep

from work_folders import budget_folder


def monitor_file(filename: str) -> None:
    """Monitor a file for changes."""
    original_mod_time = os.path.getmtime(filename)
    print(f'Monitoring {filename}, last changed {original_mod_time}')
    while True:
        mod_time = os.path.getmtime(filename)
        if mod_time != original_mod_time:
            print('Modified', mod_time)
            Beep(880, 1000)
        sleep(1)

if __name__ == '__main__':
    monitor_file(os.path.join(budget_folder, 'OBI Finance Report.xlsx'))
