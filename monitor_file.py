import os
from time import sleep
from winsound import Beep
from folders import docs_folder


def monitor_file(filename):
    """Monitor a file for changes."""
    original_mtime = os.path.getmtime(filename)
    print(f'Monitoring {filename}, last changed {original_mtime}')
    while True:
        mod_time = os.path.getmtime(filename)
        if mod_time != original_mtime:
            print('Modified', mod_time)
            Beep(880, 1000)
        sleep(1)

if __name__ == '__main__':
    monitor_file(os.path.join(docs_folder, 'Budget', 'OBI Finance Report.xlsx'))
