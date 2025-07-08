import os
import sys
import subprocess
import re
import feedparser
from natsort import natsorted
from platform import node

import folders

sys.path.append(os.path.join(folders.misc_folder, 'Scripts'))
from pushbullet import Pushbullet  # to show notifications
from pushbullet_api_key import api_key  # local file, keep secret!


def list_packages() -> dict[str, tuple[str, str]]:
    version = sys.version_info
    # guess name if env not defined: convention is e.g. py313
    env_name = os.environ.get('CONDA_DEFAULT_ENV', f'py{version.major}{version.minor}')
    # list command will fail and raise an exception here if env_name is invalid
    conda_path = os.path.join(folders.user_profile, "Miniconda3", "Scripts", "conda.exe")
    output = run_command([conda_path, 'list', '-n', env_name])
    conda_packages = {}
    for line in output:
        if line.startswith('#'):
            continue
        try:
            name, version, build_channel = re.split(r'\s+', line, maxsplit=2)
        except ValueError:
            continue
        conda_packages[name] = (version, build_channel)
    return conda_packages


def get_rss() -> dict[str, str]:
    url = 'https://repo.anaconda.com/pkgs/rss.xml'
    feed = feedparser.parse(url)
    if feed.status != 200:
        return {}
    packages = {}
    for entry in feed.entries:
        if match := re.findall(r'([\w\-_]+)-(\d[\w\.,\-]+)', entry.title):
            name, version = match[0]
            packages[name] = version
    return packages


def check_chocolatey_packages() -> str:
    """Use Chocolatey to check if any of its packages need updating."""
    if sys.platform != 'win32':  # choco is a Windows thing
        return ''
    outdated = run_command(['choco', 'outdated', '-r'])
    to_upgrade = ''
    for line in outdated:  # e.g. autohotkey|1.1.37.1|2.0.19|true
        if line.count('|') < 3:
            continue
        package, old, new, pinned = line.split('|')
        if pinned == 'true':
            continue
        to_upgrade += f'{package}: {old} ➡ {new}\n'
        # display release notes if available
        info = run_command(['choco', 'info', package])
        for info_line in info:
            notes_header = ' Release Notes:'
            if info_line.startswith(notes_header):
                print(package)
                print(info_line[len(notes_header):])
    if to_upgrade:
        print('Running Chocolatey upgrade process')
        # generate the custom event that runs Chocolatey update
        # see https://qtechbabble.wordpress.com/2021/09/09/use-system-events-to-trigger-administrator-scheduled-tasks-from-a-standard-user-account/
        # source: BenShepherdCustomEvent
        # id: 8072
        trigger_update()
    return to_upgrade


def trigger_update() -> None:
    powershell_command = 'powershell -command "Write-EventLog -LogName Application -Source ' + \
                         "'BenShepherdCustomEvent' -EntryType Information -EventId 8072 -Message 'Admin session started.'\""
    run_command(powershell_command)


def run_command(command: str | list[str]) -> list[str]:
    """Runs a command and returns the output split into lines."""
    return subprocess.check_output(command).decode('utf-8').split('\r\n')


def find_new_python_packages() -> str:
    installed_packages = list_packages()
    available_packages = get_rss()
    conda_new = []
    pip_new = []
    for name, (version, build_channel) in installed_packages.items():
        new_version = available_packages.get(name, '')
        versions = [new_version, version]
        if versions != natsorted(versions):  # natural sort that puts e.g. 3.13.1 after 3.9.21
            print(f'{name}: {new_version} available, got {version}')
            if ('numpy' in name and new_version == '2.0.0') or name == 'certifi':
                continue  # numpy upgrades aren't working right now
            (pip_new if 'pypi' in build_channel else conda_new).append(name)
    choco_new = check_chocolatey_packages()
    return ''.join([
        f'conda upgrade {" ".join(conda_new)}\n' if conda_new else '',
        f'pip install {" ".join(pip_new)} -U\n' if pip_new else '',
        choco_new
    ])


def check_updated() -> None:
    """Read the list of updated packages and notify."""
    toast = ''
    upgrade_output_file = f'choco-upgrade-{node()}.txt'
    upgrade_output = open(upgrade_output_file).read().splitlines()
    try:
        i = upgrade_output.index('Upgraded:')
    except ValueError:
        try:
            i = upgrade_output.index('Failures')
        except ValueError:
            toast = f'No upgrades, no failures\nCheck {upgrade_output_file}'
    if not toast:
        if 'pinned' in upgrade_output[-1]:  # don't need to notify that packages are pinned
            upgrade_output.pop(-1)
        if upgrade_output[-1] == 'Warnings:':  # if that's the only warning, skip the headline
            upgrade_output.pop(-1)
        toast = '\n'.join(upgrade_output[i:])
    Pushbullet(api_key).push_note('🍫 Chocolatey updates', toast)


if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == 'check_updated':  # this runs from choco-update.bat - don't change
        check_updated()
    else:
        print(trigger_update())
