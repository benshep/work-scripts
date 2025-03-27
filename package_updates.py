import os
import sys
import subprocess
import re
import feedparser
from natsort import natsorted

from folders import user_profile

def list_packages():
    version = sys.version_info
    # guess name if env not defined: convention is e.g. py313
    env_name = os.environ.get('CONDA_DEFAULT_ENV', f'py{version.major}{version.minor}')
    # list command will fail and raise an exception here if env_name is invalid
    output = run_command([os.path.join(user_profile, "Miniconda3", "Scripts", "conda.exe"), 'list', '-n', env_name])
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


def get_rss():
    url = 'https://repo.anaconda.com/pkgs/rss.xml'
    feed = feedparser.parse(url)
    if feed.status != 200:
        return None
    packages = {}
    for entry in feed.entries:
        if match := re.findall(r'([\w\-_]+)-(\d[\w\.,\-]+)', entry.title):
            name, version = match[0]
            packages[name] = version
    return packages


def check_chocolatey_packages():
    """Use Chocolatey to check if any of its packages need updating."""
    outdated = run_command(['choco', 'list', '-r'])
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
        for line in info:
            notes_header = ' Release Notes:'
            if line.startswith(notes_header):
                print(package)
                print(line[len(notes_header):])
    return to_upgrade


def run_command(command: list[str]) -> list[str]:
    """Runs a command and returns the output split into lines."""
    return subprocess.check_output(command).decode('utf-8').split('\r\n')


def find_new_python_packages():
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
    return (f'conda upgrade {" ".join(conda_new)}\n' if conda_new else '') + \
        (f'pip install {" ".join(pip_new)} -U\n' if pip_new else '') + choco_new


if __name__ == '__main__':
    print(find_new_python_packages())
