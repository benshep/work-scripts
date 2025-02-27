import os
import subprocess
import re
import feedparser
from natsort import natsorted

from folders import user_profile

def list_packages():
    env_name = os.environ['CONDA_DEFAULT_ENV']
    command = [os.path.join(user_profile, "Miniconda3", "Scripts", "conda.exe"), 'list', '-n', env_name]
    output = subprocess.check_output(command).decode('utf-8').split('\r\n')
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
    return (f'conda upgrade {" ".join(conda_new)}\n' if conda_new else '') + \
        (f'pip install {" ".join(pip_new)} -U' if pip_new else '')


if __name__ == '__main__':
    print(find_new_python_packages())
