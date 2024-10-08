import subprocess
import re
import feedparser


def list_packages():
    output = subprocess.check_output('conda list').decode('utf-8').split('\r\n')
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
        match = re.findall(r'([\w\-_]+)-(\d[\w\.,\-]+)', entry.title)
        if not match:
            continue
        name, version = match[0]
        packages[name] = version
    return packages


def find_new_python_packages():
    installed_packages = list_packages()
    available_packages = get_rss()
    toast = ''
    for name, (version, build_channel) in installed_packages.items():
        new_version = available_packages.get(name, '')
        if new_version > version:
            if 'numpy' in name and new_version == '2.0.0':
                continue  # numpy upgrades aren't working right now
            toast += (f'pip install {name} -U' if 'pypi' in build_channel else f'conda upgrade {name}') + \
                     f' # {version} --> {new_version}' + '\n'
    return toast


if __name__ == '__main__':
    find_new_python_packages()
