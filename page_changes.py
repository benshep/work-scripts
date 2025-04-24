import os
import pickle
import difflib
from datetime import datetime
from glob import glob
from types import SimpleNamespace
from selenium import webdriver
from selenium.webdriver.common.by import By
from folders import user_profile


def check_page_changes(show_window: bool = False) -> str:
    pickle_file = 'page_changes.db'
    prev_dir = os.getcwd()
    os.chdir(os.path.split(__file__)[0])  # script dir
    check_pages: list = pickle.load(open(pickle_file, 'rb'))
    url_list_filename = 'page_changes.txt'
    url_list = open(url_list_filename).read().splitlines()
    urls = []
    for line in url_list:
        name, url = line.split('\t')
        urls.append(url)
        if url not in (page.url for page in check_pages):  # new one
            check_pages.append(SimpleNamespace(name=name, url=url, old_content=None))
    check_pages = [page for page in check_pages if page.url in urls]  # remove any that aren't in the text list

    if not show_window:
        os.environ['MOZ_HEADLESS'] = '1'
    profile_dir = os.path.join(user_profile, 'AppData', 'Roaming', 'Mozilla', 'Firefox', 'Profiles')
    selenium_profile = next(folder for folder in os.listdir(profile_dir) if folder.endswith('.Selenium'))
    options = webdriver.FirefoxOptions()
    options.profile = webdriver.FirefoxProfile(os.path.join(profile_dir, selenium_profile))
    web = webdriver.Firefox(options=options)
    web.implicitly_wait(10)
    toast = ''
    for page in check_pages:
        print(page.url)
        if page.url.startswith('http'):
            web.get(page.url)
            new_content = web.find_element(By.CLASS_NAME, 'mainContent').text
            if page.old_content is None:  # new URL - no content fetched yet
                page.old_content = new_content
            elif difflib.SequenceMatcher(a=page.old_content, b=new_content).ratio() < 1:  # change
                old_lines = page.old_content.splitlines()
                new_lines = new_content.splitlines()
                html = difflib.HtmlDiff(wrapcolumn=80).make_file(old_lines, new_lines)
                open(f'page_changes {page.name}.html', 'w', encoding='utf-8').write(html)
                delta = difflib.Differ().compare(old_lines, new_lines)
                changed_lines = len([line for line in delta if not line.startswith('  ')])
                toast += f'{page.name}: {changed_lines=}\n'
                page.old_content = new_content
        else:  # file
            file_list = glob(page.url)
            if len(file_list) == 0:
                toast += f'{page.name}: no matching file found\n'
                continue
            filename = max(file_list, key=os.path.getmtime)
            modified_time = os.path.getmtime(filename)
            if page.old_content is None:  # new file - never checked
                page.old_content = modified_time
            elif modified_time > page.old_content:  # changed since last checked
                toast += f'{page.name}: modified {datetime.fromtimestamp(modified_time).strftime("%d %b at %H:%M")}\n'
            page.old_content = modified_time

    # noinspection PyTypeChecker
    pickle.dump(check_pages, open(pickle_file, 'wb'))
    web.quit()
    os.chdir(prev_dir)
    return toast


if __name__ == '__main__':
    print(check_page_changes(True))
