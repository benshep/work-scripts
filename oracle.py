import os
import time
from enum import Enum
from time import sleep
from urllib.parse import quote, urlencode

import pythoncom
import win32com.client as win32
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webdriver import WebDriver

from folders import budget_folder

user_profile = os.path.expanduser('~')

class Browser(Enum):
    firefox = 0
    edge = 1
    chrome = 2
    safari = 4


quoted_path = quote('/shared/STFC Shared/_portal/STFC Projects', safe='')
fusion_url = 'https://fa-evzn-saasfaukgovprod1.fa.ocs.oraclecloud.com'
apps: dict[tuple[str, ...], str] = {
    ('obi',): f'https://obi.ssc.rcuk.ac.uk/analytics/saw.dll?dashboard&PortalPath={quoted_path}',  # TODO: update?
    ('taleo',): "https://careersportal.taleo.net/enterprise/fluid?isNavigationCompleted=true",
    ('absences',): fusion_url + '/fscmUI/redwood/absences/existing-absences/view-summary',
    ('team_timecards',): fusion_url + '/fscmUI/redwood/human-resources/feature/launch?' + \
                         urlencode({'vbFlowStringKey': 'addTimeCard', 'action': 'MyTeam_AddTimeCardTLM',
                                    'context': 'MyTeam', 'useSessionStoredFilters': 'true', 'vbAppUi': 'time',
                                    'vbcsFlow': 'timecards', 'vbPage': 'add-timecard',
                                    'vbPageParams': 'pCurrent=false#userContext=LINE_MANAGER'}),
    ('payslips',): fusion_url + '/fscmUI/redwood/payslips/payslips/launch',
    ('home',): fusion_url
}


def go_to_oracle_page(*links: str,
                      browser: Browser = Browser.edge,  # changed from Firefox for SSO
                      manual_login: bool = False,
                      show_window: bool = False) -> WebDriver:
    """Open a selenium web driver and log in to Oracle e-Business Suite, opening the specified tuple of links.
    :param browser: Firefox, Edge, Chrome, Ie, or Safari
    :param links: text of links to follow, or 'obi' or 'taleo' to open those sites
    :param manual_login: use the standard profile, and prompt to log in manually
    :param show_window: make the browser window visible
    :return: web driver instance, so you can do more things with it"""

    # which URL to go to? OBI, Taleo, or just default to Oracle homepage
    url = apps.get(links, fusion_url + '/fscmUI/faces/FuseWelcome')
    firefox_options = webdriver.FirefoxOptions()
    firefox_options.add_argument('--MOZ_LOG=')  # try to silence errors
    edge_options = webdriver.EdgeOptions()
    edge_options.add_argument('--log-level=3')  # try to silence errors
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--log-level=3')  # try to silence errors
    if not (show_window or manual_login):  # try to open an invisible browser window
        os.environ['MOZ_HEADLESS'] = '1'
        edge_options.add_argument("--headless=new")
        chrome_options.add_argument("--headless=new")
    # avoid the following error: (https://github.com/MicrosoftEdge/EdgeWebDriver/issues/189#issuecomment-2689338112)
    # selenium.common.exceptions.SessionNotCreatedException: Message: client_session not created:
    # probably user data directory is already in use,
    # please specify a unique value for --user-data-dir argument, or don't use --user-data-dir
    edge_options.add_argument('--edge-skip-compat-layer-relaunch')
    if manual_login:
        print(f'This script will now launch a browser window ({browser.name.title()}) to log in to Oracle.')
        print('Return to this screen when you have logged in.')
        input('Press ENTER to start browser, or Ctrl-C to exit: ')

    match browser:
        case Browser.firefox:
            if not manual_login:
                # use a specifically-created Selenium profile, where I've already logged in - no need to enter credentials again
                profile_dir = os.path.join(user_profile, 'AppData', 'Roaming', 'Mozilla', 'Firefox', 'Profiles')
                selenium_profile = next(folder for folder in os.listdir(profile_dir) if folder.endswith('.Selenium'))
                firefox_options.profile = webdriver.FirefoxProfile(os.path.join(profile_dir, selenium_profile))
            # in Ubuntu, Selenium can't locate Firefox - help it out (thanks Jools Wills)
            driver = '/snap/bin/geckodriver'
            service = webdriver.FirefoxService(
                executable_path=(driver if os.path.isfile(driver) else None)
            )
            web = webdriver.Firefox(options=firefox_options, service=service)
        case Browser.edge:
            os.environ["SE_DRIVER_MIRROR_URL"] = "https://msedgedriver.microsoft.com"  # temp fix for wrong URL
            web = webdriver.Edge(options=edge_options)
        case Browser.chrome:
            web = webdriver.Chrome(options=chrome_options)
        case Browser.safari:
            web = webdriver.Safari()  # untested
        case _:
            raise ValueError(f'Invalid browser {browser}')

    web.implicitly_wait(10)  # add an automatic wait to the browser handling
    web.set_window_size(1920, 1080)  # make it big so all elements are displayed - maximize doesn't work for Edge
    web.get(url)  # go to the URL
    # sleep(2)
    for _ in range(10):  # try a few times
        for login_button in web.find_elements(By.CLASS_NAME, 'oj-button-text'):
            if login_button.text == 'UKRI':
                login_button.click()
                break
        else:  # no login button found
            continue
        break
    # 'Pick an account' screen
    sleep(5)
    # to get my name (needed if others using script!)
    # name = subprocess.check_output('net user "%USERNAME%" /domain | FIND /I "Full Name"', shell=True)
    # full_name = name.replace(b"Full Name", b"").strip().decode('utf-8')
    for cell in web.find_elements(By.CLASS_NAME, 'table-cell'):
        if cell.text.startswith('Shepherd, Ben'):  # full_name
            cell.click()
            break
    if manual_login:
        time.sleep(10)  # wait for error messages to appear
        print('\nBrowser started.')
        input('Press ENTER after logging in to Oracle, or Ctrl-C to exit: ')

    if links not in apps:  # list of links rather than an app to open
        for link in links:
            web.find_element(By.LINK_TEXT, link).click()
    # if browser == Browser.edge:
    #     web.set_permissions('clipboard-read', 'denied')  # allow
    elif links == 'obi':  # open STFC Projects dashboard
        web.find_element(By.LINK_TEXT, 'STFC Projects - Transactions').click()
    time.sleep(2)
    return web


def file_list(files: list):
    def wrapped(f):
        f.file_list = files
        return f
    return wrapped


@file_list([os.path.join(budget_folder, xls_file)
            for xls_file in ('OBI Staff Bookings.xls', 'OBI Finance Report.xls')])
def convert_obi_files():
    """Convert XLS files received from the automated OBI process to XLSX files."""
    pythoncom.CoInitialize()  # otherwise we get the "CoInitialize has not been called" error
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = 0
    # done_any = False
    for file in convert_obi_files.file_list:
        path, filename = os.path.split(file)
        new_file = os.path.join(path, 'OBI processed', filename + 'x')  # i.e. .xlsx filename
        print(new_file)
        if os.path.exists(new_file):
            # if os.path.getmtime(new_file) > os.path.getmtime(file):
            #     continue
            os.remove(new_file)  # otherwise Excel's SaveAs method will prompt about overwriting
        wb = excel.Workbooks.Open(file)
        wb.SaveAs(new_file, FileFormat=51)  # .xlsx extension
        wb.Close()
        # done_any = True
    # excel.Application.Quit()
    # if done_any:
        # trigger a Power Automate flow to run an Office Script to tidy up the files
        # mkstemp(dir=os.path.join(budget_folder, 'pa_trigger'))
    return True


if __name__ == '__main__':
    # web = go_to_oracle_page('absences', show_window=True)
    # sleep(100)
    # web.quit()
    convert_obi_files()
