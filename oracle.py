import time
import os
from urllib.parse import quote

from stfc_credentials import username, password
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as condition
from selenium.webdriver.common.by import By


def go_to_oracle_page(links=(), show_window=False):
    """Open a selenium web driver and log in to Oracle e-Business Suite, opening the specified tuple of links.
    If links is 'obi' or 'taleo', open that app instead.
    Returns the web driver instance, so you can do more things with it."""

    quoted_path = quote('/shared/STFC Shared/_portal/STFC Projects', safe='')
    apps = {'obi': f'https://obi.ssc.rcuk.ac.uk/analytics/saw.dll?dashboard&PortalPath={quoted_path}',
            'taleo': "https://careersportal.taleo.net/enterprise/fluid?isNavigationCompleted=true"}
    url = apps.get(links, 'https://ebs.ssc.rcuk.ac.uk/OA_HTML/AppsLogin')
    if not show_window:
        os.environ['MOZ_HEADLESS'] = '1'
    web = webdriver.Firefox()
    web.implicitly_wait(10)
    web.get(url)
    web.find_element(By.ID, 'ssoUKRIBtn').click()  # UKRI User Login
    fill_box('i0116', username, web)
    fill_box('i0118', password, web)
    web.find_element(By.ID, 'idSIButton9').click()  # Stay signed in
    if links not in apps:  # list of links rather than an app to open
        for link in links:
            web.find_element(By.LINK_TEXT, link).click()
    time.sleep(2)
    return web


def fill_box(box_id, value, web):
    """Wait for a username/password box to be shown, fill it, and click 'Next'."""
    wait = WebDriverWait(web, 5)
    web.find_element(By.ID, box_id).send_keys(value)
    web.find_element(By.ID, 'idSIButton9').click()  # Next
    wait.until(condition.invisibility_of_element_located((By.CLASS_NAME, 'lightbox-cover')))


if __name__ == '__main__':
    go_to_oracle_page(show_window=True, links=('STFC OTL Timecards', 'Time', 'Recent Timecards'))

