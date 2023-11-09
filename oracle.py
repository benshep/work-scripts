import time
import os
from stfc_credentials import username, password
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as condition
from selenium.webdriver.common.by import By


def go_to_oracle_page(links=(), show_window=False, use_obi=False):
    """Open a selenium web driver and log in to Oracle, opening the link(s) specified therein.
    links can be a string or a list of strings.
    Returns the web driver instance, so you can do more things with it."""
    url = 'https://obi.ssc.rcuk.ac.uk/analytics/saw.dll?dashboard&PortalPath=%2Fshared%2FSTFC%20Shared%2F_portal%2FSTFC%20Projects' \
        if use_obi else 'https://ebs.ssc.rcuk.ac.uk/OA_HTML/AppsLogin'
    if isinstance(links, str):
        links = [links, ]
    if not show_window:
        os.environ['MOZ_HEADLESS'] = '1'
    web = webdriver.Firefox()
    web.get(url)
    # UKRI User Login
    wait = WebDriverWait(web, 5)
    wait.until(condition.presence_of_element_located((By.ID, 'ssoUKRIBtn'))).click()
    fill_box('i0116', username, web)
    fill_box('i0118', password, web)
    web.find_element_by_id('idSIButton9').click()  # Stay signed in
    for link in links:
        wait.until(condition.presence_of_element_located((By.LINK_TEXT, link))).click()
    time.sleep(2)
    return web


def fill_box(box_id, value, web):
    wait = WebDriverWait(web, 5)
    wait.until(condition.presence_of_element_located((By.ID, box_id))).send_keys(value)
    web.find_element_by_id('idSIButton9').click()  # Next
    wait.until(condition.invisibility_of_element_located((By.CLASS_NAME, 'lightbox-cover')))


if __name__ == '__main__':
    web = go_to_oracle_page(show_window=True, links='STFC OTL Supervisor')

