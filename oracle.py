import time
import os
from urllib.parse import quote

from selenium import webdriver
from selenium.webdriver.common.by import By

from folders import user_profile


def go_to_oracle_page(*links: str,
                      show_window: bool = False) -> webdriver.firefox.webdriver.WebDriver:
    """Open a selenium web driver and log in to Oracle e-Business Suite, opening the specified tuple of links.
    If links is 'obi' or 'taleo', open that app instead.
    Returns the web driver instance, so you can do more things with it."""

    quoted_path = quote('/shared/STFC Shared/_portal/STFC Projects', safe='')
    apps = {('obi',): f'https://obi.ssc.rcuk.ac.uk/analytics/saw.dll?dashboard&PortalPath={quoted_path}',
            ('taleo',): "https://careersportal.taleo.net/enterprise/fluid?isNavigationCompleted=true"}
    url = apps.get(links, 'https://ebs.ssc.rcuk.ac.uk/OA_HTML/AppsLogin')
    if not show_window:
        os.environ['MOZ_HEADLESS'] = '1'
    # use a specifically-created Selenium profile, where I've already logged in - no need to enter credentials again
    profile_dir = os.path.join(user_profile, 'AppData', 'Roaming', 'Mozilla', 'Firefox', 'Profiles')
    selenium_profile = next(folder for folder in os.listdir(profile_dir) if folder.endswith('.Selenium'))
    options = webdriver.FirefoxOptions()
    options.profile = webdriver.FirefoxProfile(os.path.join(profile_dir, selenium_profile))
    web = webdriver.Firefox(options=options)
    web.implicitly_wait(10)
    web.get(url)
    web.find_element(By.ID, 'ssoUKRIBtn').click()  # UKRI User Login
    if links not in apps:  # list of links rather than an app to open
        for link in links:
            web.find_element(By.LINK_TEXT, link).click()
    elif links == 'obi':  # open STFC Projects dashboard
        web.find_element(By.LINK_TEXT, 'STFC Projects - Transactions').click()
    time.sleep(2)
    return web


if __name__ == '__main__':
    go_to_oracle_page('STFC OTL Timecards', 'Time', 'Recent Timecards', show_window=True)
