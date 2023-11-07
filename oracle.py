import time
import os
from webbot import Browser
from stfc_credentials import username, password
from selenium import webdriver
import selenium.common.exceptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as condition
from selenium.webdriver.common.by import By

# If this is run from a Python instance with no console, it will show a console window with Chrome messages.
# To prevent this, go into the site-packages\selenium\webdriver\common\service.py. At the top, add:
# from win32process import CREATE_NO_WINDOW
# In the Service.start() function, add the argument creationflags=CREATE_NO_WINDOW to the subprocess.Popen call
# See https://stackoverflow.com/questions/33983860/hide-chromedriver-console-in-python


def go_to_oracle_page(links=(), show_window=False, use_obi=False, selenium=False):
    """Open a webbot instance and log in to Oracle, opening the link(s) specified therein.
    links can be a string or a list of strings.
    Returns the webbot instance, so you can do more things with it."""
    url = 'https://obi.ssc.rcuk.ac.uk/analytics/saw.dll?dashboard&PortalPath=%2Fshared%2FSTFC%20Shared%2F_portal%2FSTFC%20Projects' \
        if use_obi else 'https://ebs.ssc.rcuk.ac.uk/OA_HTML/AppsLogin'
    if isinstance(links, str):
        links = [links, ]
    if selenium:
        if not show_window:
            os.environ['MOZ_HEADLESS'] = '1'
        web = webdriver.Firefox()
        web.get(url)
        # UKRI User Login
        wait = WebDriverWait(web, 5)
        wait.until(condition.presence_of_element_located((By.ID, 'ssoUKRIBtn'))).click()
        wait.until(condition.presence_of_element_located((By.ID, 'i0116'))).send_keys(username)
        web.find_element_by_id('idSIButton9').click()  # Next
        wait.until(condition.invisibility_of_element_located((By.CLASS_NAME, 'lightbox-cover')))
        wait.until(condition.presence_of_element_located((By.ID, 'i0118'))).send_keys(password)
        web.find_element_by_id('idSIButton9').click()  # Sign in
        wait.until(condition.invisibility_of_element_located((By.CLASS_NAME, 'lightbox-cover')))
        web.find_element_by_id('idSIButton9').click()  # Stay signed in
        for link in links:
            wait.until(condition.presence_of_element_located((By.LINK_TEXT, link))).click()
        time.sleep(2)

    else:
        web = Browser(showWindow=show_window)
        web.go_to(url)
        web.driver.execute_script('OpenUKRI()')
        wait_for(web, 'Next')
        web.type(username, 'username')
        web.click('Next')
        wait_for(web, 'Sign in')
        web.type(password, 'password')
        web.click('Sign in')
        wait_for(web, 'Yes')
        web.click('Yes')  # stay signed in

        time.sleep(2)
        for link in links:
            web.click(link)
            time.sleep(2)
    return web


def type_into(web, element, text):
    """Type this text into a given element. After finding the elements, it seems faster than webbot's 'type' method."""
    element.click()
    web.press(f'{web.Key.CONTROL}a')  # select all
    web.type(text)
    time.sleep(0.5)


def wait_for(web, text, timeout=10):
    """Wait for some text to appear in a webbot instance."""
    for _ in range(timeout):
        try:
            if web.exists(text):
                break
        except selenium.common.exceptions.StaleElementReferenceException as e:
            pass
        finally:
            time.sleep(1)
    else:
        raise TimeoutError(f'Timed out waiting for text "{text}" to appear.')


if __name__ == '__main__':
    web = go_to_oracle_page(show_window=True, selenium=True, links='STFC OTL Supervisor')

