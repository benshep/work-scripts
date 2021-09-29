import time
from webbot import Browser
from oracle_credentials import username, password

# If this is run from a Python instance with no console, it will show a console window with Chrome messages.
# To prevent this, go into the site-packages\selenium\webdriver\common\service.py. At the top, add:
# from win32process import CREATE_NO_WINDOW
# In the Service.start() function, add the argument creationflags=CREATE_NO_WINDOW to the subprocess.Popen call
# See https://stackoverflow.com/questions/33983860/hide-chromedriver-console-in-python


def go_to_oracle_page(links=[], show_window=False, use_obi=False):
    """Open a webbot instance and log in to Oracle, opening the link(s) specified therein.
    links can be a string or a list of strings.
    Returns the webbot instance so you can do more things with it."""
    web = Browser(showWindow=show_window)
    obi_url = 'https://obi.ssc.rcuk.ac.uk/analytics/saw.dll?dashboard&PortalPath=%2Fshared%2FSTFC%20Shared%2F_portal%2FSTFC%20Projects'
    ebs_url = 'https://ebs.ssc.rcuk.ac.uk/OA_HTML/AppsLogin'
    web.go_to(obi_url if use_obi else ebs_url)
    if not web.exists('Enter your Single Sign-On credentials below'):
        raise RuntimeError('Failed to load Oracle login page')
    web.type(username, 'username')
    web.type(password, 'password')
    web.click('Login')

    time.sleep(2)
    if isinstance(links, str):
        links = [links, ]
    for link in links:
        web.click(link)
        time.sleep(2)
    return web


def type_into(web, element, text):
    """Type this text into a given element. After finding the elements, it seems faster than webbot's 'type' method."""
    element.click()
    web.press(web.Key.CONTROL + 'a')  # select all
    web.type(text)
    time.sleep(0.5)


if __name__ == '__main__':
    go_to_oracle_page()
