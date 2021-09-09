import time
from webbot import Browser
from oracle_credentials import username, password


def go_to_oracle_page(links, show_window=False):
    """Open a webbot instance and log in to Oracle, opening the link(s) specified therein.
    links can be a string or a list of strings.
    Returns the webbot instance so you can do more things with it."""
    web = Browser(showWindow=show_window)
    web.go_to('https://ebs.ssc.rcuk.ac.uk/OA_HTML/AppsLogin')
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
