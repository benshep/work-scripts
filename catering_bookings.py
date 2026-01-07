import os
from datetime import datetime
from platform import node
from time import sleep

from selenium import webdriver
from selenium.webdriver.common.by import By

from hosp_credentials import username, password


def get_bookings(test_mode: bool = False) -> str | bool:
    """Open the STFC Hospitality bookings and find the DL ones booked for the next week."""
    if node() != 'DDAST0025':  # only run on desktop
        return False
    toast = ''
    if not test_mode:
        os.environ['MOZ_HEADLESS'] = '1'
    web = webdriver.Firefox()
    web.implicitly_wait(10)
    web.maximize_window()
    base_url = 'https://stfc.hospitalitybookings.com/updated'
    web.get(f'{base_url}/default.aspx')
    web.find_element(By.ID, 'menuLoginText').click()  # Sign In
    web.find_element(By.ID, 'txtUsername').send_keys(username)
    web.find_element(By.ID, 'txtPassword').send_keys(password)
    web.find_element(By.ID, 'btnSubmit').click()
    sleep(2)
    web.get(f'{base_url}/menu/ViewCurrentBookings.aspx')
    date_dropdown = web.find_element(By.ID, 'MainPlaceHolder_ddlDate')
    date_dropdown.click()
    date_options = date_dropdown.find_elements(By.TAG_NAME, 'option')
    next(option for option in date_options if option.text == 'Next Week').click()
    web.find_element(By.ID, 'MainPlaceHolder_cbAllRecords').click()  # 'all records' checkbox
    web.find_element(By.ID, 'MainPlaceHolder_Button1').click()  # Search
    for i in range(2):
        sleep(2)
        table = web.find_element(By.ID, 'MainPlaceHolder_DiaryGridView')
        rows = table.find_elements(By.TAG_NAME, 'tr')
        if i == 0:  # just do once, then refresh table and rows
            headers = rows[0].find_elements(By.TAG_NAME, 'th')
            headers[1].find_element(By.TAG_NAME, 'a').click()  # sort by start date
    now = datetime.now()
    for row in rows[1:]:  # skip header row
        cells = row.find_elements(By.TAG_NAME, 'td')
        if len(cells) < 7:
            continue
        location = cells[4].text
        # exclude RAL locations
        banned_locations = {'Diamond', 'Electron', 'Oxford Space Systems', 'EPAC Building', 'NSTF R114'}
        banned_prefixes = ('R', 'I', 'Hartree')
        if location.startswith(banned_prefixes) or location in banned_locations:
            continue
        start_time = cells[1].text
        end_time = cells[2].text
        room = cells[5].text.strip() or location  # use location if room is blank
        title = cells[6].text
        start_datetime = datetime.strptime(start_time, '%d/%m/%y %H:%M')
        end_datetime = datetime.strptime(end_time, '%d/%m/%y %H:%M')
        if start_datetime.date() <= now.date() and now < end_datetime:  # today, not finished yet?
            toast += f'{start_time[-5:]}-{end_time[-5:]} {title}, {room}\n'
        print(start_time, end_time, title, location, room, sep='\t')
        print(f'{start_time=}, {end_time=}, {title=}, {location=}, {room=}', sep='\t')

    if not test_mode:
        web.quit()
    return toast


if __name__ == '__main__':
    print(get_bookings(test_mode=False))
