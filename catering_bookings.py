import os
from datetime import datetime
from time import sleep

from selenium import webdriver
from selenium.webdriver.common.by import By
from hosp_credentials import username, password


def get_bookings(test_mode: bool = False) -> str:
    """Open the STFC Hospitality bookings and find the DL ones booked for the next week."""
    # bookings = open('catering_bookings.txt').read().splitlines()
    toast = ''
    # for booking in bookings[-1:]:
    #     _, end_time, _ = booking.split('\t', maxsplit=2)
    #     if datetime.strptime(end_time, '%d/%m/%y %H:%M') < datetime.now():  # in the past
    #         bookings.remove(booking)
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
    next(option for option in date_dropdown.find_elements(By.TAG_NAME, 'option') if option.text == 'Next Week').click()
    web.find_element(By.ID, 'MainPlaceHolder_cbAllRecords').click()  # 'all records' checkbox
    web.find_element(By.ID, 'MainPlaceHolder_Button1').click()  # Search
    sleep(2)
    table = web.find_element(By.ID, 'MainPlaceHolder_DiaryGridView')
    rows = table.find_elements(By.TAG_NAME, 'tr')
    now = datetime.now()
    for row in rows[1:]:  # skip header row
        cells = row.find_elements(By.TAG_NAME, 'td')
        if len(cells) < 7:
            continue
        location = cells[4].text
        # exclude RAL locations
        if location.startswith(('R', 'I', 'Hartree')) \
                or location in ['Diamond', 'Electron', 'Oxford Space Systems', 'EPAC Building', 'NSTF R114']:
            continue
        start_time = cells[1].text
        end_time = cells[2].text
        room = cells[5].text
        title = cells[6].text
        end_datetime = datetime.strptime(end_time, '%d/%m/%y %H:%M')
        if end_datetime.date() == now.date() and now < end_datetime:  # today, not finished yet?
            toast += f'{start_time[-5:]}-{end_time[-5:]} {title}, {room}\n'
        booking = '\t'.join((start_time, end_time, title, location, room))
        print(booking)
        # if booking not in bookings:
        #     bookings.append(booking)

    # open('catering_bookings.txt', 'w').write('\n'.join(bookings))
    return toast


if __name__ == '__main__':
    print(get_bookings(test_mode=True))
