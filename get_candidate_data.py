import os
from time import sleep
from stfc_credentials import username, password
from webbot import Browser
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from pushbullet import Pushbullet  # to show notifications
from pushbullet_api_key import api_key  # local file, keep secret!


def get_data_firefox():
    driver = webdriver.Firefox()
    driver.get("https://careersportal.taleo.net/enterprise/fluid?isNavigationCompleted=true")
    driver.find_element_by_id('ssoUKRIBtn').click()  # UKRI User Login
    driver.find_element_by_id('i0116').send_keys(username)
    driver.find_element_by_id('idSIButton9').click()  # Next
    driver.find_element_by_id('i0118').send_keys(password)
    driver.find_element_by_id('idSIButton9').click()  # Sign in
    driver.find_element_by_id('idSIButton9').click()  # Stay signed in
    driver.find_element_by_id('mainmenu-requisitions').click()  # Requisitions


def get_candidate_data():
    toast = ''
    user_profile = os.environ['UserProfile']
    downloads_folder = os.path.join(user_profile, 'Downloads')
    os.chdir(downloads_folder)
    job_title = 'Physicist Industrial Placement'
    cvs_folder = os.path.join(user_profile, 'Documents', 'Group Leader', job_title)
    url = 'https://careersportal.taleo.net/enterprise/fluid'
    while True:
        print('Opening Taleo')
        web = Browser(showWindow=True)
        # web.driver.maximize_window()
        # web.driver.minimize_window()
        web.go_to(url)
        sleep(2)
        web.click(id='ssoUKRIBtn')
        sleep(2)
        web.type(username, into='email')
        web.click('Next')
        sleep(2)
        web.type(password, into='password')
        web.click('Sign in')
        sleep(2)
        web.click('No')
        sleep(2)
        web.click('Taleo Recruiting')
        sleep(2)
        web.go_to(url)
        web.switch_to_tab(1)
        sleep(2)
        web.click('Requisitions')
        sleep(4)
        print('Clicking active candidates button')
        web.click(classname='table-requisition__cell--candidate-count--wrapper--inner-data', tag='span')
        # web.click('All Candidates')
        # sleep(4)
        # web.click('Submissions')
        sleep(4)
        candidates = web.find_elements(classname='table-application__candidate-name', tag='div')
        for candidate in candidates:
            name = candidate.text.split(' (')[0]

            cv_files = os.listdir(cvs_folder)
            if all(name not in filename for filename in cv_files):
                print(f'No files yet for {name}')
                break
            print(f'Got files for {name}')
        else:
            print('Got all files')
            break

        print(f'Getting files for {name}')
        candidate.click()
        sleep(2)
        print('Going to attachments')
        web.click('Attachments')
        sleep(8)

        download_buttons = web.find_elements(classname='attachment__icon--download', tag='span')
        for i, button in enumerate(download_buttons):
            files = os.listdir()
            button.click()
            while True:
                new_files = os.listdir()
                if len(new_files) > len(files):
                    break
                sleep(1)
            new_file = list(set(new_files) - set(files))[0]
            _, ext = os.path.splitext(new_file)
            dest_file = os.path.join(cvs_folder, f'{name} {i + 1}{ext}' if len(download_buttons) > 1 else f'{name}{ext}')
            os.rename(new_file, dest_file)

        print('Finding cover letter')
        sleep(2)
        web.click('Job Submission')
        sleep(2)
        web.click('Experience')
        sleep(2)
        if cover_letter := web.find_elements(id='Cover Letter-container'):
            print('Copying cover letter')
            open(os.path.join(cvs_folder, f'{name} cover letter.md'), 'w').write(cover_letter[0].text)

        sleep(4)
        print('Returning to list')
        web.click('Back to Submission List')
        sleep(8)
        toast += name + '\n'

    if toast:
        Pushbullet(api_key).push_note('New candidates', toast)


if __name__ == '__main__':
    get_data_firefox()
    