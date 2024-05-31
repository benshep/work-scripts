import os
import webbot
from time import sleep
from stfc_credentials import username, password


class InformationFetchFailure(Exception):
    pass


bad_chars = str.maketrans({char: None for char in '*?/\\<>:|"'})  # can't use these in filenames


def remove_bad_chars(filename: str):
    return filename.translate(bad_chars)


def get_risk_assessments():
    """Open the SHE Assure website, and bulk-download project risk assessments to the local file system."""
    user_profile = os.environ['UserProfile']
    downloads_folder = os.path.join(user_profile, 'Downloads')
    ras_folder = os.path.join(user_profile, 'OneDrive - Science and Technology Facilities Council',
                              'Documents', 'Safety', 'RAs')
    os.makedirs(ras_folder, exist_ok=True)
    os.chdir(downloads_folder)

    web = webbot.Browser()
    web.go_to('https://uk.sheassure.net/stfc/Risk/ActivityRisk/Page/1')
    wait_for(web, 'Single Sign-On')
    web.click('Microsoft')
    wait_for(web, 'Sign in')
    web.type(username, into='username')
    web.click('Next')
    wait_for(web, 'STFC Federation Service')
    web.type(password, into='password')
    web.click('Sign in')
    wait_for(web, 'Stay signed in?')
    web.click('Yes')
    wait_for(web, 'Activity Risk Assessment')

    paperclip_links = web.find_elements(tag='i', classname='fa-paperclip')
    for i in range(len(paperclip_links)):
        try:
            print(i)
            link = paperclip_links[i]
            link.click()
            # get info about RA
            info = web.find_elements(tag='ul', classname='list_information')[0]
            attributes = [tag.text for tag in info.find_elements_by_tag_name('span')]
            values = [tag.get_attribute('title') for tag in info.find_elements_by_tag_name('a')]
            info_dict = {attribute.rstrip(':'): value for attribute, value in zip(attributes, values)}
            reference_number = info_dict['Reference']
            title = info_dict['Assessment Title'].split('\n')[0]  # just take first line
            title = remove_bad_chars(title).strip()  # remove whitespace from start and end
            attachment_links = web.find_elements(tag='span', classname='fa-file')
            folder_name = f'{reference_number} {title}'
            print(folder_name)
            ra_folder = os.path.join(ras_folder, folder_name)
            os.makedirs(ra_folder, exist_ok=True)
            for attachment in attachment_links:
                parent = attachment.find_element_by_xpath('..')
                filename = parent.text
                print(filename)
                attachment.click()  # download attachment
                for _ in range(30):
                    sleep(1)
                    if os.path.exists(filename):
                        break
                else:
                    raise InformationFetchFailure('Data export not successful')
                os.rename(filename, os.path.join(ra_folder, filename))
        except InformationFetchFailure:
            print("Couldn't fetch attachment")
        finally:
            web.go_back()
            paperclip_links = web.find_elements(tag='i', classname='fa-paperclip')
    web.driver.quit()


def wait_for(web, text, timeout=10):
    """Wait for some text to appear in a webbot instance."""
    for _ in range(timeout):
        if web.exists(text):
            break
        sleep(1)
    else:
        raise TimeoutError(f'Timed out waiting for text "{text}" to appear.')


if __name__ == '__main__':
    get_risk_assessments()
