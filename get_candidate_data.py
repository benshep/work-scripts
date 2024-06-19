import os
from time import sleep
from oracle import go_to_oracle_page
from folders import user_profile, downloads_folder


def get_candidate_data(test_mode=False):
    toast = ''
    os.chdir(downloads_folder)
    cvs_folder = get_cvs_folder()
    # print('Opening Taleo')
    web = go_to_oracle_page('taleo', show_window=test_mode)
    try:
        sleep(10)
        web.find_element_by_id('mainmenu-requisitions').click()
        sleep(5)
        # print('Clicking active candidates button')
        web.find_element_by_class_name('table-requisition__cell--candidate-count--wrapper--inner-data').click()
        while True:
            # web.click('All Candidates')
            # sleep(5)
            # web.click('Submissions')
            sleep(5)
            candidates = web.find_elements_by_xpath('//div[@class="table-application__candidate-name"]/a')
            # print(f'Found {len(candidates)} candidates in list')
            for candidate in candidates:
                name = candidate.text.split(' (')[0]
                cv_files = os.listdir(cvs_folder)
                if all(name not in filename for filename in cv_files):
                    # print(f'No files yet for {name}')
                    break
                # print(f'Got files for {name}')
            else:
                # print('Got all files')
                break

            # print(f'Getting files for {name}')
            print(*(name.replace(',', '').split(' ')[:2]), sep='\t')
            sleep(15)
            candidate.click()
            sleep(5)
            # print('Going to attachments')
            web.find_element_by_link_text('Attachments').click()
            sleep(8)

            download_buttons = web.find_elements_by_class_name('attachment__icon--download')
            for i, button in enumerate(download_buttons):
                files = os.listdir()
                button.click()
                web.switch_to.window(web.window_handles[0])  # opened a new tab showing document - switch back to main one
                while True:
                    sleep(5)
                    new_files = os.listdir()
                    if len(new_files) > len(files):
                        break
                new_file = list(set(new_files) - set(files))[0]
                _, ext = os.path.splitext(new_file)
                dest_file = os.path.join(cvs_folder,
                                         f'{name} {i + 1}{ext}' if len(download_buttons) > 1 else f'{name}{ext}')
                os.rename(new_file, dest_file)

            # print('Finding cover letter')
            sleep(5)
            web.find_element_by_link_text('Job Submission').click()
            sleep(5)
            web.find_element_by_id('695-header').click()  # Experience
            sleep(5)
            if cover_letter := web.find_element_by_id('Cover Letter-container'):
                # print('Copying cover letter')
                open(os.path.join(cvs_folder, f'{name} cover letter.md'), 'w').write(cover_letter.text)

            sleep(5)
            # print('Returning to list')
            web.find_element_by_id('ui-id-1').click()  # Back to Submission List
            sleep(8)
            toast += name + '\n'
    finally:
        web.quit()
    return toast


def get_cvs_folder():
    return os.path.join(user_profile, 'STFC', 'Documents', 'Group Leader',
                        'Recruitment', 'Physicist Industrial Placement')


def list_candidates():
    files = os.listdir(get_cvs_folder())
    names = set()
    for filename in files:
        parts = filename.replace(',', '').replace('.', ' ').split(' ')
        names.add(tuple(parts[:2]))
    [print(*name, sep='\t') for name in sorted(list(names))]


if __name__ == '__main__':
    while True:
        get_candidate_data()
