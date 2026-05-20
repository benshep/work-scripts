import os
import warnings

from docling import document_converter

from work_folders import sharepoint_folder

def find_appraisals(fy: int):
    """Walk through the appraisals folder, and read in each file from a given FY."""
    fy -= 2000
    file_suffix = f'{fy}-{fy + 1}.docx'
    os.chdir(os.path.join(sharepoint_folder, 'ASTeC-APR - LM Shepherd Ben'))
    for dirpath, _, filenames in os.walk('.'):
        # print(dirpath)
        for filename in filenames:
            if not filename.lower().endswith(file_suffix):
                continue
            print(filename)
            read_appraisal(os.path.join(dirpath, filename))


def read_appraisal(filename: str):
    """Read the appraisal file."""
    warnings.filterwarnings('ignore', category=RuntimeWarning)  # otherwise docling spits out a lot of warnings
    converter = document_converter.DocumentConverter()
    doc = converter.convert(filename)
    for table in doc.document.tables:
        cells = table.data.table_cells
        header = cells[0:2]
        if header[0].text == 'Objective' and header[1].text.startswith('Status'):
            print('Objective rows', table.data.num_rows - 1)
            print(cells_word_count(cells[2::2]), 'words in objectives')
            print(cells_word_count(cells[3::2]), 'words in status')
        elif header[0].text.startswith('Development and training') and header[1].text.startswith('Status'):
            assert cells[-4].text.startswith('Employee comments')
            assert cells[-2].text.startswith('Line Manager comments')
            print('Development and training rows', table.data.num_rows - 5)
            print(cells_word_count(cells[2:-4]), 'words in L&D')
            print('L&D employee comments', item_done(cells[-3].text))
            print('L&D LM comments', item_done(cells[-1].text and cells[-1].text))
        elif header[0].text == 'Employee - Mid-Year review comments' and header[0].text == 'Line Manager - Mid-Year review comments':
            print('Mid-year employee comments', item_done(cells[1].text))
            print('Mid-year LM comments', item_done(cells[6].text))
        elif header[0].text == 'Employee - What have I identified as my highlights/successes throughout the year':
            print('Employee comments', item_done(cells[1].text and cells[1].text))
        elif header[0].text == 'Line Manager Comments':
            print('LM comments', item_done(cells[1].text and cells[1].text))
        elif header[0].text == 'Senior Line Manager Comments':
            print('LM comments', item_done(cells[1].text and cells[1].text))
    print('')


def cells_word_count(table_cells):
    return sum([len(cell.text.split()) for cell in table_cells])


def item_done(text: str) -> str:
    return '✔️' if text and text != '[Insert text here]' else '❌ '

if __name__ == '__main__':
    find_appraisals(2024)