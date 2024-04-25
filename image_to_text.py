# https://ifimageediting.com/image-to-text
# Convert multiple image files to text, now extract info and rename images

import contextlib
import os
from datetime import datetime


def image_to_text():
    user_profile = os.environ['UserProfile']
    receipts_folder = os.path.join(user_profile, 'Documents', 'Travel', 'Receipts')
    os.chdir(receipts_folder)
    screenshots = [file for file in os.listdir() if file.startswith('Screenshot') and file.endswith('.png')]
    for receipt_img in screenshots:
        text_file = f'{receipt_img[:-4]}.txt'
        receipt_text = open(text_file).read().splitlines()
        for line in receipt_text:
            with contextlib.suppress(ValueError):
                date = datetime.strptime(line, '%d %b %Y')
                break
        amount = max(float(line[1:]) for line in receipt_text if line.startswith('¬'))
        os.rename(receipt_img, f'{date.strftime("%Y-%m-%d")} Lime €{amount:.2f}.png')


if __name__ == '__main__':
    image_to_text()