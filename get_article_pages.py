import requests
import os
from PIL import Image

def fetch_pages():
    """Download parts of pages from an Article Galaxy article. Save each complete page as a single JPEG file."""
    os.chdir(os.path.join(os.environ['UserProfile'], 'Downloads'))
    parts_per_page = 3
    pages = 9
    width, height = 992, 441
    for page in range(pages):
        page_image = Image.new('RGB', (width, height * parts_per_page))
        for part in range(parts_per_page):
            part_filename = f'{page}-{part}.jpg'
            open(part_filename, 'wb').write(
                requests.get(
                    f'https://www.reprintsdesk.com/landing/rviewer_fs.ashx?o=10323699&r=156480203&pgid={page + 1}&ptid={part + 1}',
                    cookies={
                        "__zlcmid": "1IvmHnxyTM32V5x", "9B68A3865475C7A2503EC4AEE0DCD07B": "0",
                        "ASP.NET_SessionId": "bg5ubhgijgjim5l1rqttuxmq",
                        "incap_ses_1309_422718": "8oRnINAYFimN0dlZOIIqEmBmW2UAAAAAbvCXD5M53jxXEIzMn+A1kA==",
                        "MEM": "SID=56697758&RID=70308591",
                        "nlbi_422718": "qV35Ksqr5CSW6BL0/xn85gAAAACa03v3IgsJtAABxnloJcqP",
                        "SSOID": "4037-1", "TCID": "2311200600004570307",
                        "visid_incap_422718": "ZC/JYDc3SIOYYgxLi0rHLV9mW2UAAAAAQUIPAAAAAAD7uqqztjzhdQ9GJezF5api"
                    }
                ).content)
            part_image = Image.open(part_filename)
            print(part_filename, part_image.size)
            assert part_image.size == (width, height)
            page_image.paste(part_image, (0, height * part))
            os.remove(part_filename)
        page_image.save(f'page_{page + 1}.jpg')


if __name__ == '__main__':
    fetch_pages()
    