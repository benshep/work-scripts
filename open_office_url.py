import sys
import os
import subprocess
import json
from base64 import b64decode


def open_office_url(extra_info):
    """Open a SharePoint URL in the correct native app.
    The extra_info argument is the base64 encoded [EXTRA] parameter from the External Application Launcher addon.
    It decodes to a JSON object with parameters: (for context menu clicks)
    menuItemId, parentMenuItemId, viewType, frameId, editable, linkText, linkUrl, pageUrl, modifiers, button"""
    info_decoded = b64decode(extra_info).decode('utf-8')
    info_json = json.loads(info_decoded)
    # open(r'C:\Users\bjs54\Downloads\output.txt', 'w').write(str(info))
    # check for file extension: could be in text or URL, so just check JSON string
    if '.doc' in info_decoded:
        app = 'winword.exe'
    elif '.ppt' in info_decoded:
        app = 'powerpnt.exe'
    elif '.xls' in info_decoded:
        app = 'excel.exe'
    else:
        return
    office_path = r'C:\Program Files\Microsoft Office\root\Office16'
    # try leaving in the query portion of the URL (after ?) - if it doesn't work, need to be discerning
    url = info_json['linkUrl']
    # url = url.split('?')[0]
    subprocess.call([os.path.join(office_path, app), url])


if __name__ == '__main__':
    open_office_url(sys.argv[1])
