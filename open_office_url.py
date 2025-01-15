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
    info = json.loads(b64decode(extra_info))
    link_text = info['linkText']
    if '.doc' in link_text:
        app = 'winword.exe'
    elif '.ppt' in link_text:
        app = 'powerpnt.exe'
    elif '.xls' in link_text:
        app = 'excel.exe'
    else:
        return
    office_path = r'C:\Program Files\Microsoft Office\root\Office16'
    url = info['linkUrl']
    # try leaving in the query portion of the URL (after ?) - if it doesn't work, need to be discerning
    # url = url.split('?')[0]
    subprocess.call([os.path.join(office_path, app), url])


if __name__ == '__main__':
    open_office_url(sys.argv[1])
