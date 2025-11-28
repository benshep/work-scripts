import os
from shutil import copy2
from tempfile import mkstemp

import pandas


def read_excel(excel_filename: str, **kwargs):
    """Read an Excel file using pandas. First copy the workbook to a temporary file
    to get around permission errors due to workbook being open."""
    handle, temp_filename = mkstemp(suffix='.xlsx')
    os.close(handle)
    copy2(excel_filename, temp_filename)
    result = pandas.read_excel(temp_filename, **kwargs)
    os.remove(temp_filename)
    return result


class StoredData:
    """Data stored in an Excel file. Extracting it takes time, so cache the contents and refresh if the file changes."""

    def __init__(self, filename: str, **read_excel_kwargs):
        self.filename = filename
        self.modified_time = 0
        self.data = pandas.DataFrame()
        self.read_excel_kwargs = read_excel_kwargs

    def fetch(self):
        """Fetch the data, or return the cached value if the file hasn't changed."""
        modified_time = os.path.getmtime(self.filename)
        if modified_time > self.modified_time:
            self.data = read_excel(self.filename, **self.read_excel_kwargs)
            self.modified_time = modified_time
        return self.data

