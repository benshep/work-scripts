import os
from shutil import copy2
from tempfile import mkstemp

import pandas


def read_excel(excel_filename, **kwargs) -> pandas.DataFrame:
    """Read an Excel file using pandas. First copy the workbook to a temporary file
    to get around permission errors due to workbook being open."""
    handle, temp_filename = mkstemp(suffix='.xlsx')
    os.close(handle)
    copy2(excel_filename, temp_filename)
    dataframe = pandas.read_excel(temp_filename, **kwargs)
    os.remove(temp_filename)
    return dataframe
