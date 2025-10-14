import os
import pandas

from folders import docs_folder
from work_tools import read_excel

hh_excel = os.path.join(docs_folder, 'Sustainability', 'Energy', 'HH data_Daresbury_240415_FY2324.xlsx')
date_col = 'Date (UTC)'


def get_data(sheet_name: str) -> pandas.DataFrame:
    data = read_excel(hh_excel, sheet_name=sheet_name)

    data = data.drop(columns=['MPAN', 'BST?', 'Weekday', 'Total [MWh]', 'Unnamed: 53'])
    data = data.drop([0, 1, 2])

    data = data.melt([date_col], var_name='Time')
    data = data.sort_values([date_col, 'Time'])

    return data


energy, intensity = [get_data(name) for name in ('energy', 'intensity')]

