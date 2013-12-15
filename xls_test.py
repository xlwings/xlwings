import os
from win32com.client import GetObject
import numpy as np

# Reference to the calling Excel File
file_name = 'xls_test.xlsm'
dir_path = os.path.dirname(os.path.abspath(__file__))
file_path = '{0}\\{1}'.format(dir_path, file_name)
xl = GetObject(file_path)


def rand_numbers():
    sheet = xl.Sheets(1)
    n = sheet.Cells(1,2).Value
    rand_num = np.random.randn(n,n)
    sheet.Range(sheet.Cells(3,3), sheet.Cells(2 + n, 2 + n)).Value = rand_num