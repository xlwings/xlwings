import numpy as np
from xlwings import XlWings

xl = XlWings()

def rand_numbers():
    """ produces a standard normally distributed random numbers with dim (n,n)"""
    sheet = xl.xl_app.Sheets(1)
    n = sheet.Cells(1,2).Value
    rand_num = np.random.randn(n,n)
    sheet.Range(sheet.Cells(3,3), sheet.Cells(2 + n, 2 + n)).Value = rand_num