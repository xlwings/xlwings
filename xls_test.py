import numpy as np
import xlwings
import logging
from datetime import datetime

# Logging
logging.basicConfig(filename='log_xlwings.txt', level=logging.INFO)
logging.info('{0} - -------------Starting------------------'.format(datetime.now().strftime('%Y-%m-%d %H:%M:%S')))

try:
    xl = xlwings.XlWings()
    logging.info('{0} - Xls dispatched'.format(datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
except Exception as e:
    logging.error('{0} - {1}'.format(datetime.now().strftime('%Y-%m-%d %H:%M:%S'), e))

def rand_numbers():
    """ produces a standard normally distributed random numbers with dim (n,n)"""
    logging.info('{0} - function start'.format(datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
    sheet = xl.xl_app.ActiveWorkbook.Sheets(1)
    n = sheet.Cells(1,2).Value
    rand_num = np.random.randn(n,n)
    sheet.Range(sheet.Cells(3,3), sheet.Cells(2 + n, 2 + n)).Value = rand_num
    logging.info('{0} - function end'.format(datetime.now().strftime('%Y-%m-%d %H:%M:%S')))