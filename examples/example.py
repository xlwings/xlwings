# This hack is only here to make the examples work without the need to install it 
# and without the need to copy the xlwings.py file into the examples directory
import sys, os
this_path =  os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.abspath(os.path.join(this_path, os.pardir)))

import numpy as np
from xlwings import Workbook, Range

wb = Workbook()  # Creates a reference to the calling Excel workbook


def rand_numbers():
    """ produces standard normally distributed random numbers with dim (n,n)"""
    n = Range('Sheet1', 'B1').value
    rand_num = np.random.randn(n, n)
    Range('Sheet1', 'C3').table.clear_contents()
    Range('Sheet1', 'C3').value = rand_num