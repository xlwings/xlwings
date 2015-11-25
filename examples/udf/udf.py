"""
Copyright (C) 2014-2015, Zoomer Analytics LLC.
All rights reserved.

License: BSD 3-clause (see LICENSE.txt for details)
"""

from xlwings import Workbook, Range, xlsub, xlfunc, xlarg

@xlsub
def get_workbook_name():
    """Writes the name of the Workbook into Range("D1") of Sheet 1"""
    wb = Workbook.caller()
    Range(1, 'D1').value = wb.name


@xlfunc
def double_sum(x, y):
    """Returns twice the sum of the two arguments"""
    return 2 * (x + y)


@xlfunc
@xlarg('data', ndim=2)
def add_one(data):
    """Adds 1 to every cell in Range"""
    return [[cell + 1 for cell in row] for row in data]


@xlfunc
@xlarg('x', 'nparray', ndim=2)
@xlarg('y', 'nparray', ndim=2)
def matrix_mult(x, y):
    """Alternative implementation of Excel's MMULT, requires NumPy"""
    return x.dot(y)
