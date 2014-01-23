# -*- coding: utf-8 -*-
import os
import nose
from nose.tools import assert_equal
from datetime import datetime
import numpy as np
from numpy.testing import assert_array_equal
from pandas import DataFrame

from xlwings import xlwings_connect, Range

# Connect to test file and make Sheet1 the active sheet
xl_file1 = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test1.xlsx')
wb = xlwings_connect(xl_file1)
wb.Sheets('Sheet1').Activate()

# Testdata
data = [[1, 2.222, 3.333],
        ['Test1', None, u'éöà'],
        [datetime(1962, 11, 3), datetime(2020, 12, 31, 12, 12, 20), 9.999]]


def test_cell():
    params = [('A1', 22),
              ((1,1), 22),
              ('A1', 22.2222),
              ((1,1), 22.2222),
              ('A1', 'Test String'),
              ((1,1), 'Test String'),
              ('A1', u'éöà'),
              ((1,1), u'éöà'),
              ('A2', datetime(1962, 11, 3)),
              ((2,1), datetime(1962, 11, 3)),
              ('A3', datetime(2020, 12, 31, 12, 12, 20)),
              ((3,1), datetime(2020, 12, 31, 12, 12, 20))]
    for param in params:
        yield check_cell, param[0], param[1]


def check_cell(address, value):
    # Active Sheet
    Range(address).value = value
    cell = Range(address).value
    assert_equal(cell, value)

    # SheetName
    Range('Sheet2', address).value = value
    cell = Range('Sheet2', address).value
    assert_equal(cell, value)

    # SheetIndex
    Range(3, address).value = value
    cell = Range(3, address).value
    assert_equal(cell, value)


def test_range_address():
    """ Style: Range('A1:C3') """
    address = 'C1:E3'

    # Active Sheet
    Range(address[:2]).value = data  # assign to starting cell only
    cells = Range(address).value
    assert_equal(cells, data)

    # Sheetname
    Range('Sheet2', address).value = data
    cells = Range('Sheet2', address).value
    assert_equal(cells, data)

    # Sheetindex
    Range(3, address).value = data
    cells = Range(3, address).value
    assert_equal(cells, data)

def test_range_index():
    """ Style: Range((1,1), (3,3)) """
    index1 = (1,3)
    index2 = (3,5)

    # Active Sheet
    Range(index1, index2).value = data
    cells = Range(index1, index2).value
    assert_equal(cells, data)

    # Sheetname
    Range('Sheet2', index1, index2).value = data
    cells = Range('Sheet2', index1, index2).value
    assert_equal(cells, data)

    # Sheetindex
    Range(3, index1, index2).value = data
    cells = Range(3, index1, index2).value
    assert_equal(cells, data)


def test_named_range():
    value = 22.222
    # Active Sheet
    Range('cell_sheet1').value = value
    cells = Range('cell_sheet1').value
    assert_equal(cells, value)

    Range('range_sheet1').value = data
    cells = Range('range_sheet1').value
    assert_equal(cells, data)

    # Sheetname
    Range('Sheet2', 'cell_sheet2').value = value
    cells = Range('Sheet2', 'cell_sheet2').value
    assert_equal(cells, value)

    Range('Sheet2', 'range_sheet2').value = data
    cells = Range('Sheet2', 'range_sheet2').value
    assert_equal(cells, data)

    # Sheetindex
    Range(3, 'cell_sheet3').value = value
    cells = Range(3, 'cell_sheet3').value
    assert_equal(cells, value)

    Range(3, 'range_sheet3').value = data
    cells = Range(3, 'range_sheet3').value
    assert_equal(cells, data)


def test_array():
    numpy_array = np.array([[1.1, 2.2, 3.3], [4.4, 5.5, -6.6], [7.7, 8.8, np.nan]])
    Range('L1').value = numpy_array
    cells = Range('L1:N3', asarray=True).value
    assert_array_equal(cells, numpy_array)


def test_table():
    Range('Sheet4', 'A1').value = data
    cells = Range('Sheet4', 'A1').table.value
    assert_equal(cells, data)


def test_clear_content():
    Range('Sheet4', 'G1').value = 22
    Range('Sheet4', 'G1').clear_contents()
    cell = Range('Sheet4', 'G1').value
    assert_equal(cell, None)


def test_clear():
    Range('Sheet4', 'G1').value = 22
    Range('Sheet4', 'G1').clear()
    cell = Range('Sheet4', 'G1').value
    assert_equal(cell, None)


def test_dataframe():
    df_expected = DataFrame({'a': [1, 2, 3.3, np.nan], 'b': ['test1', 'test2', 'test3', None]})
    Range('Sheet5', 'A1').value = df_expected
    cells = Range('Sheet5', 'B1:C5').value
    df_result = DataFrame(cells[1:], columns=cells[0])
    assert_equal(df_expected, df_result)


if __name__ == '__main__':
    nose.main()