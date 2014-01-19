# -*- coding: utf-8 -*-
import os
import nose
from nose.tools import assert_equal
from datetime import datetime

from xlwings import xlwings_connect, Range

xl_file1 = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test1.xlsx')
wb = xlwings_connect(xl_file1)


def test_cell():
    """ Style: Range('A1') """
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

        # Sheetname
        Range('Sheet2', address).value = value
        cell = Range('Sheet2', address).value
        assert_equal(cell, value)

        # Sheetindex
        Range(3, address).value = value
        cell = Range(3, address).value
        assert_equal(cell, value)


if __name__ == '__main__':
    nose.main()