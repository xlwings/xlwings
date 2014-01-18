# -*- coding: utf-8 -*-
import sys
import os
import unittest
from datetime import datetime

sys.path.append('..')
from xlwings import xlwings_connect, Range

xl_file1 = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test1.xlsx')
wb = xlwings_connect(xl_file1)


class TestRange(unittest.TestCase):

    # Syntax: Range('A1')
    def test_activesheet_celladdress_integer(self):
        value = 22
        Range('A1').value = value
        cell = Range('A1').value
        self.assertEqual(cell, value)

    def test_activesheet_celladdress_float(self):
        value = 22.2222
        Range('A1').value = value
        cell = Range('A1').value
        self.assertEqual(cell, value)

    def test_activesheet_celladdress_string(self):
        value = 'Test String'
        Range('A1').value = value
        cell = Range('A1').value
        self.assertEqual(cell, value)

    def test_activesheet_celladdress_unicode(self):
        value = u'éöà'
        Range('A1').value = value
        cell = Range('A1').value
        self.assertEqual(cell, value)

    def test_activesheet_celladdress_date(self):
        value = datetime(1962, 11, 3)
        Range('A2').value = value
        cell = Range('A2').value
        self.assertEqual(cell, value)


    # Syntax: Range('SheetName', 'A1')
    def test_sheetname_celladdress_integer(self):
        value = 22
        Range('Sheet2', 'A1').value = value
        cell = Range('Sheet2', 'A1').value
        self.assertEqual(cell, value)

    def test_sheetname_celladdress_float(self):
        value = 22.2222
        Range('Sheet2', 'A1').value = value
        cell = Range('Sheet2', 'A1').value
        self.assertEqual(cell, value)

    def test_sheetname_celladdress_string(self):
        value = 'Test String'
        Range('Sheet2', 'A1').value = value
        cell = Range('Sheet2', 'A1').value
        self.assertEqual(cell, value)

    def test_sheetname_celladdress_unicode(self):
        value = u'éöà'
        Range('Sheet2', 'A1').value = value
        cell = Range('Sheet2', 'A1').value
        self.assertEqual(cell, value)

    # Syntax: Range(1, 'A1')
    def test_sheetindex_celladdress_integer(self):
        value = 22
        Range(3, 'A1').value = value
        cell = Range(3, 'A1').value
        self.assertEqual(cell, value)

    def test_sheetindex_celladdress_float(self):
        value = 22.2222
        Range(3, 'A1').value = value
        cell = Range(3, 'A1').value
        self.assertEqual(cell, value)

    def test_sheetindex_celladdress_string(self):
        value = 'Test String'
        Range(3, 'A1').value = value
        cell = Range(3, 'A1').value
        self.assertEqual(cell, value)

    def test_sheetindex_celladdress_unicode(self):
        value = u'éöà'
        Range(3, 'A1').value = value
        cell = Range(3, 'A1').value
        self.assertEqual(cell, value)

    # Syntax: Range((1,1))
    def test_activesheet_cellindex_integer(self):
        value = 22
        Range((1,1)).value = value
        cell = Range((1,1)).value
        self.assertEqual(cell, value)

    def test_activesheet_cellindex_float(self):
        value = 22.2222
        Range((1,1)).value = value
        cell = Range((1,1)).value
        self.assertEqual(cell, value)

    def test_activesheet_cell_cellindex_string(self):
        value = 'Test String'
        Range((1,1)).value = value
        cell = Range((1,1)).value
        self.assertEqual(cell, value)

    def test_activesheet_cellindex_unicode(self):
        value = u'éöà'
        Range((1,1)).value = value
        cell = Range((1,1)).value
        self.assertEqual(cell, value)

    # Syntax: Range('SheetName', (1,1))
    def test_sheetname_cellindex_integer(self):
        value = 22
        Range('Sheet2', (1,1)).value = value
        cell = Range('Sheet2', (1,1)).value
        self.assertEqual(cell, value)

    def test_sheetname_cellindex_float(self):
        value = 22.2222
        Range('Sheet2', (1,1)).value = value
        cell = Range('Sheet2', (1,1)).value
        self.assertEqual(cell, value)

    def test_sheetname_cellindex_string(self):
        value = 'Test String'
        Range('Sheet2', (1,1)).value = value
        cell = Range('Sheet2', (1,1)).value
        self.assertEqual(cell, value)

    def test_sheetname_cellindex_unicode(self):
        value = u'éöà'
        Range('Sheet2', (1,1)).value = value
        cell = Range('Sheet2', (1,1)).value
        self.assertEqual(cell, value)

    # Syntax: Range(1, (1,1))
    def test_sheetindex_cellindex_integer(self):
        value = 22
        Range(3, (1,1)).value = value
        cell = Range(3, (1,1)).value
        self.assertEqual(cell, value)

    def test_sheetindex_cellindex_float(self):
        value = 22.2222
        Range(3, (1,1)).value = value
        cell = Range(3, (1,1)).value
        self.assertEqual(cell, value)

    def test_sheetindex_cellindex_string(self):
        value = 'Test String'
        Range(3, (1,1)).value = value
        cell = Range(3, (1,1)).value
        self.assertEqual(cell, value)

    def test_sheetindex_cellindex_unicode(self):
        value = u'éöà'
        Range(3, (1,1)).value = value
        cell = Range(3, (1,1)).value
        self.assertEqual(cell, value)


if __name__ == '__main__':
    unittest.main()