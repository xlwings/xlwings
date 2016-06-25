# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import os
import sys

from nose.tools import assert_equal, raises, assert_raises, assert_true, assert_false, assert_not_equal

import xlwings as xw
from .common import TestBase, this_dir

# Optional dependencies
try:
    import numpy as np
    from numpy.testing import assert_array_equal
    from .test_data import array_1d, array_2d
except ImportError:
    np = None
try:
    import pandas as pd
    from pandas import DataFrame, Series
    from pandas.util.testing import assert_frame_equal, assert_series_equal
    from .test_data import series_1, timeseries_1, df_1, df_2, df_dateindex, df_multiheader, df_multiindex
except ImportError:
    pd = None
try:
    import matplotlib
    from matplotlib.figure import Figure
except ImportError:
    matplotlib = None
try:
    import PIL
except ImportError:
    PIL = None


class TestBooks(TestBase):
    def test_indexing(self):
        assert_equal(self.app1.books[0].name, self.app1.books(1).name)

    def test_len(self):
        assert_equal(len(self.app1.books), 1)

    def test_add(self):
        self.app1.books.add()
        assert_equal(len(self.app1.books), 2)

    def test_open(self):
        self.app1.books.open(os.path.join(this_dir, 'test book.xlsx'))
        assert_equal(self.app1.active_book.name, 'test book.xlsx')

    def test_iter(self):
        for ix, wb in enumerate(self.app1.books):
            assert_equal(self.app1.books[ix].name, wb.name)


class TestBook(TestBase):
    def test_instantiate_unsaved(self):
        self.wb1.sheets[0].range('B2').value = 123
        wb2 = self.app1.book(self.wb1.name)
        assert_equal(wb2.sheets[0].range('B2').value, 123)

    def test_instantiate_two_unsaved(self):
        """Covers GH Issue #63"""
        wb1 = self.wb1
        wb2 = self.app1.book()

        wb2.sheets[0].range('A1').value = 2.
        wb1.sheets[0].range('A1').value = 1.

        assert_equal(wb2.sheets[0].range('A1').value, 2.)
        assert_equal(wb1.sheets[0].range('A1').value, 1.)

    def test_instantiate_saved_by_name(self):
        wb1 = self.app1.book(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test book.xlsx'))
        wb1.sheets[0].range('A1').value = 'xx'
        wb2 = self.app1.book('test book.xlsx')
        assert_equal(wb2.sheets[0].range('A1').value, 'xx')

    def test_instantiate_saved_by_fullpath(self):
        # unicode name of book, but not unicode path
        wb = self.app1.book()
        if sys.platform.startswith('darwin') and self.app1.major_version >= 15:
            dst = os.path.join(os.path.expanduser("~") + '/Library/Containers/com.microsoft.Excel/Data/', 'üni cöde.xlsx')
        else:
            dst = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'üni cöde.xlsx')
        if os.path.isfile(dst):
            os.remove(dst)
        wb.save(dst)
        wb2 = self.app1.book(dst)  # Book is open
        wb2.sheets[0].range('A1').value = 1
        wb2.save()
        wb2.close()
        wb3 = self.app1.book(dst)  # Book is closed
        assert_equal(wb3.sheets[0].range('A1').value, 1.)
        wb3.close()
        os.remove(dst)

    def test_active_book(self):
        self.wb2.sheets[0].range('A1').value = 'active_book'  # 2nd instance
        self.wb2.activate()
        wb_active = xw.Book.active()
        assert_equal(wb_active.sheets[0].range('A1').value, 'active_book')

    def test_mock_caller(self):
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test book.xlsx')

        wb = self.app1.book(path)
        wb.set_mock_caller()
        wb2 = xw.Book.caller()
        wb2.sheets[0].range('A1').value = 333
        assert_equal(wb2.sheets[0].range('A1').value, 333)

    def test_macro(self):
        # NOTE: Uncheck Macro security check in Excel
        _none = None if sys.platform.startswith('win') else ''

        src = os.path.abspath(os.path.join(this_dir, 'macro book.xlsm'))
        wb = self.app1.book(src)

        test1 = wb.macro('Module1.Test1')
        test2 = wb.macro('Module1.Test2')
        test3 = wb.macro('Module1.Test3')
        test4 = wb.macro('Test4')

        res1 = test1('Test1a', 'Test1b')

        assert_equal(res1, 1)
        assert_equal(test2(), 2)
        assert_equal(test3('Test3a', 'Test3b'), _none)
        assert_equal(test4(), _none)
        assert_equal(wb.sheets[0].range('A1').value, 'Test1a')
        assert_equal(wb.sheets[0].range('A2').value, 'Test1b')
        assert_equal(wb.sheets[0].range('A3').value, 'Test2')
        assert_equal(wb.sheets[0].range('A4').value, 'Test3a')
        assert_equal(wb.sheets[0].range('A5').value, 'Test3b')
        assert_equal(wb.sheets[0].range('A6').value, 'Test4')

    def test_name(self):
        wb = self.app1.book(os.path.join(this_dir, 'test book.xlsx'))
        assert_equal(wb.name, 'test book.xlsx')

    def test_sheets(self):
        assert_equal(len(self.wb1.sheets), 3)

    def test_app(self):
        assert_equal(self.app1, self.wb1.app)

    def test_close(self):
        self.wb1.close()
        assert_equal(len(self.app1.books), 0)

    def test_active_sheet(self):
        assert_equal(self.wb1.active_sheet.name, self.wb1.sheets[0].name)

    def test_save_naked(self):
        if sys.platform.startswith('darwin') and self.app1.major_version >= 15:
            folder = os.path.expanduser("~") + '/Library/Containers/com.microsoft.Excel/Data/'
            if os.path.isdir(folder):
                os.chdir(folder)

        cwd = os.getcwd()
        target_file_path = os.path.join(cwd, self.wb1.name + '.xlsx')
        if os.path.isfile(target_file_path):
            os.remove(target_file_path)

        self.wb1.save()

        assert_true(os.path.isfile(target_file_path))

        self.app1.book(target_file_path).close()
        if os.path.isfile(target_file_path):
            os.remove(target_file_path)

    def test_save_path(self):
        if sys.platform.startswith('darwin') and self.app1.major_version >= 15:
            folder = os.path.expanduser("~") + '/Library/Containers/com.microsoft.Excel/Data/'
            if os.path.isdir(folder):
                os.chdir(folder)

        cwd = os.getcwd()
        target_file_path = os.path.join(cwd, 'TestFile.xlsx')
        if os.path.isfile(target_file_path):
            os.remove(target_file_path)

        self.wb1.save(target_file_path)

        assert_true(os.path.isfile(target_file_path))

        self.app1.book(target_file_path).close()
        if os.path.isfile(target_file_path):
            os.remove(target_file_path)


