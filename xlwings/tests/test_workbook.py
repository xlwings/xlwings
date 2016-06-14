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


class TestWorkbook(TestBase):
    def test_name(self):
        wb = self.app1.workbook(os.path.join(this_dir, 'test book.xlsx'))
        assert_equal(wb.name, 'test book.xlsx')

    def test_reference_two_unsaved_wb(self):
        """Covers GH Issue #63"""
        wb1 = self.wb1
        wb2 = self.app1.workbook()

        wb2.sheets[0].range('A1').value = 2.
        wb1.sheets[0].range('A1').value = 1.

        assert_equal(wb2.sheets[0].range('A1').value, 2.)
        assert_equal(wb1.sheets[0].range('A1').value, 1.)

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

        self.app1.workbook(target_file_path).close()
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

        self.app1.workbook(target_file_path).close()
        if os.path.isfile(target_file_path):
            os.remove(target_file_path)

    # def test_mock_caller(self):
    #     # Can't really run this one with app_visible=False
    #     # _skip_if_not_default_xl()
    #
    #     path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test_workbook_1.xlsx')
    #
    #     wb1 = xw.Workbook(path)  # open the wb
    #     xw.Workbook.set_mock_caller(path)
    #     wb = xw.Workbook.caller()
    #     wb[0].range('A1').value = 333
    #     assert_equal(wb[0].range('A1').value, 333)

    def test_unicode_name(self):
        wb = self.app1.workbook()
        if sys.platform.startswith('darwin') and self.app1.major_version >= 15:
            dst = os.path.join(os.path.expanduser("~") + '/Library/Containers/com.microsoft.Excel/Data/', 'ünicöde.xlsx')
        else:
            dst = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ünicöde.xlsx')
        if os.path.isfile(dst):
            os.remove(dst)
        wb.save(dst)
        wb2 = self.app1.workbook(dst)
        wb2.sheets[0].range('A1').value = 1
        wb2.close()
        os.remove(dst)

    def test_unsaved_workbook_reference(self):
        self.wb1.sheets[0].range('B2').value = 123
        wb2 = self.app1.workbook(self.wb1.name)
        assert_equal(wb2.sheets[0].range('B2').value, 123)

    def test_active_workbook(self):
        self.wb2.sheets[0].range('A1').value = 'active_workbook'  # 2nd instance
        self.wb2.activate()
        wb_active = xw.Workbook.active()
        assert_equal(wb_active.sheets[0].range('A1').value, 'active_workbook')

    def test_macro(self):
        # NOTE: Uncheck Macro security check in Excel
        _none = None if sys.platform.startswith('win') else ''

        src = os.path.abspath(os.path.join(this_dir, 'macro book.xlsm'))
        wb = self.app1.workbook(src)

        test1 = wb.macro('Module1.Test1')
        test2 = wb.macro('Module1.Test2')
        test3 = wb.macro('Module1.Test3')
        test4 = wb.macro('Test4')

        res1 = test1('Test1a', 'Test1b')

        assert_equal(res1, 1)
        assert_equal(test2(), 2)
        assert_equal(test3('Test3a', 'Test3b'), _none)
        assert_equal(test4(), _none)
        assert_equal(wb[0].range('A1').value, 'Test1a')
        assert_equal(wb[0].range('A2').value, 'Test1b')
        assert_equal(wb[0].range('A3').value, 'Test2')
        assert_equal(wb[0].range('A4').value, 'Test3a')
        assert_equal(wb[0].range('A5').value, 'Test3b')
        assert_equal(wb[0].range('A6').value, 'Test4')

