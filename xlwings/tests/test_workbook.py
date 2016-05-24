# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import os
import sys
import shutil

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
        wb = xw.Workbook(os.path.join(this_dir, 'test_workbook_1.xlsx'))
        assert_equal(wb.name, 'test_workbook_1.xlsx')

    def test_reference_two_unsaved_wb(self):
        """Covers GH Issue #63"""
        wb1 = xw.Workbook()
        wb2 = xw.Workbook()

        sht1 = wb1[0]
        sht2 = wb2[0]

        sht2.range('A1').value = 2.  # wb2
        sht1.range('A1').value = 1.  # wb1

        assert_equal(sht2.range('A1').value, 2.)
        assert_equal(sht1.range('A1').value, 1.)

        wb1.close()
        wb2.close()

    def test_save_naked(self):
        if sys.platform.startswith('darwin'):
            folder = os.path.expanduser("~") + '/Library/Containers/com.microsoft.Excel/Data/'
            if os.path.isdir(folder):
                os.chdir(folder)

        cwd = os.getcwd()
        wb1 = xw.Workbook()
        target_file_path = os.path.join(cwd, wb1.name + '.xlsx')
        if os.path.isfile(target_file_path):
            os.remove(target_file_path)

        wb1.save()

        assert_true(os.path.isfile(target_file_path))

        wb2 = xw.Workbook(target_file_path)
        wb2.close()

        if os.path.isfile(target_file_path):
            os.remove(target_file_path)

    def test_save_path(self):
        if sys.platform.startswith('darwin'):
            folder = os.path.expanduser("~") + '/Library/Containers/com.microsoft.Excel/Data/'
            if os.path.isdir(folder):
                os.chdir(folder)

        cwd = os.getcwd()
        wb1 = xw.Workbook()
        target_file_path = os.path.join(cwd, 'TestFile.xlsx')
        if os.path.isfile(target_file_path):
            os.remove(target_file_path)

        wb1.save(target_file_path)

        assert_true(os.path.isfile(target_file_path))

        wb2 = xw.Workbook(target_file_path)
        wb2.close()

        if os.path.isfile(target_file_path):
            os.remove(target_file_path)

    def test_mock_caller(self):
        # Can't really run this one with app_visible=False
        # _skip_if_not_default_xl()

        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test_workbook_1.xlsx')

        wb1 = xw.Workbook(path)  # open the wb
        xw.Workbook.set_mock_caller(path)
        wb = xw.Workbook.caller()
        wb[0].range('A1').value = 333
        assert_equal(wb[0].range('A1').value, 333)
        wb.close()

    def test_unicode_path(self):
        # pip3 seems to struggle with unicode filenames
        src = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'unicode_path.xlsx')
        if sys.platform.startswith('darwin') and os.path.isdir(os.path.expanduser("~") + '/Library/Containers/com.microsoft.Excel/Data/'):
            dst = os.path.join(os.path.expanduser("~") + '/Library/Containers/com.microsoft.Excel/Data/',
                           'ünicödé_päth.xlsx')
        else:
            dst = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ünicödé_päth.xlsx')
        shutil.copy(src, dst)
        wb = xw.Workbook(dst)
        wb[0].range('A1').value = 1
        wb.close()
        os.remove(dst)

    def test_unsaved_workbook_reference(self):
        wb = xw.Workbook()
        wb[0].range('B2').value = 123
        wb2 = xw.Workbook(wb.name)
        assert_equal(wb2[0].range('B2').value, 123)
        wb2.close()

    def test_active_workbook(self):
        app2 = xw.Application()
        wb2 = app2.workbook()
        wb2[0].range('A1').value = 'active_workbook'
        wb_active = xw.Workbook.active()
        assert_equal(wb_active[0].range('A1').value, 'active_workbook')

    def test_macro(self):
        src = os.path.realpath(os.path.join(this_dir, 'macro book.xlsm'))
        wb1 = xw.Workbook(src)

        test1 = wb1.macro('Module1.Test1')
        test2 = wb1.macro('Module1.Test2')
        test3 = wb1.macro('Module1.Test3')
        test4 = wb1.macro('Test4')

        res1 = test1('Test1a', 'Test1b')

        assert_equal(res1, 1)
        assert_equal(test2(), 2)
        if sys.platform.startswith('win'):
            assert_equal(test3('Test3a', 'Test3b'), None)
        else:
            assert_equal(test3('Test3a', 'Test3b'), '')
        if sys.platform.startswith('win'):
            assert_equal(test4(), None)
        else:
            assert_equal(test4(), '')
        assert_equal(wb1[0].range('A1').value, 'Test1a')
        assert_equal(wb1[0].range('A2').value, 'Test1b')
        assert_equal(wb1[0].range('A3').value, 'Test2')
        assert_equal(wb1[0].range('A4').value, 'Test3a')
        assert_equal(wb1[0].range('A5').value, 'Test3b')
        assert_equal(wb1[0].range('A6').value, 'Test4')

        wb1.close()
