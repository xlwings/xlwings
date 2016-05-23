# -*- coding: utf-8 -*-

from __future__ import unicode_literals
import os
import sys
import shutil
from datetime import datetime, date

import pytz
import inspect
import nose
from nose.tools import assert_equal, raises, assert_raises, assert_true, assert_false, assert_not_equal

import xlwings as xw
from xlwings.constants import Calculation

from .test_data import data, test_date_1, test_date_2, list_row_1d, list_row_2d, list_col, chart_data


this_dir = os.path.abspath(os.path.dirname(inspect.getfile(inspect.currentframe())))

# Mac imports
if sys.platform.startswith('darwin'):
    from appscript import k as kw

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


# Test skips and fixtures
def _skip_if_no_numpy():
    if np is None:
        raise nose.SkipTest('numpy missing')


def _skip_if_no_pandas():
    if pd is None:
        raise nose.SkipTest('pandas missing')


def _skip_if_no_matplotlib():
    if matplotlib is None:
        raise nose.SkipTest('matplotlib missing')


class TestBase:
    def setUp(self, xlsx=None):
        self.app = xw.Application(visible=False)
        self.wb = self.app.workbook()

    def tearDown(self):
        self.wb.close()
        if sys.platform.startswith('win'):
            self.app.quit()


class TestApplication(TestBase):
    def setUp(self):
        super(TestApplication, self).setUp()

    def test_screen_updating(self):
        self.app.screen_updating = False
        assert_equal(self.app.screen_updating, False)

        self.app.screen_updating = True
        assert_equal(self.app.screen_updating, True)

    def test_calculation(self):
        sht = self.wb[0]
        sht.range('A1').value = 2
        sht.range('B1').formula = '=A1 * 2'

        self.app.calculation = Calculation.xlCalculationManual
        sht.range('A1').value = 4
        assert_equal(self.wb[0].range('B1').value, 4)

        self.app.calculation = Calculation.xlCalculationAutomatic
        self.app.calculate()  # This is needed on Mac Excel 2016 but not on Mac Excel 2011 (changed behaviour)
        assert_equal(sht.range('B1').value, 8)

        sht.range('A1').value = 2
        assert_equal(sht.range('B1').value, 4)

    def test_version(self):
        assert_true(int(self.app.version.split('.')[0]) > 0)

    def test_apps(self):
        n_original = len(list(xw.apps))
        app2 = xw.Application()
        wb2 = app2.workbook()
        assert_equal(n_original + 1, len(list(xw.apps)))
        assert_equal(xw.apps[0], app2)
        wb2.close()
        app2.quit()

    def test_wb_across_instances(self):
        app2 = xw.Application()
        wb2 = self.app.workbook()
        wb3 = app2.workbook()
        wb4 = app2.workbook()
        wb5 = app2.workbook()

        assert_equal(len(self.app), 2)
        assert_equal(len(app2), 3)

        wb2.close()
        wb3.close()
        wb4.close()
        wb5.close()

        app2.quit()

    # def test_close_all_wbs(self):
    #     app2 = xw.Application()
    #     wb1 = app2.workbook()
    #     wb2 = app2.workbook()
    #
    #     for wb in app2:
    #         wb.close()
    #
    #     app2.quit()



