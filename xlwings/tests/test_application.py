# -*- coding: utf-8 -*-
from __future__ import unicode_literals

from nose.tools import assert_equal, raises, assert_raises, assert_true, assert_false, assert_not_equal

import xlwings as xw
from xlwings.constants import Calculation
from .common import TestBase

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


class TestApplication(TestBase):
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
        assert_equal(sht.range('B1').value, 4)

        self.app.calculation = Calculation.xlCalculationAutomatic
        self.app.calculate()  # This is needed on Mac Excel 2016 but not on Mac Excel 2011 (changed behaviour)
        assert_equal(sht.range('B1').value, 8)

        sht.range('A1').value = 2
        assert_equal(sht.range('B1').value, 4)

    def test_version(self):
        assert_true(int(self.app.version.split('.')[0]) > 0)

    def test_apps(self):
        n_original = len(xw.apps)
        app2 = xw.Application()
        wb2 = app2.workbook()
        assert_equal(n_original + 1, len(xw.apps))
        assert_equal(xw.apps[0], app2)
        wb2.close()
        app2.quit()

    def test_wb_across_instances(self):
        app2 = xw.Application()

        app1_wb_count = len(self.app)
        app2_wb_count = len(app2)

        wb2 = self.app.workbook()
        wb3 = app2.workbook()
        wb4 = app2.workbook()
        wb5 = app2.workbook()

        assert_equal(len(self.app), app1_wb_count + 1)
        assert_equal(len(app2), app2_wb_count + 3)

        wb2.close()
        wb3.close()
        wb4.close()
        wb5.close()

        app2.quit()

    def test_selection(self):
        assert_equal(self.app.selection.address, '$A$1')

    def test_close_all_wbs(self):
        app2 = xw.Application()
        wb1 = app2.workbook()
        wb2 = app2.workbook()

        for wb in app2:
            wb.close()

        app2.quit()



