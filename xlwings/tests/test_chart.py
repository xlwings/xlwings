# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import os
import nose
from nose.tools import assert_equal
from datetime import datetime
from xlwings import Workbook, Range, Chart, ChartType

# Optional imports
try:
    import numpy as np
    from numpy.testing import assert_array_equal
except ImportError:
    np = None
try:
    import pandas as pd
    from pandas import DataFrame, Series
    from pandas.util.testing import assert_frame_equal, assert_series_equal
except ImportError:
    pd = None

# # Connect to test file and make Sheet1 the active sheet
# xl_file_1 = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test_chart_1.xlsx')
# wb = Workbook(xl_file_1)
# wb.activate('Sheet1')

# Test skips and fixtures
def _skip_if_no_numpy():
    if np is None:
        raise nose.SkipTest('numpy missing')


def _skip_if_no_pandas():
    if pd is None:
        raise nose.SkipTest('pandas missing')


# Test Data
chart_data = [['one', 'two'], [1.1, 2.2]]


class TestRange:
    def setUp(self):
        # Connect to test file and make Sheet1 the active sheet
        xl_file1 = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test_chart_1.xlsx')
        # global wb
        self.wb = Workbook(xl_file1)
        self.wb.activate('Sheet1')

    def tearDown(self):
        self.wb.close()

    def test_add_keywords(self):
        name = 'My Chart'
        chart_type = ChartType.xlLine
        Range('A1').value = chart_data
        chart = Chart().add(chart_type=chart_type, name=name, source_data=Range('A1').table)

        chart_actual = Chart(name)
        name_actual = chart_actual.name
        chart_type_actual = chart_actual.chart_type
        assert_equal(name, name_actual)
        assert_equal(chart_type, chart_type_actual)


    def test_add_properties(self):
        name = 'My Chart'
        chart_type = ChartType.xlLine
        Range('Sheet2', 'A1').value = chart_data
        chart = Chart().add('Sheet2')
        chart.chart_type = chart_type
        chart.name = name
        chart.set_source_data(Range('Sheet2', 'A1').table)

        chart_actual = Chart('Sheet2', name)
        name_actual = chart_actual.name
        chart_type_actual = chart_actual.chart_type
        assert_equal(name, name_actual)
        assert_equal(chart_type, chart_type_actual)

if __name__ == '__main__':
    nose.main()