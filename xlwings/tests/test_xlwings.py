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


# Test data
data = [[1, 2.222, 3.333],
        ['Test1', None, 'éöà'],
        [datetime(1962, 11, 3), datetime(2020, 12, 31, 12, 12, 20), 9.999]]

test_date_1 = datetime(1962, 11, 3)
test_date_2 = datetime(2020, 12, 31, 12, 12, 20)

list_row_1d = [1.1, None, 3.3]
list_row_2d = [[1.1, None, 3.3]]
list_col = [[1.1], [None], [3.3]]
chart_data = [['one', 'two'], [1.1, 2.2]]

if np is not None:
    array_1d = np.array([1.1, 2.2, np.nan, -4.4])
    array_2d = np.array([[1.1, 2.2, 3.3], [-4.4, 5.5, np.nan]])

if pd is not None:
    series_1 = pd.Series([1.1, 3.3, 5., np.nan, 6., 8.])

    rng = pd.date_range('1/1/2012', periods=10, freq='D')
    timeseries_1 = pd.Series(np.arange(len(rng)) + 0.1, rng)
    timeseries_1[1] = np.nan

    df_1 = pd.DataFrame([[1, 'test1'],
                         [2, 'test2'],
                         [np.nan, None],
                         [3.3, 'test3']], columns=['a', 'b'])

    df_2 = pd.DataFrame([1, 3, 5, np.nan, 6, 8], columns=['col1'])

    # MultiIndex (Index)
    tuples = list(zip(*[['bar', 'bar', 'baz', 'baz', 'foo', 'foo', 'qux', 'qux'],
                        ['one', 'two', 'one', 'two', 'one', 'two', 'one', 'two'],
                        ['x', 'x', 'x', 'x', 'y', 'y', 'y', 'y']]))
    index = pd.MultiIndex.from_tuples(tuples, names=['first', 'second', 'third'])
    df_multiindex = pd.DataFrame([[1.1, 2.2], [3.3, 4.4], [5.5, 6.6], [7.7, 8.8], [9.9, 10.10],
                                  [11.11, 12.12],[13.13, 14.14], [15.15, 16.16]], index=index)

    # MultiIndex (Header)
    header = [['Foo', 'Foo', 'Bar', 'Bar', 'Baz'], ['A', 'B', 'C', 'D', 'E']]

    df_multiheader = pd.DataFrame([[0.0, 1.0, 2.0, 3.0, 4.0],
                                   [0.0, 1.0, 2.0, 3.0, 4.0],
                                   [0.0, 1.0, 2.0, 3.0, 4.0],
                                   [0.0, 1.0, 2.0, 3.0, 4.0],
                                   [0.0, 1.0, 2.0, 3.0, 4.0],
                                   [0.0, 1.0, 2.0, 3.0, 4.0]], columns=pd.MultiIndex.from_arrays(header))


# Test skips and fixtures
def _skip_if_no_numpy():
    if np is None:
        raise nose.SkipTest('numpy missing')


def _skip_if_no_pandas():
    if pd is None:
        raise nose.SkipTest('pandas missing')


class TestRange:
    def setUp(self):
        # Connect to test file and make Sheet1 the active sheet
        xl_file1 = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test_range_1.xlsx')
        self.wb = Workbook(xl_file1)
        self.wb.activate('Sheet1')

    def tearDown(self):
        self.wb.close()

    def test_cell(self):
        params = [('A1', 22),
                  ((1,1), 22),
                  ('A1', 22.2222),
                  ((1,1), 22.2222),
                  ('A1', 'Test String'),
                  ((1,1), 'Test String'),
                  ('A1', 'éöà'),
                  ((1,1), 'éöà'),
                  ('A2', test_date_1),
                  ((2,1), test_date_1),
                  ('A3', test_date_2),
                  ((3,1), test_date_2)]
        for param in params:
            yield self.check_cell, param[0], param[1]

    def check_cell(self, address, value):
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

    def test_range_address(self):
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

    def test_range_index(self):
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

    def test_named_range(self):
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

    def test_array(self):
        _skip_if_no_numpy()

        # 1d array
        Range('Sheet6', 'A1').value = array_1d
        cells = Range('Sheet6', 'A1:D1', asarray=True).value
        assert_array_equal(cells, array_1d)

        # 2d array
        Range('Sheet6', 'A4').value = array_2d
        cells = Range('Sheet6', 'A4', asarray=True).table.value
        assert_array_equal(cells, array_2d)

        # 1d array (atleast_2d)
        Range('Sheet6', 'A10').value = array_1d
        cells = Range('Sheet6', 'A10:D10', asarray=True, atleast_2d=True).value
        assert_array_equal(cells, np.atleast_2d(array_1d))

        # 2d array (atleast_2d)
        Range('Sheet6', 'A12').value = array_2d
        cells = Range('Sheet6', 'A12', asarray=True, atleast_2d=True).table.value
        assert_array_equal(cells, array_2d)

    def test_vertical(self):
        Range('Sheet4', 'A10').value = data
        cells = Range('Sheet4', 'A10').vertical.value
        assert_equal(cells, [row[0] for row in data])

    def test_horizontal(self):
        Range('Sheet4', 'A20').value = data
        cells = Range('Sheet4', 'A20').horizontal.value
        assert_equal(cells, data[0])

    def test_table(self):
        Range('Sheet4', 'A1').value = data
        cells = Range('Sheet4', 'A1').table.value
        assert_equal(cells, data)

    def test_list(self):

        # 1d List Row
        Range('Sheet4', 'A27').value = list_row_1d
        cells = Range('Sheet4', 'A27:C27').value
        assert_equal(list_row_1d, cells)

        # 2d List Row
        Range('Sheet4', 'A29').value = list_row_2d
        cells = Range('Sheet4', 'A29:C29', atleast_2d=True).value
        assert_equal(list_row_2d, cells)

        # 1d List Col
        Range('Sheet4', 'A31').value = list_col
        cells = Range('Sheet4', 'A31:A33').value
        assert_equal([i[0] for i in list_col], cells)
        # 2d List Col
        cells = Range('Sheet4', 'A31:A33', atleast_2d=True).value
        assert_equal(list_col, cells)

    def test_clear_content(self):
        Range('Sheet4', 'G1').value = 22
        Range('Sheet4', 'G1').clear_contents()
        cell = Range('Sheet4', 'G1').value
        assert_equal(cell, None)

    def test_clear(self):
        Range('Sheet4', 'G1').value = 22
        Range('Sheet4', 'G1').clear()
        cell = Range('Sheet4', 'G1').value
        assert_equal(cell, None)

    def test_dataframe_1(self):
        _skip_if_no_pandas()

        df_expected = df_1
        Range('Sheet5', 'A1').value = df_expected
        cells = Range('Sheet5', 'B1:C5').value
        df_result = DataFrame(cells[1:], columns=cells[0])
        assert_frame_equal(df_expected, df_result)

    def test_dataframe_2(self):
        """ Covers GH Issue #31"""
        _skip_if_no_pandas()

        df_expected = df_2
        Range('Sheet5', 'A9').value = df_expected
        cells = Range('Sheet5', 'B9:B15').value
        df_result = DataFrame(cells[1:], columns=[cells[0]])
        assert_frame_equal(df_expected, df_result)

    def test_dataframe_multiindex(self):
        _skip_if_no_pandas()

        df_expected = df_multiindex
        Range('Sheet5', 'A20').value = df_expected
        cells = Range('Sheet5', 'D20').table.value
        multiindex = Range('Sheet5', 'A20:C28').value
        ix = pd.MultiIndex.from_tuples(multiindex[1:], names=multiindex[0])
        df_result = DataFrame(cells[1:], columns=cells[0], index=ix)
        assert_frame_equal(df_expected, df_result)

    def test_dataframe_multiheader(self):
        _skip_if_no_pandas()

        df_expected = df_multiheader
        Range('Sheet5', 'A52').value = df_expected
        cells = Range('Sheet5', 'B52').table.value
        df_result = DataFrame(cells[2:], columns=pd.MultiIndex.from_arrays(cells[:2]))
        assert_frame_equal(df_expected, df_result)

    def test_series_1(self):
        _skip_if_no_pandas()

        series_expected = series_1
        Range('Sheet5', 'A32').value = series_expected
        cells = Range('Sheet5', 'B32:B37').value
        series_result = Series(cells)
        assert_series_equal(series_expected, series_result)

    def test_timeseries_1(self):
        _skip_if_no_pandas()

        series_expected = timeseries_1
        Range('Sheet5', 'A40').value = series_expected
        cells = Range('Sheet5', 'B40:B49').value
        date_index = Range('Sheet5', 'A40:A49').value
        series_result = Series(cells, index=date_index)
        assert_series_equal(series_expected, series_result)

    def test_none(self):
        """ Covers GH Issue #16"""
        # None
        Range('Sheet1', 'A7').value = None
        assert_equal(None, Range('Sheet1', 'A7').value)
        # List
        Range('Sheet1', 'A7').value = [None, None]
        assert_equal(None, Range('Sheet1', 'A7').horizontal.value)

    def test_scalar_nan(self):
        """Covers GH Issue #15"""
        _skip_if_no_numpy()

        Range('Sheet1', 'A20').value = np.nan
        assert_equal(None, Range('Sheet1', 'A20').value)


class TestChart:
    def setUp(self):
        # Connect to test file and make Sheet1 the active sheet
        xl_file1 = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test_chart_1.xlsx')
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


class TestWorkbook:
    def setUp(self):
        # Connect to test file and make Sheet1 the active sheet
        xl_file1 = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test_workbook_1.xlsx')
        self.wb = Workbook(xl_file1)
        self.wb.activate('Sheet1')

    def tearDown(self):
        self.wb.close()

    def test_clear_content_active_sheet(self):
        Range('G10').value = 22
        self.wb.clear_contents()
        cell = Range('G10').value
        assert_equal(cell, None)

    def test_clear_active_sheet(self):
        Range('G10').value = 22
        self.wb.clear()
        cell = Range('G10').value
        assert_equal(cell, None)

    def test_clear_content(self):
        Range('Sheet2', 'G10').value = 22
        self.wb.clear_contents('Sheet2')
        cell = Range('Sheet2', 'G10').value
        assert_equal(cell, None)

    def test_clear(self):
        Range('Sheet2', 'G10').value = 22
        self.wb.clear('Sheet2')
        cell = Range('Sheet2', 'G10').value
        assert_equal(cell, None)


if __name__ == '__main__':
    nose.main()