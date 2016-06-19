# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import sys
import os
import datetime as dt

import pytz
from nose.tools import assert_equal, raises, assert_raises, assert_true, assert_false, assert_not_equal

import xlwings as xw
from xlwings.constants import RgbColor
from .common import TestBase, this_dir, _skip_if_no_numpy, _skip_if_no_pandas
from .test_data import data, list_col, list_row_1d, list_row_2d, test_date_1, test_date_2

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

# Mac imports
if sys.platform.startswith('darwin'):
    from appscript import k as kw


class TestRange(TestBase):
    def test_cell(self):
        params = [('A1', 22),
                  ((1, 1), 22),
                  ('A1', 22.2222),
                  ((1, 1), 22.2222),
                  ('A1', 'Test String'),
                  ((1, 1), 'Test String'),
                  ('A1', 'éöà'),
                  ((1, 1), 'éöà'),
                  ('A2', test_date_1),
                  ((2, 1), test_date_1),
                  ('A3', test_date_2),
                  ((3, 1), test_date_2)]
        for param in params:
            yield self.check_cell, param[0], param[1]

    def check_cell(self, address, value):
        # Active Sheet
        self.wb1.sheets[0].range(address).value = value
        cell = self.wb1.sheets[0].range(address).value
        assert_equal(cell, value)

        # SheetName
        self.wb1.sheets['Sheet2'].range(address).value = value
        cell = self.wb1.sheets['Sheet2'].range(address).value
        assert_equal(cell, value)

        # SheetIndex
        self.wb1.sheets[2].range(address).value = value
        cell = self.wb1.sheets[2].range(address).value
        assert_equal(cell, value)

    def test_range_address(self):
        """ test_range_address: Style: Range('A1:C3') """
        address = 'C1:E3'

        # Active Sheet
        xw.Range(address[:2]).value = data  # assign to starting cell only
        cells = xw.Range(address).value
        assert_equal(cells, data)

        # Sheetname
        self.wb1.sheets['Sheet2'].range(address).value = data
        cells = self.wb1.sheets['Sheet2'].range(address).value
        assert_equal(cells, data)

        # Sheetindex
        self.wb1.sheets[2].range(address).value = data
        cells = self.wb1.sheets[2].range(address).value
        assert_equal(cells, data)

    def test_range_index(self):
        """ test_range_index: Style: Range((1,1), (3,3)) """
        index1 = (1, 3)
        index2 = (3, 5)

        # Active Sheet
        xw.Range(index1, index2).value = data
        cells = xw.Range(index1, index2).value
        assert_equal(cells, data)

        # Sheetname
        self.wb1.sheets['Sheet2'].range(index1, index2).value = data
        cells = self.wb1.sheets['Sheet2'].range(index1, index2).value
        assert_equal(cells, data)

        # Sheetindex
        self.wb1.sheets[2].range(index1, index2).value = data
        cells = self.wb1.sheets[2].range(index1, index2).value
        assert_equal(cells, data)

    def test_named_range_value(self):
        value = 22.222
        # Active Sheet
        xw.Range('F1').name = 'cell_sheet1'
        xw.Range('cell_sheet1').value = value
        cells = xw.Range('cell_sheet1').value
        assert_equal(cells, value)

        xw.Range('A1:C3').name = 'range_sheet1'
        xw.Range('range_sheet1').value = data
        cells = xw.Range('range_sheet1').value
        assert_equal(cells, data)

        # Sheetname
        self.wb1.sheets['Sheet2'].range('F1').name = 'cell_sheet2'
        self.wb1.sheets['Sheet2'].range('cell_sheet2').value = value
        cells = self.wb1.sheets['Sheet2'].range('cell_sheet2').value
        assert_equal(cells, value)

        self.wb1.sheets['Sheet2'].range('A1:C3').name = 'range_sheet2'
        self.wb1.sheets['Sheet2'].range('range_sheet2').value = data
        cells = self.wb1.sheets['Sheet2'].range('range_sheet2').value
        assert_equal(cells, data)

        # Sheetindex
        self.wb1.sheets[2].range('F3').name = 'cell_sheet3'
        self.wb1.sheets[2].range('cell_sheet3').value = value
        cells = self.wb1.sheets[2].range('cell_sheet3').value
        assert_equal(cells, value)

        self.wb1.sheets[2].range('A1:C3').name = 'range_sheet3'
        self.wb1.sheets[2].range('range_sheet3').value = data
        cells = self.wb1.sheets[2].range('range_sheet3').value
        assert_equal(cells, data)

    def test_array(self):
        _skip_if_no_numpy()

        # 1d array
        self.wb1.sheets[0].range('A1').value = array_1d
        cells = self.wb1.sheets[0].range('A1:D1').options(np.array).value
        assert_array_equal(cells, array_1d)

        # 2d array
        self.wb1.sheets[0].range('A4').value = array_2d
        cells = self.wb1.sheets[0].range('A4').options(np.array, expand='table').value
        assert_array_equal(cells, array_2d)

        # 1d array (atleast_2d)
        self.wb1.sheets[0].range('A10').value = array_1d
        cells = self.wb1.sheets[0].range('A10:D10').options(np.array, ndim=2).value
        assert_array_equal(cells, np.atleast_2d(array_1d))

        # 2d array (atleast_2d)
        self.wb1.sheets[0].range('A12').value = array_2d
        cells = self.wb1.sheets[0].range('A12').options(np.array, ndim=2, expand='table').value
        assert_array_equal(cells, array_2d)

    def test_vertical(self):
        self.wb1.sheets[0].range('A10').value = data
        if sys.platform.startswith('win') and self.wb1.app.version == '14.0':
            self.wb1.sheets[0].range('A12:B12').number_format = 'dd/mm/yyyy'  # Hack for Excel 2010 bug, see GH #43
        cells = self.wb1.sheets[0].range('A10').vertical.value
        assert_equal(cells, [row[0] for row in data])

    def test_horizontal(self):
        self.wb1.sheets[0].range('A20').value = data
        cells = self.wb1.sheets[0].range('A20').horizontal.value
        assert_equal(cells, data[0])

    def test_table(self):
        self.wb1.sheets[0].range('A1').value = data
        if sys.platform.startswith('win') and self.wb1.app.version == '14.0':
            self.wb1.sheets[0].range('A3:B3').number_format = 'dd/mm/yyyy'  # Hack for Excel 2010 bug, see GH #43
        cells = self.wb1.sheets[0].range('A1').table.value
        assert_equal(cells, data)

    def test_list(self):
        # 1d List Row
        self.wb1.sheets[0].range('A27').value = list_row_1d
        cells = self.wb1.sheets[0].range('A27:C27').value
        assert_equal(list_row_1d, cells)

        # 2d List Row
        self.wb1.sheets[0].range('A29').value = list_row_2d
        cells = self.wb1.sheets[0].range('A29:C29').options(ndim=2).value
        assert_equal(list_row_2d, cells)

        # 1d List Col
        self.wb1.sheets[0].range('A31').value = list_col
        cells = self.wb1.sheets[0].range('A31:A33').value
        assert_equal([i[0] for i in list_col], cells)
        # 2d List Col
        cells = self.wb1.sheets[0].range('A31:A33').options(ndim=2).value
        assert_equal(list_col, cells)

    def test_formula(self):
        self.wb1.sheets[0].range('A1').formula = '=SUM(A2:A10)'
        assert_equal(self.wb1.sheets[0].range('A1').formula, '=SUM(A2:A10)')

    def test_formula_array(self):
        self.wb1.sheets[0].range('A1').value = [[1, 4], [2, 5], [3, 6]]
        self.wb1.sheets[0].range('D1').formula_array = '=SUM(A1:A3*B1:B3)'
        assert_equal(self.wb1.sheets[0].range('D1').value, 32.)

    def test_current_region(self):
        values = [[1., 2.], [3., 4.]]
        self.wb1.sheets[0].range('A20').value = values
        assert_equal(self.wb1.sheets[0].range('B21').current_region.value, values)

    def test_clear_content(self):
        self.wb1.sheets[0].range('G1').value = 22
        self.wb1.sheets[0].range('G1').clear_contents()
        cell = self.wb1.sheets[0].range('G1').value
        assert_equal(cell, None)

    def test_clear(self):
        self.wb1.sheets[0].range('G1').value = 22
        self.wb1.sheets[0].range('G1').clear()
        cell = self.wb1.sheets[0].range('G1').value
        assert_equal(cell, None)

    def test_dataframe_1(self):
        _skip_if_no_pandas()

        df_expected = df_1
        self.wb1.sheets[0].range('A1').value = df_expected
        df_result = self.wb1.sheets[0].range('A1:C5').options(pd.DataFrame).value
        df_result.index = pd.Int64Index(df_result.index)
        assert_frame_equal(df_expected, df_result)

    def test_dataframe_2(self):
        """ test_dataframe_2: Covers GH Issue #31"""
        _skip_if_no_pandas()

        df_expected = df_2
        self.wb1.sheets[0].range('A9').value = df_expected
        cells = self.wb1.sheets[0].range('B9:B15').value
        df_result = DataFrame(cells[1:], columns=[cells[0]])
        assert_frame_equal(df_expected, df_result)

    def test_dataframe_multiindex(self):
        _skip_if_no_pandas()

        df_expected = df_multiindex
        self.wb1.sheets[0].range('A20').value = df_expected
        cells = self.wb1.sheets[0].range('D20').table.value
        multiindex = self.wb1.sheets[0].range('A20:C28').value
        ix = pd.MultiIndex.from_tuples(multiindex[1:], names=multiindex[0])
        df_result = DataFrame(cells[1:], columns=cells[0], index=ix)
        assert_frame_equal(df_expected, df_result)

    def test_dataframe_multiheader(self):
        _skip_if_no_pandas()

        df_expected = df_multiheader
        self.wb1.sheets[0].range('A52').value = df_expected
        cells = self.wb1.sheets[0].range('B52').table.value
        df_result = DataFrame(cells[2:], columns=pd.MultiIndex.from_arrays(cells[:2]))
        assert_frame_equal(df_expected, df_result)

    def test_dataframe_dateindex(self):
        _skip_if_no_pandas()

        df_expected = df_dateindex
        self.wb1.sheets[0].range('A100').value = df_expected
        if sys.platform.startswith('win') and self.wb1.app.version == '14.0':
            self.wb1.sheets[0].range('A100').vertical.number_format = 'dd/mm/yyyy'  # Hack for Excel 2010 bug, see GH #43
        cells = self.wb1.sheets[0].range('B100').table.value
        index = self.wb1.sheets[0].range('A101').vertical.value
        df_result = DataFrame(cells[1:], index=index, columns=cells[0])
        assert_frame_equal(df_expected, df_result)

    def test_read_df_0header_0index(self):
        _skip_if_no_pandas()

        self.wb1.sheets[0].range('A1').value = [[1, 2, 3],
                                        [4, 5, 6],
                                        [7, 8, 9]]

        df1 = pd.DataFrame([[1., 2., 3.], [4., 5., 6.], [7., 8., 9.]])

        df2 = self.wb1.sheets[0].range('A1:C3').options(pd.DataFrame, header=0, index=False).value
        assert_frame_equal(df1, df2)

    def test_df_1header_0index(self):
        _skip_if_no_pandas()

        self.wb1.sheets[0].range('A1').options(pd.DataFrame, index=False, header=True).value = pd.DataFrame([[1., 2.], [3., 4.]], columns=['a', 'b'])
        df = self.wb1.sheets[0].range('A1').options(pd.DataFrame, index=False, header=True,
                                            expand='table').value
        assert_frame_equal(df, pd.DataFrame([[1., 2.], [3., 4.]], columns=['a', 'b']))

    def test_df_0header_1index(self):
        _skip_if_no_pandas()

        self.wb1.sheets[0].range('A1').options(pd.DataFrame, index=True, header=False).value = pd.DataFrame([[1., 2.], [3., 4.]], index=[10., 20.])
        df = self.wb1.sheets[0].range('A1').options(pd.DataFrame, index=True, header=False,
                                            expand='table').value
        assert_frame_equal(df, pd.DataFrame([[1., 2.], [3., 4.]], index=[10., 20.]))

    def test_read_df_1header_1namedindex(self):
        _skip_if_no_pandas()

        self.wb1.sheets[0].range('A1').value = [['ix1', 'c', 'd', 'c'],
                                        [1, 1, 2, 3],
                                        [2, 4, 5, 6]]

        df1 = pd.DataFrame([[1., 2., 3.], [4., 5., 6.]],
                           index=[1., 2.],
                           columns=['c', 'd', 'c'])
        df1.index.name = 'ix1'

        df2 = self.wb1.sheets[0].range('A1:D3').options(pd.DataFrame).value
        assert_frame_equal(df1, df2)

    def test_read_df_1header_1unnamedindex(self):
        _skip_if_no_pandas()

        self.wb1.sheets[0].range('A1').value = [[None, 'c', 'd', 'c'],
                                        [1, 1, 2, 3],
                                        [2, 4, 5, 6]]

        df1 = pd.DataFrame([[1., 2., 3.], [4., 5., 6.]],
                           index=pd.Index([1., 2.]),
                           columns=['c', 'd', 'c'])

        df2 = self.wb1.sheets[0].range('A1:D3').options(pd.DataFrame).value

        assert_frame_equal(df1, df2)

    def test_read_df_2header_1namedindex(self):
        _skip_if_no_pandas()

        self.wb1.sheets[0].range('A1').value = [[None, 'a', 'a', 'b'],
                                        ['ix1', 'c', 'd', 'c'],
                                        [1, 1, 2, 3],
                                        [2, 4, 5, 6]]

        df1 = pd.DataFrame([[1., 2., 3.], [4., 5., 6.]],
                           index=[1., 2.],
                           columns=pd.MultiIndex.from_arrays([['a', 'a', 'b'], ['c', 'd', 'c']]))
        df1.index.name = 'ix1'

        df2 = self.wb1.sheets[0].range('A1:D4').options(pd.DataFrame, header=2).value

        assert_frame_equal(df1, df2)

    def test_read_df_2header_1unnamedindex(self):
        _skip_if_no_pandas()

        self.wb1.sheets[0].range('A1').value = [[None, 'a', 'a', 'b'],
                                        [None, 'c', 'd', 'c'],
                                        [1, 1, 2, 3],
                                        [2, 4, 5, 6]]

        df1 = pd.DataFrame([[1., 2., 3.], [4., 5., 6.]],
                           index=pd.Index([1, 2]),
                           columns=pd.MultiIndex.from_arrays([['a', 'a', 'b'], ['c', 'd', 'c']]))

        df2 = self.wb1.sheets[0].range('A1:D4').options(pd.DataFrame, header=2).value
        df2.index = pd.Int64Index(df2.index)

        assert_frame_equal(df1, df2)

    def test_read_df_2header_2namedindex(self):
        _skip_if_no_pandas()

        self.wb1.sheets[0].range('A1').value = [[None, None, 'a', 'a', 'b'],
                                        ['x1', 'x2', 'c', 'd', 'c'],
                                        ['a', 1, 1, 2, 3],
                                        ['a', 2, 4, 5, 6],
                                        ['b', 1, 7, 8, 9]]

        df1 = pd.DataFrame([[1., 2., 3.], [4., 5., 6.], [7., 8., 9.]],
                           index=pd.MultiIndex.from_arrays([['a', 'a', 'b'], [1., 2., 1.]], names=['x1', 'x2']),
                           columns=pd.MultiIndex.from_arrays([['a', 'a', 'b'], ['c', 'd', 'c']]))

        df2 = self.wb1.sheets[0].range('A1:E5').options(pd.DataFrame, header=2, index=2).value
        assert_frame_equal(df1, df2)

    def test_read_df_2header_2unnamedindex(self):
        _skip_if_no_pandas()

        self.wb1.sheets[0].range('A1').value = [[None, None, 'a', 'a', 'b'],
                                        [None, None, 'c', 'd', 'c'],
                                        ['a', 1, 1, 2, 3],
                                        ['a', 2, 4, 5, 6],
                                        ['b', 1, 7, 8, 9]]

        df1 = pd.DataFrame([[1., 2., 3.], [4., 5., 6.], [7., 8., 9.]],
                           index=pd.MultiIndex.from_arrays([['a', 'a', 'b'], [1., 2., 1.]]),
                           columns=pd.MultiIndex.from_arrays([['a', 'a', 'b'], ['c', 'd', 'c']]))

        df2 = self.wb1.sheets[0].range('A1:E5').options(pd.DataFrame, header=2, index=2).value
        assert_frame_equal(df1, df2)

    def test_read_df_1header_2namedindex(self):
        _skip_if_no_pandas()

        self.wb1.sheets[0].range('A1').value = [['x1', 'x2', 'a', 'd', 'c'],
                                        ['a', 1, 1, 2, 3],
                                        ['a', 2, 4, 5, 6],
                                        ['b', 1, 7, 8, 9]]

        df1 = pd.DataFrame([[1., 2., 3.], [4., 5., 6.], [7., 8., 9.]],
                           index=pd.MultiIndex.from_arrays([['a', 'a', 'b'], [1., 2., 1.]], names=['x1', 'x2']),
                           columns=['a', 'd', 'c'])

        df2 = self.wb1.sheets[0].range('A1:E4').options(pd.DataFrame, header=1, index=2).value
        assert_frame_equal(df1, df2)

    def test_timeseries_1(self):
        _skip_if_no_pandas()

        series_expected = timeseries_1
        self.wb1.sheets[0].range('A40').options(header=False).value = series_expected
        if sys.platform.startswith('win') and self.wb1.app.version == '14.0':
            self.wb1.sheets[0].range('A40').vertical.number_format = 'dd/mm/yyyy'  # Hack for Excel 2010 bug, see GH #43
        series_result = self.wb1.sheets[0].range('A40:B49').options(pd.Series, header=False).value
        assert_series_equal(series_expected, series_result)

    def test_read_series_noheader_noindex(self):
        _skip_if_no_pandas()

        self.wb1.sheets[0].range('A1').value = [[1.],
                                        [2.],
                                        [3.]]
        s = self.wb1.sheets[0].range('A1:A3').options(pd.Series, index=False, header=False).value
        assert_series_equal(s, pd.Series([1., 2., 3.]))

    def test_read_series_noheader_index(self):
        _skip_if_no_pandas()

        self.wb1.sheets[0].range('A1').value = [[10., 1.],
                                        [20., 2.],
                                        [30., 3.]]
        s = self.wb1.sheets[0].range('A1:B3').options(pd.Series, index=True, header=False).value
        assert_series_equal(s, pd.Series([1., 2., 3.], index=[10., 20., 30.]))

    def test_read_series_header_noindex(self):
        _skip_if_no_pandas()

        self.wb1.sheets[0].range('A1').value = [['name'],
                                        [1.],
                                        [2.],
                                        [3.]]
        s = self.wb1.sheets[0].range('A1:A4').options(pd.Series, index=False, header=True).value
        assert_series_equal(s, pd.Series([1., 2., 3.], name='name'))

    def test_read_series_header_index(self):
        _skip_if_no_pandas()

        # Named index
        self.wb1.sheets[0].range('A1').value = [['ix', 'name'],
                                        [10., 1.],
                                        [20., 2.],
                                        [30., 3.]]
        s = self.wb1.sheets[0].range('A1:B4').options(pd.Series, index=True, header=True).value
        assert_series_equal(s, pd.Series([1., 2., 3.], name='name', index=pd.Index([10., 20., 30.], name='ix')))

        # Nameless index
        self.wb1.sheets[0].range('A1').value = [[None, 'name'],
                                        [10., 1.],
                                        [20., 2.],
                                        [30., 3.]]
        s = self.wb1.sheets[0].range('A1:B4').options(pd.Series, index=True, header=True).value
        assert_series_equal(s, pd.Series([1., 2., 3.], name='name', index=[10., 20., 30.]))

    def test_write_series_noheader_noindex(self):
        _skip_if_no_pandas()

        self.wb1.sheets[0].range('A1').options(index=False).value = pd.Series([1., 2., 3.])
        assert_equal([[1.],[2.],[3.]], self.wb1.sheets[0].range('A1').options(ndim=2, expand='table').value)

    def test_write_series_noheader_index(self):
        _skip_if_no_pandas()

        self.wb1.sheets[0].range('A1').options(index=True).value = pd.Series([1., 2., 3.], index=[10., 20., 30.])
        assert_equal([[10., 1.],[20., 2.],[30., 3.]], self.wb1.sheets[0].range('A1').options(ndim=2, expand='table').value)

    def test_write_series_header_noindex(self):
        _skip_if_no_pandas()

        self.wb1.sheets[0].range('A1').options(index=False).value = pd.Series([1., 2., 3.], name='name')
        assert_equal([['name'],[1.],[2.],[3.]], self.wb1.sheets[0].range('A1').options(ndim=2, expand='table').value)

    def test_write_series_header_index(self):
        _skip_if_no_pandas()

        # Named index
        self.wb1.sheets[0].range('A1').value = pd.Series([1., 2., 3.], name='name', index=pd.Index([10., 20., 30.], name='ix'))
        assert_equal([['ix', 'name'],[10., 1.],[20., 2.],[30., 3.]], self.wb1.sheets[0].range('A1').options(ndim=2, expand='table').value)

        # Nameless index
        self.wb1.sheets[0].range('A1').value = pd.Series([1., 2., 3.], name='name', index=[10., 20., 30.])
        assert_equal([[None, 'name'],[10., 1.],[20., 2.],[30., 3.]], self.wb1.sheets[0].range('A1:B4').options(ndim=2).value)

    def test_none(self):
        """ test_none: Covers GH Issue #16"""
        # None
        self.wb1.sheets[0].range( 'A7').value = None
        assert_equal(None, self.wb1.sheets[0].range('A7').value)
        # List
        self.wb1.sheets[0].range('A7').value = [None, None]
        assert_equal(None, self.wb1.sheets[0].range('A7').horizontal.value)

    def test_scalar_nan(self):
        """test_scalar_nan: Covers GH Issue #15"""
        _skip_if_no_numpy()

        self.wb1.sheets[0].range('A20').value = np.nan
        assert_equal(None, self.wb1.sheets[0].range('A20').value)

    def test_atleast_2d_scalar(self):
        """test_atleast_2d_scalar: Covers GH Issue #53a"""
        self.wb1.sheets[0].range('A50').value = 23
        result = self.wb1.sheets[0].range('A50').options(ndim=2).value
        assert_equal([[23]], result)

    def test_atleast_2d_scalar_as_array(self):
        """test_atleast_2d_scalar_as_array: Covers GH Issue #53b"""
        _skip_if_no_numpy()

        self.wb1.sheets[0].range('A50').value = 23
        result = self.wb1.sheets[0].range('A50').options(np.array, ndim=2).value
        assert_equal(np.array([[23]]), result)

    def test_column_width(self):
        self.wb1.sheets[0].range('A1:B2').column_width = 10.0
        result = self.wb1.sheets[0].range('A1').column_width
        assert_equal(10.0, result)

        self.wb1.sheets[0].range('A1:B2').value = 'ensure cells are used'
        self.wb1.sheets[0].range('B2').column_width = 20.0
        result = self.wb1.sheets[0].range('A1:B2').column_width
        if sys.platform.startswith('win'):
            assert_equal(None, result)
        else:
            assert_equal(kw.missing_value, result)

    def test_row_height(self):
        self.wb1.sheets[0].range('A1:B2').row_height = 15.0
        result = self.wb1.sheets[0].range('A1').row_height
        assert_equal(15.0, result)

        self.wb1.sheets[0].range('A1:B2').value = 'ensure cells are used'
        self.wb1.sheets[0].range('B2').row_height = 20.0
        result = self.wb1.sheets[0].range('A1:B2').row_height
        if sys.platform.startswith('win'):
            assert_equal(None, result)
        else:
            assert_equal(kw.missing_value, result)

    def test_width(self):
        """test_width: Width depends on default style text size, so do not test absolute widths"""
        self.wb1.sheets[0].range('A1:D4').column_width = 10.0
        result_before = self.wb1.sheets[0].range('A1').width
        self.wb1.sheets[0].range('A1:D4').column_width = 12.0
        result_after = self.wb1.sheets[0].range('A1').width
        assert_true(result_after > result_before)

    def test_height(self):
        self.wb1.sheets[0].range('A1:D4').row_height = 60.0
        result = self.wb1.sheets[0].range('A1:D4').height
        assert_equal(240.0, result)

    def test_left(self):
        assert_equal(self.wb1.sheets[0].range('A1').left, 0.0)
        self.wb1.sheets[0].range('A1').column_width = 20.0
        assert_equal(self.wb1.sheets[0].range('B1').left, self.wb1.sheets[0].range('A1').width)

    def test_top(self):
        assert_equal(self.wb1.sheets[0].range('A1').top, 0.0)
        self.wb1.sheets[0].range('A1').row_height = 20.0
        assert_equal(self.wb1.sheets[0].range('A2').top, self.wb1.sheets[0].range('A1').height)

    def test_autofit_range(self):
        self.wb1.sheets[0].range('A1:D4').value = 'test_string'

        self.wb1.sheets[0].range('A1:D4').row_height = 40
        self.wb1.sheets[0].range('A1:D4').column_width = 40
        assert_equal(40, self.wb1.sheets[0].range('A1:D4').row_height)
        assert_equal(40, self.wb1.sheets[0].range('A1:D4').column_width)
        self.wb1.sheets[0].range('A1:D4').autofit()
        assert_true(40 != self.wb1.sheets[0].range('A1:D4').column_width)
        assert_true(40 != self.wb1.sheets[0].range('A1:D4').row_height)

        self.wb1.sheets[0].range('A1:D4').row_height = 40
        assert_equal(40, self.wb1.sheets[0].range('A1:D4').row_height)
        self.wb1.sheets[0].range('A1:D4').autofit('r')
        assert_true(40 != self.wb1.sheets[0].range('A1:D4').row_height)

        self.wb1.sheets[0].range('A1:D4').column_width = 40
        assert_equal(40, self.wb1.sheets[0].range('A1:D4').column_width)
        self.wb1.sheets[0].range('A1:D4').autofit('c')
        assert_true(40 != self.wb1.sheets[0].range('A1:D4').column_width)

        self.wb1.sheets[0].range('A1:D4').autofit('rows')
        self.wb1.sheets[0].range('A1:D4').autofit('columns')

    def test_autofit_col(self):
        self.wb1.sheets[0].range('A1:D4').value = 'test_string'
        self.wb1.sheets[0].range('A:D').column_width = 40
        assert_equal(40, self.wb1.sheets[0].range('A:D').column_width)
        self.wb1.sheets[0].range('A:D').autofit()
        assert_true(40 != self.wb1.sheets[0].range('A:D').column_width)

        # Just checking if they don't throw an error
        self.wb1.sheets[0].range('A:D').autofit('r')
        self.wb1.sheets[0].range('A:D').autofit('c')
        self.wb1.sheets[0].range('A:D').autofit('rows')
        self.wb1.sheets[0].range('A:D').autofit('columns')

    def test_autofit_row(self):
        self.wb1.sheets[0].range('A1:D4').value = 'test_string'
        self.wb1.sheets[0].range('1:10').row_height = 40
        assert_equal(40, self.wb1.sheets[0].range('1:10').row_height)
        self.wb1.sheets[0].range('1:10').autofit()
        assert_true(40 != self.wb1.sheets[0].range('1:10').row_height)

        # Just checking if they don't throw an error
        self.wb1.sheets[0].range('1:1000000').autofit('r')
        self.wb1.sheets[0].range('1:1000000').autofit('c')
        self.wb1.sheets[0].range('1:1000000').autofit('rows')
        self.wb1.sheets[0].range('1:1000000').autofit('columns')

    def test_number_format_cell(self):
        format_string = "mm/dd/yy;@"
        self.wb1.sheets[0].range('A1').number_format = format_string
        result = self.wb1.sheets[0].range('A1').number_format
        assert_equal(format_string, result)

    def test_number_format_range(self):
        format_string = "mm/dd/yy;@"
        self.wb1.sheets[0].range('A1:D4').number_format = format_string
        result = self.wb1.sheets[0].range('A1:D4').number_format
        assert_equal(format_string, result)

    def test_get_address(self):
        wb1 = xw.Book(os.path.join(this_dir, 'test book.xlsx'))

        res = wb1.sheets[0].range((1, 1), (3, 3)).get_address()
        assert_equal(res, '$A$1:$C$3')

        res = wb1.sheets[0].range((1, 1), (3, 3)).get_address(False)
        assert_equal(res, '$A1:$C3')

        res = wb1.sheets[0].range((1, 1), (3, 3)).get_address(True, False)
        assert_equal(res, 'A$1:C$3')

        res = wb1.sheets[0].range((1, 1), (3, 3)).get_address(False, False)
        assert_equal(res, 'A1:C3')

        res = wb1.sheets[0].range((1, 1), (3, 3)).get_address(include_sheetname=True)
        assert_equal(res, "'Sheet1'!$A$1:$C$3")

        res = wb1.sheets[1].range((1, 1), (3, 3)).get_address(include_sheetname=True)
        assert_equal(res, "'Sheet2'!$A$1:$C$3")

        res = wb1.sheets[0].range((1, 1), (3, 3)).get_address(external=True)
        assert_equal(res, "'[test book.xlsx]Sheet1'!$A$1:$C$3")

        wb1.close()

    def test_hyperlink(self):
        address = 'www.xlwings.org'
        # Naked address
        self.wb1.sheets[0].range('A1').add_hyperlink(address)
        assert_equal(self.wb1.sheets[0].range('A1').value, address)
        hyperlink = self.wb1.sheets[0].range('A1').hyperlink
        if not hyperlink.endswith('/'):
            hyperlink += '/'
        assert_equal(hyperlink, 'http://' + address + '/')

        # Address + FriendlyName
        self.wb1.sheets[0].range('A2').add_hyperlink(address, 'test_link')
        assert_equal(self.wb1.sheets[0].range('A2').value, 'test_link')
        hyperlink = self.wb1.sheets[0].range('A2').hyperlink
        if not hyperlink.endswith('/'):
            hyperlink += '/'
        assert_equal(hyperlink, 'http://' + address + '/')

    def test_hyperlink_formula(self):
        self.wb1.sheets[0].range('B10').formula = '=HYPERLINK("http://xlwings.org", "xlwings")'
        assert_equal(self.wb1.sheets[0].range('B10').hyperlink, 'http://xlwings.org')

    def test_color(self):
        rgb = (30, 100, 200)

        self.wb1.sheets[0].range('A1').color = rgb
        assert_equal(rgb, self.wb1.sheets[0].range('A1').color)

        self.wb1.sheets[0].range('A2').color = RgbColor.rgbAqua
        assert_equal((0, 255, 255), self.wb1.sheets[0].range('A2').color)

        self.wb1.sheets[0].range('A2').color = None
        assert_equal(self.wb1.sheets[0].range('A2').color, None)

        self.wb1.sheets[0].range('A1:D4').color = rgb
        assert_equal(rgb, self.wb1.sheets[0].range('A1:D4').color)

    def test_size(self):
        assert_equal(self.wb1.sheets[0].range('A1:C4').size, 12)

    def test_shape(self):
        assert_equal(self.wb1.sheets[0].range('A1:C4').shape, (4, 3))

    def test_len(self):
        assert_equal(len(self.wb1.sheets[0].range('A1:C4')), 12)

    def test_len_rows(self):
        assert_equal(len(self.wb1.sheets[0].range('A1:C4').rows), 4)

    def test_len_cols(self):
        assert_equal(len(self.wb1.sheets[0].range('A1:C4').columns), 3)

    def test_iterator(self):
        self.wb1.sheets[0].range('A20').value = [[1., 2.], [3., 4.]]
        r = self.wb1.sheets[0].range('A20:B21')

        assert_equal([c.value for c in r], [1., 2., 3., 4.])

        # check that reiterating on same range works properly
        assert_equal([c.value for c in r], [1., 2., 3., 4.])

    def test_resize(self):
        r = self.wb1.sheets[0].range('A1').resize(4, 5)
        assert_equal(r.shape, (4, 5))

        r = self.wb1.sheets[0].range('A1').resize(row_size=4)
        assert_equal(r.shape, (4, 1))

        r = self.wb1.sheets[0].range('A1:B4').resize(column_size=5)
        assert_equal(r.shape, (4, 5))

        r = self.wb1.sheets[0].range('A1:B4').resize(row_size=5)
        assert_equal(r.shape, (5, 2))

        r = self.wb1.sheets[0].range('A1:B4').resize()
        assert_equal(r.shape, (4, 2))

        assert_raises(AssertionError, self.wb1.sheets[0].range('A1:B4').resize, row_size=0)
        assert_raises(AssertionError, self.wb1.sheets[0].range('A1:B4').resize, column_size=0)

    def test_offset(self):
        o = self.wb1.sheets[0].range('A1:B3').offset(3, 4)
        assert_equal(o.address, '$E$4:$F$6')

        o = self.wb1.sheets[0].range('A1:B3').offset(row_offset=3)
        assert_equal(o.address, '$A$4:$B$6')

        o = self.wb1.sheets[0].range('A1:B3').offset(column_offset=4)
        assert_equal(o.address, '$E$1:$F$3')

    def test_date(self):
        date_1 = dt.date(2000, 12, 3)
        self.wb1.sheets[0].range('X1').value = date_1
        date_2 = self.wb1.sheets[0].range('X1').value
        assert_equal(date_1, dt.date(date_2.year, date_2.month, date_2.day))

    def test_row(self):
        assert_equal(self.wb1.sheets[0].range('B3:F5').row, 3)

    def test_column(self):
        assert_equal(self.wb1.sheets[0].range('B3:F5').column, 2)

    def test_last_cell(self):
        assert_equal(self.wb1.sheets[0].range('B3:F5').last_cell.row, 5)
        assert_equal(self.wb1.sheets[0].range('B3:F5').last_cell.column, 6)

    def test_integers(self):
        """test_integers: Covers GH 227"""
        self.wb1.sheets[0].range('A99').value = 2147483647  # max SInt32
        assert_equal(self.wb1.sheets[0].range('A99').value, 2147483647)

        self.wb1.sheets[0].range('A100').value = 2147483648  # SInt32 < x < SInt64
        assert_equal(self.wb1.sheets[0].range('A100').value, 2147483648)

        self.wb1.sheets[0].range('A101').value = 10000000000000000000  # long
        assert_equal(self.wb1.sheets[0].range('A101').value, 10000000000000000000)

    def test_numpy_datetime(self):
        _skip_if_no_numpy()

        self.wb1.sheets[0].range('A55').value = np.datetime64('2005-02-25T03:30Z')
        assert_equal(self.wb1.sheets[0].range('A55').value, dt.datetime(2005, 2, 25, 3, 30))

    def test_dataframe_timezone(self):
        _skip_if_no_pandas()

        np_dt = np.datetime64(1434149887000, 'ms')
        ix = pd.DatetimeIndex(data=[np_dt], tz='GMT')
        df = pd.DataFrame(data=[1], index=ix, columns=['A'])
        self.wb1.sheets[0].range('A1').value = df
        assert_equal(self.wb1.sheets[0].range('A2').value, dt.datetime(2015, 6, 12, 22, 58, 7))

    def test_datetime_timezone(self):
        eastern = pytz.timezone('US/Eastern')
        dt_naive = dt.datetime(2002, 10, 27, 6, 0, 0)
        dt_tz = eastern.localize(dt_naive)
        self.wb1.sheets[0].range('F34').value = dt_tz
        assert_equal(self.wb1.sheets[0].range('F34').value, dt_naive)

    @raises(IndexError)
    def test_zero_based_index1(self):
        self.wb1.sheets[0].range((0, 1)).value = 123

    @raises(IndexError)
    def test_zero_based_index2(self):
        a = self.wb1.sheets[0].range((1, 1), (1, 0)).value

    @raises(IndexError)
    def test_zero_based_index3(self):
        xw.Range((1, 0)).value = 123

    @raises(IndexError)
    def test_zero_based_index4(self):
        a = xw.Range((1, 0), (1, 0)).value

    def test_dictionary(self):
        d = {'a': 1., 'b': 2.}
        self.wb1.sheets[0].range('A1').value = d
        assert_equal(d, self.wb1.sheets[0].range('A1:B2').options(dict).value)

    def test_write_to_multicell_range(self):
        self.wb1.sheets[0].range('A1:B2').value = 5
        assert_equal(self.wb1.sheets[0].range('A1:B2').value, [[5., 5.],[5., 5.]])

    def test_transpose(self):
        self.wb1.sheets[0].range('A1').options(transpose=True).value = [[1., 2.], [3., 4.]]
        assert_equal(self.wb1.sheets[0].range('A1:B2').value, [[1., 3.], [2., 4.]])