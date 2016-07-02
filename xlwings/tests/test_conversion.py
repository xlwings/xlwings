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


class TestConverter(TestBase):
    def test_transpose(self):
        self.wb1.sheets[0].range('A1').options(transpose=True).value = [[1., 2.], [3., 4.]]
        assert_equal(self.wb1.sheets[0].range('A1:B2').value, [[1., 3.], [2., 4.]])

    def test_dictionary(self):
        d = {'a': 1., 'b': 2.}
        self.wb1.sheets[0].range('A1').value = d
        assert_equal(d, self.wb1.sheets[0].range('A1:B2').options(dict).value)

    def test_integers(self):
        """test_integers: Covers GH 227"""
        self.wb1.sheets[0].range('A99').value = 2147483647  # max SInt32
        assert_equal(self.wb1.sheets[0].range('A99').value, 2147483647)

        self.wb1.sheets[0].range('A100').value = 2147483648  # SInt32 < x < SInt64
        assert_equal(self.wb1.sheets[0].range('A100').value, 2147483648)

        self.wb1.sheets[0].range('A101').value = 10000000000000000000  # long
        assert_equal(self.wb1.sheets[0].range('A101').value, 10000000000000000000)

    def test_datetime_timezone(self):
        eastern = pytz.timezone('US/Eastern')
        dt_naive = dt.datetime(2002, 10, 27, 6, 0, 0)
        dt_tz = eastern.localize(dt_naive)
        self.wb1.sheets[0].range('F34').value = dt_tz
        assert_equal(self.wb1.sheets[0].range('F34').value, dt_naive)

    def test_date(self):
        date_1 = dt.date(2000, 12, 3)
        self.wb1.sheets[0].range('X1').value = date_1
        date_2 = self.wb1.sheets[0].range('X1').value
        assert_equal(date_1, dt.date(date_2.year, date_2.month, date_2.day))

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

    def test_none(self):
        """ test_none: Covers GH Issue #16"""
        # None
        self.wb1.sheets[0].range('A7').value = None
        assert_equal(None, self.wb1.sheets[0].range('A7').value)
        # List
        self.wb1.sheets[0].range('A7').value = [None, None]
        assert_equal(None, self.wb1.sheets[0].range('A7').expand('horizontal').value)

    def test_ndim2_scalar(self):
        """test_atleast_2d_scalar: Covers GH Issue #53a"""
        self.wb1.sheets[0].range('A50').value = 23
        result = self.wb1.sheets[0].range('A50').options(ndim=2).value
        assert_equal([[23]], result)

    def test_write_single_value_to_multicell_range(self):
        self.wb1.sheets[0].range('A1:B2').value = 5
        assert_equal(self.wb1.sheets[0].range('A1:B2').value, [[5., 5.], [5., 5.]])


class TestNumpy(TestBase):
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

        # 1d array (ndim=2)
        self.wb1.sheets[0].range('A10').value = array_1d
        cells = self.wb1.sheets[0].range('A10:D10').options(np.array, ndim=2).value
        assert_array_equal(cells, np.atleast_2d(array_1d))

        # 2d array (ndim=2)
        self.wb1.sheets[0].range('A12').value = array_2d
        cells = self.wb1.sheets[0].range('A12').options(np.array, ndim=2, expand='table').value
        assert_array_equal(cells, array_2d)

    def test_numpy_datetime(self):
        _skip_if_no_numpy()

        self.wb1.sheets[0].range('A55').value = np.datetime64('2005-02-25T03:30Z')
        assert_equal(self.wb1.sheets[0].range('A55').value, dt.datetime(2005, 2, 25, 3, 30))

    def test_scalar_nan(self):
        """test_scalar_nan: Covers GH Issue #15"""
        _skip_if_no_numpy()

        self.wb1.sheets[0].range('A20').value = np.nan
        assert_equal(None, self.wb1.sheets[0].range('A20').value)

    def test_ndim2_scalar_as_array(self):
        """test_atleast_2d_scalar_as_array: Covers GH Issue #53b"""
        _skip_if_no_numpy()

        self.wb1.sheets[0].range('A50').value = 23
        result = self.wb1.sheets[0].range('A50').options(np.array, ndim=2).value
        assert_equal(np.array([[23]]), result)

    def test_float64(self):
        self.wb1.sheets[0].range('A1').value = np.float64(2)
        assert_equal(self.wb1.sheets[0].range('A1').value, 2.)

    def test_int64(self):
        self.wb1.sheets[0].range('A1').value = np.int64(2)
        assert_equal(self.wb1.sheets[0].range('A1').value, 2.)


class TestPandas(TestBase):
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
        cells = self.wb1.sheets[0].range('D20').expand('table').value
        multiindex = self.wb1.sheets[0].range('A20:C28').value
        ix = pd.MultiIndex.from_tuples(multiindex[1:], names=multiindex[0])
        df_result = DataFrame(cells[1:], columns=cells[0], index=ix)
        assert_frame_equal(df_expected, df_result)

    def test_dataframe_multiheader(self):
        _skip_if_no_pandas()

        df_expected = df_multiheader
        self.wb1.sheets[0].range('A52').value = df_expected
        cells = self.wb1.sheets[0].range('B52').expand('table').value
        df_result = DataFrame(cells[2:], columns=pd.MultiIndex.from_arrays(cells[:2]))
        assert_frame_equal(df_expected, df_result)

    def test_dataframe_dateindex(self):
        _skip_if_no_pandas()

        df_expected = df_dateindex
        self.wb1.sheets[0].range('A100').value = df_expected
        if sys.platform.startswith('win') and self.wb1.app.version == '14.0':
            self.wb1.sheets[0].range(
                'A100').expand('vertical').number_format = 'dd/mm/yyyy'  # Hack for Excel 2010 bug, see GH #43
        cells = self.wb1.sheets[0].range('B100').expand('table').value
        index = self.wb1.sheets[0].range('A101').expand('vertical').value
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

        self.wb1.sheets[0].range('A1').options(pd.DataFrame, index=False, header=True).value = pd.DataFrame(
            [[1., 2.], [3., 4.]], columns=['a', 'b'])
        df = self.wb1.sheets[0].range('A1').options(pd.DataFrame, index=False, header=True,
                                                    expand='table').value
        assert_frame_equal(df, pd.DataFrame([[1., 2.], [3., 4.]], columns=['a', 'b']))

    def test_df_0header_1index(self):
        _skip_if_no_pandas()

        self.wb1.sheets[0].range('A1').options(pd.DataFrame, index=True, header=False).value = pd.DataFrame(
            [[1., 2.], [3., 4.]], index=[10., 20.])
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
            self.wb1.sheets[0].range('A40').expand('vertical').number_format = 'dd/mm/yyyy'  # Hack for Excel 2010 bug, see GH #43
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
        assert_equal([[1.], [2.], [3.]], self.wb1.sheets[0].range('A1').options(ndim=2, expand='table').value)

    def test_write_series_noheader_index(self):
        _skip_if_no_pandas()

        self.wb1.sheets[0].range('A1').options(index=True).value = pd.Series([1., 2., 3.], index=[10., 20., 30.])
        assert_equal([[10., 1.], [20., 2.], [30., 3.]],
                     self.wb1.sheets[0].range('A1').options(ndim=2, expand='table').value)

    def test_write_series_header_noindex(self):
        _skip_if_no_pandas()

        self.wb1.sheets[0].range('A1').options(index=False).value = pd.Series([1., 2., 3.], name='name')
        assert_equal([['name'], [1.], [2.], [3.]], self.wb1.sheets[0].range('A1').options(ndim=2, expand='table').value)

    def test_write_series_header_index(self):
        _skip_if_no_pandas()

        # Named index
        self.wb1.sheets[0].range('A1').value = pd.Series([1., 2., 3.], name='name',
                                                         index=pd.Index([10., 20., 30.], name='ix'))
        assert_equal([['ix', 'name'], [10., 1.], [20., 2.], [30., 3.]],
                     self.wb1.sheets[0].range('A1').options(ndim=2, expand='table').value)

        # Nameless index
        self.wb1.sheets[0].range('A1').value = pd.Series([1., 2., 3.], name='name', index=[10., 20., 30.])
        assert_equal([[None, 'name'], [10., 1.], [20., 2.], [30., 3.]],
                     self.wb1.sheets[0].range('A1:B4').options(ndim=2).value)

    def test_dataframe_timezone(self):
        _skip_if_no_pandas()

        np_dt = np.datetime64(1434149887000, 'ms')
        ix = pd.DatetimeIndex(data=[np_dt], tz='GMT')
        df = pd.DataFrame(data=[1], index=ix, columns=['A'])
        self.wb1.sheets[0].range('A1').value = df
        assert_equal(self.wb1.sheets[0].range('A2').value, dt.datetime(2015, 6, 12, 22, 58, 7))

    def test_NaT(self):
        df = pd.DataFrame([pd.Timestamp('20120102'), np.nan], index=[0., 1.], columns=['one'])
        self.wb1.sheets[0].range('A1').value = df
        assert_frame_equal(df, self.wb1.sheets[0].range('A1').options(pd.DataFrame, expand='table').value)