from datetime import datetime, date
from nose.tools import assert_dict_equal
import xlwings as xw
try:
    import numpy as np
    from numpy.testing import assert_array_equal

    def nparray_equal(a, b):
        try:
            assert_array_equal(a, b)
        except AssertionError:
            return False
        return True

except ImportError:
    np = None
try:
    import pandas as pd
    from pandas import DataFrame, Series
    from pandas.util.testing import assert_frame_equal, assert_series_equal

    def frame_equal(a, b):
        try:
            assert_frame_equal(a, b)
        except AssertionError:
            return False
        return True

    def series_equal(a, b):
        try:
            assert_series_equal(a, b)
        except AssertionError:
            return False
        return True

except ImportError:
    pd = None


def dict_equal(a, b):
    try:
        assert_dict_equal(a, b)
    except AssertionError:
        return False
    return True

# Defaults
@xw.func
def read_float(x):
    return x == 2.

@xw.func
def write_float():
    return 2.

@xw.func
def read_string(x):
    return x == 'xlwings'

@xw.func
def write_string():
    return 'xlwings'

@xw.func
def read_empty(x):
    return x is None

@xw.func
def read_date(x):
    return x == datetime(2015, 1, 15)

@xw.func
def write_date():
    return datetime(1969, 12, 31)

@xw.func
def read_datetime(x):
    return x == datetime(1976, 2, 15, 13, 6, 22)

@xw.func
def write_datetime():
    return datetime(1976, 2, 15, 13, 6, 23)

@xw.func
def read_horizontal_list(x):
    return x == [1., 2.]

@xw.func
def write_horizontal_list():
    return [1., 2.]

@xw.func
def read_vertical_list(x):
    return x == [1., 2.]

@xw.func
def write_vertical_list():
    return [[1.], [2.]]

@xw.func
def read_2dlist(x):
    return x == [[1., 2.], [3., 4.]]

@xw.func
def write_2dlist():
    return [[1., 2.], [3., 4.]]

# Keyword args on default converters

@xw.func
@xw.arg('x', ndim=1)
def read_ndim1(x):
    return x == [2.]

@xw.func
@xw.arg('x', ndim=2)
def read_ndim2(x):
    return x == [[2.]]

@xw.func
@xw.arg('x', transpose=True)
def read_transpose(x):
    return x == [[1., 3.], [2., 4.]]

@xw.func
@xw.ret(transpose=True)
def write_transpose():
    return [[1., 2.], [3., 4.]]

@xw.func
@xw.arg('x', dates_as=date)
def read_dates_as1(x):
    return x == [[1., date(2015, 1, 13)], [date(2000, 12, 1), 4.]]

@xw.func
@xw.arg('x', dates_as=date)
def read_dates_as2(x):
    return x == date(2005, 1, 15)

@xw.func
@xw.arg('x', dates_as=datetime)
def read_dates_as3(x):
    return x == [[1., datetime(2015, 1, 13)], [datetime(2000, 12, 1), 4.]]

@xw.func
@xw.arg('x', empty_as='empty')
def read_empty_as(x):
    return x == [[1., 'empty'], ['empty', 4.]]

# Dicts
@xw.func
@xw.arg('x', as_=dict)
def read_dict(x):
    return dict_equal(x, {'a': 1., 'b': 'c'})

@xw.func
@xw.arg('x', as_=dict, transpose=True)
def read_dict_transpose(x):
    return dict_equal(x, {1.0: 'c', 'a': 'b'})

@xw.func
def write_dict():
    return {'a': 1., 'b': 'c'}

# Numpy Array
@xw.func
@xw.arg('x', as_=np.array)
def read_scalar_nparray(x):
    return nparray_equal(x, np.array(1.))

@xw.func
@xw.arg('x', as_=np.array)
def read_empty_nparray(x):
    return nparray_equal(x, np.array(np.nan))

@xw.func
@xw.arg('x', as_=np.array)
def read_horizontal_nparray(x):
    return nparray_equal(x, np.array([1., 2.]))

@xw.func
@xw.arg('x', as_=np.array)
def read_vertical_nparray(x):
    return nparray_equal(x, np.array([1., 2.]))

@xw.func
@xw.arg('x', as_=np.array)
def read_date_nparray(x):
    return nparray_equal(x, np.array(datetime(2000, 12, 20)))

# Keyword args on Numpy arrays

@xw.func
@xw.arg('x', as_=np.array, ndim=1)
def read_ndim1_nparray(x):
    return nparray_equal(x, np.array([2.]))

@xw.func
@xw.arg('x', as_=np.array, ndim=2)
def read_ndim2_nparray(x):
    return nparray_equal(x, np.array([[2.]]))

@xw.func
@xw.arg('x', as_=np.array, transpose=True)
def read_transpose_nparray(x):
    return nparray_equal(x, np.array([[1., 3.], [2., 4.]]))

@xw.func
@xw.ret(transpose=True)
def write_transpose_nparray():
    return np.array([[1., 2.], [3., 4.]])

@xw.func
@xw.arg('x', as_=np.array, dates_as=date)
def read_dates_as_nparray(x):
    return nparray_equal(x, np.array(date(2000, 12, 20)))

@xw.func
@xw.arg('x', as_=np.array, empty_as='empty')
def read_empty_as_nparray(x):
    return nparray_equal(x, np.array('empty'))

# DataFrame

@xw.func
@xw.arg('x', as_=pd.Series, header=False, index=False)
def read_series_noheader_noindex(x):
    return series_equal(x, pd.Series([1., 2.]))

@xw.func
@xw.arg('x', as_=pd.Series, header=False, index=True)
def read_series_noheader_index(x):
    return series_equal(x, pd.Series([1., 2.], index=[10., 20.]))

@xw.func
@xw.arg('x', as_=pd.Series, header=True, index=False)
def read_series_header_noindex(x):
    return series_equal(x, pd.Series([1., 2.], name='name'))

@xw.func
@xw.arg('x', as_=pd.Series, header=True, index=True)
def read_series_header_named_index(x):
    return series_equal(x, pd.Series([1., 2.], name='name', index=pd.Index([10., 20.], name='ix')))

@xw.func
@xw.arg('x', as_=pd.Series, header=True, index=True)
def read_series_header_nameless_index(x):
    return series_equal(x, pd.Series([1., 2.], name='name', index=[10., 20.]))


@xw.func
@xw.ret(as_=pd.Series, index=False)
def write_series_noheader_noindex():
    return pd.Series([1., 2.])

@xw.func
@xw.ret(as_=pd.Series, index=True)
def write_series_noheader_index():
    return pd.Series([1., 2.], index=[10., 20.])

@xw.func
@xw.ret(as_=pd.Series, index=False)
def write_series_header_noindex():
    return pd.Series([1., 2.], name='name')

@xw.func
def write_series_header_named_index():
    return pd.Series([1., 2.], name='name', index=pd.Index([10., 20.], name='ix'))

@xw.func
@xw.ret(as_=pd.Series, index=True, header=True)
def write_series_header_nameless_index():
    return pd.Series([1., 2.], name='name', index=[10., 20.])

@xw.func
@xw.arg('x', as_=pd.Series)
def read_timeseries(x):
    return series_equal(x, pd.Series([1.5, 2.5], name='ts', index=[datetime(2000, 12, 20), datetime(2000, 12, 21)]))

@xw.func
@xw.ret(as_=pd.Series)
def write_timeseries():
    return pd.Series([1.5, 2.5], name='ts', index=[datetime(2000, 12, 20), datetime(2000, 12, 21)])

@xw.func
@xw.arg('x', as_=pd.DataFrame, index=False, header=False)
def read_df_0header_0index(x):
    return frame_equal(x, pd.DataFrame([[1., 2.], [3., 4.]]))

@xw.func
@xw.arg('x', as_=pd.DataFrame, index=1, header=1)
def read_df_1header_1namedindex(x):

    df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.]],
                      index=[1., 2.],
                      columns=['c', 'd', 'c'])
    df.index.name = 'ix1'
    return frame_equal(x, df)

@xw.func
@xw.arg('x', as_=pd.DataFrame, index=1, header=1)
def read_df_1header_1unnamedindex(x):
    df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.]],
                      index=[1., 2.],
                      columns=['c', 'd', 'c'])
    return frame_equal(x, df)

@xw.func
@xw.arg('x', as_=pd.DataFrame, index=1, header=2)
def read_df_2header_1namedindex(x):
    df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.]],
                      index=[1., 2.],
                      columns=pd.MultiIndex.from_arrays([['a', 'a', 'b'], ['c', 'd', 'c']]))
    df.index.name = 'ix1'
    return frame_equal(x, df)

@xw.func
@xw.arg('x', as_=pd.DataFrame, index=1, header=2)
def read_df_2header_1unnamedindex(x):
    df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.]],
                      index=[1., 2.],
                      columns=pd.MultiIndex.from_arrays([['a', 'a', 'b'], ['c', 'd', 'c']]))
    return frame_equal(x, df)

@xw.func
@xw.arg('x', as_=pd.DataFrame, index=2, header=2)
def read_df_2header_2namedindex(x):
    df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.], [7., 8., 9.]],
                      index=pd.MultiIndex.from_arrays([['a', 'a', 'b'], [1., 2., 1.]], names=['x1', 'x2']),
                      columns=pd.MultiIndex.from_arrays([['a', 'a', 'b'], ['c', 'd', 'c']]))
    return frame_equal(x, df)

@xw.func
@xw.arg('x', as_=pd.DataFrame, index=2, header=2)
def read_df_2header_2unnamedindex(x):
    df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.], [7., 8., 9.]],
                      index=pd.MultiIndex.from_arrays([['a', 'a', 'b'], [1., 2., 1.]]),
                      columns=pd.MultiIndex.from_arrays([['a', 'a', 'b'], ['c', 'd', 'c']]))
    return frame_equal(x, df)

@xw.func
@xw.arg('x', as_=pd.DataFrame, index=2, header=1)
def read_df_1header_2namedindex(x):
    df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.], [7., 8., 9.]],
                      index=pd.MultiIndex.from_arrays([['a', 'a', 'b'], [1., 2., 1.]], names=['x1', 'x2']),
                      columns=['a', 'd', 'c'])
    return frame_equal(x, df)

@xw.func
@xw.arg('x', as_=pd.DataFrame)
def read_df_date_index(x):
    df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.]],
                      index=[datetime(1999,12,13), datetime(1999,12,14)],
                      columns=['c', 'd', 'c'])
    return frame_equal(x, df)