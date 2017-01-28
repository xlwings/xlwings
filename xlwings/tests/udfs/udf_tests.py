from datetime import datetime, date
import sys
if sys.version_info >= (2, 7):
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
@xw.arg('x', dates=date)
def read_dates_as1(x):
    return x == [[1., date(2015, 1, 13)], [date(2000, 12, 1), 4.]]

@xw.func
@xw.arg('x', dates=date)
def read_dates_as2(x):
    return x == date(2005, 1, 15)

@xw.func
@xw.arg('x', dates=datetime)
def read_dates_as3(x):
    return x == [[1., datetime(2015, 1, 13)], [datetime(2000, 12, 1), 4.]]

@xw.func
@xw.arg('x', empty='empty')
def read_empty_as(x):
    return x == [[1., 'empty'], ['empty', 4.]]

if sys.version_info >= (2, 7):
    # assert_dict_equal isn't available on nose for PY 2.6

    # Dicts
    @xw.func
    @xw.arg('x', dict)
    def read_dict(x):
        return dict_equal(x, {'a': 1., 'b': 'c'})

    @xw.func
    @xw.arg('x', dict, transpose=True)
    def read_dict_transpose(x):
        return dict_equal(x, {1.0: 'c', 'a': 'b'})

@xw.func
def write_dict():
    return {'a': 1., 'b': 'c'}

# Numpy Array
if np:

    @xw.func
    @xw.arg('x', np.array)
    def read_scalar_nparray(x):
        return nparray_equal(x, np.array(1.))

    @xw.func
    @xw.arg('x', np.array)
    def read_empty_nparray(x):
        return nparray_equal(x, np.array(np.nan))

    @xw.func
    @xw.arg('x', np.array)
    def read_horizontal_nparray(x):
        return nparray_equal(x, np.array([1., 2.]))

    @xw.func
    @xw.arg('x', np.array)
    def read_vertical_nparray(x):
        return nparray_equal(x, np.array([1., 2.]))

    @xw.func
    @xw.arg('x', np.array)
    def read_date_nparray(x):
        return nparray_equal(x, np.array(datetime(2000, 12, 20)))

    # Keyword args on Numpy arrays

    @xw.func
    @xw.arg('x', np.array, ndim=1)
    def read_ndim1_nparray(x):
        return nparray_equal(x, np.array([2.]))

    @xw.func
    @xw.arg('x', np.array, ndim=2)
    def read_ndim2_nparray(x):
        return nparray_equal(x, np.array([[2.]]))

    @xw.func
    @xw.arg('x', np.array, transpose=True)
    def read_transpose_nparray(x):
        return nparray_equal(x, np.array([[1., 3.], [2., 4.]]))

    @xw.func
    @xw.ret(transpose=True)
    def write_transpose_nparray():
        return np.array([[1., 2.], [3., 4.]])

    @xw.func
    @xw.arg('x', np.array, dates=date)
    def read_dates_as_nparray(x):
        return nparray_equal(x, np.array(date(2000, 12, 20)))

    @xw.func
    @xw.arg('x', np.array, empty='empty')
    def read_empty_as_nparray(x):
        return nparray_equal(x, np.array('empty'))

    @xw.func
    def write_np_scalar():
        return np.float64(2)

# Pandas Series

if pd:

    @xw.func
    @xw.arg('x', pd.Series, header=False, index=False)
    def read_series_noheader_noindex(x):
        return series_equal(x, pd.Series([1., 2.]))

    @xw.func
    @xw.arg('x', pd.Series, header=False, index=True)
    def read_series_noheader_index(x):
        return series_equal(x, pd.Series([1., 2.], index=[10., 20.]))

    @xw.func
    @xw.arg('x', pd.Series, header=True, index=False)
    def read_series_header_noindex(x):
        return series_equal(x, pd.Series([1., 2.], name='name'))

    @xw.func
    @xw.arg('x', pd.Series, header=True, index=True)
    def read_series_header_named_index(x):
        return series_equal(x, pd.Series([1., 2.], name='name', index=pd.Index([10., 20.], name='ix')))

    @xw.func
    @xw.arg('x', pd.Series, header=True, index=True)
    def read_series_header_nameless_index(x):
        return series_equal(x, pd.Series([1., 2.], name='name', index=[10., 20.]))

    @xw.func
    @xw.arg('x', pd.Series, header=True, index=2)
    def read_series_header_nameless_2index(x):
        ix = pd.MultiIndex.from_arrays([['a', 'a'], [10., 20.]])
        return series_equal(x, pd.Series([1., 2.], name='name', index=ix))

    @xw.func
    @xw.arg('x', pd.Series, header=True, index=2)
    def read_series_header_named_2index(x):
        ix = pd.MultiIndex.from_arrays([['a', 'a'], [10., 20.]], names=['ix1', 'ix2'])
        return series_equal(x, pd.Series([1., 2.], name='name', index=ix))

    @xw.func
    @xw.arg('x', pd.Series, header=False, index=2)
    def read_series_noheader_2index(x):
        ix = pd.MultiIndex.from_arrays([['a', 'a'], [10., 20.]])
        return series_equal(x, pd.Series([1., 2.], index=ix))

    @xw.func
    @xw.ret(pd.Series, index=False)
    def write_series_noheader_noindex():
        return pd.Series([1., 2.])

    @xw.func
    @xw.ret(pd.Series, index=True)
    def write_series_noheader_index():
        return pd.Series([1., 2.], index=[10., 20.])

    @xw.func
    @xw.ret(pd.Series, index=False)
    def write_series_header_noindex():
        return pd.Series([1., 2.], name='name')

    @xw.func
    def write_series_header_named_index():
        return pd.Series([1., 2.], name='name', index=pd.Index([10., 20.], name='ix'))

    @xw.func
    @xw.ret(pd.Series, index=True, header=True)
    def write_series_header_nameless_index():
        return pd.Series([1., 2.], name='name', index=[10., 20.])

    @xw.func
    @xw.ret(pd.Series, header=True, index=2)
    def write_series_header_nameless_2index():
        ix = pd.MultiIndex.from_arrays([['a', 'a'], [10., 20.]])
        return pd.Series([1., 2.], name='name', index=ix)

    @xw.func
    @xw.ret(pd.Series, header=True, index=2)
    def write_series_header_named_2index():
        ix = pd.MultiIndex.from_arrays([['a', 'a'], [10., 20.]], names=['ix1', 'ix2'])
        return pd.Series([1., 2.], name='name', index=ix)

    @xw.func
    @xw.ret(pd.Series, header=False, index=2)
    def write_series_noheader_2index():
        ix = pd.MultiIndex.from_arrays([['a', 'a'], [10., 20.]])
        return pd.Series([1., 2.], index=ix)

    @xw.func
    @xw.arg('x', pd.Series)
    def read_timeseries(x):
        return series_equal(x, pd.Series([1.5, 2.5], name='ts', index=[datetime(2000, 12, 20), datetime(2000, 12, 21)]))

    @xw.func
    @xw.ret(pd.Series)
    def write_timeseries():
        return pd.Series([1.5, 2.5], name='ts', index=[datetime(2000, 12, 20), datetime(2000, 12, 21)])

    @xw.func
    @xw.ret(pd.Series, index=False)
    def write_series_nan():
        return pd.Series([1, np.nan, 3])

# Pandas DataFrame

if pd:

    @xw.func
    @xw.arg('x', pd.DataFrame, index=False, header=False)
    def read_df_0header_0index(x):
        return frame_equal(x, pd.DataFrame([[1., 2.], [3., 4.]]))

    @xw.func
    @xw.ret(pd.DataFrame, index=False, header=False)
    def write_df_0header_0index():
        return pd.DataFrame([[1., 2.], [3., 4.]])

    @xw.func
    @xw.arg('x', pd.DataFrame, index=False, header=True)
    def read_df_1header_0index(x):
        return frame_equal(x, pd.DataFrame([[1., 2.], [3., 4.]], columns=['a', 'b']))

    @xw.func
    @xw.ret(pd.DataFrame, index=False, header=True)
    def write_df_1header_0index():
        return pd.DataFrame([[1., 2.], [3., 4.]], columns=['a', 'b'])

    @xw.func
    @xw.arg('x', pd.DataFrame, index=True, header=False)
    def read_df_0header_1index(x):
        return frame_equal(x, pd.DataFrame([[1., 2.], [3., 4.]], index=[10., 20.]))

    @xw.func
    @xw.ret(pd.DataFrame, index=True, header=False)
    def write_df_0header_1index():
        return pd.DataFrame([[1., 2.], [3., 4.]], index=[10, 20])

    @xw.func
    @xw.arg('x', pd.DataFrame, index=2, header=False)
    def read_df_0header_2index(x):
        df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.], [7., 8., 9.]],
                          index=pd.MultiIndex.from_arrays([['a', 'a', 'b'], [1., 2., 1.]]))
        return frame_equal(x, df)

    @xw.func
    @xw.ret(pd.DataFrame, index=2, header=False)
    def write_df_0header_2index():
        df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.], [7., 8., 9.]],
                          index=pd.MultiIndex.from_arrays([['a', 'a', 'b'], [1., 2., 1.]]))
        return df

    @xw.func
    @xw.arg('x', pd.DataFrame, index=1, header=1)
    def read_df_1header_1namedindex(x):
        df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.]],
                          index=[1., 2.],
                          columns=['c', 'd', 'c'])
        df.index.name = 'ix1'
        return frame_equal(x, df)

    @xw.func
    def write_df_1header_1namedindex():
        df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.]],
                          index=[1., 2.],
                          columns=['c', 'd', 'c'])
        df.index.name = 'ix1'
        return df

    @xw.func
    @xw.arg('x', pd.DataFrame, index=1, header=1)
    def read_df_1header_1unnamedindex(x):
        df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.]],
                          index=[1., 2.],
                          columns=['c', 'd', 'c'])
        return frame_equal(x, df)

    @xw.func
    def write_df_1header_1unnamedindex():
        df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.]],
                          index=[1., 2.],
                          columns=['c', 'd', 'c'])
        return df

    @xw.func
    @xw.arg('x', pd.DataFrame, index=False, header=2)
    def read_df_2header_0index(x):
        df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.]],
                          columns=pd.MultiIndex.from_arrays([['a', 'a', 'b'], ['c', 'd', 'c']]))
        return frame_equal(x, df)

    @xw.func
    @xw.ret(pd.DataFrame, index=False, header=2)
    def write_df_2header_0index():
        df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.]],
                          columns=pd.MultiIndex.from_arrays([['a', 'a', 'b'], ['c', 'd', 'c']]))
        return df

    @xw.func
    @xw.arg('x', pd.DataFrame, index=1, header=2)
    def read_df_2header_1namedindex(x):
        df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.]],
                          index=[1., 2.],
                          columns=pd.MultiIndex.from_arrays([['a', 'a', 'b'], ['c', 'd', 'c']]))
        df.index.name = 'ix1'
        return frame_equal(x, df)

    @xw.func
    def write_df_2header_1namedindex():
        df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.]],
                          index=[1., 2.],
                          columns=pd.MultiIndex.from_arrays([['a', 'a', 'b'], ['c', 'd', 'c']]))
        df.index.name = 'ix1'
        return df

    @xw.func
    @xw.arg('x', pd.DataFrame, index=1, header=2)
    def read_df_2header_1unnamedindex(x):
        df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.]],
                          index=[1., 2.],
                          columns=pd.MultiIndex.from_arrays([['a', 'a', 'b'], ['c', 'd', 'c']]))
        return frame_equal(x, df)

    @xw.func
    def write_df_2header_1unnamedindex():
        df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.]],
                          index=[1., 2.],
                          columns=pd.MultiIndex.from_arrays([['a', 'a', 'b'], ['c', 'd', 'c']]))
        return df

    @xw.func
    @xw.arg('x', pd.DataFrame, index=2, header=2)
    def read_df_2header_2namedindex(x):
        df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.], [7., 8., 9.]],
                          index=pd.MultiIndex.from_arrays([['a', 'a', 'b'], [1., 2., 1.]], names=['x1', 'x2']),
                          columns=pd.MultiIndex.from_arrays([['a', 'a', 'b'], ['c', 'd', 'c']]))
        return frame_equal(x, df)

    @xw.func
    def write_df_2header_2namedindex():
        df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.], [7., 8., 9.]],
                          index=pd.MultiIndex.from_arrays([['a', 'a', 'b'], [1., 2., 1.]], names=['x1', 'x2']),
                          columns=pd.MultiIndex.from_arrays([['a', 'a', 'b'], ['c', 'd', 'c']]))
        return df

    @xw.func
    @xw.arg('x', pd.DataFrame, index=2, header=2)
    def read_df_2header_2unnamedindex(x):
        df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.], [7., 8., 9.]],
                          index=pd.MultiIndex.from_arrays([['a', 'a', 'b'], [1., 2., 1.]]),
                          columns=pd.MultiIndex.from_arrays([['a', 'a', 'b'], ['c', 'd', 'c']]))
        return frame_equal(x, df)

    @xw.func
    def write_df_2header_2unnamedindex():
        df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.], [7., 8., 9.]],
                          index=pd.MultiIndex.from_arrays([['a', 'a', 'b'], [1., 2., 1.]]),
                          columns=pd.MultiIndex.from_arrays([['a', 'a', 'b'], ['c', 'd', 'c']]))
        return df

    @xw.func
    @xw.arg('x', pd.DataFrame, index=2, header=1)
    def read_df_1header_2namedindex(x):
        df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.], [7., 8., 9.]],
                          index=pd.MultiIndex.from_arrays([['a', 'a', 'b'], [1., 2., 1.]], names=['x1', 'x2']),
                          columns=['a', 'd', 'c'])
        return frame_equal(x, df)

    @xw.func
    def write_df_1header_2namedindex():
        df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.], [7., 8., 9.]],
                          index=pd.MultiIndex.from_arrays([['a', 'a', 'b'], [1., 2., 1.]], names=['x1', 'x2']),
                          columns=['a', 'd', 'c'])
        return df

    @xw.func
    @xw.arg('x', pd.DataFrame)
    def read_df_date_index(x):
        df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.]],
                          index=[datetime(1999,12,13), datetime(1999,12,14)],
                          columns=['c', 'd', 'c'])
        return frame_equal(x, df)

    @xw.func
    def write_df_date_index():
        df = pd.DataFrame([[1., 2., 3.], [4., 5., 6.]],
                          index=[datetime(1999,12,13), datetime(1999,12,14)],
                          columns=['c', 'd', 'c'])
        return df

    @xw.func
    def read_workbook_caller():
        wb = xw.Book.caller()
        return xw.Range('E277').value == 1.


@xw.func
def default_args(x, y="hello", z=20):
    return 2 * x + 3 * len(y) + 7 * z


@xw.func
def variable_args(x, *z):
    return 2 * x + 3 * len(z) + 7 * z[0]


@xw.func
def optional_args(x, y=None):
    if y is None:
        y = 10
    return x * y


@xw.func
def write_none():
    return None


@xw.func(category=1)
def category_1():
    return 'category 1'


@xw.func(category=2)
def category_2():
    return 'category 2'


@xw.func(category=3)
def category_3():
    return 'category 3'


@xw.func(category=4)
def category_4():
    return 'category 4'


@xw.func(category=5)
def category_5():
    return 'category 5'


@xw.func(category=6)
def category_6():
    return 'category 6'


@xw.func(category=7)
def category_7():
    return 'category 7'


@xw.func(category=8)
def category_8():
    return 'category 8'


@xw.func(category=9)
def category_9():
    return 'category 9'


@xw.func(category=10)
def category_10():
    return 'category 10'


@xw.func(category=11)
def category_11():
    return 'category 11'


@xw.func(category=12)
def category_12():
    return 'category 12'


@xw.func(category=13)
def category_13():
    return 'category 13'


@xw.func(category=14)
def category_14():
    return 'category 14'

try:
    @xw.func(category=15)
    def category_15():
        return 'category 15'
except Exception as e:
    assert e.args[0] == 'There is only 14 build-in categories available in Excel. Please use a string value to specify a custom category.'
else:
    assert False

try:
    @xw.func(category=0)
    def category_0():
        return 'category 0'
except Exception as e:
    assert e.args[0] == 'There is only 14 build-in categories available in Excel. Please use a string value to specify a custom category.'
else:
    assert False


@xw.func(category='custom category')
def custom_category():
    return 'custom category'


try:
    @xw.func(category=1.1)
    def object_category():
        return 'object category'
except Exception as e:
    assert e.args[0] == 'Category 1.1 should either be a predefined Excel category (int value) or a custom one (str value).'
else:
    assert False


@xw.func
def method_signature_with_less_than_1024_characters(
        very_long_parameter_name_1=None,
        very_long_parameter_name_2=None,
        very_long_parameter_name_3=None,
        very_long_parameter_name_4=None,
        very_long_parameter_name_5=None,
        very_long_parameter_name_6=None,
        very_long_parameter_name_7=None,
        very_long_parameter_name_8=None,
        very_long_parameter_name_9=None,
        very_long_parameter_name_10=None,
        very_long_parameter_name_11=None,
        very_long_parameter_name_12=None,
        very_long_parameter_name_13=None,
        very_long_parameter_name_14=None,
        very_long_parameter_name_15=None,
        very_long_parameter_name_16=None,
        very_long_parameter_name_17=None,
        very_long_parameter_name_18=None,
        very_long_parameter_name_19=None,
        very_long_parameter_name_20=None,
        very_long_parameter_name_21=None,
        very_long_parameter_name_22=None,
        very_long_parameter_name_23=None,
        very_long_parameter_name_24=None,
        very_long_parameter_name_25=None,
        paramet_name_26=None
):
    return 'non splitted signature'


@xw.func
def method_signature_with_more_than_1024_characters(
        very_long_parameter_name_1=None,
        very_long_parameter_name_2=None,
        very_long_parameter_name_3=None,
        very_long_parameter_name_4=None,
        very_long_parameter_name_5=None,
        very_long_parameter_name_6=None,
        very_long_parameter_name_7=None,
        very_long_parameter_name_8=None,
        very_long_parameter_name_9=None,
        very_long_parameter_name_10=None,
        very_long_parameter_name_11=None,
        very_long_parameter_name_12=None,
        very_long_parameter_name_13=None,
        very_long_parameter_name_14=None,
        very_long_parameter_name_15=None,
        very_long_parameter_name_16=None,
        very_long_parameter_name_17=None,
        very_long_parameter_name_18=None,
        very_long_parameter_name_19=None,
        very_long_parameter_name_20=None,
        very_long_parameter_name_21=None,
        very_long_parameter_name_22=None,
        very_long_parameter_name_23=None,
        very_long_parameter_name_24=None,
        very_long_parameter_name_25=None,
        very_long_parameter_name_26=None
):
    return 'splitted signature'


if __name__ == "__main__":
    xw.serve()
