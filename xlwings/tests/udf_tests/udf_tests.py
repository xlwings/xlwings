from datetime import datetime, date
import xlwings as xw
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

# Default
@xw.func
def test_read_float(x):
    return x == 2.

@xw.func
def test_write_float():
    return 2.

@xw.func
def test_read_string(x):
    return x == 'xlwings'

@xw.func
def test_write_string():
    return 'xlwings'

@xw.func
def test_read_empty(x):
    return x is None

@xw.func
def test_read_date(x):
    return x == datetime(2015, 1, 15)

@xw.func
def test_write_date():
    return datetime(1969, 12, 31)

@xw.func
def test_read_horizontal_list(x):
    return x == [1., 2.]

@xw.func
def test_write_horizontal_list():
    return [1., 2.]

@xw.func
def test_read_vertical_list(x):
    return x == [1., 2.]

@xw.func
def test_write_vertical_list():
    return [[1.], [2.]]

@xw.func
def test_read_2dlist(x):
    return x == [[1., 2.], [3., 4.]]

@xw.func
def test_write_2dlist():
    return [[1., 2.], [3., 4.]]

@xw.func
@xw.arg('x', ndim=1)
def test_read_ndim1(x):
    return x == [2.]

@xw.func
@xw.arg('x', ndim=2)
def test_read_ndim2(x):
    return x == [[2.]]

@xw.func
@xw.arg('x', transpose=True)
def test_read_transpose(x):
    return x == [[1., 3.], [2., 4.]]

@xw.func
@xw.ret(transpose=True)
def test_write_transpose():
    return [[1., 2.], [3., 4.]]

@xw.func
@xw.arg('x', dates_as=date)
def test_read_as_date(x):
    return x == [[1., date(2015, 1, 13)], [date(2000, 12, 1), 4.]]

@xw.func
@xw.arg('x', dates_as=datetime)
def test_read_as_datetime(x):
    return x == [[1., datetime(2015, 1, 13)], [datetime(2000, 12, 1), 4.]]

@xw.func
@xw.arg('x', empty_as='empty')
def test_read_empty_as(x):
    return x == [[1., 'empty'], ['empty', 4.]]



# Numpy Array
@xw.func
@xw.arg('x', as_=np.array)
def test_nparray1(x):
    return x * 2


@xw.func
@xw.arg('x', as_=np.array)
@xw.ret(as_=np.array, transpose=True)
def test_nparray_transpose1(x):
    return x


# DataFrame
@xw.func
@xw.arg('x', as_=pd.DataFrame, index=False, header=False)
def test_df_read_noindex_noheader(x):
    return x.equals(pd.DataFrame([[1., 2.], [3., 4.]]))

