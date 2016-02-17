from datetime import datetime, date
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
except ImportError:
    pd = None





# Defaults
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
def test_read_datetime(x):
    return x == datetime(1976, 2, 15, 13, 6, 22)

@xw.func
def test_write_datetime():
    return datetime(1976, 2, 15, 13, 6, 23)

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

# Keyword args on default converters

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
def test_read_dates_as1(x):
    return x == [[1., date(2015, 1, 13)], [date(2000, 12, 1), 4.]]

@xw.func
@xw.arg('x', dates_as=date)
def test_read_dates_as2(x):
    return x == date(2005, 1, 15)

@xw.func
@xw.arg('x', dates_as=datetime)
def test_read_dates_as3(x):
    return x == [[1., datetime(2015, 1, 13)], [datetime(2000, 12, 1), 4.]]

@xw.func
@xw.arg('x', empty_as='empty')
def test_read_empty_as(x):
    return x == [[1., 'empty'], ['empty', 4.]]



# Numpy Array
@xw.func
@xw.arg('x', as_=np.array)
def test_read_scalar_nparray(x):
    return nparray_equal(x, np.array(1.))

@xw.func
@xw.arg('x', as_=np.array)
def test_read_empty_nparray(x):
    return nparray_equal(x, np.array(np.nan))

@xw.func
@xw.arg('x', as_=np.array)
def test_read_horizontal_nparray(x):
    return nparray_equal(x, np.array([1., 2.]))

@xw.func
@xw.arg('x', as_=np.array)
def test_read_vertical_nparray(x):
    return nparray_equal(x, np.array([1., 2.]))

@xw.func
@xw.arg('x', as_=np.array)
def test_read_date_nparray(x):
    return nparray_equal(x, np.array(datetime(2000, 12, 20)))

# Keyword args on Numpy arrays

@xw.func
@xw.arg('x', as_=np.array, ndim=1)
def test_read_ndim1_nparray(x):
    return nparray_equal(x, np.array([2.]))

@xw.func
@xw.arg('x', as_=np.array, ndim=2)
def test_read_ndim2_nparray(x):
    return nparray_equal(x, np.array([[2.]]))

@xw.func
@xw.arg('x', as_=np.array, transpose=True)
def test_read_transpose_nparray(x):
    return nparray_equal(x, np.array([[1., 3.], [2., 4.]]))

@xw.func
@xw.ret(transpose=True)
def test_write_transpose_nparray():
    return np.array([[1., 2.], [3., 4.]])

@xw.func
@xw.arg('x', as_=np.array, dates_as=date)
def test_read_dates_as_nparray(x):
    return nparray_equal(x, np.array(date(2000, 12, 20)))

@xw.func
@xw.arg('x', as_=np.array, empty_as='empty')
def test_read_empty_as_nparray(x):
    return nparray_equal(x, np.array('empty'))



