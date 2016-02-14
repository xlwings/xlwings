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


@xw.func
def test_simple(x):
    return x * 2


@xw.func
@xw.arg('x', ndim=1)
def test_ndim1(x):
    return x == [2.]


@xw.func
@xw.arg('x', ndim=2)
def test_ndim2(x):
    return x == [[2.]]


@xw.func
def test_2dlist(x):
    return [[cell + 1 for cell in row] for row in x] 


@xw.func
@xw.ret(transpose=True)
def test_transpose(x):
    return x


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

