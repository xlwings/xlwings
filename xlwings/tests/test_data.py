from datetime import datetime

try:
    import numpy as np
    from numpy.testing import assert_array_equal
except ImportError:
    np = None
try:
    import pandas as pd
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

if np:
    array_1d = np.array([1.1, 2.2, np.nan, -4.4])
    array_2d = np.array([[1.1, 2.2, 3.3], [-4.4, 5.5, np.nan]])

if pd:
    series_1 = pd.Series([1.1, 3.3, 5., np.nan, 6., 8.])

    rng = pd.date_range('1/1/2012', periods=10, freq='D')
    timeseries_1 = pd.Series(np.arange(len(rng)) + 0.1, rng)
    timeseries_1[1] = np.nan

    df_1 = pd.DataFrame([[1, 'test1'],
                         [2, 'test2'],
                         [np.nan, None],
                         [3.3, 'test3']], columns=['a', 'b'])

    df_2 = pd.DataFrame([1, 3, 5, np.nan, 6, 8], columns=['col1'])

    df_dateindex = pd.DataFrame(np.arange(50).reshape(10, 5) + 0.1, index=rng,
                                columns=['one', 'two', 'three', 'four', 'five'])

    # MultiIndex (Index)
    tuples = list(zip(*[['bar', 'bar', 'baz', 'baz', 'foo', 'foo', 'qux', 'qux'],
                        ['one', 'two', 'one', 'two', 'one', 'two', 'one', 'two'],
                        ['x', 'x', 'x', 'x', 'y', 'y', 'y', 'y']]))
    index = pd.MultiIndex.from_tuples(tuples, names=['first', 'second', 'third'])
    df_multiindex = pd.DataFrame([[1.1, 2.2], [3.3, 4.4], [5.5, 6.6], [7.7, 8.8], [9.9, 10.10],
                                  [11.11, 12.12], [13.13, 14.14], [15.15, 16.16]], index=index, columns=['one', 'two'])

    # MultiIndex (Header)
    header = [['Foo', 'Foo', 'Bar', 'Bar', 'Baz'], ['A', 'B', 'C', 'D', 'E']]

    df_multiheader = pd.DataFrame([[0.0, 1.0, 2.0, 3.0, 4.0],
                                   [0.0, 1.0, 2.0, 3.0, 4.0],
                                   [0.0, 1.0, 2.0, 3.0, 4.0],
                                   [0.0, 1.0, 2.0, 3.0, 4.0],
                                   [0.0, 1.0, 2.0, 3.0, 4.0],
                                   [0.0, 1.0, 2.0, 3.0, 4.0]], columns=pd.MultiIndex.from_arrays(header))