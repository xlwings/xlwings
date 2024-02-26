"""
TODO: why is this here and and under root tests folder?

Key differences with COM UDFs:
* respects ints (COM always returns floats)
* returns 0 for empty cells. To get None like in COM, you need to set the formula to: =""
* caller range object not supported (caller address would easy to get though)
* reading datetime must be explicitly converted via dt.date / dt.datetime or parse_dates (pandas)
* writing datetime is now automatically formatting it as date in Excel
* categories aren't supported: replaced by namespaces
"""
import datetime as dt
from datetime import date, datetime

import xlwings as xw
from xlwings.server import arg, func, ret

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
    from pandas.testing import assert_frame_equal, assert_series_equal

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


# Defaults
@func
def read_float(x):
    return x == 2


@func
def write_float():
    return 2


@func
def read_string(x):
    return x == "xlwings"


@func
def write_string():
    return "xlwings"


@func
def read_empty(x):
    return x is None


@func
@arg("x", dt.datetime)
def read_date(x):
    print(x)
    return x == datetime(2015, 1, 15)


@func
def write_date():
    return datetime(1969, 12, 31)


@func
@arg("x", dt.datetime)
def read_datetime(x):
    return x == datetime(1976, 2, 15, 13, 6, 22)


@func
def write_datetime():
    return datetime(1976, 2, 15, 13, 6, 23)


@func
def read_horizontal_list(x):
    return x == [1, 2]


@func
def write_horizontal_list():
    return [1, 2]


@func
def read_vertical_list(x):
    return x == [1, 2]


@func
def write_vertical_list():
    return [[1], [2]]


@func
def read_2dlist(x):
    return x == [[1, 2], [3, 4]]


@func
def write_2dlist():
    return [[1, 2], [3, 4]]


# Keyword args on default converters


@func
@arg("x", ndim=1)
def read_ndim1(x):
    return x == [2]


@func
@arg("x", ndim=2)
def read_ndim2(x):
    return x == [[2]]


@func
@arg("x", transpose=True)
def read_transpose(x):
    return x == [[1, 3], [2, 4]]


@func
@ret(transpose=True)
def write_transpose():
    return [[1, 2], [3, 4]]


@func
def read_dates_as1(x):
    x[0][1] = xw.to_datetime(x[0][1]).date()
    x[1][0] = xw.to_datetime(x[1][0]).date()
    return x == [[1, date(2015, 1, 13)], [date(2000, 12, 1), 4]]


@func
@arg("x", dt.date)
def read_dates_as2(x):
    return x == date(2005, 1, 15)


@func
def read_dates_as3(x):
    x[0][1] = xw.to_datetime(x[0][1])
    x[1][0] = xw.to_datetime(x[1][0])
    return x == [[1, datetime(2015, 1, 13)], [datetime(2000, 12, 1), 4]]


@func
@arg("x", empty="empty")
def read_empty_as(x):
    return x == [[1, "empty"], ["empty", 4]]


# Dicts
@func
@arg("x", dict)
def read_dict(x):
    return x == {"a": 1, "b": "c"}


@func
@arg("x", dict, transpose=True)
def read_dict_transpose(x):
    return x == {1: "c", "a": "b"}


@func
def write_dict():
    return {"a": 1, "b": "c"}


# Numpy Array
if np:

    @func
    @arg("x", np.array)
    def read_scalar_nparray(x):
        return nparray_equal(x, np.array(1))

    @func
    @arg("x", np.array)
    def read_empty_nparray(x):
        return nparray_equal(x, np.array(np.nan))

    @func
    @arg("x", np.array)
    def read_horizontal_nparray(x):
        return nparray_equal(x, np.array([1, 2]))

    @func
    @arg("x", np.array)
    def read_vertical_nparray(x):
        return nparray_equal(x, np.array([1, 2]))

    @func
    @arg("x", dt.datetime)
    def read_date_nparray(x):
        x = np.array(x)
        return nparray_equal(x, np.array(datetime(2000, 12, 20)))

    # Keyword args on Numpy arrays

    @func
    @arg("x", np.array, ndim=1)
    def read_ndim1_nparray(x):
        return nparray_equal(x, np.array([2]))

    @func
    @arg("x", np.array, ndim=2)
    def read_ndim2_nparray(x):
        return nparray_equal(x, np.array([[2]]))

    @func
    @arg("x", np.array, transpose=True)
    def read_transpose_nparray(x):
        return nparray_equal(x, np.array([[1, 3], [2, 4]]))

    @func
    @ret(transpose=True)
    def write_transpose_nparray():
        return np.array([[1, 2], [3, 4]])

    @func
    @arg("x", dt.date)
    def read_dates_as_nparray(x):
        x = np.array(x)
        return nparray_equal(x, np.array(date(2000, 12, 20)))

    @func
    @arg("x", np.array, empty="empty")
    def read_empty_as_nparray(x):
        return nparray_equal(x, np.array("empty"))

    @func
    def write_np_scalar():
        return np.float64(2)


# Pandas Series

if pd:

    @func
    @arg("x", pd.Series, header=False, index=False)
    def read_series_noheader_noindex(x):
        return series_equal(x, pd.Series([1, 2]))

    @func
    @arg("x", pd.Series, header=False, index=True)
    def read_series_noheader_index(x):
        return series_equal(x, pd.Series([1, 2], index=[10, 20]))

    @func
    @arg("x", pd.Series, header=True, index=False)
    def read_series_header_noindex(x):
        return series_equal(x, pd.Series([1, 2], name="name"))

    @func
    @arg("x", pd.Series, header=True, index=True)
    def read_series_header_named_index(x):
        return series_equal(
            x,
            pd.Series([1, 2], name="name", index=pd.Index([10, 20], name="ix")),
        )

    @func
    @arg("x", pd.Series, header=True, index=True)
    def read_series_header_nameless_index(x):
        print(x)
        return series_equal(x, pd.Series([1, 2], name="name", index=[10, 20]))

    @func
    @arg("x", pd.Series, header=True, index=2)
    def read_series_header_nameless_2index(x):
        ix = pd.MultiIndex.from_arrays([["a", "a"], [10, 20]])
        return series_equal(x, pd.Series([1, 2], name="name", index=ix))

    @func
    @arg("x", pd.Series, header=True, index=2)
    def read_series_header_named_2index(x):
        ix = pd.MultiIndex.from_arrays([["a", "a"], [10, 20]], names=["ix1", "ix2"])
        return series_equal(x, pd.Series([1, 2], name="name", index=ix))

    @func
    @arg("x", pd.Series, header=False, index=2)
    def read_series_noheader_2index(x):
        ix = pd.MultiIndex.from_arrays([["a", "a"], [10, 20]])
        return series_equal(x, pd.Series([1, 2], index=ix))

    @func
    @ret(pd.Series, index=False)
    def write_series_noheader_noindex():
        return pd.Series([1, 2])

    @func
    @ret(pd.Series, index=True)
    def write_series_noheader_index():
        return pd.Series([1, 2], index=[10, 20])

    @func
    @ret(pd.Series, index=False)
    def write_series_header_noindex():
        return pd.Series([1, 2], name="name")

    @func
    def write_series_header_named_index():
        return pd.Series([1, 2], name="name", index=pd.Index([10, 20], name="ix"))

    @func
    @ret(pd.Series, index=True, header=True)
    def write_series_header_nameless_index():
        return pd.Series([1, 2], name="name", index=[10, 20])

    @func
    @ret(pd.Series, header=True, index=2)
    def write_series_header_nameless_2index():
        ix = pd.MultiIndex.from_arrays([["a", "a"], [10, 20]])
        return pd.Series([1, 2], name="name", index=ix)

    @func
    @ret(pd.Series, header=True, index=2)
    def write_series_header_named_2index():
        ix = pd.MultiIndex.from_arrays([["a", "a"], [10, 20]], names=["ix1", "ix2"])
        return pd.Series([1, 2], name="name", index=ix)

    @func
    @ret(pd.Series, header=False, index=2)
    def write_series_noheader_2index():
        ix = pd.MultiIndex.from_arrays([["a", "a"], [10, 20]])
        return pd.Series([1, 2], index=ix)

    @func
    @arg("x", pd.Series, parse_dates=True)
    def read_timeseries(x):
        return series_equal(
            x,
            pd.Series(
                [1.5, 2.5],
                name="ts",
                index=[datetime(2000, 12, 20), datetime(2000, 12, 21)],
            ),
        )

    @func
    @ret(pd.Series)
    def write_timeseries():
        return pd.Series(
            [1.5, 2.5],
            name="ts",
            index=[datetime(2000, 12, 20), datetime(2000, 12, 21)],
        )

    @func
    @ret(pd.Series, index=False)
    def write_series_nan():
        return pd.Series([1, np.nan, 3])


# Pandas DataFrame

if pd:

    @func
    @arg("x", pd.DataFrame, index=False, header=False)
    def read_df_0header_0index(x):
        return frame_equal(x, pd.DataFrame([[1, 2], [3, 4]]))

    @func
    @ret(pd.DataFrame, index=False, header=False)
    def write_df_0header_0index():
        return pd.DataFrame([[1, 2], [3, 4]])

    @func
    @arg("x", pd.DataFrame, index=False, header=True)
    def read_df_1header_0index(x):
        return frame_equal(x, pd.DataFrame([[1, 2], [3, 4]], columns=["a", "b"]))

    @func
    @ret(pd.DataFrame, index=False, header=True)
    def write_df_1header_0index():
        return pd.DataFrame([[1, 2], [3, 4]], columns=["a", "b"])

    @func
    @arg("x", pd.DataFrame, index=True, header=False)
    def read_df_0header_1index(x):
        return frame_equal(x, pd.DataFrame([[1, 2], [3, 4]], index=[10, 20]))

    @func
    @ret(pd.DataFrame, index=True, header=False)
    def write_df_0header_1index():
        return pd.DataFrame([[1, 2], [3, 4]], index=[10, 20])

    @func
    @arg("x", pd.DataFrame, index=2, header=False)
    def read_df_0header_2index(x):
        df = pd.DataFrame(
            [[1, 2, 3], [4, 5, 6], [7, 8, 9]],
            index=pd.MultiIndex.from_arrays([["a", "a", "b"], [1, 2, 1]]),
        )
        return frame_equal(x, df)

    @func
    @ret(pd.DataFrame, index=2, header=False)
    def write_df_0header_2index():
        df = pd.DataFrame(
            [[1, 2, 3], [4, 5, 6], [7, 8, 9]],
            index=pd.MultiIndex.from_arrays([["a", "a", "b"], [1, 2, 1]]),
        )
        return df

    @func
    @arg("x", pd.DataFrame, index=1, header=1)
    def read_df_1header_1namedindex(x):
        df = pd.DataFrame(
            [[1, 2, 3], [4, 5, 6]],
            index=[1, 2],
            columns=["c", "d", "c"],
        )
        df.index.name = "ix1"
        return frame_equal(x, df)

    @func
    def write_df_1header_1namedindex():
        df = pd.DataFrame(
            [[1, 2, 3], [4, 5, 6]],
            index=[1, 2],
            columns=["c", "d", "c"],
        )
        df.index.name = "ix1"
        return df

    @func
    @arg("x", pd.DataFrame, index=1, header=1)
    def read_df_1header_1unnamedindex(x):
        df = pd.DataFrame(
            [[1, 2, 3], [4, 5, 6]],
            index=[1, 2],
            columns=["c", "d", "c"],
        )
        return frame_equal(x, df)

    @func
    def write_df_1header_1unnamedindex():
        df = pd.DataFrame(
            [[1, 2, 3], [4, 5, 6]],
            index=[1, 2],
            columns=["c", "d", "c"],
        )
        return df

    @func
    @arg("x", pd.DataFrame, index=False, header=2)
    def read_df_2header_0index(x):
        df = pd.DataFrame(
            [[1, 2, 3], [4, 5, 6]],
            columns=pd.MultiIndex.from_arrays([["a", "a", "b"], ["c", "d", "c"]]),
        )
        return frame_equal(x, df)

    @func
    @ret(pd.DataFrame, index=False, header=2)
    def write_df_2header_0index():
        df = pd.DataFrame(
            [[1, 2, 3], [4, 5, 6]],
            columns=pd.MultiIndex.from_arrays([["a", "a", "b"], ["c", "d", "c"]]),
        )
        return df

    @func
    @arg("x", pd.DataFrame, index=1, header=2)
    def read_df_2header_1namedindex(x):
        df = pd.DataFrame(
            [[1, 2, 3], [4, 5, 6]],
            index=[1, 2],
            columns=pd.MultiIndex.from_arrays([["a", "a", "b"], ["c", "d", "c"]]),
        )
        df.index.name = "ix1"
        return frame_equal(x, df)

    @func
    def write_df_2header_1namedindex():
        df = pd.DataFrame(
            [[1, 2, 3], [4, 5, 6]],
            index=[1, 2],
            columns=pd.MultiIndex.from_arrays([["a", "a", "b"], ["c", "d", "c"]]),
        )
        df.index.name = "ix1"
        return df

    @func
    @arg("x", pd.DataFrame, index=1, header=2)
    def read_df_2header_1unnamedindex(x):
        df = pd.DataFrame(
            [[1, 2, 3], [4, 5, 6]],
            index=[1, 2],
            columns=pd.MultiIndex.from_arrays([["a", "a", "b"], ["c", "d", "c"]]),
        )
        return frame_equal(x, df)

    @func
    def write_df_2header_1unnamedindex():
        df = pd.DataFrame(
            [[1, 2, 3], [4, 5, 6]],
            index=[1, 2],
            columns=pd.MultiIndex.from_arrays([["a", "a", "b"], ["c", "d", "c"]]),
        )
        return df

    @func
    @arg("x", pd.DataFrame, index=2, header=2)
    def read_df_2header_2namedindex(x):
        df = pd.DataFrame(
            [[1, 2, 3], [4, 5, 6], [7, 8, 9]],
            index=pd.MultiIndex.from_arrays(
                [["a", "a", "b"], [1, 2, 1]], names=["x1", "x2"]
            ),
            columns=pd.MultiIndex.from_arrays([["a", "a", "b"], ["c", "d", "c"]]),
        )
        return frame_equal(x, df)

    @func
    def write_df_2header_2namedindex():
        df = pd.DataFrame(
            [[1, 2, 3], [4, 5, 6], [7, 8, 9]],
            index=pd.MultiIndex.from_arrays(
                [["a", "a", "b"], [1, 2, 1]], names=["x1", "x2"]
            ),
            columns=pd.MultiIndex.from_arrays([["a", "a", "b"], ["c", "d", "c"]]),
        )
        return df

    @func
    @arg("x", pd.DataFrame, index=2, header=2)
    def read_df_2header_2unnamedindex(x):
        df = pd.DataFrame(
            [[1, 2, 3], [4, 5, 6], [7, 8, 9]],
            index=pd.MultiIndex.from_arrays([["a", "a", "b"], [1, 2, 1]]),
            columns=pd.MultiIndex.from_arrays([["a", "a", "b"], ["c", "d", "c"]]),
        )
        return frame_equal(x, df)

    @func
    def write_df_2header_2unnamedindex():
        df = pd.DataFrame(
            [[1, 2, 3], [4, 5, 6], [7, 8, 9]],
            index=pd.MultiIndex.from_arrays([["a", "a", "b"], [1, 2, 1]]),
            columns=pd.MultiIndex.from_arrays([["a", "a", "b"], ["c", "d", "c"]]),
        )
        return df

    @func
    @arg("x", pd.DataFrame, index=2, header=1)
    def read_df_1header_2namedindex(x):
        df = pd.DataFrame(
            [[1, 2, 3], [4, 5, 6], [7, 8, 9]],
            index=pd.MultiIndex.from_arrays(
                [["a", "a", "b"], [1, 2, 1]], names=["x1", "x2"]
            ),
            columns=["a", "d", "c"],
        )
        return frame_equal(x, df)

    @func
    def write_df_1header_2namedindex():
        df = pd.DataFrame(
            [[1, 2, 3], [4, 5, 6], [7, 8, 9]],
            index=pd.MultiIndex.from_arrays(
                [["a", "a", "b"], [1, 2, 1]], names=["x1", "x2"]
            ),
            columns=["a", "d", "c"],
        )
        return df

    @func
    @arg("x", pd.DataFrame, parse_dates=True)
    def read_df_date_index(x):
        df = pd.DataFrame(
            [[1, 2, 3], [4, 5, 6]],
            index=[datetime(1999, 12, 13), datetime(1999, 12, 14)],
            columns=["c", "d", "c"],
        )
        return frame_equal(x, df)

    @func
    def write_df_date_index():
        df = pd.DataFrame(
            [[1, 2, 3], [4, 5, 6]],
            index=[datetime(1999, 12, 13), datetime(1999, 12, 14)],
            columns=["c", "d", "c"],
        )
        return df

    @func
    def read_workbook_caller():
        wb = xw.Book.caller()
        return wb.sheets.active["E277"].value == 1


@func
def default_args(x, y="hello", z=20):
    return 2 * x + 3 * len(y) + 7 * z


@func
def variable_args(x, *z):
    return 2 * x + 3 * len(z) + 7 * z[0]


@func
def optional_args(x, y=None):
    if y is None:
        y = 10
    return x * y


@func
def write_none():
    return None


@func
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
    paramet_name_26=None,
):
    return "non splitted signature"


@func
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
    very_long_parameter_name_26=None,
):
    return "splitted signature"


@func
def return_pd_nat():
    return pd.DataFrame(data=[pd.NaT], columns=[1], index=[1])


@func
@arg("df", pd.DataFrame, parse_dates=[0, 2])
def parse_dates_index(df):
    expected = pd.DataFrame(
        [
            [1, dt.datetime(2021, 1, 1, 11, 11, 11), 4],
            [2, dt.datetime(2021, 1, 2, 22, 22, 22), 5],
            [3, dt.datetime(2021, 1, 3), 6],
        ],
        columns=["one", "two", "three"],
        index=[
            dt.datetime(2021, 1, 1, 11, 11, 11),
            dt.datetime(2021, 1, 2, 22, 22, 22),
            dt.datetime(2021, 1, 3),
        ],
    )
    assert_frame_equal(df, expected)
    return True


@func
@arg("df", pd.DataFrame, parse_dates=["ix", "two"])
def parse_dates_names(df):
    expected = pd.DataFrame(
        [
            [1, dt.datetime(2021, 1, 1, 11, 11, 11), 4],
            [2, dt.datetime(2021, 1, 2, 22, 22, 22), 5],
            [3, dt.datetime(2021, 1, 3), 6],
        ],
        columns=["one", "two", "three"],
        index=[
            dt.datetime(2021, 1, 1, 11, 11, 11),
            dt.datetime(2021, 1, 2, 22, 22, 22),
            dt.datetime(2021, 1, 3),
        ],
    )
    expected.index.name = "ix"
    assert_frame_equal(df, expected)
    return True


@func
@arg("df", pd.DataFrame, parse_dates=True)
def parse_dates_true(df):
    expected = pd.DataFrame(
        [[1], [2], [3]],
        columns=["one"],
        index=[
            dt.datetime(2021, 1, 1, 11, 11, 11),
            dt.datetime(2021, 1, 2, 22, 22, 22),
            dt.datetime(2021, 1, 3),
        ],
    )
    assert_frame_equal(df, expected)
    return True


@func
@ret(transpose=True)
def write_error_cells():
    return ["#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!"]


@func
def read_error_cells(errors):
    assert [None] * 7 == errors
    return True


@func
@arg("errors", err_to_str=True)
def read_error_cells_str(errors):
    assert [
        "#DIV/0!",
        "#N/A",
        "#NAME?",
        "#NULL!",
        "#NUM!",
        "#REF!",
        "#VALUE!",
    ] == errors
    return True


@func
@ret(date_format="yyyy-m-d")
def explicit_date_format():
    return dt.datetime(2022, 1, 13)


@func(namespace="subname")
def namespace():
    return True


@func(volatile=True)
def volatile():
    return True


@func
@arg("x", pd.DataFrame, index=False)
@arg("*params", pd.DataFrame, index=False)
def varargs_arg_decorator(x, *params):
    return pd.concat(params + (x,))
