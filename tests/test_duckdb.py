import datetime
from pathlib import Path

import pandas as pd
import pytest
from pandas.testing import assert_frame_equal

import xlwings as xw

this_dir = Path(__file__).resolve().parent


try:
    import duckdb
except ImportError:
    duckdb = None


@pytest.fixture(scope="module")
def app():
    with xw.App(visible=False) as app:
        yield app


@pytest.fixture(scope="function")
def sheet(app):
    sheet = app.books[0].sheets[0]
    sheet.clear()
    yield sheet
    sheet.clear()


@pytest.mark.skipif(duckdb is None, reason="DuckDB not installed")
def test_duckdb_write_read_roundtrip(sheet):
    """Test writing a DuckDB Relation and reading it back."""
    data = {
        "col_str": ["a", "b", "c", "d"],
        "col_int": [1, None, 3, 4],
        "col_float": [1.1, 2.2, None, 4.4],
        "col_bool": [True, False, True, False],
        "col_datetime": [
            datetime.datetime(2023, 1, 1, 10, 0, 0),
            datetime.datetime(2023, 1, 2, 11, 0, 0),
            None,
            datetime.datetime(2023, 1, 4, 13, 0, 0),
        ],
    }
    df = pd.DataFrame(data)  # noqa: F841
    sheet["A1"].value = duckdb.sql("SELECT * FROM df")
    rel_read = sheet["A1"].expand().options("duckdb").value
    df_read = rel_read.df()
    df_read["col_datetime"] = df_read["col_datetime"].astype("datetime64[ns]")

    assert_frame_equal(df, df_read)
