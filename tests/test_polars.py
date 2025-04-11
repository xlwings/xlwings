import datetime
from pathlib import Path

import pytest

import xlwings as xw

this_dir = Path(__file__).resolve().parent


try:
    import polars as pl
    from polars.testing import assert_frame_equal
except ImportError:
    pl = None


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


@pytest.mark.skipif(pl is None, reason="Polars not installed")
def test_polars_write_read_roundtrip(sheet):
    """Test writing a Polars DataFrame and reading it back."""
    data = {
        "col_str": ["a", "b", "c", "d"],
        "col_int": [1, None, 3, 4],
        "col_float": [1.1, 2.2, None, 4.4],
        "col_bool": [True, False, True, None],
        "col_datetime": [
            datetime.datetime(2023, 1, 1, 10, 0, 0),
            datetime.datetime(2023, 1, 2, 11, 0, 0),
            None,
            datetime.datetime(2023, 1, 4, 13, 0, 0),
        ],
        "col_date": [
            datetime.date(2023, 1, 1),
            datetime.date(2023, 1, 2),
            datetime.date(2023, 1, 3),
            None,
        ],
    }

    sheet["A1"].value = pl.DataFrame(data)
    df_read = sheet["A1"].expand().options(pl.DataFrame).value

    expected_df = pl.DataFrame(
        data,
        schema={
            "col_str": pl.Utf8,
            "col_int": pl.Float64,  # classic xlwings, otherwise pl.Int64
            "col_float": pl.Float64,
            "col_bool": pl.Boolean,
            "col_datetime": pl.Datetime,
            "col_date": pl.Datetime,  # xlwings delivers datetime by default
        },
    )

    assert_frame_equal(df_read, expected_df, check_dtype=True)
