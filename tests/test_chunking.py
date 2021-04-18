from pathlib import Path

import pytest
import pandas as pd
from pandas.testing import assert_frame_equal
from numpy.testing import assert_array_equal
import numpy as np
import xlwings as xw

this_dir = Path(__file__).resolve().parent

nrows, ncols = 5, 4
np_data = np.arange(nrows * ncols).reshape(nrows, ncols).astype(float)
list_data = np_data.tolist()
df = pd.DataFrame(data=np_data, columns=[f'c{i}' for i in range(ncols)],
                  index=[f'c{i}' for i in range(nrows)])


@pytest.fixture(scope="module")
def app():
    app = xw.App(visible=False)
    yield app

    for book in reversed(app.books):
        book.close()
    app.kill()


def test_read_write_df(app):
    sheet = app.books[0].sheets[0]

    sheet['A1'].value = df
    assert_frame_equal(df, sheet['A1'].expand().options(pd.DataFrame).value)
    sheet['A1'].clear()

    sheet['A1'].options(chunksize=2).value = df
    assert_frame_equal(df, sheet['A1'].expand().options(pd.DataFrame, chunksize=2).value)
    sheet['A1'].clear()

    sheet['A1'].options(chunksize=nrows - 1).value = df
    assert_frame_equal(df, sheet['A1'].expand().options(pd.DataFrame, chunksize=nrows - 1).value)
    sheet['A1'].clear()

    sheet['A1'].options(chunksize=nrows + 1).value = df
    assert_frame_equal(df, sheet['A1'].expand().options(pd.DataFrame, chunksize=nrows + 2).value)
    sheet['A1'].clear()


def test_read_write_np(app):
    sheet = app.books[0].sheets[0]

    sheet['A1'].value = np_data
    assert_array_equal(np_data, sheet['A1'].expand().options(np.array).value)
    sheet['A1'].clear()

    sheet['A1'].options(chunksize=2).value = np_data
    assert_array_equal(np_data, sheet['A1'].expand().options(np.array, chunksize=2).value)
    sheet['A1'].clear()

    sheet['A1'].options(chunksize=nrows - 1).value = np_data
    assert_array_equal(np_data, sheet['A1'].expand().options(np.array, chunksize=nrows - 1).value)
    sheet['A1'].clear()

    sheet['A1'].options(chunksize=nrows + 1).value = np_data
    assert_array_equal(np_data, sheet['A1'].expand().options(np.array, chunksize=nrows + 2).value)
    sheet['A1'].clear()


def test_read_write_list(app):
    sheet = app.books[0].sheets[0]

    sheet['A1'].value = list_data
    for i in range(len(list_data)):
        assert list_data[i] == sheet['A1'].expand().value[i]
    sheet['A1'].clear()

    sheet['A1'].options(chunksize=2).value = list_data
    for i in range(len(list_data)):
        assert list_data[i] == sheet['A1'].expand().options(chunksize=2).value[i]
    sheet['A1'].clear()

    sheet['A1'].options(chunksize=nrows - 1).value = list_data
    for i in range(len(list_data)):
        assert list_data[i] == sheet['A1'].expand().options(chunksize=nrows - 1).value[i]
    sheet['A1'].clear()

    sheet['A1'].options(chunksize=nrows + 1).value = list_data
    for i in range(len(list_data)):
        assert list_data[i] == sheet['A1'].expand().options(chunksize=nrows + 2).value[i]
    sheet['A1'].clear()
