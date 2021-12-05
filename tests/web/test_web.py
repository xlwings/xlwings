from pathlib import Path

import pytest
import numpy as np
import pandas as pd

import xlwings as xw
from xlwings.web import Book

this_dir = Path(__file__).resolve().parent

data = [
    {
        'name': 'Sheet1',
        'values': [['a', 'b', 'c', ''], [1, 2, 3, ''], [4, 5, 6, ''], ['', '', '', '']],
    },
    {'name': 'Sheet2', 'values': [['aa', 'bb'], [11, 22]]},
]


@pytest.fixture(scope="module")
def book():
    book = Book(json=data)
    yield book


# range.value
def test_range_index(book):
    sheet = book.sheets[0]
    assert sheet.range((1, 1)).value == 'a'
    assert sheet.range((1, 3), (3, 3)).value == ['c', 3, 6]
    assert sheet.range((1, 1), (3, 3)).value == [['a', 'b', 'c'], [1, 2, 3], [4, 5, 6]]
    assert sheet.range((2, 2), (3, 3)).value == [[2, 3], [5, 6]]


def test_range_address(book):
    sheet = book.sheets[0]
    assert sheet.range('A1').value == 'a'
    assert sheet.range('C1:C3').value == ['c', 3, 6]
    assert sheet.range('A1:C3').value == [['a', 'b', 'c'], [1, 2, 3], [4, 5, 6]]
    assert sheet.range('B2:C3').value == [[2, 3], [5, 6]]


def test_range_from_sheet(book):
    sheet = book.sheets[0]
    assert sheet['A1'].value == 'a'
    assert sheet['C1:C3'].value == ['c', 3, 6]
    assert sheet['A1:C3'].value == [['a', 'b', 'c'], [1, 2, 3], [4, 5, 6]]
    assert sheet['B2:C3'].value == [[2, 3], [5, 6]]


def test_numpy_array(book):
    sheet = book.sheets[0]
    np.testing.assert_array_equal(
        sheet['B2:C3'].options(np.array).value, np.array([[2, 3], [5, 6]])
    )


def test_pandas_df(book):
    sheet = book.sheets[0]
    pd.testing.assert_frame_equal(
        sheet['A1:C3'].options(pd.DataFrame, index=False).value,
        pd.DataFrame(data=[[1, 2, 3], [4, 5, 6]], columns=['a', 'b', 'c']),
    )


# sheet indexing
def test_sheet_access(book):
    assert book.sheets[0] == book.sheets['Sheet1']
    assert book.sheets[1] == book.sheets['Sheet2']
    assert book.sheets[0].name == 'Sheet1'
    assert book.sheets[1].name == 'Sheet2'

