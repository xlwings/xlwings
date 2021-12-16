"""TODO
* clear_contents
* accept dicts and json
"""

from pathlib import Path

import pytest
import numpy as np
import pandas as pd

import xlwings as xw
from xlwings import web

this_dir = Path(__file__).resolve().parent

data = {
    'book': {'name': 'web.xlsx',
             'active_sheet_index': 0},
    'sheets': [
        {
            'name': 'Sheet1',
            'values': [['a', 'b', 'c', ''], [1., 2., 3., '2021-01-01T00:00:00.000Z'], [4., 5., 6., ''], ['', '', '', '']],
        },
        {'name': 'Sheet2', 'values': [['aa', 'bb'], [11., 22.]]},
    ],
}

# data = {'book': {'name': 'ExcelPython.xlsx', 'active_sheet_index': 0},
#         'sheets': [{'name': 'Sheet1', 'values': [[1.1, '2021-01-01T00:00:00.000Z', "string", True]]},
#                     {'name': 'xlwings.conf',
#                      'values': [['AUTH_TOKEN', 'Bearer MYTOKEN'],
#                                 ['URL2', 'https://w77zc253d1.execute-api.us-east-1.amazonaws.com/api'],
#                                 ['URL', 'https://459c-217-22-135-239.ngrok.io']]}]}


@pytest.fixture(scope="module")
def book():
    book = web.Book(json=data)
    # book = xw.Book('web.xlsx')
    yield book


# range.value
def test_range_index(book):
    sheet = book.sheets[0]
    assert sheet.range((1, 1)).value == 'a'
    assert sheet.range((1, 1), (3, 1)).value == ['a', 1., 4.]
    assert sheet.range((1, 3), (3, 3)).value == ['c', 3., 6.]
    assert sheet.range((1, 1), (3, 3)).value == [['a', 'b', 'c'], [1., 2., 3.], [4., 5., 6.]]
    assert sheet.range((2, 2), (3, 3)).value == [[2., 3.], [5., 6.]]


def test_range_a1(book):
    sheet = book.sheets[0]
    assert sheet.range('A1').value == 'a'
    assert sheet.range('A1:A3').value == ['a', 1., 4.]
    assert sheet.range('C1:C3').value == ['c', 3., 6.]
    assert sheet.range('A1:C3').value == [['a', 'b', 'c'], [1., 2., 3.], [4., 5., 6.]]
    assert sheet.range('B2:C3').value == [[2., 3.], [5., 6.]]


def test_range_shortcut(book):
    sheet = book.sheets[0]
    assert sheet['A1'].value == 'a'
    assert sheet['A1:A3'].value == ['a', 1., 4.]
    assert sheet['C1:C3'].value == ['c', 3., 6.]
    assert sheet['A1:C3'].value == [['a', 'b', 'c'], [1., 2., 3.], [4., 5., 6.]]
    assert sheet['B2:C3'].value == [[2., 3.], [5., 6.]]


def test_range_from_range(book):
    sheet = book.sheets[0]
    assert sheet.range(sheet.range((1, 1)), sheet.range((3, 1))).value == ['a', 1., 4.]
    assert sheet.range(sheet.range('C1'), sheet.range('C3')).value == ['c', 3., 6.]
    assert sheet.range(sheet.range('A1'), sheet.range('C3')).value == [['a', 'b', 'c'], [1., 2., 3.], [4., 5., 6.]]
    assert sheet.range(sheet.range('B2'), sheet.range('C3')).value == [[2., 3.], [5., 6.]]


def test_range_round_indexing(book):
    sheet = book.sheets[0]
    assert sheet['B2:C3'](1, 1).value == 2.
    assert sheet['B2:C3'](1, 1).address == '$B$2'
    assert sheet['B2:C3'](2, 1).value == 5.
    assert sheet['B2:C3'](2, 1).address == '$B$3'


def test_range_square_indexing(book):
    sheet = book.sheets[0]
    assert sheet['B2:C3'][0, 0].value == 2.
    assert sheet['B2:C3'][1, 0].value == 5.


def test_range_resize(book):
    sheet1 = book.sheets[0]
    assert sheet1['A1'].resize(row_size=2, column_size=3).address == '$A$1:$C$2'
    assert sheet1['A1'].resize(row_size=4, column_size=5).address == '$A$1:$E$4'  # outside of used range


def test_range_offset(book):
    pass


def test_expand(book):
    sheet1 = book.sheets[0]
    assert sheet1['A1'].expand().address == '$A$1:$C$3'
    assert sheet1['A1'].expand().value == [['a', 'b', 'c'], [1., 2., 3.], [4., 5., 6.]]
    assert sheet1['B1'].expand().address == '$B$1:$C$3'
    assert sheet1['B1'].expand().value == [['b', 'c'], [2., 3.], [5., 6.]]
    assert sheet1['C3'].expand().address == '$C$3'
    assert sheet1['C3'].expand().value == 6.

    # Edge case (no more rows/cols after expanded range
    sheet2 = book.sheets[1]
    assert sheet2['A1'].expand().value == [['aa', 'bb'], [11., 22.]]
    assert sheet2['A1'].expand().address == '$A$1:$B$2'


# Conversion
def test_numpy_array(book):
    sheet = book.sheets[0]
    np.testing.assert_array_equal(
        sheet['B2:C3'].options(np.array).value, np.array([[2., 3.], [5., 6.]])
    )


def test_pandas_df(book):
    sheet = book.sheets[0]
    pd.testing.assert_frame_equal(
        sheet['A1:C3'].options(pd.DataFrame, index=False).value,
        pd.DataFrame(data=[[1., 2., 3.], [4., 5., 6.]], columns=['a', 'b', 'c']),
    )


# sheets
def test_sheet_access(book):
    assert book.sheets[0] == book.sheets['Sheet1']
    assert book.sheets[1] == book.sheets['Sheet2']
    assert book.sheets[0].name == 'Sheet1'
    assert book.sheets[1].name == 'Sheet2'


def test_sheet_active(book):
    assert book.sheets.active == book.sheets[0]


def test_sheets_iteration(book):
    for ix, sheet in enumerate(book.sheets):
        assert sheet.name == 'Sheet1' if ix == 0 else 'Sheet2'


# book name
def test_book(book):
    assert book.name == 'web.xlsx'
