"""
Tests for Sheet.used_range on the remote (Office.js) backend.

The Office.js client sends cell values as "A1:<last cell of the used range>",
so the used range is derived from the shape of that payload and always starts
at A1 --- unlike the COM API, where it starts at the used range's real
top-left corner.
"""

import pytest

import xlwings as xw
from xlwings import XlwingsError
from xlwings.pro import _xlremote as R


def _book_json(values):
    return {
        "client": "Office.js",
        "version": xw.__version__,
        "book": {"name": "B", "active_sheet_index": 0, "selection": "A1"},
        "names": [],
        "sheets": [
            {
                "name": "S1",
                "values": values,
                "pictures": [],
                "tables": [],
            }
        ],
    }


def _book(values, lazy=False):
    apps = R.Apps()
    impl = R.App(apps, add_book=False).books.open(_book_json(values), lazy=lazy)
    return xw.Book(impl=impl)


def test_used_range_address():
    book = _book([[1, 2, 3], [4, 5, 6]])
    assert book.sheets[0].used_range.address == "$A$1:$C$2"


def test_used_range_shape():
    book = _book([[1, 2, 3], [4, 5, 6]])
    assert book.sheets[0].used_range.shape == (2, 3)


def test_used_range_single_cell():
    book = _book([[1]])
    used = book.sheets[0].used_range
    assert used.address == "$A$1"
    assert used.shape == (1, 1)


def test_used_range_values():
    book = _book([[1, 2], [3, 4]])
    assert book.sheets[0].used_range.value == [[1, 2], [3, 4]]


@pytest.mark.parametrize("values", [[], [[]]])
def test_used_range_empty_sheet(values):
    # Excel reports A1 as the used range of an empty sheet
    used = _book(values).sheets[0].used_range
    assert used.address == "$A$1"
    assert used.shape == (1, 1)


def test_used_range_ragged_rows_uses_widest():
    book = _book([[1], [2, 3, 4]])
    assert book.sheets[0].used_range.address == "$A$1:$C$2"


def test_used_range_is_a_range_on_the_right_sheet():
    book = _book([[1, 2]])
    used = book.sheets[0].used_range
    assert isinstance(used, xw.Range)
    assert used.sheet.name == "S1"


def test_used_range_raises_on_unloaded_lazy_book():
    book = _book([[1, 2]], lazy=True)
    with pytest.raises(XlwingsError, match="haven't been loaded"):
        book.sheets[0].used_range


def test_used_range_works_on_lazy_book_once_values_loaded():
    book = _book([[1, 2]], lazy=True)
    R._mark_sheet_values_loaded(book.sheets[0].impl.api)
    assert book.sheets[0].used_range.address == "$A$1:$B$1"
