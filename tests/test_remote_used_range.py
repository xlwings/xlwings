"""
Tests for Sheet.used_range on the remote (Office.js) backend.

The Office.js client sends the used range's address as `used_range_address`,
which is what's used when present. Clients that don't send it yet fall back to
deriving the extent from the shape of the `values` payload --- which is always
anchored at A1, so the real top-left corner is lost in that case.
"""

import pytest

import xlwings as xw
from xlwings import XlwingsError
from xlwings.pro import _xlremote as R

_MISSING = object()


def _book_json(values, used_range_address=_MISSING):
    sheet = {
        "name": "S1",
        "values": values,
        "pictures": [],
        "tables": [],
    }
    if used_range_address is not _MISSING:
        sheet["used_range_address"] = used_range_address
    return {
        "client": "Office.js",
        "version": xw.__version__,
        "book": {"name": "B", "active_sheet_index": 0, "selection": "A1"},
        "names": [],
        "sheets": [sheet],
    }


def _book(values, lazy=False, used_range_address=_MISSING):
    apps = R.Apps()
    impl = R.App(apps, add_book=False).books.open(
        _book_json(values, used_range_address), lazy=lazy
    )
    return xw.Book(impl=impl)


# --- with used_range_address in the payload (current clients) ---


def test_used_range_uses_the_reported_address():
    book = _book([[None, None], [None, 1]], used_range_address="B2")
    assert book.sheets[0].used_range.address == "$B$2"


def test_used_range_keeps_the_real_origin():
    # The values payload is anchored at A1, but the used range isn't
    values = [[None] * 4 for _ in range(10)]
    book = _book(values, used_range_address="C5:D10")
    used = book.sheets[0].used_range
    assert used.address == "$C$5:$D$10"
    assert used.shape == (6, 2)


def test_used_range_null_address_means_empty_sheet():
    used = _book([[]], used_range_address=None).sheets[0].used_range
    assert used.address == "$A$1"
    assert used.shape == (1, 1)


def test_used_range_address_wins_over_values_shape():
    book = _book([[1, 2, 3], [4, 5, 6]], used_range_address="A1:B2")
    assert book.sheets[0].used_range.address == "$A$1:$B$2"


def test_used_range_address_is_used_on_lazy_book_without_values():
    # The address doesn't need the values payload, so no need to raise
    book = _book([[]], lazy=True, used_range_address="B2:C3")
    assert book.sheets[0].used_range.address == "$B$2:$C$3"


def test_used_range_on_lazy_book_doesnt_ask_to_load_values():
    # Regression: the client must send used_range_address in lazy mode too,
    # since it's metadata, not cell values. Reading used_range on an async book
    # used to raise "haven't been loaded".
    book = _book([[]], lazy=True, used_range_address="A1:D10")
    used = book.sheets[0].used_range  # must not raise
    assert used.address == "$A$1:$D$10"


def test_used_range_null_address_on_lazy_book_means_empty_sheet():
    book = _book([[]], lazy=True, used_range_address=None)
    assert book.sheets[0].used_range.address == "$A$1"


# --- fallback: clients that don't send used_range_address ---


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
