"""
Tests for the Range setters and action methods added to the remote (Office.js)
backend.

These only assert the JSON actions that the backend queues up: the actual Excel
side effects happen in the Office.js client, which isn't available here. Reading
these properties back isn't supported (the values aren't part of the payload
sent to Python), so the getters still raise NotImplementedError.
"""

import pytest

import xlwings as xw
from xlwings import XlwingsError
from xlwings.pro import _xlremote as R


def _book_json(n_sheets=1):
    return {
        "client": "Office.js",
        "version": xw.__version__,
        "book": {"name": "B", "active_sheet_index": 0, "selection": "A1"},
        "names": [],
        "sheets": [
            {
                "name": f"S{i + 1}",
                "values": [[None] * 5] * 5,
                "pictures": [],
                "tables": [],
            }
            for i in range(n_sheets)
        ],
    }


@pytest.fixture
def book():
    """A Book on the remote engine, wrapped in the public xlwings API."""
    impl = R.App(R.Apps(), add_book=False).books.open(_book_json(n_sheets=2))
    return xw.Book(impl=impl)


def actions(book):
    return book.impl.json()["actions"]


def last_action(book):
    return actions(book)[-1]


# --- setters ---


def test_formula_single_cell(book):
    book.sheets[0]["A1"].formula = "=1+1"
    action = last_action(book)
    assert action["func"] == "setFormula"
    assert action["values"] == [["=1+1"]]
    assert (action["row_count"], action["column_count"]) == (1, 1)


def test_formula_broadcasts_scalar_to_range(book):
    book.sheets[0]["A1:B2"].formula = "=A1"
    action = last_action(book)
    assert action["values"] == [["=A1", "=A1"], ["=A1", "=A1"]]


def test_formula_accepts_nested_list(book):
    book.sheets[0]["A1:B1"].formula = [["=1", "=2"]]
    assert last_action(book)["values"] == [["=1", "=2"]]


def test_formula_normalizes_flat_list_for_row(book):
    book.sheets[0]["A1:B1"].formula = ["=1", "=2"]
    assert last_action(book)["values"] == [["=1", "=2"]]


def test_formula_normalizes_flat_list_for_column(book):
    book.sheets[0]["A1:A2"].formula = ["=1", "=2"]
    assert last_action(book)["values"] == [["=1"], ["=2"]]


def test_formula_expands_single_cell_to_fit_flat_list(book):
    # Like `.value`, the data wins over the target's shape.
    book.sheets[0]["A1"].formula = ["=1", "=2"]
    action = last_action(book)
    assert action["values"] == [["=1", "=2"]]
    assert (action["row_count"], action["column_count"]) == (1, 2)


def test_formula_expands_single_cell_to_fit_nested_list(book):
    book.sheets[0]["A1"].formula = [["=1", "=2"], ["=3", "=4"]]
    action = last_action(book)
    assert action["values"] == [["=1", "=2"], ["=3", "=4"]]
    assert (action["row_count"], action["column_count"]) == (2, 2)


def test_formula_flat_list_writes_a_row_on_a_multi_row_range(book):
    book.sheets[0]["A1:B2"].formula = ["=1", "=2"]
    action = last_action(book)
    assert action["values"] == [["=1", "=2"]]
    assert (action["row_count"], action["column_count"]) == (1, 2)


def test_formula_resizes_when_nested_list_is_smaller_than_range(book):
    book.sheets[0]["A1:B2"].formula = [["=1", "=2"]]
    action = last_action(book)
    assert action["values"] == [["=1", "=2"]]
    assert (action["row_count"], action["column_count"]) == (1, 2)


def test_formula_expansion_keeps_the_ranges_origin(book):
    book.sheets[0]["B2"].formula = ["=1", "=2", "=3"]
    action = last_action(book)
    assert (action["start_row"], action["start_column"]) == (1, 1)
    assert (action["row_count"], action["column_count"]) == (1, 3)


def test_formula_ignores_empty_list(book):
    before = len(actions(book))
    book.sheets[0]["A1"].formula = []
    assert len(actions(book)) == before


def test_formula2_delegates_to_formula(book):
    book.sheets[0]["A1"].formula2 = "=SEQUENCE(3)"
    action = last_action(book)
    assert action["func"] == "setFormula"
    assert action["values"] == [["=SEQUENCE(3)"]]


def test_formula_array_targets_the_full_range(book):
    book.sheets[0]["B2:B4"].formula_array = "=TRANSPOSE(A1:C1)"
    action = last_action(book)
    assert action["func"] == "setFormulaArray"
    assert action["args"] == ["=TRANSPOSE(A1:C1)"]
    assert (action["start_row"], action["start_column"]) == (1, 1)
    assert (action["row_count"], action["column_count"]) == (3, 1)


def test_column_width(book):
    book.sheets[0]["A1:C1"].column_width = 12
    action = last_action(book)
    assert action["func"] == "setColumnWidth"
    assert action["args"] == [12]
    assert action["column_count"] == 3


@pytest.mark.parametrize("value", [-1, 256, "12", True])
def test_column_width_rejects_invalid_values(book, value):
    with pytest.raises(ValueError, match="between 0 and 255"):
        book.sheets[0]["A1"].column_width = value


def test_row_height(book):
    book.sheets[0]["A1"].row_height = 30
    action = last_action(book)
    assert action["func"] == "setRowHeight"
    assert action["args"] == [30]


@pytest.mark.parametrize("value", [True, False])
def test_wrap_text_sends_real_booleans(book, value):
    # Must stay a JSON boolean: the Office.js side does Boolean(args[0]), and
    # the string "false" would be truthy.
    book.sheets[0]["A1"].wrap_text = value
    assert last_action(book)["args"] == [value]


# --- action methods ---


def test_merge(book):
    book.sheets[0]["A1:B2"].merge()
    action = last_action(book)
    assert action["func"] == "rangeMerge"
    assert action["args"] == [False]
    assert (action["row_count"], action["column_count"]) == (2, 2)


def test_merge_across(book):
    book.sheets[0]["A1:C1"].merge(across=True)
    assert last_action(book)["args"] == [True]


def test_merge_restores_display_alerts(book):
    # Range.merge() runs inside app.properties(display_alerts=False)
    book.sheets[0]["A1:B2"].merge()
    assert book.app.display_alerts is True


def test_unmerge(book):
    book.sheets[0]["A1:B2"].unmerge()
    assert last_action(book)["func"] == "rangeUnmerge"


def test_autofill(book):
    sheet = book.sheets[0]
    sheet["A1:A2"].autofill(sheet["A1:A10"], "fill_series")
    action = last_action(book)
    assert action["func"] == "rangeAutofill"
    assert action["args"] == ["$A$1:$A$10", "FillSeries"]
    # the action targets the source range
    assert action["row_count"] == 2


def test_autofill_defaults_to_fill_default(book):
    sheet = book.sheets[0]
    sheet["B1"].autofill(sheet["B1:B5"])
    assert last_action(book)["args"] == ["$B$1:$B$5", "FillDefault"]


def test_autofill_rejects_unknown_type(book):
    sheet = book.sheets[0]
    with pytest.raises(XlwingsError, match="Invalid autofill type"):
        sheet["A1"].autofill(sheet["A1:A5"], "nonsense")


def test_autofill_rejects_destination_on_other_sheet(book):
    with pytest.raises(XlwingsError, match="same sheet"):
        book.sheets[0]["A1"].autofill(book.sheets[1]["A1:A5"], "fill_series")


def test_autofill_rejects_same_sheet_index_in_another_book(book):
    other_impl = R.App(R.Apps(), add_book=False).books.open(_book_json(n_sheets=2))
    other_book = xw.Book(impl=other_impl)
    with pytest.raises(XlwingsError, match="same sheet"):
        book.sheets[0]["A1"].autofill(other_book.sheets[0]["A1:A5"], "fill_series")


# --- getters remain unimplemented ---


@pytest.mark.parametrize(
    "attribute",
    [
        "formula",
        "formula2",
        "formula_array",
        "column_width",
        "row_height",
        "wrap_text",
    ],
)
def test_getters_still_raise(book, attribute):
    with pytest.raises(NotImplementedError):
        getattr(book.sheets[0]["A1"], attribute)
