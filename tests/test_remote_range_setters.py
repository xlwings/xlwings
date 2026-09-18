"""
Tests for the Range setters and action methods added to the remote (Office.js)
backend.

These only assert the JSON actions that the backend queues up: the actual Excel
side effects happen in the Office.js client, which isn't available here. Reading
these properties back isn't supported (the values aren't part of the payload
sent to Python), so the getters still raise NotImplementedError.
"""

import datetime as dt
import sys
from types import ModuleType, SimpleNamespace

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


def test_data_validation_literal_list(book):
    book.sheets[0]["A1:A3"].data_validation.set_list(
        ["Open", 2, 3.5, True], in_cell_dropdown=False
    )
    action = last_action(book)
    assert action["func"] == "setDataValidationList"
    assert action["args"] == [
        {"type": "literal", "values": ["Open", "2", "3.5", "TRUE"]},
        False,
    ]
    assert (action["start_row"], action["start_column"]) == (0, 0)
    assert (action["row_count"], action["column_count"]) == (3, 1)


def test_data_validation_range_source(book):
    book.sheets[0]["A1:A3"].data_validation.set_list(book.sheets[1]["C2:C4"])
    assert last_action(book)["args"] == [
        {
            "type": "range",
            "sheet_position": 1,
            "start_row": 1,
            "start_column": 2,
            "row_count": 3,
            "column_count": 1,
        },
        True,
    ]


def test_data_validation_named_range_source(book):
    name = book.names.add("Statuses", "=S2!$C$2:$C$4")
    book.sheets[0]["A1:A3"].data_validation.set_list(name)
    assert last_action(book)["args"] == [
        {"type": "name", "name": "Statuses"},
        True,
    ]


def test_data_validation_delete(book):
    book.sheets[0]["A1:A3"].data_validation.delete()
    assert last_action(book)["func"] == "deleteDataValidation"


@pytest.mark.parametrize(
    "method,args,expected",
    [
        (
            "set_whole_number",
            ("between", 1, 10),
            {
                "type": "whole_number",
                "operator": "between",
                "formula1": "1",
                "formula2": "10",
            },
        ),
        (
            "set_decimal",
            ("greater_than", 0.5),
            {
                "type": "decimal",
                "operator": "greater_than",
                "formula1": "0.5",
                "formula2": None,
            },
        ),
        (
            "set_date",
            ("greater_than_or_equal", dt.date(2026, 1, 1)),
            {
                "type": "date",
                "operator": "greater_than_or_equal",
                "formula1": "=DATE(2026,1,1)",
                "formula2": None,
            },
        ),
        (
            "set_time",
            ("less_than", dt.time(6)),
            {
                "type": "time",
                "operator": "less_than",
                "formula1": "=TIME(6,0,0)",
                "formula2": None,
            },
        ),
        (
            "set_text_length",
            ("less_than_or_equal", 40),
            {
                "type": "text_length",
                "operator": "less_than_or_equal",
                "formula1": "40",
                "formula2": None,
            },
        ),
    ],
)
def test_data_validation_comparison_rule_actions(book, method, args, expected):
    getattr(book.sheets[0]["A1:A3"].data_validation, method)(*args)
    action = last_action(book)
    assert action["func"] == "setDataValidationRule"
    assert action["args"] == [expected]


def test_data_validation_custom_rule_action(book):
    book.sheets[0]["A1:A3"].data_validation.set_custom("=COUNTIF(A:A,A1)=1")
    assert last_action(book)["args"] == [
        {
            "type": "custom",
            "operator": None,
            "formula1": "=COUNTIF(A:A,A1)=1",
            "formula2": None,
        }
    ]


@pytest.mark.parametrize(
    "method,args,error",
    [
        ("set_decimal", ("unknown", 1), "Invalid data-validation operator"),
        ("set_decimal", ("between", 1), "formula2 is required"),
        ("set_decimal", ("greater_than", 1, 2), "formula2 isn't valid"),
        ("set_decimal", ("greater_than", True), "must not be a bool"),
        ("set_decimal", ("greater_than", float("inf")), "finite number"),
        (
            "set_date",
            ("greater_than", dt.datetime.now(dt.UTC)),
            "timezone-naive datetime",
        ),
        (
            "set_time",
            ("greater_than", dt.time(9, tzinfo=dt.UTC)),
            "timezone-naive time",
        ),
        ("set_custom", ("A1>0",), "starting with '='"),
        ("set_custom", ("=" + "x" * 255,), "255 characters"),
    ],
)
def test_data_validation_rule_rejects_invalid_arguments(book, method, args, error):
    with pytest.raises((TypeError, ValueError), match=error):
        getattr(book.sheets[0]["A1"].data_validation, method)(*args)
    assert actions(book) == []


@pytest.mark.anyio
async def test_data_validation_async_snapshot(book, monkeypatch):
    entry = {
        "type": "whole_number",
        "operator": "between",
        "formula1": "=1",
        "formula2": "=10",
        "formula": None,
        "source": None,
        "in_cell_dropdown": None,
        "ignore_blank": True,
        "input_title": "Quantity",
        "input_message": "Enter 1 through 10",
        "show_input": True,
        "error_title": "Invalid",
        "error_message": "Use a whole number",
        "show_error": True,
        "alert_style": "stop",
    }

    async def get_range_data(self, key, method=None):
        assert key == "data_validation"
        return entry

    monkeypatch.setattr(R.Range, "_get_range_data", get_range_data)
    validation = await book.sheets[0]["A1:A3"].get_data_validation()
    assert validation.type == "whole_number"
    assert validation.operator == "between"
    assert validation.formula1 == "=1"
    assert validation.formula2 == "=10"
    assert validation.ignore_blank is True
    assert validation.input_title == "Quantity"
    assert validation.error_message == "Use a whole number"
    assert validation.alert_style == "stop"

    validation.set_decimal("greater_than", 0)
    assert validation.type == "whole_number"
    assert last_action(book)["args"][0]["type"] == "decimal"


def test_data_validation_sync_snapshot_requires_lite(book):
    with pytest.raises(NotImplementedError, match="get_data_validation"):
        _ = book.sheets[0]["A1"].data_validation.type


@pytest.mark.parametrize(
    "source,error",
    [
        ([], "must not be empty"),
        (["a,b"], "cannot contain"),
        ([None], "strings, numbers, or booleans"),
        ([float("inf")], "strings, numbers, or booleans"),
        (["x" * 256], "255 characters"),
        ("Open,Closed", "non-empty sequence"),
    ],
)
def test_data_validation_rejects_invalid_literal_sources(book, source, error):
    with pytest.raises((TypeError, ValueError), match=error):
        book.sheets[0]["A1"].data_validation.set_list(source)
    assert actions(book) == []


def test_data_validation_rejects_two_dimensional_range(book):
    with pytest.raises(ValueError, match="one-dimensional"):
        book.sheets[0]["A1"].data_validation.set_list(book.sheets[0]["C1:D2"])
    assert actions(book) == []


def test_data_validation_rejects_range_from_another_book(book):
    other_impl = R.App(R.Apps(), add_book=False).books.open(_book_json())
    other_book = xw.Book(impl=other_impl)
    with pytest.raises(ValueError, match="same workbook"):
        book.sheets[0]["A1"].data_validation.set_list(other_book.sheets[0]["A1:A2"])
    assert actions(book) == []


def test_data_validation_rejects_non_boolean_dropdown_flag(book):
    with pytest.raises(TypeError, match="must be a bool"):
        book.sheets[0]["A1"].data_validation.set_list(
            ["Open", "Closed"], in_cell_dropdown=1
        )
    assert actions(book) == []


@pytest.mark.anyio
async def test_flush_does_not_replay_actions_after_partial_failure(book, monkeypatch):
    monkeypatch.setattr(sys, "platform", "emscripten")
    dispatched = []

    async def run_actions(payload):
        dispatched.append(payload)
        if len(dispatched) == 1:
            raise RuntimeError("partial dispatch")

    js = ModuleType("js")
    js.Object = SimpleNamespace(fromEntries=lambda entries: dict(entries))
    js.xlwings = SimpleNamespace(runActions=run_actions)
    monkeypatch.setitem(sys.modules, "js", js)

    pyodide = ModuleType("pyodide")
    ffi = ModuleType("pyodide.ffi")
    ffi.to_js = lambda value, **kwargs: value
    pyodide.ffi = ffi
    monkeypatch.setitem(sys.modules, "pyodide", pyodide)
    monkeypatch.setitem(sys.modules, "pyodide.ffi", ffi)

    target = book.sheets[0]["A1:A3"]
    target.data_validation.set_list(["Open", "Closed"])
    with pytest.raises(RuntimeError, match="partial dispatch"):
        await book.flush()
    assert actions(book) == []

    target.data_validation.delete()
    await book.flush()
    assert [action["func"] for action in dispatched[1]["actions"]] == [
        "deleteDataValidation"
    ]


@pytest.fixture
def anyio_backend():
    return "asyncio"


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
