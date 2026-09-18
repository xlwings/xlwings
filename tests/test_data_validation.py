"""Round-trip tests for data validation on desktop Excel."""

import datetime as dt
import sys
from pathlib import Path

import pytest

import xlwings as xw

this_dir = Path(__file__).resolve().parent


@pytest.fixture(scope="module")
def app():
    with xw.App(visible=False) as app:
        app.books.open(this_dir / "test book.xlsx")
        yield app


@pytest.fixture
def validation_range(app):
    sheet = app.books[0].sheets[0]
    rng = sheet["B2:B4"]
    rng.clear()
    yield rng
    rng.clear()


def _property(validation, windows_name, mac_name):
    if sys.platform.startswith("win"):
        return getattr(validation.api, windows_name)
    return getattr(validation.api, mac_name).get()


def _set_property(validation, windows_name, mac_name, value):
    if sys.platform.startswith("win"):
        setattr(validation.api, windows_name, value)
    else:
        getattr(validation.api, mac_name).set(value)


def test_literal_list(validation_range):
    validation_range.data_validation.set_list(
        ["Open", "Closed", "Pending"], in_cell_dropdown=False
    )
    formula = _property(validation_range.data_validation, "Formula1", "formula1")
    assert all(value in formula for value in ("Open", "Closed", "Pending"))
    assert (
        _property(
            validation_range.data_validation,
            "InCellDropdown",
            "in_cell_dropdown",
        )
        is False
    )
    assert validation_range.data_validation.type == "list"
    assert validation_range.data_validation.source is not None


def test_range_source(validation_range):
    source_sheet = validation_range.sheet.book.sheets.add()
    try:
        source = source_sheet["D1:D3"]
        source.value = [["Open"], ["Closed"], ["Pending"]]
        validation_range.data_validation.set_list(source)
        formula = _property(validation_range.data_validation, "Formula1", "formula1")
        assert source_sheet.name in formula
        assert "$D$1:$D$3" in formula
    finally:
        source_sheet.delete()


def test_named_range_source(validation_range):
    book = validation_range.sheet.book
    source = validation_range.sheet["D1:D3"]
    source.value = [["Open"], ["Closed"], ["Pending"]]
    name = book.names.add(
        "XlwingsValidationStatuses", f"={source.sheet.name}!$D$1:$D$3"
    )
    try:
        validation_range.data_validation.set_list(name)
        formula = _property(validation_range.data_validation, "Formula1", "formula1")
        assert "XlwingsValidationStatuses" in formula
    finally:
        name.delete()


def test_update_preserves_prompt_and_error_alert(validation_range):
    validation = validation_range.data_validation
    validation.set_list(["Open", "Closed"])
    _set_property(validation, "InputTitle", "input_title", "Choose")
    _set_property(validation, "InputMessage", "input_message", "Pick a status")
    _set_property(validation, "ErrorTitle", "error_title", "Invalid")
    _set_property(validation, "ErrorMessage", "error_message", "Use the list")
    _set_property(validation, "ShowInput", "show_input", True)
    _set_property(validation, "ShowError", "show_error", True)

    validation.set_list(["Open", "Closed", "Pending"])

    assert _property(validation, "InputTitle", "input_title") == "Choose"
    assert _property(validation, "InputMessage", "input_message") == "Pick a status"
    assert _property(validation, "ErrorTitle", "error_title") == "Invalid"
    assert _property(validation, "ErrorMessage", "error_message") == "Use the list"
    assert _property(validation, "ShowInput", "show_input") is True
    assert _property(validation, "ShowError", "show_error") is True
    assert validation.input_title == "Choose"
    assert validation.input_message == "Pick a status"
    assert validation.error_title == "Invalid"
    assert validation.error_message == "Use the list"
    assert validation.show_input is True
    assert validation.show_error is True
    assert validation.alert_style == "stop"


def test_mixed_rules_are_not_overwritten(validation_range):
    validation_range[0, 0].data_validation.set_list(["Open"])
    validation_range[1, 0].data_validation.set_list(["Closed"])
    assert validation_range.data_validation.type in {
        "inconsistent",
        "mixed_criteria",
    }
    with pytest.raises(xw.XlwingsError, match="different validation rules"):
        validation_range.data_validation.set_list(["Pending"])


def test_delete_then_recreate(validation_range):
    validation = validation_range.data_validation
    validation.set_list(["Open", "Closed"])
    validation.delete()
    validation.set_list(["Pending"])
    assert "Pending" in _property(validation, "Formula1", "formula1")


@pytest.mark.parametrize(
    "method,args,rule_type,operator",
    [
        ("set_whole_number", ("between", 1, 10), "whole_number", "between"),
        ("set_decimal", ("greater_than", 0.5), "decimal", "greater_than"),
        (
            "set_date",
            ("between", dt.date(2026, 1, 1), dt.date(2026, 12, 31)),
            "date",
            "between",
        ),
        ("set_time", ("less_than", dt.time(17)), "time", "less_than"),
        (
            "set_text_length",
            ("less_than_or_equal", 40),
            "text_length",
            "less_than_or_equal",
        ),
    ],
)
def test_comparison_rules(validation_range, method, args, rule_type, operator):
    validation = validation_range.data_validation
    getattr(validation, method)(*args)
    assert validation.type == rule_type
    assert validation.operator == operator
    assert validation.formula1 is not None
    if operator == "between":
        assert validation.formula2 is not None
    else:
        assert validation.formula2 is None


def test_custom_rule(validation_range):
    validation = validation_range.data_validation
    validation.set_custom("=COUNTIF($B$2:$B$4,B2)=1")
    assert validation.type == "custom"
    assert "COUNTIF" in validation.formula
    assert validation.operator is None


def test_no_validation_snapshot(validation_range):
    validation_range.data_validation.delete()
    validation = validation_range.data_validation
    assert validation.type == "none"
    assert validation.formula1 is None
    assert validation.ignore_blank is None
