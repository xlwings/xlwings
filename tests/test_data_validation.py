"""Round-trip tests for list data validation on desktop Excel."""

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


def test_mixed_rules_are_not_overwritten(validation_range):
    validation_range[0, 0].data_validation.set_list(["Open"])
    validation_range[1, 0].data_validation.set_list(["Closed"])
    with pytest.raises(xw.XlwingsError, match="different validation rules"):
        validation_range.data_validation.set_list(["Pending"])


def test_delete_then_recreate(validation_range):
    validation = validation_range.data_validation
    validation.set_list(["Open", "Closed"])
    validation.delete()
    validation.set_list(["Pending"])
    assert "Pending" in _property(validation, "Formula1", "formula1")
