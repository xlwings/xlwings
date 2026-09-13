"""Round-trip tests for Range.horizontal_alignment and Range.vertical_alignment
against a real Excel (Windows and macOS).

Unlike the border getters, these report None for a range whose cells disagree
on both hosts: Windows COM returns None directly, and macOS returns an
undocumented enum that the keyword mapping doesn't cover. Measured on Excel
for Mac on 2026-09-13; re-measure before assuming other hosts or versions
behave the same.
"""

from pathlib import Path

import pytest

import xlwings as xw

this_dir = Path(__file__).resolve().parent

HORIZONTAL = [
    "general",
    "left",
    "center",
    "right",
    "fill",
    "justify",
    "center_across_selection",
    "distributed",
]
VERTICAL = ["top", "center", "bottom", "justify", "distributed"]


@pytest.fixture(scope="module")
def app():
    with xw.App(visible=False) as app:
        app.books.open(this_dir / "test book.xlsx")
        yield app


@pytest.fixture
def rng(app):
    rng = app.books[0].sheets[0]["B2:D4"]
    _reset(rng)
    yield rng
    _reset(rng)


def _reset(rng):
    rng.horizontal_alignment = "general"
    rng.vertical_alignment = "bottom"


def _invalid(value):
    """Feed a value the type checker would reject to the runtime validation."""
    return value


def test_defaults(app):
    # The rng fixture sets alignment explicitly, so use a fresh workbook to
    # measure Excel's defaults without first invoking either setter.
    book = app.books.add()
    try:
        cell = book.sheets[0]["A1"]
        assert cell.horizontal_alignment == "general"
        assert cell.vertical_alignment == "bottom"
    finally:
        book.close()


@pytest.mark.parametrize("value", HORIZONTAL)
def test_horizontal_round_trip(rng, value):
    cell = rng[0, 0]
    cell.horizontal_alignment = value
    assert cell.horizontal_alignment == value


@pytest.mark.parametrize("value", VERTICAL)
def test_vertical_round_trip(rng, value):
    cell = rng[0, 0]
    cell.vertical_alignment = value
    assert cell.vertical_alignment == value


def test_axes_are_independent(rng):
    cell = rng[0, 0]
    cell.horizontal_alignment = "center"
    cell.vertical_alignment = "top"
    assert cell.horizontal_alignment == "center"
    assert cell.vertical_alignment == "top"


def test_set_applies_to_every_cell(rng):
    rng.horizontal_alignment = "center"
    rng.vertical_alignment = "top"
    for cell in rng:
        assert cell.horizontal_alignment == "center"
        assert cell.vertical_alignment == "top"
    assert rng.horizontal_alignment == "center"
    assert rng.vertical_alignment == "top"


def test_mixed_horizontal_is_none(rng):
    rng[0, 0].horizontal_alignment = "left"
    rng[1, 0].horizontal_alignment = "right"
    assert rng[0:2, 0].horizontal_alignment is None
    # the individual cells still report their own value
    assert rng[0, 0].horizontal_alignment == "left"
    assert rng[1, 0].horizontal_alignment == "right"


def test_mixed_vertical_is_none(rng):
    rng[0, 0].vertical_alignment = "top"
    rng[1, 0].vertical_alignment = "bottom"
    assert rng[0:2, 0].vertical_alignment is None
    assert rng[0, 0].vertical_alignment == "top"
    assert rng[1, 0].vertical_alignment == "bottom"


@pytest.mark.parametrize("value", ["centre", "middle", "", "Center"])
def test_invalid_horizontal_names_the_options(rng, value):
    with pytest.raises(ValueError, match="Invalid horizontal_alignment"):
        rng.horizontal_alignment = _invalid(value)


@pytest.mark.parametrize("value", ["middle", "centre", "", "Top"])
def test_invalid_vertical_names_the_options(rng, value):
    with pytest.raises(ValueError, match="Invalid vertical_alignment"):
        rng.vertical_alignment = _invalid(value)


def test_none_is_rejected(rng):
    # There's no "unset" alignment: 'general'/'bottom' are the neutral values
    with pytest.raises(ValueError, match="Invalid horizontal_alignment"):
        rng.horizontal_alignment = _invalid(None)
    with pytest.raises(ValueError, match="Invalid vertical_alignment"):
        rng.vertical_alignment = _invalid(None)


def test_axis_specific_values_are_rejected(rng):
    # 'fill' is horizontal-only, 'top' vertical-only
    with pytest.raises(ValueError, match="Invalid vertical_alignment"):
        rng.vertical_alignment = _invalid("fill")
    with pytest.raises(ValueError, match="Invalid horizontal_alignment"):
        rng.horizontal_alignment = _invalid("top")


def test_invalid_value_leaves_alignment_unchanged(rng):
    rng.horizontal_alignment = "center"
    with pytest.raises(ValueError):
        rng.horizontal_alignment = _invalid("centre")
    assert rng.horizontal_alignment == "center"
