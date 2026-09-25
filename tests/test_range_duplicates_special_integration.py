"""Round-trip range data operation tests against desktop Excel."""

import inspect
import sys

import pytest

import xlwings as xw


@pytest.fixture(scope="module")
def app():
    with xw.App(visible=False) as app:
        yield app


@pytest.fixture
def sheet(app):
    book = app.books.add()
    try:
        yield book.sheets[0]
    finally:
        book.close()


def cells(areas):
    assert all(isinstance(area, xw.Range) for area in areas)
    return {
        (row, column)
        for area in areas
        for row in range(area.row, area.row + area.shape[0])
        for column in range(area.column, area.column + area.shape[1])
    }


@pytest.mark.skipif(
    sys.platform != "win32", reason="RemoveDuplicates needs Excel for Windows"
)
def test_remove_duplicates_keeps_header_and_only_changes_selected_columns(sheet):
    sheet["A2:D6"].value = [
        ["left-1", "Name", "ID", "right-1"],
        ["left-2", "A", 1, "right-2"],
        ["left-3", "A", 1, "right-3"],
        ["left-4", "A", 2, "right-4"],
        ["left-5", "B", 1, "right-5"],
    ]

    result = sheet["B2:C6"].remove_duplicates([1, 2], has_headers=True)

    assert result is None
    assert sheet["B2:C6"].value == [
        ["Name", "ID"],
        ["A", 1],
        ["A", 2],
        ["B", 1],
        [None, None],
    ]
    assert sheet["A2:A6"].options(ndim=2).value == [
        ["left-1"],
        ["left-2"],
        ["left-3"],
        ["left-4"],
        ["left-5"],
    ]
    assert sheet["D2:D6"].options(ndim=2).value == [
        ["right-1"],
        ["right-2"],
        ["right-3"],
        ["right-4"],
        ["right-5"],
    ]


@pytest.mark.skipif(
    sys.platform != "win32", reason="RemoveDuplicates needs Excel for Windows"
)
def test_remove_duplicates_uses_columns_relative_to_the_range(sheet):
    sheet["B2:D5"].value = [
        ["first", "x", 10],
        ["second", "x", 20],
        ["third", "y", 30],
        ["fourth", "x", 40],
    ]

    sheet["B2:D5"].remove_duplicates(2)

    assert sheet["B2:D5"].value == [
        ["first", "x", 10],
        ["third", "y", 30],
        [None, None, None],
        [None, None, None],
    ]


def test_get_special_cells_filters_constants_and_formulas(sheet):
    sheet["B2"].value = 1
    sheet["C2"].value = "word"
    sheet["D2"].value = True
    sheet["B3"].formula = "=2+2"
    sheet["C3"].formula = '="word"'
    sheet["D3"].formula = "=1/0"
    selected = sheet["B2:D3"]

    areas = selected.get_special_cells("constants", "numbers")

    assert not inspect.isawaitable(areas)
    assert cells(areas) == {(2, 2)}
    assert cells(selected.get_special_cells("constants", "text")) == {(2, 3)}
    assert cells(selected.get_special_cells("constants", "logical")) == {(2, 4)}
    assert cells(selected.get_special_cells("formulas", "numbers")) == {(3, 2)}
    assert cells(selected.get_special_cells("formulas", "text")) == {(3, 3)}
    assert cells(selected.get_special_cells("formulas", "errors")) == {(3, 4)}
    assert cells(selected.get_special_cells("formulas")) == {
        (3, 2),
        (3, 3),
        (3, 4),
    }


def test_get_special_cells_respects_scope_and_empty_result(sheet):
    sheet["B2:D3"].value = [[1, "word", None], [2, "other", None]]
    selected = sheet["B2:D3"]

    assert cells(selected.get_special_cells("blanks")) == {(2, 4), (3, 4)}
    assert selected.get_special_cells("formulas") == []
    assert cells(sheet["D2"].get_special_cells("blanks")) == {(2, 4)}
    assert sheet["D2"].get_special_cells("constants") == []
    assert cells(sheet["B2"].get_special_cells("constants")) == {(2, 2)}
    assert sheet["B2"].get_special_cells("blanks") == []


def test_get_special_cells_visible_excludes_filtered_rows(sheet):
    selected = sheet["B2:C5"]
    selected.value = [
        ["Region", "Value"],
        ["keep", 1],
        ["drop", 2],
        ["keep", 3],
    ]
    selected.autofilter.apply_values(1, ["keep"])

    assert cells(selected.get_special_cells("visible")) == {
        (row, column) for row in (2, 3, 5) for column in (2, 3)
    }
