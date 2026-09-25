"""Round-trip find and replace tests against desktop Excel on Windows and macOS."""

import inspect

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


@pytest.mark.parametrize(
    ("order", "direction", "address"),
    [
        ("rows", "forward", "$C$2"),
        ("rows", "backward", "$C$4"),
        ("columns", "forward", "$B$3"),
        ("columns", "backward", "$D$3"),
    ],
)
def test_find_traversal_order(sheet, order, direction, address):
    selected = sheet["B2:D4"]
    selected.value = [
        [None, "needle", None],
        ["needle", None, "needle"],
        [None, "needle", None],
    ]

    found = selected.find("needle", whole=True, order=order, direction=direction)

    assert isinstance(found, xw.Range)
    assert not inspect.isawaitable(found)
    assert found.address == address
    assert found.sheet.name == sheet.name


def test_find_respects_scope_and_match_options(sheet):
    sheet["A1:E2"].value = [
        ["Needle", "needle suffix", "needle", "needle", "needle"],
        ["other", "other", "other", "other", "needle"],
    ]

    assert sheet["B1:C1"].find("needle", whole=True).address == "$C$1"
    assert sheet["B1:C1"].find("needle", match_case=True).address == "$B$1"
    assert sheet["B1:C1"].find("needle", whole=True, match_case=True).address == (
        "$C$1"
    )
    assert sheet["B1:C1"].find("absent") is None
    assert sheet["D1"].find("needle", whole=True).address == "$D$1"
    assert sheet["D1"].find("needle", direction="backward").address == "$D$1"
    assert sheet["D2"].find("needle") is None


def test_replace_all_respects_scope_and_options(sheet):
    sheet["A1:E2"].value = [
        ["cat", "cat", "Cat", "catfish", "cat"],
        ["cat", "CAT", "cat", "other", "other"],
    ]

    alerts_state = sheet.book.app.display_alerts
    assert sheet["B1:D2"].replace_all("cat", "dog", whole=True, match_case=True) is None
    assert sheet.book.app.display_alerts == alerts_state
    assert sheet["A1:E2"].value == [
        ["cat", "dog", "Cat", "catfish", "cat"],
        ["cat", "CAT", "dog", "other", "other"],
    ]

    sheet["D1"].replace_all("cat", "", match_case=True)
    assert sheet["D1"].value == "fish"
    assert sheet["A1"].value == "cat"
    assert sheet["E1"].value == "cat"

    sheet["D1"].replace_all("FISH", "dog", whole=True)
    assert sheet["D1"].value == "dog"


def test_replace_all_preserves_single_cell_formula(sheet):
    cell = sheet["D1"]
    cell.formula = '="cat"&"fish"'

    cell.replace_all("cat", "dog")

    assert cell.formula == '="dog"&"fish"'
    assert cell.value == "dogfish"
