"""Round-trip Range.move tests against desktop Excel."""

import pytest

import xlwings as xw


@pytest.fixture(scope="module")
def app():
    with xw.App(visible=False) as app:
        yield app


@pytest.fixture
def book(app):
    book = app.books.add()
    try:
        yield book
    finally:
        book.close()


def test_move_to_overlapping_destination(book):
    sheet = book.sheets[0]
    sheet["O1:S2"].value = [["a", "b", "c", "d", "e"], [1, 2, 3, 4, 5]]
    sheet["O1"].color = (255, 0, 0)
    sheet["A1"].formula = "=S2"

    # Overlaps the source in O:P; only the destination's top-left cell counts.
    sheet["O1:S2"].move(sheet["L1:Z9"])

    assert sheet["L1:S2"].value == [
        ["a", "b", "c", "d", "e", None, None, None],
        [1, 2, 3, 4, 5, None, None, None],
    ]
    assert sheet["L1"].color == (255, 0, 0)
    assert sheet["O1"].color is None
    assert sheet["A1"].formula == "=P2"


def test_move_to_another_sheet(book):
    source = book.sheets[0]
    target = book.sheets.add()
    source["A1:B2"].value = [["x", "y"], [1, 2]]

    source["A1:B2"].move(target["C3"])

    assert target["C3:D4"].value == [["x", "y"], [1, 2]]
    assert source["A1:B2"].value == [[None, None], [None, None]]


def test_move_rejects_destination_in_another_book(app, book):
    other = app.books.add()
    try:
        with pytest.raises(ValueError, match="same book"):
            book.sheets[0]["A1"].move(other.sheets[0]["A1"])
    finally:
        other.close()
