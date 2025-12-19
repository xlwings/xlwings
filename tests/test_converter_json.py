import datetime as dt
from pathlib import Path

import pytest

import xlwings as xw

this_dir = Path(__file__).resolve().parent


@pytest.fixture(scope="module")
def app():
    with xw.App(visible=False) as app:
        app.books.add()
        yield app


def test_read_single_value_float(app):
    sheet = app.books[0].sheets[0]
    sheet["A1"].clear()
    sheet["A1"].value = 3.14
    result = sheet["A1"].options("json").value
    assert result == "3.14"


def test_read_single_value_str(app):
    sheet = app.books[0].sheets[0]
    sheet["A1"].clear()
    sheet["A1"].value = "hello"
    result = sheet["A1"].options("json").value
    assert result == '"hello"'


def test_read_single_value_bool(app):
    sheet = app.books[0].sheets[0]
    sheet["A1"].clear()
    sheet["A1"].value = True
    result = sheet["A1"].options("json").value
    assert result == "true"
    sheet["A1"].clear()
    sheet["A1"].value = False
    result = sheet["A1"].options("json").value
    assert result == "false"


def test_read_single_value_empty(app):
    sheet = app.books[0].sheets[0]
    sheet["A1"].clear()
    sheet["A1"].value = None
    result = sheet["A1"].options("json").value
    assert result == "null"


def test_read_single_value_date(app):
    sheet = app.books[0].sheets[0]
    sheet["A1"].clear()
    sheet["A1"].value = dt.datetime(2024, 1, 15, 10, 30, 0)
    result = sheet["A1"].options("json").value
    assert result == '"2024-01-15T10:30:00"'


def test_read_horizontal_1d_range(app):
    sheet = app.books[0].sheets[0]
    sheet["A1:E1"].clear()
    sheet["A1"].value = [1, 2.5, "test", True, None]
    result = sheet["A1:E1"].options("json").value
    assert result == '[1.0, 2.5, "test", true, null]'


def test_read_vertical_1d_range(app):
    sheet = app.books[0].sheets[0]
    sheet["A1:A3"].clear()
    sheet["A1"].value = [[1], [2], [3]]
    result = sheet["A1:A3"].options("json").value
    assert result == "[1.0, 2.0, 3.0]"


def test_read_2d_range(app):
    sheet = app.books[0].sheets[0]
    sheet["A1"].value = [[1, 2, 3], [4, 5, 6]]
    result = sheet["A1:C2"].options("json").value
    assert result == "[[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]]"


def test_read_2d_range_mixed_types(app):
    sheet = app.books[0].sheets[0]
    sheet["A1"].value = [[1, "a", True], [2.5, "b", False]]
    result = sheet["A1:C2"].options("json").value
    assert result == '[[1.0, "a", true], [2.5, "b", false]]'


def test_read_2d_range_with_empty_cells(app):
    sheet = app.books[0].sheets[0]
    sheet["A1:C2"].clear_contents()
    sheet["A1"].value = 1
    sheet["B1"].value = 2
    sheet["A2"].value = 3
    result = sheet["A1:C2"].options("json").value
    assert result == "[[1.0, 2.0, null], [3.0, null, null]]"


def test_read_2d_range_with_dates(app):
    sheet = app.books[0].sheets[0]
    test_date1 = dt.datetime(2024, 1, 1)
    test_date2 = dt.datetime(2024, 12, 31)
    sheet["A1"].value = [[test_date1, "text"], [test_date2, 42]]
    result = sheet["A1:B2"].options("json").value
    assert result == '[["2024-01-01T00:00:00", "text"], ["2024-12-31T00:00:00", 42.0]]'


def test_write_json_single_value(app):
    sheet = app.books[0].sheets[0]
    sheet["A1"].clear()
    sheet["A1"].options("json").value = "42"
    assert sheet["A1"].value == 42


def test_write_json_1d_array(app):
    sheet = app.books[0].sheets[0]
    sheet["A1:C1"].clear()
    sheet["A1"].options("json").value = "[1, 2, 3]"
    result = sheet["A1:C1"].value
    assert result == [1.0, 2.0, 3.0]


def test_write_json_2d_array(app):
    sheet = app.books[0].sheets[0]
    sheet["A1:B2"].clear()
    sheet["A1"].options("json").value = "[[1, 2], [3, 4]]"
    result = sheet["A1:B2"].value
    assert result == [[1.0, 2.0], [3.0, 4.0]]


def test_write_json_with_datetime(app):
    sheet = app.books[0].sheets[0]
    sheet["A1"].clear()
    sheet["A1"].options("json").value = '"2024-01-15T10:30:00"'
    result = sheet["A1"].value
    assert result == dt.datetime(2024, 1, 15, 10, 30, 0)


def test_write_json_jagged_array(app):
    """Test that jagged arrays are padded with None"""
    sheet = app.books[0].sheets[0]
    sheet["A1"].options("json").value = '[["a", "b"], ["c"]]'
    result = sheet["A1:B2"].value
    assert result == [["a", "b"], ["c", None]]


def test_write_json_with_markdown_code_block(app):
    """Test that markdown code blocks (```json...```) are stripped from AI responses"""
    sheet = app.books[0].sheets[0]
    sheet["A1:D2"].clear()
    # Simulate AI-generated response with markdown code block
    ai_response = """```json
[
    ["Conglomerate", "US", 95000, 32.8],
    ["Biotechnology", "US", 50000, 56.2]
]
```"""
    sheet["A1"].options("json").value = ai_response
    result = sheet["A1:D2"].value
    assert result == [
        ["Conglomerate", "US", 95000.0, 32.8],
        ["Biotechnology", "US", 50000.0, 56.2],
    ]
