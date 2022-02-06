from pathlib import Path


import pytest
import xlwings as xw

this_dir = Path(__file__).resolve().parent


@pytest.fixture(scope="module")
def app():
    with xw.App(visible=False) as app:
        app.books.open(this_dir / "test book.xlsx")
        yield app


# Range
def test_range_font_bold(app):
    sheet = app.books[0].sheets[0]
    sheet["A1"].value = "text"
    assert sheet["A1"].font.bold is False
    sheet["A1"].font.bold = True
    assert sheet["A1"].font.bold is True
    sheet["A1"].font.bold = False
    assert sheet["A1"].font.bold is False


def test_range_font_italic(app):
    sheet = app.books[0].sheets[0]
    sheet["A1"].value = "text"
    assert sheet["A1"].font.italic is False
    sheet["A1"].font.italic = True
    assert sheet["A1"].font.italic is True
    sheet["A1"].font.italic = False
    assert sheet["A1"].font.italic is False


def test_range_font_size(app):
    sheet = app.books[0].sheets[0]
    sheet["A1"].value = "text"
    assert sheet["A1"].font.size != 0
    sheet["A1"].font.size = 33
    assert sheet["A1"].font.size == 33


def test_range_font_color(app):
    sheet = app.books[0].sheets[0]
    sheet["A1"].value = "text"
    assert sheet["A1"].font.color == (0, 0, 0)
    sheet["A1"].font.color = (255, 0, 0)
    assert sheet["A1"].font.color == (255, 0, 0)


def test_range_font_name(app):
    sheet = app.books[0].sheets[0]
    sheet["A1"].value = "text"
    sheet["A1"].font.name = "Calibri"
    assert sheet["A1"].font.name == "Calibri"
    sheet["A1"].font.name = "Arial"
    assert sheet["A1"].font.name == "Arial"


# Shape
def test_shape_font_bold(app):
    shape = xw.Book("test book.xlsx").sheets["shape"].shapes[0]
    shape.text = "text"
    assert shape.font.bold is False
    shape.font.bold = True
    assert shape.font.bold is True
    shape.font.bold = False
    assert shape.font.bold is False


def test_shape_font_italic(app):
    shape = xw.Book("test book.xlsx").sheets["shape"].shapes[0]
    shape.text = "text"
    assert shape.font.italic is False
    shape.font.italic = True
    assert shape.font.italic is True
    shape.font.italic = False
    assert shape.font.italic is False


def test_shape_font_size(app):
    shape = xw.Book("test book.xlsx").sheets["shape"].shapes[0]
    shape.text = "text"
    assert shape.font.size != 33
    shape.font.size = 33
    assert shape.font.size == 33


def test_shape_font_color(app):
    shape = xw.Book("test book.xlsx").sheets["shape"].shapes[0]
    shape.text = "text"
    assert shape.font.color == (255, 255, 255)
    shape.font.color = (255, 0, 0)
    assert shape.font.color == (255, 0, 0)


def test_shape_font_name(app):
    shape = xw.Book("test book.xlsx").sheets["shape"].shapes[0]
    shape.text = "text"
    shape.font.name = "Calibri"
    assert shape.font.name == "Calibri"
    shape.font.name = "Arial"
    assert shape.font.name == "Arial"
