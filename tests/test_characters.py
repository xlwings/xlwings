# Characters are currently not properly supported
# on macOS due to an Excel/AppleScript bug
from pathlib import Path

import pytest
import xlwings as xw

this_dir = Path(__file__).resolve().parent


@pytest.fixture(scope="module")
def app():
    with xw.App(visible=False) as app:
        app.books.open(this_dir / "test book.xlsx")
        yield app


# Range no indexing/slicing
def test_range_characters_font_bold(app):
    sheet = app.books[0].sheets[0]
    sheet['A1'].value = 'text'
    assert sheet['A1'].characters.font.bold is False
    sheet['A1'].characters.font.bold = True
    assert sheet['A1'].characters.font.bold is True
    sheet['A1'].characters.font.bold = False
    assert sheet['A1'].characters.font.bold is False


def test_range_font_characters_italic(app):
    sheet = app.books[0].sheets[0]
    sheet['A1'].value = 'text'
    assert sheet['A1'].characters.font.italic is False
    sheet['A1'].characters.font.italic = True
    assert sheet['A1'].characters.font.italic is True
    sheet['A1'].characters.font.italic = False
    assert sheet['A1'].characters.font.italic is False


def test_range_font_characters_size(app):
    sheet = app.books[0].sheets[0]
    sheet['A1'].value = 'text'
    assert sheet['A1'].characters.font.size != 0
    sheet['A1'].characters.font.size = 33
    assert sheet['A1'].characters.font.size == 33


def test_range_font_characters_color(app):
    sheet = app.books[0].sheets[0]
    sheet['A1'].value = 'text'
    assert sheet['A1'].characters.font.color == (0, 0, 0)
    sheet['A1'].characters.font.color = (255, 0, 0)
    assert sheet['A1'].characters.font.color == (255, 0, 0)


# Shape no slicing/indexing
def test_shape_characters_font_bold(app):
    shape = xw.Book("test book.xlsx").sheets["shape"].shapes[0]
    shape.text = "text"
    shape.characters.font.bold = False
    assert shape.characters.font.bold is False
    shape.characters.font.bold = True
    assert shape.characters.font.bold is True
    shape.characters.font.bold = False
    assert shape.characters.font.bold is False


def test_shape_characters_font_italic(app):
    shape = xw.Book("test book.xlsx").sheets["shape"].shapes[0]
    shape.text = "text"
    assert shape.characters.font.italic is False
    shape.characters.font.italic = True
    assert shape.characters.font.italic is True
    shape.characters.font.italic = False
    assert shape.characters.font.italic is False


def test_shape_characters_font_size(app):
    shape = xw.Book("test book.xlsx").sheets["shape"].shapes[0]
    shape.text = "text"
    assert shape.characters.font.size != 33
    shape.characters.font.size = 33
    assert shape.characters.font.size == 33


def test_shape_characters_font_color(app):
    shape = xw.Book("test book.xlsx").sheets["shape"].shapes[0]
    shape.text = "text"
    assert shape.characters.font.color == (255, 255, 255)
    shape.characters.font.color = (255, 0, 0)
    assert shape.characters.font.color == (255, 0, 0)


# Range indexing
def test_range_characters_index_font_bold(app):
    sheet = app.books[0].sheets[0]
    sheet['A1'].value = 'text'
    assert sheet['A1'].characters[1].font.bold is False
    sheet['A1'].characters[1].font.bold = True
    assert sheet['A1'].characters[1].font.bold is True
    sheet['A1'].characters[1].font.bold = False
    assert sheet['A1'].characters[1].font.bold is False


def test_range_font_characters_index_italic(app):
    sheet = app.books[0].sheets[0]
    sheet['A1'].value = 'text'
    sheet['A1'].characters[1].font.italic = False
    assert sheet['A1'].characters[1].font.italic is False
    sheet['A1'].characters[1].font.italic = True
    assert sheet['A1'].characters[1].font.italic is True
    sheet['A1'].characters[1].font.italic = False
    assert sheet['A1'].characters[1].font.italic is False


def test_range_font_characters_index_size(app):
    sheet = app.books[0].sheets[0]
    sheet['A1'].value = 'text'
    assert sheet['A1'].characters[1].font.size != 0
    sheet['A1'].characters[1].font.size = 33
    assert sheet['A1'].characters[1].font.size == 33


def test_range_font_characters_index_color(app):
    sheet = app.books[0].sheets[0]
    sheet['A1'].value = 'text'
    sheet['A1'].characters[1].font.color = (0, 0, 0)
    assert sheet['A1'].characters[1].font.color == (0, 0, 0)
    sheet['A1'].characters[1].font.color = (255, 0, 0)
    assert sheet['A1'].characters[1].font.color == (255, 0, 0)


# Shape indexing
def test_shape_characters_index_font_bold(app):
    shape = xw.Book("test book.xlsx").sheets["shape"].shapes[0]
    shape.text = "text"
    assert shape.characters[1].font.bold is False
    shape.characters[1].font.bold = True
    assert shape.characters[1].font.bold is True
    shape.characters[1].font.bold = False
    assert shape.characters[1].font.bold is False


def test_shape_characters_index_font_italic(app):
    shape = xw.Book("test book.xlsx").sheets["shape"].shapes[0]
    shape.text = "text"
    assert shape.characters[1].font.italic is False
    shape.characters[1].font.italic = True
    assert shape.characters[1].font.italic is True
    shape.characters[1].font.italic = False
    assert shape.characters[1].font.italic is False


def test_shape_characters_index_font_size(app):
    shape = xw.Book("test book.xlsx").sheets["shape"].shapes[0]
    shape.text = "text"
    shape.characters[1].font.size = 30
    assert shape.characters[1].font.size == 30
    shape.characters[1].font.size = 33
    assert shape.characters[1].font.size == 33


def test_shape_characters_index_font_color(app):
    shape = xw.Book("test book.xlsx").sheets["shape"].shapes[0]
    shape.text = "text"
    shape.characters[1].font.color = (255, 255, 255)
    assert shape.characters[1].font.color == (255, 255, 255)
    shape.characters[1].font.color = (255, 0, 0)
    assert shape.characters[1].font.color == (255, 0, 0)


# Range slicing
def test_range_characters_slicing_font_bold(app):
    sheet = app.books[0].sheets[0]
    sheet['A1'].value = 'text'
    assert sheet['A1'].characters[2:4].font.bold is False
    sheet['A1'].characters[2:4].font.bold = True
    assert sheet['A1'].characters[2:4].font.bold is True
    sheet['A1'].characters[2:4].font.bold = False
    assert sheet['A1'].characters[2:4].font.bold is False


def test_range_font_characters_slicing_italic(app):
    sheet = app.books[0].sheets[0]
    sheet['A1'].value = 'text'
    sheet['A1'].characters[2:4].font.italic = False
    assert sheet['A1'].characters[2:4].font.italic is False
    sheet['A1'].characters[2:4].font.italic = True
    assert sheet['A1'].characters[2:4].font.italic is True
    sheet['A1'].characters[2:4].font.italic = False
    assert sheet['A1'].characters[2:4].font.italic is False


def test_range_font_characters_slicing_size(app):
    sheet = app.books[0].sheets[0]
    sheet['A1'].value = 'text'
    assert sheet['A1'].characters[2:4].font.size != 0
    sheet['A1'].characters[2:4].font.size = 33
    assert sheet['A1'].characters[2:4].font.size == 33


def test_range_font_characters_slicing_color(app):
    sheet = app.books[0].sheets[0]
    sheet['A1'].value = 'text'
    sheet['A1'].characters[2:4].font.color = (0, 0, 0)
    assert sheet['A1'].characters[2:4].font.color == (0, 0, 0)
    sheet['A1'].characters[2:4].font.color = (255, 0, 0)
    assert sheet['A1'].characters[2:4].font.color == (255, 0, 0)


# Shape slicing
def test_shape_characters_slicing_font_bold(app):
    shape = xw.Book("test book.xlsx").sheets["shape"].shapes[0]
    shape.text = "text"
    assert shape.characters[2:4].font.bold is False
    shape.characters[2:4].font.bold = True
    assert shape.characters[2:4].font.bold is True
    shape.characters[2:4].font.bold = False
    assert shape.characters[2:4].font.bold is False


def test_shape_characters_slicing_font_italic(app):
    shape = xw.Book("test book.xlsx").sheets["shape"].shapes[0]
    shape.text = "text"
    assert shape.characters[2:4].font.italic is False
    shape.characters[2:4].font.italic = True
    assert shape.characters[2:4].font.italic is True
    shape.characters[2:4].font.italic = False
    assert shape.characters[2:4].font.italic is False


def test_shape_characters_slicing_font_size(app):
    shape = xw.Book("test book.xlsx").sheets["shape"].shapes[0]
    shape.text = "text"
    shape.characters[2:4].font.size = 30
    assert shape.characters[2:4].font.size == 30
    shape.characters[2:4].font.size = 33
    assert shape.characters[2:4].font.size == 33


def test_shape_characters_slicing_font_color(app):
    shape = xw.Book("test book.xlsx").sheets["shape"].shapes[0]
    shape.text = "text"
    shape.characters[2:4].font.color = (255, 255, 255)
    assert shape.characters[2:4].font.color == (255, 255, 255)
    shape.characters[2:4].font.color = (255, 0, 0)
    assert shape.characters[2:4].font.color == (255, 0, 0)
