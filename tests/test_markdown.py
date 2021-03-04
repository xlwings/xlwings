# Characters are currently not properly supported
# on macOS due to an Excel/AppleScript bug
from pathlib import Path

import pytest
import xlwings as xw
from xlwings.pro import Markdown, MarkdownStyle

this_dir = Path(__file__).resolve().parent


text1 = """\
# Title

Text **bold** and *italic*

* a first bullet
* a second bullet

# Another title

this has a line break
new line
"""


@pytest.fixture(scope="module")
def app():
    app = xw.App(visible=True)
    app.books.open(this_dir / "test book.xlsx")
    yield app

    for book in app.books:
        book.close()
    app.kill()


# Range
def test_markdown_cell_defaults_formatting(app):
    cell = app.books["test book.xlsx"].sheets[0]['A1']
    cell.clear()
    cell.value = Markdown(text1)
    assert cell.font.name == 'Calibri'
    assert cell.font.size == 11
    assert cell.characters[:5].font.bold is True
    assert cell.characters[7:11].font.bold is False
    assert cell.characters[7:11].font.italic is False
    assert cell.characters[21:27].font.italic is True
    assert cell.characters[65:78].font.bold is True


def test_markdown_cell_defaults_value(app):
    cell = app.books["test book.xlsx"].sheets[0]['A1']
    cell.clear()
    cell.value = Markdown(text1)
    assert cell.value == 'Title\n\nText bold and italic\n\n• a first bullet\n• a second bullet\n\nAnother title\n\nthis has a line break\nnew line'


def test_markdown_cell_h1(app):
    cell = app.books["test book.xlsx"].sheets[0]['A1']
    cell.clear()
    style = MarkdownStyle()
    style.h1.blank_lines_after = 2
    style.h1.font.color = (255, 0, 0)
    style.h1.font.bold = False
    style.h1.font.size = 20
    style.h1.font.italic = True
    style.h1.font.name = 'Arial'
    cell.value = Markdown(text1, style)

    for selection in [slice(0, 5), slice(68, 81)]:
        assert cell.characters[selection].font.color == (255, 0, 0)
        assert cell.characters[selection].font.bold is False
        assert cell.characters[selection].font.size == 20
        assert cell.characters[selection].font.italic is True
        assert cell.characters[selection].font.name == 'Arial'


def test_markdown_cell_strong(app):
    cell = app.books["test book.xlsx"].sheets[0]['A1']
    cell.clear()
    style = MarkdownStyle()
    style.strong.color = (255, 0, 0)
    style.strong.bold = False
    style.strong.size = 20
    style.strong.italic = True
    style.strong.name = 'Arial'
    cell.value = Markdown(text1, style)

    assert cell.characters[12:16].font.color == (255, 0, 0)
    assert cell.characters[12:16].font.bold is False
    assert cell.characters[12:16].font.size == 20
    assert cell.characters[12:16].font.italic is True
    assert cell.characters[12:16].font.name == 'Arial'


def test_markdown_cell_emphasis(app):
    cell = app.books["test book.xlsx"].sheets[0]['A1']
    cell.clear()
    style = MarkdownStyle()
    style.emphasis.color = (255, 0, 0)
    style.emphasis.bold = False
    style.emphasis.size = 20
    style.emphasis.italic = True
    style.emphasis.name = 'Arial'
    cell.value = Markdown(text1, style)

    assert cell.characters[21:27].font.color == (255, 0, 0)
    assert cell.characters[21:27].font.bold is False
    assert cell.characters[21:27].font.size == 20
    assert cell.characters[21:27].font.italic is True
    assert cell.characters[21:27].font.name == 'Arial'


def test_markdown_cell_unordered_list(app):
    cell = app.books["test book.xlsx"].sheets[0]['A1']
    cell.clear()
    style = MarkdownStyle()
    style.unordered_list.bullet_character = '-'
    style.unordered_list.blank_lines_after = 1
    cell.value = Markdown(text1, style)

    assert cell.characters[29].text == '-'
    assert cell.characters[65:72].text == 'Another'

    style.unordered_list.blank_lines_after = 2
    cell.clear()
    cell.value = Markdown(text1, style)
    assert cell.characters[66:73].text == 'Another'


# Shape
def test_markdown_shape_defaults_formatting(app):
    shape = app.books["test book.xlsx"].sheets['shape'].shapes[0]
    shape.text = ''
    shape.text = Markdown(text1)
    assert shape.font.name == 'Calibri'
    assert shape.font.size == 11
    assert shape.characters[:5].font.bold is True
    assert shape.characters[7:11].font.bold is False
    assert shape.characters[7:11].font.italic is False
    assert shape.characters[21:27].font.italic is True
    assert shape.characters[65:78].font.bold is True


def test_markdown_shape_defaults_value(app):
    shape = app.books["test book.xlsx"].sheets['shape'].shapes[0]
    shape.text = ''
    shape.text = Markdown(text1)
    assert shape.text == 'Title\n\nText bold and italic\n\n• a first bullet\n• a second bullet\n\nAnother title\n\nthis has a line break\nnew line'


def test_markdown_shape_h1(app):
    shape = app.books["test book.xlsx"].sheets['shape'].shapes[0]
    shape.text = ''
    style = MarkdownStyle()
    style.h1.blank_lines_after = 2
    style.h1.font.color = (255, 0, 0)
    style.h1.font.bold = False
    style.h1.font.size = 20
    style.h1.font.italic = True
    style.h1.font.name = 'Arial'
    shape.text = Markdown(text1, style)

    for selection in [slice(0, 5), slice(68, 81)]:
        assert shape.characters[selection].font.color == (255, 0, 0)
        assert shape.characters[selection].font.bold is False
        assert shape.characters[selection].font.size == 20
        assert shape.characters[selection].font.italic is True
        assert shape.characters[selection].font.name == 'Arial'


def test_markdown_shape_strong(app):
    shape = app.books["test book.xlsx"].sheets['shape'].shapes[0]
    shape.text = ''
    style = MarkdownStyle()
    style.strong.color = (255, 0, 0)
    style.strong.bold = False
    style.strong.size = 20
    style.strong.italic = True
    style.strong.name = 'Arial'
    shape.text = Markdown(text1, style)

    assert shape.characters[12:16].font.color == (255, 0, 0)
    assert shape.characters[12:16].font.bold is False
    assert shape.characters[12:16].font.size == 20
    assert shape.characters[12:16].font.italic is True
    assert shape.characters[12:16].font.name == 'Arial'


def test_markdown_shape_emphasis(app):
    shape = app.books["test book.xlsx"].sheets['shape'].shapes[0]
    shape.text = ''
    style = MarkdownStyle()
    style.emphasis.color = (255, 0, 0)
    style.emphasis.bold = False
    style.emphasis.size = 20
    style.emphasis.italic = True
    style.emphasis.name = 'Arial'
    shape.text = Markdown(text1, style)

    assert shape.characters[21:27].font.color == (255, 0, 0)
    assert shape.characters[21:27].font.bold is False
    assert shape.characters[21:27].font.size == 20
    assert shape.characters[21:27].font.italic is True
    assert shape.characters[21:27].font.name == 'Arial'


def test_markdown_shape_unordered_list(app):
    shape = app.books["test book.xlsx"].sheets['shape'].shapes[0]
    shape.text = ''
    style = MarkdownStyle()
    style.unordered_list.bullet_character = '-'
    style.unordered_list.blank_lines_after = 1
    shape.text = Markdown(text1, style)

    assert shape.characters[29].text == '-'
    assert shape.characters[65:72].text == 'Another'

    style.unordered_list.blank_lines_after = 2
    shape.text = ''
    shape.text = Markdown(text1, style)
    assert shape.characters[66:73].text == 'Another'
