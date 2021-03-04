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
    app = xw.App(visible=False)
    app.books.open(this_dir / "test book.xlsx")
    yield app

    for book in app.books:
        book.close()
    app.kill()


def test_markdown_cell_defaults_formatting(app):
    cell = app.books["test book.xlsx"].sheets[0]['A1']
    cell.clear()
    cell.value = Markdown(text1)
    assert cell.font.name == 'Calibri'
    assert cell.font.size == 11
    assert cell.characters[:5].font.bold is True
    assert cell.characters[7:11].font.bold is False
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
    style.h1.font.bold = False
    style.h1.font.italic = True
    style.h1.font.size = 20
    style.h1.font.color = (255, 0, 0)
    style.h1.font.name = 'Arial'
    cell.value = Markdown(text1, style)

    for selection in [slice(0, 5), slice(68, 81)]:
        assert cell.characters[selection].font.bold is False
        assert cell.characters[selection].font.italic is True
        assert cell.characters[selection].font.size == 20
        assert cell.characters[selection].font.color == (255, 0, 0)
        assert cell.characters[selection].font.name == 'Arial'
