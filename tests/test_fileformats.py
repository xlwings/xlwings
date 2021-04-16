import os
from pathlib import Path

import pytest
import xlwings as xw

this_dir = Path(__file__).resolve().parent


@pytest.fixture(scope="module")
def app():
    app = xw.App(visible=False)
    yield app

    for book in reversed(app.books):
        book.close()
    app.kill()
    for f in Path('.').glob('tempfile*'):
        f.unlink()
    for f in Path('/tmp').glob('tempfile*'):
        f.unlink()


def test_save_new_book_defaults(app):
    if Path('Book1.xlsx').is_file():
        Path('Book1.xlsx').unlink()
    book = app.books.add()
    book.save()

    assert Path('Book1.xlsx').is_file()


# TODO: xlam fails

@pytest.mark.parametrize("name", ["tempfile.xlsx", "tempfile.xlsm", "tempfile.xlsb", "tempfile.xltm", "tempfile.xltx", "tempfile.xls", "tempfile.xlt", "tempfile.xla"])
def test_save_new_book_no_path(app, name):
    book = app.books.add()
    book.save(name)
    assert book.name == name
    assert Path(name).is_file()


@pytest.mark.parametrize("name", ["tempfile2.xlsx", "tempfile2.xlsm", "tempfile2.xlsb", "tempfile2.xltm", "tempfile2.xltx", "tempfile2.xls", "tempfile2.xlt", "tempfile2.xla"])
def test_save_new_book_with_path(app, name):
    Path('temp').mkdir(exist_ok=True)
    book = app.books.add()
    fullname = Path('.').resolve() / 'temp' / name
    book.save(fullname)
    assert book.fullname == str(fullname)
    assert Path(fullname).is_file()


@pytest.mark.parametrize("name", ["tempfile1.xlsx", "tempfile1.xlsm", "tempfile1.xlsb", "tempfile1.xltm", "tempfile1.xltx", "tempfile1.xls", "tempfile1.xlt", "tempfile1.xla"])
def test_save_existing_book_no_path(app, name):
    book = app.books.open(this_dir / "test book.xlsx")
    book.save(name)
    assert book.name == name
    assert Path(name).is_file()


@pytest.mark.parametrize("name", ["tempfile1.xlsx", "tempfile1.xlsm", "tempfile1.xlsb", "tempfile1.xltm", "tempfile1.xltx", "tempfile1.xls", "tempfile1.xlt", "tempfile1.xla"])
def test_save_existing_book_with_path(app, name):
    Path('temp').mkdir(exist_ok=True)
    book = app.books.open(this_dir / "test book.xlsx")
    fullname = Path('.').resolve() / 'temp' / name
    book.save(fullname)
    assert book.fullname == str(fullname)
    assert Path(fullname).is_file()
