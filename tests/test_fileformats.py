from pathlib import Path

import pytest
import xlwings as xw

this_dir = Path(__file__).resolve().parent


@pytest.fixture(scope="module")
def app():
    with xw.App(visible=False) as app:
        yield app

    for f in Path('.').glob('tempfile*'):
        f.unlink()
    for f in Path('temp').glob('tempfile*'):
        f.unlink()


def test_save_new_book_defaults(app):
    book = app.books.add()
    if Path(book.name + '.xlsx').is_file():
        Path(book.name + '.xlsx').unlink()
    book.save()

    assert Path(book.name).is_file()


# TODO: xlam and xltx fail

@pytest.mark.parametrize("name", ["tempfile.xlsx", "tempfile.xlsm", "tempfile.xlsb", "tempfile.xltm", "tempfile.xls", "tempfile.xlt", "tempfile.xla"])
def test_save_new_book_no_path(app, name):
    book = app.books.add()
    book.save(name)
    assert book.name == name
    assert Path(name).is_file()


@pytest.mark.parametrize("name", ["tempfile2.xlsx", "tempfile2.xlsm", "tempfile2.xlsb", "tempfile2.xltm", "tempfile2.xls", "tempfile2.xlt", "tempfile2.xla"])
def test_save_new_book_with_path(app, name):
    Path('temp').mkdir(exist_ok=True)
    book = app.books.add()
    fullname = Path('.').resolve() / 'temp' / name
    book.save(fullname)
    assert book.fullname == str(fullname)
    assert Path(fullname).is_file()


@pytest.mark.parametrize("name", ["tempfile3.xlsx", "tempfile3.xlsm", "tempfile3.xlsb", "tempfile3.xltm", "tempfile3.xls", "tempfile3.xlt", "tempfile3.xla"])
def test_save_existing_book_no_path(app, name):
    book = app.books.open(this_dir / "test book.xlsx")
    book.save(name)
    book.save()
    assert book.name == name
    assert Path(name).is_file()


@pytest.mark.parametrize("name", ["tempfile4.xlsx", "tempfile4.xlsm", "tempfile4.xlsb", "tempfile4.xltm", "tempfile4.xls", "tempfile4.xlt", "tempfile4.xla"])
def test_save_existing_book_with_path(app, name):
    Path('temp').mkdir(exist_ok=True)
    book = app.books.open(this_dir / "test book.xlsx")
    fullname = Path('.').resolve() / 'temp' / name
    book.save(fullname)
    book.save()
    assert book.fullname == str(fullname)
    assert Path(fullname).is_file()
