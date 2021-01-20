import os
from pathlib import Path
import shutil
from shlex import split
import subprocess

import pytest
import xlwings as xw

this_dir = Path(__file__).resolve().parent


@pytest.fixture
def app():
    app = xw.App(visible=False)
    yield app
    app.kill()


@pytest.fixture
def clear_user_config():
    if (Path.home() / '.backup_xlwings').exists():
        shutil.rmtree(Path.home() / '.backup_xlwings')
    if (Path.home() / '.xlwings').exists():
        shutil.copytree(Path.home() / '.xlwings', Path.home() / '.backup_xlwings')
        shutil.rmtree(Path.home() / '.xlwings')
    yield
    if (Path.home() / '.xlwings').exists():
        shutil.rmtree(Path.home() / '.xlwings')
    if (Path.home() / '.backup_xlwings').exists():
        shutil.copytree(Path.home() / '.backup_xlwings', Path.home() / '.xlwings')


@pytest.fixture
def addin(app):
    return app.books.open(this_dir.parent / 'xlwings' / 'addin' / 'xlwings.xlam')


@pytest.fixture
def quickstart_book(app, tmpdir):
    os.chdir(tmpdir)
    subprocess.run(split('xlwings quickstart testproject'))
    return app.books.open(Path(tmpdir) / 'testproject' / 'testproject.xlsm')


def test_config(clear_user_config, app, addin):
    get_config = addin.macro('GetConfig')
    assert get_config('PYTHONPATH') == ''

    # Workbook sheet config
    book = app.books.open(this_dir / 'test book.xlsx')
    sheet = book.sheets[0]
    sheet.name = 'xlwings.conf'
    sheet['A1'].value = ['PYTHONPATH', 'workbook sheet']

    # Addin sheet config
    addin.sheets[0].name = 'xlwings.conf'
    addin.sheets[0]['A1'].value = ['PYTHONPATH', 'addin sheet']

    # Config file workbook directory
    with open(this_dir / 'xlwings.conf', 'w') as config:
        config.write('"PYTHONPATH","directory config"')

    # Config file user home directory
    os.makedirs(Path.home() / '.xlwings', exist_ok=True)
    with open((Path.home() / '.xlwings' / 'xlwings.conf'), 'w') as config:
        config.write('"PYTHONPATH","user config"')

    assert get_config('PYTHONPATH') == 'workbook sheet'
    sheet.name = '_xlwings.conf'
    assert get_config('PYTHONPATH') == 'addin sheet'
    addin.sheets[0].name = '_xlwings.conf'
    assert get_config('PYTHONPATH') == 'directory config'
    (this_dir / 'xlwings.conf').unlink()
    assert get_config('PYTHONPATH') == 'user config'


def test_runpython(addin, quickstart_book):
    quickstart_book.sheets['_xlwings.conf'].name = 'xlwings.conf'
    sample_call = quickstart_book.macro('Module1.SampleCall')
    sample_call()
    assert quickstart_book.sheets[0]['A1'].value == 'Hello xlwings!'
    sample_call()
    assert quickstart_book.sheets[0]['A1'].value == 'Bye xlwings!'


def test_runpython_server(addin, quickstart_book):
    sample_call = quickstart_book.macro('Module1.SampleCall')
    quickstart_book.sheets['_xlwings.conf'].name = 'xlwings.conf'
    quickstart_book.sheets['xlwings.conf']['B8'].value = True
    sample_call()
    assert quickstart_book.sheets[0]['A1'].value == 'Hello xlwings!'
    sample_call()
    assert quickstart_book.sheets[0]['A1'].value == 'Bye xlwings!'


def test_embedded_code(clear_user_config, addin, quickstart_book):
    os.makedirs(Path.home() / '.xlwings')
    with open((Path.home() / '.xlwings' / 'xlwings.conf'), 'w') as config:
        config.write(f'"LICENSE_KEY","{os.getenv("TEST_LICENSE_KEY")}"')
    os.chdir(Path(quickstart_book.fullname).parent)
    subprocess.run(split('xlwings code embed'))
    (Path(quickstart_book.fullname).parent / 'testproject.py').unlink()
    sample_call = quickstart_book.macro('Module1.SampleCall')
    sample_call()
    assert quickstart_book.sheets[0]['A1'].value == 'Hello xlwings!'
    sample_call()
    assert quickstart_book.sheets[0]['A1'].value == 'Bye xlwings!'
