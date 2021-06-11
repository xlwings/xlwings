"""
Requires the following env var: TEST_XLWINGS_LICENSE_KEY
If you run this on a built/installed package, make sure to cd out of the xlwings source
directory, copy the test folder next to the install xlwings package,then run:

* all tests (this relies on the settings in pytest.ini):

pytest test_e2e.py

* single test:
pytest test_e2e.py::test_name
"""

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
    for book in app.books:
        book.close()
    app.kill()  # test_addin_installation currently leaves Excel hanging otherwise


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
    return app.books.open(Path(xw.__path__[0]) / 'addin' / 'xlwings.xlam')


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


def test_runpython_embedded_code(clear_user_config, addin, quickstart_book):
    os.makedirs(Path.home() / '.xlwings')
    with open((Path.home() / '.xlwings' / 'xlwings.conf'), 'w') as config:
        config.write(f'"LICENSE_KEY","{os.getenv("TEST_XLWINGS_LICENSE_KEY")}"')
    os.chdir(Path(quickstart_book.fullname).parent)
    subprocess.run(split('xlwings code embed'))
    (Path(quickstart_book.fullname).parent / 'testproject.py').unlink()
    sample_call = quickstart_book.macro('Module1.SampleCall')
    sample_call()
    assert quickstart_book.sheets[0]['A1'].value == 'Hello xlwings!'
    sample_call()
    assert quickstart_book.sheets[0]['A1'].value == 'Bye xlwings!'


def test_udf(clear_user_config, addin, quickstart_book):
    addin.macro('ImportPythonUDFs')()
    quickstart_book.sheets[0]['A1'].value = '=hello("test")'
    assert quickstart_book.sheets[0]['A1'].value == 'Hello test!'


def test_udf_embedded_code(clear_user_config, addin, quickstart_book):
    os.makedirs(Path.home() / '.xlwings')
    with open((Path.home() / '.xlwings' / 'xlwings.conf'), 'w') as config:
        config.write(f'"LICENSE_KEY","{os.getenv("TEST_XLWINGS_LICENSE_KEY")}"')
    os.chdir(Path(quickstart_book.fullname).parent)
    subprocess.run(split('xlwings code embed'))
    (Path(quickstart_book.fullname).parent / 'testproject.py').unlink()
    addin.macro('ImportPythonUDFs')()
    quickstart_book.sheets[0]['A1'].value = '=hello("test")'
    assert quickstart_book.sheets[0]['A1'].value == 'Hello test!'
    (Path.home() / '.xlwings' / 'xlwings.conf').unlink()
    quickstart_book.app.api.CalculateFull()
    assert 'xlwings.LicenseError: Embedded code requires a valid LICENSE_KEY.' in quickstart_book.sheets[0]['A1'].value


def test_can_use_xlwings_without_license_key(clear_user_config, tmp_path):
    import xlwings
    os.chdir(tmp_path)
    subprocess.run(split('xlwings quickstart testproject'))


def test_can_use_xlwings_with_wrong_license_key(clear_user_config, tmp_path):
    os.makedirs(Path.home() / '.xlwings')
    with open((Path.home() / '.xlwings' / 'xlwings.conf'), 'w') as config:
        config.write(f'"LICENSE_KEY","xxx"')
    import xlwings
    os.chdir(tmp_path)
    subprocess.run(split('xlwings quickstart testproject'))


def test_cant_use_xlwings_pro_without_license_key(clear_user_config):
    with pytest.raises(xw.LicenseError):
        import xlwings.pro


def test_addin_installation(app):
    assert not (Path(app.startup_path) / 'xlwings.xlam').exists()
    subprocess.run(split('xlwings addin install'))
    assert (Path(app.startup_path) / 'xlwings.xlam').exists()
    subprocess.run(split('xlwings addin remove'))
    assert not (Path(app.startup_path) / 'xlwings.xlam').exists()

    # Custom file
    assert not (Path(app.startup_path) / 'test book.xlsx').exists()
    os.chdir(this_dir)
    subprocess.run(split('xlwings addin install -f "test book.xlsx"'))
    assert (Path(app.startup_path) / 'test book.xlsx').exists()
    subprocess.run(split('xlwings addin remove -f "test book.xlsx"'))
    assert not (Path(app.startup_path) / 'test book.xlsx').exists()


def test_update_license_key(clear_user_config):
    subprocess.run(split('xlwings license update -k test_key'))
    with open(Path.home() / '.xlwings' / 'xlwings.conf', 'r') as f:
        assert f.read() == '"LICENSE_KEY","test_key"\n'


@pytest.mark.skipif(xw.__version__ == 'dev', reason='requires a built package')
def test_standalone(clear_user_config, app, tmp_path):
    os.chdir(tmp_path)
    subprocess.run(split('xlwings quickstart testproject --standalone'))
    standalone_book = app.books.open(tmp_path / 'testproject' / 'testproject.xlsm')
    sample_call = standalone_book.macro('Module1.SampleCall')
    sample_call()
    assert standalone_book.sheets[0]['A1'].value == 'Hello xlwings!'
    sample_call()
    assert standalone_book.sheets[0]['A1'].value == 'Bye xlwings!'


@pytest.mark.skipif(xw.__version__ == 'dev', reason='requires a built package')
def test_runpython_embedded_code_standalone(app, clear_user_config, tmp_path):
    os.chdir(tmp_path)
    subprocess.run(split(f'xlwings quickstart testproject --standalone'))
    quickstart_book = app.books.open(tmp_path / 'testproject' / 'testproject.xlsm')

    os.makedirs(Path.home() / '.xlwings')
    with open((Path.home() / '.xlwings' / 'xlwings.conf'), 'w') as config:
        config.write(f'"LICENSE_KEY","{os.getenv("TEST_XLWINGS_LICENSE_KEY")}"')

    os.chdir(tmp_path / 'testproject')
    subprocess.run(split('xlwings code embed'))
    (tmp_path / 'testproject' / f'testproject.py').unlink()
    sample_call = quickstart_book.macro('Module1.SampleCall')
    sample_call()
    assert quickstart_book.sheets[0]['A1'].value == 'Hello xlwings!'
    sample_call()
    assert quickstart_book.sheets[0]['A1'].value == 'Bye xlwings!'


@pytest.mark.skipif(xw.__version__ == 'dev', reason='requires a built package')
def test_udf_embedded_code_standalone(clear_user_config, app, tmp_path):
    os.chdir(tmp_path)
    subprocess.run(split(f'xlwings quickstart testproject --standalone'))
    quickstart_book = app.books.open(tmp_path / 'testproject' / 'testproject.xlsm')

    os.makedirs(Path.home() / '.xlwings')
    with open((Path.home() / '.xlwings' / 'xlwings.conf'), 'w') as config:
        config.write(f'"LICENSE_KEY","{os.getenv("TEST_XLWINGS_LICENSE_KEY")}"')

    os.chdir(tmp_path / 'testproject')
    subprocess.run(split('xlwings code embed'))
    (tmp_path / 'testproject' / f'testproject.py').unlink()

    quickstart_book.macro('ImportPythonUDFs')()
    quickstart_book.sheets[0]['A1'].value = '=hello("test")'
    assert quickstart_book.sheets[0]['A1'].value == 'Hello test!'
    (Path.home() / '.xlwings' / 'xlwings.conf').unlink()
    quickstart_book.app.api.CalculateFull()
    assert 'xlwings.LicenseError: Embedded code requires a valid LICENSE_KEY.' in quickstart_book.sheets[0]['A1'].value
