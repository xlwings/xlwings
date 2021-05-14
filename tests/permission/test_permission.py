"""
Requires the following env var: TEST_XLWINGS_LICENSE_KEY
Requires permission_server.py running

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
    app = xw.App(visible=True)
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


@pytest.mark.parametrize("method", ["POST", "GET"])
def test_permission_success(clear_user_config, app, addin, method):
    book = app.books.open(Path('.') / 'permission.xlsm')

    config = ['"PERMISSION_CHECK_ENABLED","True"\n',
              '"PERMISSION_CHECK_URL","http://localhost:5000/success"\n',
              f'"PERMISSION_CHECK_METHOD","{method}"\n',
              '"UDF Modules","permission;permission2"\n',
              f'"LICENSE_KEY","{os.environ["TEST_XLWINGS_LICENSE_KEY"]}"']

    os.makedirs(Path.home() / '.xlwings')
    with open((Path.home() / '.xlwings' / 'xlwings.conf'), 'w') as f:
        f.writelines(config)

    # UDF 1
    book.macro('ImportPythonUDFs')()
    book.sheets[0]['A10'].value = '=hello("test")'
    assert book.sheets[0]['A10'].value == 'Hello test!'

    # UDF 2
    book.macro('ImportPythonUDFs')()
    book.sheets[0]['A11'].value = '=hello2("test")'
    assert book.sheets[0]['A11'].value == 'Hello2 test!'

    # RunPython 1
    sample_call = book.macro('Module1.Main')
    sample_call()
    assert book.sheets[0]['A1'].value == 'Hello xlwings!'
    sample_call()
    assert book.sheets[0]['A1'].value == 'Bye xlwings!'

    # RunPython 2
    book.sheets[0]['A1'].clear_contents()

    sample_call = book.macro('Module1.Main2')
    sample_call()
    assert book.sheets[0]['A1'].value == 'Hello2 xlwings!'
    sample_call()
    assert book.sheets[0]['A1'].value == 'Bye2 xlwings!'


@pytest.mark.parametrize("method", ["POST", "GET"])
def test_permission_runpython_server_success(clear_user_config, app, addin, method):
    book = app.books.open(Path('.') / 'permission.xlsm')

    config = ['"PERMISSION_CHECK_ENABLED","True"\n',
              '"PERMISSION_CHECK_URL","http://localhost:5000/success"\n',
              f'"PERMISSION_CHECK_METHOD","{method}"\n',
              '"UDF Modules","permission;permission2"\n',
              f'"LICENSE_KEY","{os.environ["TEST_XLWINGS_LICENSE_KEY"]}"\n',
              '"USE UDF SERVER","True"']

    os.makedirs(Path.home() / '.xlwings')
    with open((Path.home() / '.xlwings' / 'xlwings.conf'), 'w') as f:
        f.writelines(config)

    # RunPython 1
    sample_call = book.macro('Module1.Main')
    sample_call()
    assert book.sheets[0]['A1'].value == 'Hello xlwings!'
    sample_call()
    assert book.sheets[0]['A1'].value == 'Bye xlwings!'

    # RunPython 2
    book.sheets[0]['A1'].clear_contents()

    sample_call = book.macro('Module1.Main2')
    sample_call()
    assert book.sheets[0]['A1'].value == 'Hello2 xlwings!'
    sample_call()
    assert book.sheets[0]['A1'].value == 'Bye2 xlwings!'


@pytest.mark.parametrize("method, endpoint", [("POST", "fail-machinename"),
                                              ("GET", "fail-machinename"), ("GET", "fail-hash"), ("GET", "fail-filename")])
def test_permission_udf_calc_fails(clear_user_config, app, addin, method, endpoint):
    # UDFs have to be already imported
    book = app.books.open(Path('.') / 'permission.xlsm')

    config = ['"PERMISSION_CHECK_ENABLED","True"\n',
              f'"PERMISSION_CHECK_URL","http://localhost:5000/{endpoint}"\n',
              f'"PERMISSION_CHECK_METHOD","{method}"\n',
              '"UDF Modules","permission;permission2"\n',
              f'"LICENSE_KEY","{os.environ["TEST_XLWINGS_LICENSE_KEY"]}"']

    os.makedirs(Path.home() / '.xlwings')
    with open((Path.home() / '.xlwings' / 'xlwings.conf'), 'w') as f:
        f.writelines(config)

    # UDF 1
    book.sheets[0]['A10'].value = '=hello("test")'
    assert 'Failed to get permission' in book.sheets[0]['A10'].value

    # UDF 2
    book.sheets[0]['A11'].value = '=hello2("test")'
    assert 'Failed to get permission' in book.sheets[0]['A11'].value


def test_permission_cant_override_config_file(clear_user_config, app, addin):
    book = app.books.open(Path('.') / 'permission.xlsm')

    config = ['"PERMISSION_CHECK_ENABLED","True"\n',
              f'"PERMISSION_CHECK_URL","http://localhost:5000/fail-machinename"\n',
              f'"PERMISSION_CHECK_METHOD","POST"\n',
              '"UDF Modules","permission;permission2"\n',
              f'"LICENSE_KEY","{os.environ["TEST_XLWINGS_LICENSE_KEY"]}"']

    os.makedirs(Path.home() / '.xlwings')
    with open((Path.home() / '.xlwings' / 'xlwings.conf'), 'w') as f:
        f.writelines(config)

    sheet = book.sheets.add("xlwings.conf")
    sheet['A1'].value = ['PERMISSION_CHECK_ENABLED', False]

    # UDF 1
    book.sheets[0]['A10'].value = '=hello("test")'
    assert 'Failed to get permission' in book.sheets[0]['A10'].value


@pytest.mark.parametrize("method", ["POST", "GET"])
def test_permission_udf_cant_find_file(clear_user_config, app, addin, method):
    # TODO: has to fail currently, need to disable Excel pop-up errors for testing
    book = app.books.open(Path('.') / 'permission.xlsm')

    config = ['"PERMISSION_CHECK_ENABLED","True"\n',
              f'"PERMISSION_CHECK_URL","http://localhost:5000/success"\n',
              f'"PERMISSION_CHECK_METHOD","{method}"\n',
              '"UDF Modules","permission;doesnexist"\n',
              f'"LICENSE_KEY","{os.environ["TEST_XLWINGS_LICENSE_KEY"]}"']

    os.makedirs(Path.home() / '.xlwings')
    with open((Path.home() / '.xlwings' / 'xlwings.conf'), 'w') as f:
        f.writelines(config)

    with pytest.raises(FileNotFoundError):
        book.macro('ImportPythonUDFs')()


@pytest.mark.parametrize("method, endpoint", [("POST", "fail-machinename"),
                                              ("GET", "fail-machinename"), ("GET", "fail-hash"), ("GET", "fail-filename")])
def test_permission_udf_import_fails(clear_user_config, app, addin, method, endpoint):
    # TODO: has to fail currently, need to disable Excel pop-up errors for testing
    book = app.books.open(Path('.') / 'permission.xlsm')

    config = ['"PERMISSION_CHECK_ENABLED","True"\n',
              f'"PERMISSION_CHECK_URL","http://localhost:5000/{endpoint}"\n',
              f'"PERMISSION_CHECK_METHOD","{method}"\n',
              '"UDF Modules","permission;permission2"\n',
              f'"LICENSE_KEY","{os.environ["TEST_XLWINGS_LICENSE_KEY"]}"']

    os.makedirs(Path.home() / '.xlwings')
    with open((Path.home() / '.xlwings' / 'xlwings.conf'), 'w') as f:
        f.writelines(config)

    with pytest.raises(xw.XlwingsError):
        book.macro('ImportPythonUDFs')()


@pytest.mark.parametrize("method", ["POST", "GET"])
def test_permission_runpython_cant_find_file(clear_user_config, app, addin, method):
    # TODO: has to fail currently, need to disable Excel pop-up errors for testing
    book = app.books.open(Path('.') / 'permission.xlsm')

    config = ['"PERMISSION_CHECK_ENABLED","True"\n',
              f'"PERMISSION_CHECK_URL","http://localhost:5000/success"\n',
              f'"PERMISSION_CHECK_METHOD","{method}"\n',
              f'"LICENSE_KEY","{os.environ["TEST_XLWINGS_LICENSE_KEY"]}"']

    os.makedirs(Path.home() / '.xlwings')
    with open((Path.home() / '.xlwings' / 'xlwings.conf'), 'w') as f:
        f.writelines(config)

    with pytest.raises(FileNotFoundError):
        book.macro('Module1.CantFindFile')()


@pytest.mark.parametrize("method", ["POST", "GET"])
def test_permission_runpython_cant_use_from_imports(clear_user_config, app, addin, method):
    # TODO: has to fail currently, need to disable Excel pop-up errors for testing
    book = app.books.open(Path('.') / 'permission.xlsm')

    config = ['"PERMISSION_CHECK_ENABLED","True"\n',
              f'"PERMISSION_CHECK_URL","http://localhost:5000/success"\n',
              f'"PERMISSION_CHECK_METHOD","{method}"\n',
              f'"LICENSE_KEY","{os.environ["TEST_XLWINGS_LICENSE_KEY"]}"']

    os.makedirs(Path.home() / '.xlwings')
    with open((Path.home() / '.xlwings' / 'xlwings.conf'), 'w') as f:
        f.writelines(config)

    with pytest.raises(xw.XlwingsError):
        book.macro('Module1.FromImport')()


@pytest.mark.parametrize("method, endpoint", [("POST", "fail-machinename"),
                                              ("GET", "fail-machinename"), ("GET", "fail-hash"), ("GET", "fail-filename")])
def test_permission_runpython_fails(clear_user_config, app, addin, method, endpoint):
    # TODO: has to fail currently, need to disable Excel pop-up errors for testing
    book = app.books.open(Path('.') / 'permission.xlsm')

    config = ['"PERMISSION_CHECK_ENABLED","True"\n',
              f'"PERMISSION_CHECK_URL","http://localhost:5000/{endpoint}"\n',
              f'"PERMISSION_CHECK_METHOD","{method}"\n',
              f'"LICENSE_KEY","{os.environ["TEST_XLWINGS_LICENSE_KEY"]}"']

    os.makedirs(Path.home() / '.xlwings')
    with open((Path.home() / '.xlwings' / 'xlwings.conf'), 'w') as f:
        f.writelines(config)

    with pytest.raises(xw.XlwingsError):
        book.macro('Module1.Main')()

    with pytest.raises(xw.XlwingsError):
        book.macro('Module1.Main2')()


@pytest.mark.parametrize("method, endpoint", [("POST", "fail-machinename"),
                                              ("GET", "fail-machinename"), ("GET", "fail-hash"), ("GET", "fail-filename")])
def test_permission_runpython_server_fails(clear_user_config, app, addin, method, endpoint):
    # TODO: has to fail currently, need to disable Excel pop-up errors for testing
    book = app.books.open(Path('.') / 'permission.xlsm')

    config = ['"PERMISSION_CHECK_ENABLED","True"\n',
              f'"PERMISSION_CHECK_URL","http://localhost:5000/{endpoint}"\n',
              f'"PERMISSION_CHECK_METHOD","{method}"\n',
              f'"LICENSE_KEY","{os.environ["TEST_XLWINGS_LICENSE_KEY"]}"\n'
              '"USE UDF SERVER","True"']

    os.makedirs(Path.home() / '.xlwings')
    with open((Path.home() / '.xlwings' / 'xlwings.conf'), 'w') as f:
        f.writelines(config)

    with pytest.raises(xw.XlwingsError):
        book.macro('Module1.Main')()

    with pytest.raises(xw.XlwingsError):
        book.macro('Module1.Main2')()


@pytest.mark.parametrize("method", ["GET", "POST"])
def test_permission_embedded_code_success(clear_user_config, app, addin, method):
    book = app.books.open(Path('.') / 'permission.xlsm')

    config = ['"PERMISSION_CHECK_ENABLED","True"\n',
              f'"PERMISSION_CHECK_URL","http://localhost:5000/success-embedded"\n',
              f'"PERMISSION_CHECK_METHOD","{method}"\n',
              f'"LICENSE_KEY","{os.environ["TEST_XLWINGS_LICENSE_KEY"]}"']

    os.makedirs(Path.home() / '.xlwings')
    with open((Path.home() / '.xlwings' / 'xlwings.conf'), 'w') as f:
        f.writelines(config)

    subprocess.run(split('xlwings code embed  -f permission.py'))
    subprocess.run(split('xlwings code embed  -f permission2.py'))

    # UDF 1
    book.macro('ImportPythonUDFs')()
    book.sheets[0]['A10'].value = '=hello("test")'
    assert book.sheets[0]['A10'].value == 'Hello test!'

    # UDF 2
    book.macro('ImportPythonUDFs')()
    book.sheets[0]['A11'].value = '=hello2("test")'
    assert book.sheets[0]['A11'].value == 'Hello2 test!'

    # RunPython 1
    sample_call = book.macro('Module1.Main')
    sample_call()
    assert book.sheets[0]['A1'].value == 'Hello xlwings!'
    sample_call()
    assert book.sheets[0]['A1'].value == 'Bye xlwings!'

    # RunPython 2
    book.sheets[0]['A1'].clear_contents()

    sample_call = book.macro('Module1.Main2')
    sample_call()
    assert book.sheets[0]['A1'].value == 'Hello2 xlwings!'
    sample_call()
    assert book.sheets[0]['A1'].value == 'Bye2 xlwings!'


@pytest.mark.parametrize("method", ["GET", "POST"])
def test_permission_runpython_embedded_code_server_success(clear_user_config, app, addin, method):
    book = app.books.open(Path('.') / 'permission.xlsm')

    config = ['"PERMISSION_CHECK_ENABLED","True"\n',
              f'"PERMISSION_CHECK_URL","http://localhost:5000/success-embedded"\n',
              f'"PERMISSION_CHECK_METHOD","{method}"\n',
              f'"LICENSE_KEY","{os.environ["TEST_XLWINGS_LICENSE_KEY"]}"\n',
              '"USE UDF SERVER","True"']

    os.makedirs(Path.home() / '.xlwings')
    with open((Path.home() / '.xlwings' / 'xlwings.conf'), 'w') as f:
        f.writelines(config)

    subprocess.run(split('xlwings code embed  -f permission.py'))
    subprocess.run(split('xlwings code embed  -f permission2.py'))

    # RunPython 1
    sample_call = book.macro('Module1.Main')
    sample_call()
    assert book.sheets[0]['A1'].value == 'Hello xlwings!'
    sample_call()
    assert book.sheets[0]['A1'].value == 'Bye xlwings!'

    # RunPython 2
    book.sheets[0]['A1'].clear_contents()

    sample_call = book.macro('Module1.Main2')
    sample_call()
    assert book.sheets[0]['A1'].value == 'Hello2 xlwings!'
    sample_call()
    assert book.sheets[0]['A1'].value == 'Bye2 xlwings!'


@pytest.mark.parametrize("method, endpoint", [("POST", "fail-machinename"),
                                              ("GET", "fail-machinename"), ("GET", "fail-hash"), ("GET", "fail-filename")])
def test_permission_embedded_runpython_fails(clear_user_config, app, addin, method, endpoint):
    # TODO: has to fail currently, need to disable Excel pop-up errors for testing
    book = app.books.open(Path('.') / 'permission.xlsm')

    config = ['"PERMISSION_CHECK_ENABLED","True"\n',
              f'"PERMISSION_CHECK_URL","http://localhost:5000/{endpoint}"\n',
              f'"PERMISSION_CHECK_METHOD","{method}"\n',
              f'"LICENSE_KEY","{os.environ["TEST_XLWINGS_LICENSE_KEY"]}"']

    os.makedirs(Path.home() / '.xlwings')
    with open((Path.home() / '.xlwings' / 'xlwings.conf'), 'w') as f:
        f.writelines(config)

    subprocess.run(split('xlwings code embed  -f permission.py'))
    subprocess.run(split('xlwings code embed  -f permission2.py'))

    with pytest.raises(xw.XlwingsError):
        book.macro('Module1.Main')()

    with pytest.raises(xw.XlwingsError):
        book.macro('Module1.Main2')()


@pytest.mark.parametrize("method, endpoint", [("POST", "fail-machinename"),
                                              ("GET", "fail-machinename"), ("GET", "fail-hash"), ("GET", "fail-filename")])
def test_permission_embedded_runpython_server_fails(clear_user_config, app, addin, method, endpoint):
    # TODO: has to fail currently, need to disable Excel pop-up errors for testing
    book = app.books.open(Path('.') / 'permission.xlsm')

    config = ['"PERMISSION_CHECK_ENABLED","True"\n',
              f'"PERMISSION_CHECK_URL","http://localhost:5000/{endpoint}"\n',
              f'"PERMISSION_CHECK_METHOD","{method}"\n',
              f'"LICENSE_KEY","{os.environ["TEST_XLWINGS_LICENSE_KEY"]}"\n'
              '"USE UDF SERVER","True"']

    os.makedirs(Path.home() / '.xlwings')
    with open((Path.home() / '.xlwings' / 'xlwings.conf'), 'w') as f:
        f.writelines(config)

    subprocess.run(split('xlwings code embed  -f permission.py'))
    subprocess.run(split('xlwings code embed  -f permission2.py'))

    with pytest.raises(xw.XlwingsError):
        book.macro('Module1.Main')()

    with pytest.raises(xw.XlwingsError):
        book.macro('Module1.Main2')()


@pytest.mark.parametrize("method, endpoint", [("POST", "fail-machinename"),
                                              ("GET", "fail-machinename"), ("GET", "fail-hash"), ("GET", "fail-filename")])
def test_permission_embedded_udf_import_fails(clear_user_config, app, addin, method, endpoint):
    # TODO: has to fail currently, need to disable Excel pop-up errors for testing
    book = app.books.open(Path('.') / 'permission.xlsm')

    config = ['"PERMISSION_CHECK_ENABLED","True"\n',
              f'"PERMISSION_CHECK_URL","http://localhost:5000/{endpoint}"\n',
              f'"PERMISSION_CHECK_METHOD","{method}"\n',
              '"UDF Modules","permission;permission2"\n',
              f'"LICENSE_KEY","{os.environ["TEST_XLWINGS_LICENSE_KEY"]}"']

    os.makedirs(Path.home() / '.xlwings')
    with open((Path.home() / '.xlwings' / 'xlwings.conf'), 'w') as f:
        f.writelines(config)

    subprocess.run(split('xlwings code embed  -f permission.py'))
    subprocess.run(split('xlwings code embed  -f permission2.py'))

    with pytest.raises(xw.XlwingsError):
        book.macro('ImportPythonUDFs')()


@pytest.mark.parametrize("method, endpoint", [("POST", "fail-machinename"),
                                              ("GET", "fail-machinename"), ("GET", "fail-hash"), ("GET", "fail-filename")])
def test_permission_embedded_udf_calc_fails(clear_user_config, app, addin, method, endpoint):
    # UDFs have to be already imported
    book = app.books.open(Path('.') / 'permission.xlsm')

    config = ['"PERMISSION_CHECK_ENABLED","True"\n',
              f'"PERMISSION_CHECK_URL","http://localhost:5000/{endpoint}"\n',
              f'"PERMISSION_CHECK_METHOD","{method}"\n',
              '"UDF Modules","permission;permission2"\n',
              f'"LICENSE_KEY","{os.environ["TEST_XLWINGS_LICENSE_KEY"]}"']

    os.makedirs(Path.home() / '.xlwings')
    with open((Path.home() / '.xlwings' / 'xlwings.conf'), 'w') as f:
        f.writelines(config)

    subprocess.run(split('xlwings code embed  -f permission.py'))
    subprocess.run(split('xlwings code embed  -f permission2.py'))

    # UDF 1
    book.sheets[0]['A10'].value = '=hello("test")'
    assert 'Failed to get permission' in book.sheets[0]['A10'].value

    # UDF 2
    book.sheets[0]['A11'].value = '=hello2("test")'
    assert 'Failed to get permission' in book.sheets[0]['A11'].value