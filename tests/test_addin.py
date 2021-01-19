# Run with pytest
import os
from pathlib import Path
import shutil

import xlwings as xw

this_dir = Path(__file__).resolve().parent


def test_config():
    if (Path.home() / '.xlwings').exists():
        shutil.rmtree(Path.home() / '.xlwings')
    app = xw.App(visible=False)
    addin = app.books.open(this_dir.parent / 'xlwings' / 'addin' / 'xlwings.xlam')
    get_config = addin.macro('GetConfig')
    assert get_config('PYTHONPATH') == ''

    # Workbook sheet config
    book = app.books.open(this_dir / 'test book.xlsx')
    sheet = app.books.active.sheets.active
    sheet.name = 'xlwings.conf'
    sheet['A1'].value = ['PYTHONPATH', 'workbook sheet']

    # Addin sheet config
    addin.sheets[0].name = 'xlwings.conf'
    addin.sheets[0]['A1'].value = ['PYTHONPATH', 'addin sheet']

    # Config file same directory
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
    app.quit()

