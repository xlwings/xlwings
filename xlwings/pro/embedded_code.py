import os
import sys
import tempfile

from .utils import LicenseHandler
from ..main import Book

LicenseHandler.validate_license('pro')


def dump_embedded_code(book, target_dir):
    for sheet in book.sheets:
        if sheet.name.endswith(".py"):
            last_cell = sheet.used_range.last_cell
            sheet_content = sheet.range((1, 1), (last_cell.row, 1)).options(ndim=1).value

            with open(os.path.join(target_dir, sheet.name), 'w') as f:
                for row in sheet_content:
                    if row is None:
                        f.write('\n')
                    else:
                        f.write(row + "\n")
    sys.path[0:0] = [target_dir]


def runpython_embedded_code(command):
    with tempfile.TemporaryDirectory(prefix='xlwings-') as tempdir:
        dump_embedded_code(Book.caller(), tempdir)
        exec(command)
