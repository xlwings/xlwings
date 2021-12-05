import os
import glob
import time
import shutil
import sys
import tempfile

from .utils import LicenseHandler
from .module_permissions import verify_execute_permission
from ..main import Book
from ..utils import get_cached_user_config

LicenseHandler.validate_license('pro')


def dump_embedded_code(book, target_dir):
    for sheet in book.sheets:
        if sheet.name.endswith(".py"):
            last_cell = sheet.used_range.last_cell
            sheet_content = sheet.range((1, 1), (last_cell.row, 1)).options(ndim=1).value

            with open(os.path.join(target_dir, sheet.name), 'w', encoding='utf-8', newline='\n') as f:
                for row in sheet_content:
                    if row is None:
                        f.write('\n')
                    else:
                        f.write(row + "\n")
    sys.path[0:0] = [target_dir]


def runpython_embedded_code(command):
    with tempfile.TemporaryDirectory(prefix='xlwings-') as tempdir:
        dump_embedded_code(Book.caller(), tempdir)
        if (get_cached_user_config('permission_check_enabled')
                and get_cached_user_config('permission_check_enabled').lower() == 'true'):
            verify_execute_permission(command=command)
        exec(command)


def get_udf_temp_dir():
    tmp_base_path = os.path.join(tempfile.gettempdir(), 'xlwingsudfs')
    os.makedirs(tmp_base_path, exist_ok=True)
    try:
        # HACK: Clean up directories that are older than 30 days
        # This should be done in the C++ part when the Python process is killed from there
        for subdir in glob.glob(tmp_base_path + '/*/'):
            if os.path.getmtime(subdir) < time.time() - 30 * 86400:
                shutil.rmtree(subdir, ignore_errors=True)
    except Exception:
        pass  # we don't care if it fails
    tempdir = tempfile.TemporaryDirectory(dir=tmp_base_path)
    return tempdir
