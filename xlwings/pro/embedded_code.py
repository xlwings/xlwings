"""
Required Notice: Copyright (C) Zoomer Analytics GmbH.

xlwings PRO is dual-licensed under one of the following licenses:

* PolyForm Noncommercial License 1.0.0 (for noncommercial use):
  https://polyformproject.org/licenses/noncommercial/1.0.0
* xlwings PRO License (for commercial use):
  https://github.com/xlwings/xlwings/blob/main/LICENSE_PRO.txt

Commercial licenses can be purchased at https://www.xlwings.org
"""

import os
import sys
from functools import lru_cache

from .utils import LicenseHandler
from .module_permissions import verify_execute_permission
from ..main import Book
from .utils import get_embedded_code_temp_dir
from ..utils import get_cached_user_config

LicenseHandler.validate_license("pro")

TEMPDIR = get_embedded_code_temp_dir()


@lru_cache()
def dump_embedded_code(book, target_dir):
    for sheet in book.sheets:
        if sheet.name.endswith(".py"):
            last_cell = sheet.used_range.last_cell
            sheet_content = (
                sheet.range((1, 1), (last_cell.row, 1)).options(ndim=1).value
            )

            with open(
                os.path.join(target_dir, sheet.name),
                "w",
                encoding="utf-8",
                newline="\n",
            ) as f:
                for row in sheet_content:
                    if row is None:
                        f.write("\n")
                    else:
                        f.write(row + "\n")
    sys.path[0:0] = [target_dir]


def runpython_embedded_code(command):
    dump_embedded_code(Book.caller(), TEMPDIR)
    if (
        get_cached_user_config("permission_check_enabled")
        and get_cached_user_config("permission_check_enabled").lower() == "true"
    ):
        verify_execute_permission(command=command)
    exec(command)
