"""
Required Notice: Copyright (C) Zoomer Analytics GmbH.

xlwings PRO is dual-licensed under one of the following licenses:

* PolyForm Noncommercial License 1.0.0 (for noncommercial use):
  https://polyformproject.org/licenses/noncommercial/1.0.0
* xlwings PRO License (for commercial use):
  https://github.com/xlwings/xlwings/blob/main/LICENSE_PRO.txt

Commercial licenses can be purchased at https://www.xlwings.org
"""

import json
import os
import sys
from functools import lru_cache
from pathlib import Path

from ..main import Book
from ..utils import read_config_sheet
from .utils import LicenseHandler, get_embedded_code_temp_dir

LicenseHandler.validate_license("pro")

TEMPDIR = get_embedded_code_temp_dir()


@lru_cache()
def dump_embedded_code(book, target_dir):
    code_map = read_config_sheet(book).get("RELEASE_EMBED_CODE_MAP", "{}")
    sheetname_to_path = json.loads(code_map)
    for sheet in book.sheets:
        if sheet.name.endswith(".py"):
            last_cell = sheet.used_range.last_cell
            sheet_content = (
                sheet.range((1, 1), (last_cell.row, 1)).options(ndim=1).value
            )
            if sheetname_to_path:
                (Path(target_dir) / sheetname_to_path[sheet.name]).parent.mkdir(
                    exist_ok=True
                )
            with open(
                os.path.join(
                    target_dir,
                    sheetname_to_path[sheet.name] if sheetname_to_path else sheet.name,
                ),
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
    exec(command)
