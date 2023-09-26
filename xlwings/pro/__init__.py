"""
Required Notice: Copyright (C) Zoomer Analytics GmbH.

xlwings PRO is dual-licensed under one of the following licenses:

* PolyForm Noncommercial License 1.0.0 (for noncommercial use):
  https://polyformproject.org/licenses/noncommercial/1.0.0
* xlwings PRO License (for commercial use):
  https://github.com/xlwings/xlwings/blob/main/LICENSE_PRO.txt

Commercial licenses can be purchased at https://www.xlwings.org
"""

from .embedded_code import dump_embedded_code, runpython_embedded_code
from .reports import Markdown, MarkdownStyle
from .udfs_officejs import (
    custom_functions_call,
    custom_functions_code,
    custom_functions_meta,
    xlarg as arg,
    xlfunc as func,
    xlret as ret,
)
from .utils import LicenseHandler

__all__ = (
    "dump_embedded_code",
    "runpython_embedded_code",
    "verify_execute_permission",
    "Markdown",
    "MarkdownStyle",
    "arg",
    "func",
    "ret",
    "custom_functions_code",
    "custom_functions_meta",
    "custom_functions_call",
)

LicenseHandler.validate_license("pro")
