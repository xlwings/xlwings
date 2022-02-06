"""
Required Notice: Copyright (C) Zoomer Analytics GmbH.

xlwings PRO is dual-licensed under one of the following licenses:

* PolyForm Noncommercial License 1.0.0 (for noncommercial use):
  https://polyformproject.org/licenses/noncommercial/1.0.0
* xlwings PRO License (for commercial use):
  https://github.com/xlwings/xlwings/blob/main/LICENSE_PRO.txt

Commercial licenses can be purchased at https://www.xlwings.org
"""

from .utils import LicenseHandler
from .embedded_code import runpython_embedded_code, dump_embedded_code
from .reports import Markdown, MarkdownStyle
from .module_permissions import verify_execute_permission

LicenseHandler.validate_license("pro")
