"""
Required Notice: Copyright (C) Zoomer Analytics GmbH.

xlwings PRO is dual-licensed under one of the following licenses:

* PolyForm Noncommercial License 1.0.0 (for noncommercial use):
  https://polyformproject.org/licenses/noncommercial/1.0.0
* xlwings PRO License (for commercial use):
  https://github.com/xlwings/xlwings/blob/main/LICENSE_PRO.txt

Commercial licenses can be purchased at https://www.xlwings.org
"""

from ..utils import LicenseHandler

# __all__ is required by Sphinx so it won't produce things like
# xlwings.pro.main.render_template (?) and to have undocumented functions
__all__ = ["render_template", "create_report", "Image", "Markdown", "MarkdownStyle"]
LicenseHandler.validate_license("reports")

# API
from .main import render_template, create_report
from .markdown import MarkdownStyle, Markdown
from .image import Image
