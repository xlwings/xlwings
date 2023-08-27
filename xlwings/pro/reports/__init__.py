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

__all__ = (
    "render_template",
    "create_report",
    "Image",
    "Markdown",
    "MarkdownStyle",
    "register_formatter",
    "formatter",
)

LicenseHandler.validate_license("reports")

# API
from .image import Image
from .main import create_report, render_template
from .markdown import Markdown, MarkdownStyle

format_callbacks = {}


def formatter(func):
    """Decorator"""
    format_callbacks[func.__name__] = func
    return func


# Deprecated
register_formatter = formatter
