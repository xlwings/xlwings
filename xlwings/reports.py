# Allows you to use xlwings.reports instead of xlwings.pro.reports

from .pro.reports import *  # noqa: F401,F403

__all__ = (  # noqa: F405
    "render_template",
    "create_report",
    "Image",
    "Markdown",
    "MarkdownStyle",
    "formatter",
)
