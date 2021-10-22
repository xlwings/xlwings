from ..utils import LicenseHandler

# __all__ is required by Sphinx so it won't produce things like xlwings.pro.main.render_template (?)
# and to have undocumented functions
__all__ = ['render_template', 'create_report', 'Image', 'Markdown', 'MarkdownStyle']
LicenseHandler.validate_license('reports')

# API
from .main import render_template, create_report
from .markdown import MarkdownStyle, Markdown
from .image import Image
