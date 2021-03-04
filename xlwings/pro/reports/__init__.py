from ..utils import LicenseHandler

# __all__ is required by Sphinx so it won't produce things like xlwings.pro.main.create_report (?)
# and to have undocumented functions
__all__ = ['create_report']
LicenseHandler.validate_license('reports')

# API
from .main import create_report
from .markdown import MarkdownStyle, Markdown
