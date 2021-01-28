from ..utils import LicenseHandler

# __all__ is required by Sphinx so it won't produce things like xlwings.reports.main.create_pdf
# and to have undocumented functions
__all__ = ['create_report']
LicenseHandler.validate_license('reports')

# API
from .main import create_report
