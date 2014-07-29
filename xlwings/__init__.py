from __future__ import absolute_import
import sys

__version__ = '0.2.0'

# Python 2 vs 3
PY3 = sys.version_info[0] == 3

# Platform specific imports
if sys.platform.startswith('win'):
    import xlwings._xlwindows as xlplatform
else:
    import xlwings._xlmac as xlplatform

# API
from .main import Workbook, Range, Chart
from .constants import *