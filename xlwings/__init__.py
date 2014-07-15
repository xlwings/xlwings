from __future__ import absolute_import
import sys

# Python 2 vs 3
PY3 = sys.version_info[0] == 3

# Platform specific imports
if sys.platform.startswith('win'):
    import xlwings._xlwindows as xlplatform

if sys.platform.startswith('darwin'):
    import xlwings._xlmac as xlplatform

# API
from .main import Workbook, Range, Chart, __version__
from .constants import *