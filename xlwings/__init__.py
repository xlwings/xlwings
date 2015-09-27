from __future__ import absolute_import
import sys

__version__ = '0.4.1'

# Python 2 vs 3
PY3 = sys.version_info[0] == 3

if PY3:
    string_types = str
    xrange = range
else:
    string_types = basestring
    xrange = xrange

# Platform specifics
if sys.platform.startswith('win'):
    from . import _xlwindows as xlplatform
else:
    from . import _xlmac as xlplatform

time_types = xlplatform.time_types

# API
from .main import Application, Workbook, Range, Chart, Sheet
from .constants import *