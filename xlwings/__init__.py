from __future__ import absolute_import
import sys

__version__ = '0.6.2'

# Python 2 vs 3
PY3 = sys.version_info[0] == 3

if PY3:
    string_types = str
    xrange = range
    from builtins import map
else:
    string_types = basestring
    xrange = xrange
    from future_builtins import map

# Platform specifics
if sys.platform.startswith('win'):
    from . import _xlwindows as xlplatform
else:
    from . import _xlmac as xlplatform

time_types = xlplatform.time_types

# Errors
class ShapeAlreadyExists(Exception):
    pass

# API
from .main import Application, Workbook, Range, Chart, Sheet, Picture, Shape, Plot
from .constants import *

# UDFs
from .udfs import xlfunc, xlsub, xlret, xlarg, udf_script, import_udfs
