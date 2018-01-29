from __future__ import absolute_import
import sys


__version__ = 'dev'

# Python 2 vs 3
PY3 = sys.version_info[0] == 3

if PY3:
    string_types = str
    xrange = range
    from builtins import map
    import builtins
else:
    string_types = basestring
    xrange = xrange
    from future_builtins import map
    builtins = __builtins__

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
from .main import App, Book, Range, Chart, Sheet, Picture, Shape, Name, view, RangeRows, RangeColumns
from .main import apps, books, sheets

# UDFs
if sys.platform.startswith('win'):
    from .udfs import xlfunc as func, xlsub as sub, xlret as ret, xlarg as arg, xlevent as event, xleventexc as eventexc, get_udf_module, import_udfs
else:
    def func(*args, **kwargs):
        def real_decorator(f):
            return f
        return real_decorator

    def sub(*args, **kwargs):
        def real_decorator(f):
            return f
        return real_decorator

    def ret(*args, **kwargs):
        def real_decorator(f):
            return f
        return real_decorator

    def arg(*args, **kwargs):
        def real_decorator(f):
            return f
        return real_decorator

    def event(*args, **kwargs):
        def real_decorator(f):
            return f
        return real_decorator

    def eventexc(*args, **kwargs):
        def real_decorator(f):
            return f
        return real_decorator


def xlfunc(*args, **kwargs):
    raise Exception("Deprecation: 'xlfunc' has been renamed to 'func' - use 'import xlwings as xw' and decorate your function with '@xw.func'.")


def xlsub(*args, **kwargs):
    raise Exception("Deprecation: 'xlsub' has been renamed to 'sub' - use 'import xlwings as xw' and decorate your function with '@xw.sub'.")


def xlret(*args, **kwargs):
    raise Exception("Deprecation: 'xlret' has been renamed to 'ret' - use 'import xlwings as xw' and decorate your function with '@xw.ret'.")


def xlarg(*args, **kwargs):
    raise Exception("Deprecation: 'xlarg' has been renamed to 'arg' - use 'import xlwings as xw' and decorate your function with '@xw.arg'.")

# Events
from .event_dispatcher import EventDispatcher, EventHandler

# Server
if sys.platform.startswith('win'):
    from .server import serve
