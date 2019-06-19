from __future__ import absolute_import
from functools import wraps
import sys


__version__ = '0.15.8'

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

    from logging import Logger
    def hasHandlers(self):
        """
        logging backport from Python 3.2
        """
        c = self
        rv = False
        while c:
            if c.handlers:
                rv = True
                break
            if not c.propagate:
                break
            else:
                c = c.parent
        return rv
    Logger.hasHandlers = hasHandlers

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
    from .udfs import xlfunc as func, xlsub as sub, xlret as ret, xlarg as arg, get_udf_module, import_udfs
else:
    def func(f=None, *args, **kwargs):
        @wraps(f)
        def inner(f):
            return f
        if f is None:
            return inner
        else:
            return inner(f)

    def sub(f=None, *args, **kwargs):
        @wraps(f)
        def inner(f):
            return f
        if f is None:
            return inner
        else:
            return inner(f)

    def ret(*args, **kwargs):
        def inner(f):
            return f
        return inner

    def arg(*args, **kwargs):
        def inner(f):
            return f
        return inner


def xlfunc(*args, **kwargs):
    raise Exception("Deprecation: 'xlfunc' has been renamed to 'func' - use 'import xlwings as xw' and decorate your function with '@xw.func'.")


def xlsub(*args, **kwargs):
    raise Exception("Deprecation: 'xlsub' has been renamed to 'sub' - use 'import xlwings as xw' and decorate your function with '@xw.sub'.")


def xlret(*args, **kwargs):
    raise Exception("Deprecation: 'xlret' has been renamed to 'ret' - use 'import xlwings as xw' and decorate your function with '@xw.ret'.")


def xlarg(*args, **kwargs):
    raise Exception("Deprecation: 'xlarg' has been renamed to 'arg' - use 'import xlwings as xw' and decorate your function with '@xw.arg'.")


# Server
if sys.platform.startswith('win'):
    from .server import serve
