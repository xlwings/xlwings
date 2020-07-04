import sys
from functools import wraps


__version__ = 'dev'

# Platform specifics
if sys.platform.startswith('win'):
    from . import _xlwindows as xlplatform
else:
    from . import _xlmac as xlplatform

time_types = xlplatform.time_types
USER_CONFIG_FILE = xlplatform.USER_CONFIG_FILE


# Errors
class ShapeAlreadyExists(Exception):
    pass


class LicenseError(Exception):
    pass


# API
from .main import App, Book, Range, Chart, Sheet, Picture, Shape, Name, view, RangeRows, RangeColumns
from .main import apps, books, sheets

try:
    from . import pro
    PRO = True
except (ImportError, LicenseError):
    PRO = False

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

# Server
if sys.platform.startswith('win'):
    from .server import serve
