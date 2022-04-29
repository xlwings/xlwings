import os
import sys
from functools import wraps


__version__ = "dev"

# Platform specifics
if sys.platform.startswith("darwin"):
    USER_CONFIG_FILE = os.path.join(
        os.path.expanduser("~"),
        "Library",
        "Containers",
        "com.microsoft.Excel",
        "Data",
        "xlwings.conf",
    )
else:
    USER_CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".xlwings", "xlwings.conf")


# Errors
class XlwingsError(Exception):
    pass


class LicenseError(XlwingsError):
    pass


class ShapeAlreadyExists(XlwingsError):
    pass


# API
from .main import (
    App,
    Book,
    Range,
    Chart,
    Sheet,
    Picture,
    Shape,
    Name,
    view,
    load,
    RangeRows,
    RangeColumns,
    Engine,
)
from .main import apps, books, sheets, engines

# Populate engines list
if sys.platform.startswith("win"):
    try:
        from . import _xlwindows

        engines.add(Engine(impl=_xlwindows.engine))
        engines.active = engines[0]
    except ImportError:
        pass
if sys.platform.startswith("darwin"):
    try:
        from . import _xlmac

        engines.add(Engine(impl=_xlmac.engine))
        engines.active = engines[0]
    except ImportError:
        pass

try:
    from .pro import _xljson

    engines.add(Engine(impl=_xljson.engine))
    PRO = True
except (ImportError, LicenseError):
    PRO = False

# UDFs
if sys.platform.startswith("win"):
    from .udfs import (
        xlfunc as func,
        xlsub as sub,
        xlret as ret,
        xlarg as arg,
        get_udf_module,
        import_udfs,
    )

    # This generates the modules for early-binding under %TEMP%\gen_py\3.x
    # generated via makepy.py -i, but using an old minor=2, as it still seems to
    # generate the most recent version of it whereas it would fail if the minor is
    # higher than what exists on the machine. Allowing it to fail silently, as this is
    # only a hard requirement for ComRange in udf.py which is only used for async and
    # legacy dynamic arrays.
    try:
        from win32com.client import gencache

        gencache.EnsureModule(
            "{00020813-0000-0000-C000-000000000046}", lcid=0, major=1, minor=2
        )
    except:
        pass
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
if sys.platform.startswith("win"):
    from .server import serve
