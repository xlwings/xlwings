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


class NoSuchObjectError(XlwingsError):
    pass


# API
from .main import (
    App,
    Book,
    Chart,
    Engine,
    Name,
    Picture,
    Range,
    RangeColumns,
    RangeRows,
    Shape,
    Sheet,
    apps,
    books,
    engines,
    load,
    sheets,
    view,
)

__all__ = (
    "App",
    "Book",
    "Chart",
    "Engine",
    "Name",
    "Picture",
    "Range",
    "RangeColumns",
    "RangeRows",
    "Shape",
    "Sheet",
    "apps",
    "books",
    "engines",
    "load",
    "sheets",
    "view",
)

# Populate engines list
has_pywin32 = False
if sys.platform.startswith("win"):
    try:
        from . import _xlwindows

        engines.add(Engine(impl=_xlwindows.engine))
        has_pywin32 = True
    except ImportError:
        pass
if sys.platform.startswith("darwin"):
    try:
        from . import _xlmac

        engines.add(Engine(impl=_xlmac.engine))
    except ImportError:
        pass

try:
    from .pro import _xlremote

    engines.add(Engine(impl=_xlremote.engine))
    PRO = True
except (ImportError, LicenseError):
    PRO = False

try:
    # Separately handled in case the Rust extension is missing
    from .pro import _xlcalamine

    engines.add(Engine(impl=_xlcalamine.engine))
except (ImportError, LicenseError):
    pass

if engines:
    engines.active = engines[0]

# UDFs
if sys.platform.startswith("win") and has_pywin32:
    from .server import serve
    from .udfs import (
        get_udf_module,
        import_udfs,
        xlarg as arg,
        xlfunc as func,
        xlret as ret,
        xlsub as sub,
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
    except:  # noqa: E722
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

    def raise_missing_pywin32():
        raise ImportError(
            "Couldn't find 'pywin32'. Install it via"
            "'pip install pywin32' or 'conda install pywin32'."
        )

    serve = raise_missing_pywin32
    get_udf_module = raise_missing_pywin32
    import_udfs = raise_missing_pywin32
