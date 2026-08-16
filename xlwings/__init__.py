from __future__ import annotations

import json
import os
import sys
from typing import TYPE_CHECKING, Annotated, Any, Callable, TypeVar, overload

__version__ = "0.0.0"

# TypeVar for the CachedObject[T] type alias (defined after the ObjectHandle class).
_CachedT = TypeVar("_CachedT")

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


class ObjectCacheMissError(XlwingsError):
    """Raised when an object handle can't be resolved from the object cache, e.g.,
    because the cached object has expired or been evicted. Carries the cache ``key`` so
    that the caller can render a "stale" object handle instead of a hard error."""

    def __init__(self, key=None):
        self.key = key
        super().__init__(
            "Object handle is no longer cached"
            if key is None
            else f"Object handle is no longer cached: {key}"
        )


class ObjectHandle:
    """Wraps a return value of a custom function so that it's stored as an object handle
    with custom presentation. Use it to override the cell text, icon, and the properties
    shown on the object handle's card on a per-object basis::

        import xlwings as xw

        @xw.func
        def make_obj() -> xw.ObjectHandle:
            df = build()
            return xw.ObjectHandle(
                df,
                text=f"{len(df)} rows",
                icon=xw.ObjectHandleIcons.table,
                properties={
                    "Region": {
                        "type": "String",
                        "basicValue": str(df["region"].iloc[0]),
                    },
                },
            )

    The wrapped object (``df``) is what gets cached; ``text``, ``icon``, and
    ``properties`` only shape the object handle's appearance. ``properties`` follows the
    Excel entity property format (see the example above) and, when provided, replaces
    the automatically derived properties (such as type and shape) as the complete set
    shown on the card.

    Use ``ObjectHandle`` (bare) as the return type hint to store the returned object as an
    object handle (``object`` works too). To type a custom function *argument* that is an
    object handle, use :data:`CachedObject` instead, which keeps the wrapped type visible
    to editors and type checkers.
    """

    def __init__(self, obj, *, text=None, icon=None, properties=None):
        self.obj = obj
        self.text = text
        self.icon = icon
        self.properties = properties or {}

    def __class_getitem__(cls, item):
        from typing import Annotated

        return Annotated[item, cls]


class WithScript:
    """Wraps the return value of a custom function to request that a custom script runs
    after the function returns::

        import xlwings as xw

        @xw.func
        def hello_with_script(name):
            return xw.WithScript(
                f"Hello {name}!",
                "hello_args",
                args=[name, 42],
            )

    ``value`` is converted and written to the cell exactly as it would be without the
    wrapper. ``script`` is either the custom script function itself or its name as a
    string. ``args`` are passed on to the script and must be JSON-serializable.

    The script runs at the next calculation boundary after the custom function returns
    (or, on hosts below ExcelApi 1.8, on a best-effort basis right after it returns, with
    no guarantee that the cell value has been committed). A custom function may not write
    to the grid during its own invocation, which is why the script can't run inline. Note
    that every invocation that returns a ``WithScript`` triggers the script, so filling a
    formula down N rows runs it N times. Not supported in streaming functions.
    """

    def __init__(self, value, script, *, args=None):
        if isinstance(value, WithScript):
            raise XlwingsError(
                "Cannot nest WithScript objects: only one script can be requested per "
                "return value."
            )
        if args is None:
            args = []
        if not isinstance(args, (list, tuple)):
            raise XlwingsError("WithScript: 'args' must be a list.")
        args = list(args)
        try:
            # allow_nan=False: NaN/Infinity are accepted by json.dumps but aren't valid
            # JSON, so they'd pass here only to blow up when the HTTP response is
            # serialized - far away from the line that caused it.
            json.dumps(args, allow_nan=False)
        except (TypeError, ValueError) as e:
            raise XlwingsError(
                f"WithScript: 'args' must be JSON-serializable: {e}"
            ) from None
        self.value = value
        self.script_name = script if isinstance(script, str) else script.__name__
        self.args = args

    @property
    def payload(self):
        """The script request as sent to the client. Excludes ``value``, which travels
        separately as the (converted) cell value."""
        return {
            "script_name": self.script_name,
            "args": self.args,
        }


class CustomFunctionResult:
    """Envelope returned by ``custom_functions_call`` when the custom function requested
    a follow-up script via :class:`WithScript`. Carries the already-converted cell value
    alongside the script request, so the two stay coupled instead of being passed via
    ambient state. Functions that don't use ``WithScript`` return their converted value
    directly, unchanged.
    """

    def __init__(self, value, script):
        self.value = value
        self.script = script


if TYPE_CHECKING:
    # For type checkers, ``CachedObject[T]`` is just ``T``, so editors show the real type
    # of an object-handle argument (e.g. CachedObject[pd.DataFrame] -> pd.DataFrame).
    CachedObject = Annotated[_CachedT, ObjectHandle]
else:
    # At runtime it's ObjectHandle, so CachedObject[T] == ObjectHandle[T] ==
    # Annotated[T, ObjectHandle], which custom functions convert via the object cache.
    CachedObject = ObjectHandle


# API
from .main import (
    App,
    Book,
    BookAsync,
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
from .utils import xlserial_to_datetime as to_datetime

__all__ = (
    "App",
    "Book",
    "BookAsync",
    "Chart",
    "Engine",
    "Name",
    "CachedObject",
    "CustomFunctionResult",
    "ObjectCacheMissError",
    "ObjectHandle",
    "WithScript",
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
    "to_datetime",
    # UDF decorators, defined further below depending on platform/environment.
    # 'sub' is deliberately left out: it's a deprecated alias of 'script' and
    # doesn't exist when running on the server.
    "arg",
    "func",
    "ret",
    "script",
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
    from .pro import _xlofficejs, _xlremote

    engines.add(Engine(impl=_xlremote.engine))
    engines.add(Engine(impl=_xlofficejs.engine))
    __pro__ = True
except (ImportError, LicenseError, AttributeError):
    __pro__ = False

try:
    # The Caller type hint for custom functions. Imported here (rather than only from
    # xlwings.server) so it's available as xw.Caller, like the other custom-function
    # types above. Importing it also registers it as an injectable typehint.
    from .pro.caller import Caller  # noqa: F401

    # Unlike the other names in __all__, Caller lives under pro/ and is therefore
    # license-gated: importing it runs LicenseHandler.validate_license("pro"). Only
    # advertise it once it's actually bound, or `from xlwings import *` would raise
    # AttributeError on an install without a valid PRO license.
    __all__ += ("Caller",)
except (ImportError, LicenseError, AttributeError):
    pass

try:
    # Separately handled in case the Rust extension is missing
    from .pro import _xlcalamine

    engines.add(Engine(impl=_xlcalamine.engine))
except (ImportError, LicenseError, AttributeError):
    pass

if "excel" in [engine.name for engine in engines]:
    # An active engine only really makes sense for the interactive mode with a desktop
    # installation of Excel. Still, you could activate an engine explicitly via
    # xw.engines["engine_name"].activate() which might be useful for testing purposes.
    engines.active = engines["excel"]

# UDFs
_F = TypeVar("_F", bound=Callable[..., Any])

on_server = os.environ.get("XLWINGS_ON_SERVER") == "true"

if on_server:
    from xlwings.server import arg, func, ret, script  # noqa: F401
elif sys.platform.startswith("win") and has_pywin32:
    from .com_server import serve
    from .udfs import (
        get_udf_module,
        import_udfs,
        xlarg as arg,
        xlfunc as func,
        xlret as ret,
        xlsub as script,
        xlsub as sub,
    )

    # This generates the modules for early-binding under %TEMP%\gen_py\3.x
    # generated via makepy.py -i, but using an old minor=2, as it still seems to
    # generate the most recent version of it whereas it would fail if the minor is
    # higher than what exists on the machine. Allowing it to fail silently, as this is
    # only a hard requirement for ComRange in udf.py which is only used for async funcs,
    # legacy dynamic arrays, and the 'caller' argument.
    try:
        from win32com.client import gencache

        gencache.EnsureModule(
            "{00020813-0000-0000-C000-000000000046}", lcid=0, major=1, minor=2
        )
    except:  # noqa: E722
        pass
else:

    @overload
    def func(f: _F) -> _F:
        ...

    @overload
    def func(f: None = ..., *args: Any, **kwargs: Any) -> Callable[[_F], _F]:
        ...

    def func(f: _F | None = None, *args: Any, **kwargs: Any) -> _F | Callable[[_F], _F]:
        def inner(f: _F) -> _F:
            return f

        if f is None:
            return inner
        else:
            return inner(f)

    @overload
    def sub(f: _F) -> _F:
        ...

    @overload
    def sub(f: None = ..., *args: Any, **kwargs: Any) -> Callable[[_F], _F]:
        ...

    def sub(f: _F | None = None, *args: Any, **kwargs: Any) -> _F | Callable[[_F], _F]:
        def inner(f: _F) -> _F:
            return f

        if f is None:
            return inner
        else:
            return inner(f)

    @overload
    def script(f: _F) -> _F:
        ...

    @overload
    def script(f: None = ..., *args: Any, **kwargs: Any) -> Callable[[_F], _F]:
        ...

    def script(
        f: _F | None = None, *args: Any, **kwargs: Any
    ) -> _F | Callable[[_F], _F]:
        def inner(f: _F) -> _F:
            return f

        if f is None:
            return inner
        else:
            return inner(f)

    def ret(*args: Any, **kwargs: Any) -> Callable[[_F], _F]:
        def inner(f: _F) -> _F:
            return f

        return inner

    def arg(*args: Any, **kwargs: Any) -> Callable[[_F], _F]:
        def inner(f: _F) -> _F:
            return f

        return inner

    def raise_missing_pywin32() -> None:
        raise ImportError(
            "Couldn't find 'pywin32'. Install it via"
            "'pip install pywin32' or 'conda install pywin32'."
        )

    serve = raise_missing_pywin32
    get_udf_module = raise_missing_pywin32
    import_udfs = raise_missing_pywin32

# This follows the Office Script/Office.js convention to make the constants available
# in the top-level namespace. Should be done for all constants with xlwings 1.0.
from .constants import ObjectHandleIcons  # noqa: F401
