"""
Required Notice: Copyright (C) Zoomer Analytics GmbH.

xlwings PRO is dual-licensed under one of the following licenses:

* PolyForm Noncommercial License 1.0.0 (for noncommercial use):
  https://polyformproject.org/licenses/noncommercial/1.0.0
* xlwings PRO License (for commercial use):
  https://github.com/xlwings/xlwings/blob/main/LICENSE_PRO.txt

Commercial licenses can be purchased at https://www.xlwings.org
"""

from __future__ import annotations

import asyncio
import contextlib
import datetime as dt
import inspect
import logging
import os
import re
import types
import warnings
from functools import wraps
from pathlib import Path
from textwrap import dedent
from typing import (
    Annotated,
    Any,
    Callable,
    Literal,
    TypeVar,
    Union,
    get_args,
    get_origin,
    get_type_hints,
    overload,
)

_F = TypeVar("_F", bound=Callable[..., Any])

import xlwings as xw

from .. import ObjectHandle, XlwingsError, __version__, conversion
from . import object_handles

logger = logging.getLogger(__name__)

# Tasks started by streaming functions
background_tasks = {}

MODULE_NAMESPACE_ATTRIBUTE = "__xlwings_func_namespace__"

# Typehints whose values are injected into custom functions by the framework instead of
# being provided by Excel, e.g. xlwings Server's CurrentUser. Such parameters are hidden
# from the custom function's Excel-facing signature, are skipped when converting the
# incoming args, and may appear in any position - including keyword-only, i.e. after
# *args. Frameworks built on top of xlwings register their own types via
# register_injectable_typehint().
# NOTE: this registry only applies to custom functions. xw.Book must NOT be registered
# here: it is injected into custom *scripts* (via custom_scripts_call), which handle it
# separately, while custom functions never receive a book. Registering it would hide
# xw.Book params from the Excel-facing signature and silently shift the remaining args.
_injectable_typehints = set()


def register_injectable_typehint(type_hint) -> None:
    """Mark a typehint as framework-injected so that parameters annotated with it are
    excluded from the Excel-facing signature and may be placed anywhere in the signature.
    """
    _injectable_typehints.add(type_hint)


def _unwrap_optional_hint(hint):
    """Return X for Optional[X] / X | None, otherwise the hint unchanged.

    Restricted to unions on purpose: unwrapping any generic would turn
    `list[int]` into `int`.
    """
    # types.UnionType (the `X | None` form) only exists on Python 3.10+.
    union_types = {Union, getattr(types, "UnionType", Union)}
    if get_origin(hint) not in union_types:
        return hint
    members = [arg for arg in get_args(hint) if arg is not type(None)]
    return members[0] if len(members) == 1 else hint


def is_injectable_typehint(type_hint) -> bool:
    # A `= None` default makes get_type_hints() report Optional[X] rather than X, so
    # unwrap before the lookup - otherwise `user: CurrentUser = None` wouldn't be
    # recognized as injectable.
    type_hint = _unwrap_optional_hint(type_hint)
    try:
        return type_hint in _injectable_typehints
    except TypeError:
        # Unhashable typehints (e.g. some parametrized generics) are never injectable
        return False


def get_custom_function_namespace(function, module=None):
    """Return the explicit or defining-module namespace for a custom function."""
    namespace = function.__xlfunc__.get("namespace")
    if namespace is not None:
        return namespace

    defining_module = inspect.getmodule(function)
    if defining_module is None:
        defining_module = module
    return getattr(defining_module, MODULE_NAMESPACE_ATTRIBUTE, None)


def func_sig(f):
    sig = inspect.signature(f)
    # Resolved lazily and tolerantly: a UDF may carry typehints that can't be resolved
    # here (e.g. under `from __future__ import annotations` with local names), which
    # must not break the signature check itself.
    try:
        type_hints = get_type_hints(f)
    except Exception:
        type_hints = {}
    vararg = None
    args = []
    defaults = []
    injected = []
    for param in sig.parameters.values():
        # Framework-injected params (e.g. CurrentUser) are never supplied by Excel, so
        # they're allowed in any position, including keyword-only after *args.
        if is_injectable_typehint(type_hints.get(param.name)):
            injected.append(param.name)
            continue
        if param.kind is inspect.Parameter.POSITIONAL_OR_KEYWORD:
            args.append(param.name)
            if param.default is not inspect.Signature.empty:
                defaults.append(param.default)
        elif param.kind is inspect.Parameter.VAR_POSITIONAL:
            args.append(param.name)
            vararg = param.name
        else:
            raise XlwingsError("xlwings does not support UDFs with keyword arguments")
    return {
        "args": args,
        "defaults": defaults,
        "vararg": vararg,
        "injected": injected,
    }


def check_bool(kw, default, **func_kwargs):
    if kw in func_kwargs:
        check = func_kwargs.pop(kw)
        if isinstance(check, bool):
            return check
        raise XlwingsError(f'{kw} only takes boolean values. ("{check}" provided).')
    return default


def extract_type_and_annotations(type_hint):
    """Extracts only the top-level type, i.e., list for type_hint=list[list[int]]
    so that the ValueAccessor doesn't have to register all possibilities of nested types
    TODO: it would, however, be great to make list[list[dt.datetime]] work as well as
    use list[list] as equivalent to ndim=2
    """
    origin = get_origin(type_hint)
    if origin is Annotated:
        base_type, *annotations = get_args(type_hint)
        top_level_type = get_origin(base_type) or base_type
        # ObjectHandle[T] (i.e. Annotated[T, ObjectHandle]) marks an object handle while
        # keeping T as the type seen by editors/type checkers. Convert via the object
        # cache (registered for `object`), not via T's own converter.
        if ObjectHandle in annotations:
            top_level_type = object
        # Only the dict-style annotations carry conversion options; drop markers such as
        # ObjectHandle so downstream code can assume annotations are option dicts.
        annotations = [a for a in annotations if isinstance(a, dict)]
    else:
        top_level_type = origin or type_hint
        annotations = []
    # A bare ObjectHandle (e.g. `-> ObjectHandle`) is an alias for `object`, i.e. it's
    # converted via the object cache.
    if top_level_type is ObjectHandle:
        top_level_type = object
    return top_level_type, annotations


def extract_enum_descriptor(type_hint, func_name, param_name):
    """If type_hint is Literal[...] (optionally wrapped in Annotated[..., {"tooltips":
    ...}]), return a descriptor dict for an Office.js custom-function enum. Otherwise
    None.

    Descriptor shape:
        {"id": str, "type": "string"|"number", "values": list,
         "tooltips": dict[value, str]}
    """
    base = type_hint
    companion = {}
    if get_origin(type_hint) is Annotated:
        base, *annotations = get_args(type_hint)
        for ann in annotations:
            if isinstance(ann, dict):
                companion = ann
                break
    if get_origin(base) is not Literal:
        return None
    values = list(get_args(base))
    if not values:
        return None
    if all(isinstance(v, str) for v in values):
        enum_type = "string"
    elif all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in values):
        enum_type = "number"
    else:
        raise XlwingsError(
            f"Literal values for parameter '{param_name}' of '{func_name}' must be "
            "all strings or all numbers. Mixed types are not supported."
        )
    tooltips = {k: v for k, v in companion.get("tooltips", {}).items() if k in values}
    return {
        "id": f"{func_name}.{param_name}".upper(),
        "type": enum_type,
        "values": values,
        "tooltips": tooltips,
    }


def _validate_excel_name(name):
    # Office custom function names must start with a letter and may only contain
    # letters, numbers, periods, and underscores (max 128 characters)
    if name is None:
        return None
    if len(name) > 128 or not re.fullmatch(r"[^\W\d_][\w.]*", name):
        raise XlwingsError(
            f"Invalid custom function name '{name}': it must start with a letter "
            "and contain only letters, numbers, periods, and underscores "
            "(max. 128 characters)."
        )
    return name


@overload
def xlfunc(f: _F) -> _F:
    ...


@overload
def xlfunc(f: None = ..., **kwargs: Any) -> Callable[[_F], _F]:
    ...


def xlfunc(f: _F | None = None, **kwargs: Any) -> _F | Callable[[_F], _F]:
    def inner(f: _F) -> _F:
        if not hasattr(f, "__xlfunc__"):
            type_hints = get_type_hints(f, include_extras=True)  # requires Python 3.9
            xlf = f.__xlfunc__ = {}
            xlf["name"] = f.__name__
            xlargs = xlf["args"] = []
            xlargmap = xlf["argmap"] = {}
            sig = func_sig(f)
            num_args = len(sig["args"])
            num_defaults = len(sig["defaults"])
            num_required_args = num_args - num_defaults
            if sig["vararg"] and num_defaults > 0:
                raise XlwingsError(
                    "xlwings does not support UDFs "
                    "with both optional and variable length arguments"
                )
            for var_pos, var_name in enumerate(sig["args"]):
                arg_info = {
                    "name": var_name,
                    "pos": var_pos,
                    "doc": f"Positional argument {var_pos + 1}",
                    "vararg": var_name == sig["vararg"],
                    "options": {},
                }
                if var_name in type_hints:
                    enum_descriptor = extract_enum_descriptor(
                        type_hints[var_name], f.__name__, var_name
                    )
                    type_hint, annotations = extract_type_and_annotations(
                        type_hints[var_name]
                    )
                    if enum_descriptor is not None:
                        arg_info["options"]["enum"] = enum_descriptor
                        if annotations and "doc" in annotations[0]:
                            arg_info["doc"] = annotations[0]["doc"]
                    else:
                        arg_info["options"]["convert"] = type_hint
                        if annotations:
                            for key, value in annotations[0].items():
                                if key == "doc":
                                    arg_info["doc"] = value
                                else:
                                    arg_info["options"][key] = value
                if var_pos >= num_required_args:
                    arg_info["optional"] = sig["defaults"][var_pos - num_required_args]
                xlargs.append(arg_info)
                xlargmap[var_name] = xlargs[-1]
            xlf["ret"] = {
                "doc": (
                    f.__doc__
                    if f.__doc__ is not None
                    else f"Python function '{f.__name__}'"
                ),
                "options": {},
            }
            if "return" in type_hints:
                type_hint, annotations = extract_type_and_annotations(
                    type_hints["return"]
                )
                xlf["ret"]["options"]["convert"] = type_hint
                if annotations:
                    xlf["ret"]["options"].update(annotations[0])

        f.__xlfunc__["volatile"] = check_bool("volatile", default=False, **kwargs)
        # The Excel-facing function name (case-preserved); the metadata id and
        # the dispatch key remain the Python function name
        f.__xlfunc__["excel_name"] = _validate_excel_name(kwargs.get("name"))
        # If there's a global namespace defined in the manifest, this will be the
        # sub-namespace, i.e. NAMESPACE.SUBNAMESPACE.FUNCTIONNAME
        f.__xlfunc__["namespace"] = kwargs.get("namespace")
        f.__xlfunc__["help_url"] = kwargs.get("help_url")
        f.__xlfunc__["required_roles"] = kwargs.get("required_roles")
        return f

    if f is None:
        return inner
    else:
        return inner(f)


def xlret(convert: Any = None, **kwargs: Any) -> Callable[[_F], _F]:
    if convert is not None:
        kwargs["convert"] = convert

    def inner(f: _F) -> _F:
        xlf = xlfunc(f).__xlfunc__
        xlr = xlf["ret"]
        xlr["options"].update(kwargs)
        return f

    return inner


def xlarg(arg: str, convert: Any = None, **kwargs: Any) -> Callable[[_F], _F]:
    if convert is not None:
        kwargs["convert"] = convert

    def inner(f: _F) -> _F:
        xlf = xlfunc(f).__xlfunc__
        if arg.lstrip("*") not in xlf["argmap"]:
            raise Exception(f"Invalid argument name '{arg}'.")
        xla = xlf["argmap"][arg.lstrip("*")]
        if "doc" in kwargs:
            xla["doc"] = kwargs.pop("doc")
        xla["options"].update(kwargs)
        return f

    return inner


def js_to_none(arg):
    # Pyodide >= 0.28 surfaces JS `null` as a distinct `JsNull` sentinel rather
    # than Python `None`, which breaks `arg is None` checks (e.g. empty Excel
    # cells no longer trigger optional-argument defaults). Normalize it back.
    try:
        from pyodide.ffi import JsNull

        if isinstance(arg, JsNull):
            return None
    except ImportError:
        pass
    return arg


def to_scalar(arg):
    if isinstance(arg, (list, tuple)) and len(arg) == 1:
        if isinstance(arg[0], (list, tuple)) and len(arg[0]) == 1:
            arg = arg[0][0]
        else:
            arg = arg[0]
    return js_to_none(arg)


date_format_language_map = {
    # This is currently missing unusual locales such as de-IT (gg.mm.aaaa) but it's
    # covering native locales and those starting with en-, which probably covers 99%
    # of use cases. If needed, specific locales such as de-IT can always be added.
    "cs": {"r": "y"},
    "da": {"å": "y"},
    "de": {"j": "y", "t": "d"},
    "el": {"ε": "y", "μ": "m", "η": "d"},
    "en-at": {"j": "y", "t": "d"},
    "en-de": {"j": "y", "t": "d"},
    "en-dk": {"å": "y"},
    "en-fi": {"v": "y", "k": "m", "p": "d"},
    "en-nl": {"j": "y"},
    "en-se": {"å": "y"},
    "es": {"a": "y"},
    "fi": {"v": "y", "k": "m", "p": "d"},
    "fr": {"a": "y", "j": "d"},
    "hu": {"é": "y", "h": "m", "n": "d"},
    "it": {"a": "y", "g": "d"},
    "nb": {"å": "y"},
    "nl": {"j": "y"},
    "pl": {"r": "y"},
    "pt": {"a": "y"},
    "ru": {"г": "y", "м": "m", "д": "d"},
    "sv": {"å": "y"},
    "tr": {"a": "m", "g": "d"},
}


async def convert(result, ret_info, data):
    options = ret_info["options"].copy()
    date_format = (
        options.get("date_format")  # @ret decorator
        or os.getenv("XLWINGS_DATE_FORMAT")  # env var
        or data.get("date_format")  # Excel cultureInfo
    )

    # Handle international locales, which are completely inconsistent. Examples:
    # en-DE: TT/MM/JJJJ
    # de-DE: TT.MM.JJJJ
    # en-CH: dd.mm.yyyy
    # de-CH: TT.MM.JJJJ
    #
    # The main issue is that Office.js delivers date_format a.k.a
    # context.application.cultureInfo.datetimeFormat.shortDatePattern
    # sometimes in a localized version, which in turn isn't accepted when setting the
    # values. To change the default datetime format for Excel:
    # WIN: Windows Settings > Time & Language > Language & Region > Regional Format.
    # Note that the available selection depends on the added languages under Language.
    # MAC: Mac System Settings > Language & Region. Select Microsoft Excel under
    # Applications.
    # WEB: File > Options > Regional Format Settings
    if date_format and data.get("culture_info_name"):
        if any(c not in "dmy" for c in date_format.lower() if c.isalpha()):
            locale = data["culture_info_name"]
            replacements = date_format_language_map.get(locale.lower())

            if replacements is None:
                language = locale.split("-")[0]
                replacements = date_format_language_map.get(language.lower())

            if replacements:
                for old, new in replacements.items():
                    date_format = date_format.lower().replace(old, new)
            else:
                date_format = None

    options.update({"date_format": date_format, "runtime": data["runtime"]})
    result = await conversion.async_write(result, None, options, engine_name="officejs")
    return result


def provide_values_for_special_args(func, args, typehint_to_value: dict) -> tuple:
    """Inject framework-provided values (e.g. CurrentUser, xw.Book) into the call.

    Returns (args, kwargs): params declared before *args are inserted positionally at
    their signature index, while keyword-only params (i.e. those after *args) are
    returned as kwargs since they can't be passed positionally.
    """
    if typehint_to_value is None:
        typehint_to_value = {}

    type_hints = get_type_hints(func)
    parameters = inspect.signature(func).parameters
    args_list = list(args)
    kwargs = {}
    for index, (param, spec) in enumerate(parameters.items()):
        # Unwrap Optional[X] so that `user: CurrentUser = None` resolves to the same
        # key as `user: CurrentUser` (see is_injectable_typehint).
        hint = _unwrap_optional_hint(type_hints.get(param))
        try:
            value_provided = hint in typehint_to_value
        except TypeError:
            value_provided = False
        if not value_provided:
            continue
        if spec.kind is inspect.Parameter.KEYWORD_ONLY:
            kwargs[param] = typehint_to_value[hint]
        else:
            args_list.insert(index, typehint_to_value[hint])
    return tuple(args_list), kwargs


async def check_user_roles(current_user, required_roles):
    has_required_roles = await current_user.has_required_roles(required_roles)
    if not has_required_roles:
        error_message = (
            f"Access Denied. {current_user.name} is missing the following roles: "
            f"{', '.join(set(required_roles).difference(current_user.roles))}"
        )
        logger.error(error_message)
        raise XlwingsError(error_message)


async def custom_functions_call(
    data,
    module,
    current_user=None,
    sio=None,
    typehint_to_value: dict = None,
    streaming_callback=None,
    streaming_context=None,
):
    """
    sio : socketio.AsyncServer instance
    streaming_callback : callable, used by Lite/Pyodide to push streaming results directly
    streaming_context : context manager applied inside the streaming task (e.g., stdout redirect)
    """
    func_name = data["func_name"]
    args = data["args"]
    func = getattr(module, func_name)
    func_info = func.__xlfunc__
    args_info = func_info["args"]
    ret_info = func_info["ret"]
    required_roles = func_info["required_roles"]

    # Compute the producer discriminator only for handle-producing calls that carry a
    # caller address - the same guard as the evict_superseded() call below - to avoid the
    # JSON+hash over the raw args on every (typically non-producing) custom function call.
    # Captured here from the raw args before they're mutated below (varargs flattened,
    # scalars converted, object handles resolved): distinguishes handle-producing calls
    # that share a caller address, e.g. the two MAKE calls in =CONSUME(MAKE("a"), MAKE("b")).
    caller_address = data.get("caller_address")
    produces_handles = (
        ret_info["options"].get("convert") in object_handles.CONVERTER_KEYS
    )
    producer_scope = (
        object_handles.producer_discriminator(func_name, args)
        if caller_address and produces_handles
        else None
    )

    if current_user:
        await check_user_roles(current_user, required_roles)

    if data["version"] != __version__ and data["client"] != "Office.js":
        raise XlwingsError(
            f"xlwings version mismatch (client: {data['version']} backend: {__version__}): please restart Excel or "
            "right-click on the task pane and select 'reload'!"
        )

    # Turn varargs into regular arguments
    args = list(args)
    new_args = []
    new_args_info = []
    for i, arg in enumerate(args):
        arg_info = args_info[min(i, len(args_info) - 1)]
        if arg_info["vararg"]:
            new_args.extend(arg)
            for _ in range(len(arg)):
                new_args_info.append(arg_info)
        else:
            new_args.append(arg)
            new_args_info.append(arg_info)
    args = new_args
    args_info = new_args_info

    for i, arg in enumerate(args):
        arg_info = args_info[i]
        arg = to_scalar(arg)
        if arg is None:
            args[i] = arg_info.get("optional", None)
        else:
            args[i] = conversion.read(
                None, arg, arg_info["options"], engine_name="officejs"
            )

    # Handle function args that are provided behind the scenes and not via Excel
    args, kwargs = provide_values_for_special_args(func, args, typehint_to_value)

    if inspect.isasyncgenfunction(func):
        # Streaming functions
        task_key = data["task_key"]

        async def task():
            ctx = streaming_context or contextlib.nullcontext()
            with ctx:
                try:
                    async for result in func(*args, **kwargs):
                        result = await convert(result, ret_info, data)
                        if streaming_callback:
                            streaming_callback(result)
                        else:
                            await sio.emit(
                                f"xlwings:set-result-{task_key}",
                                {"result": result},
                            )
                except Exception as e:  # noqa: E722
                    error_result = [[f"ERROR: {repr(e)}"]]
                    if streaming_callback:
                        streaming_callback(error_result)
                        logger.exception(f"Error in custom function '{func_name}'")
                    else:
                        await sio.emit(
                            f"xlwings:set-result-{task_key}",
                            {"result": error_result},
                        )
                        logger.exception(f"Error in custom function '{func_name}'")
                        raise

        # For xlwings Lite (streaming_callback), always restart the task since
        # re-registration invalidates the old invocation/callback.
        if task_key in background_tasks and streaming_callback:
            old_task = background_tasks.pop(task_key)
            old_task.cancel()
            with contextlib.suppress(asyncio.CancelledError):
                await old_task
        if task_key not in background_tasks:
            mytask = asyncio.create_task(task(), name=f"xlwings-{task_key}")
            background_tasks[task_key] = mytask
            logger.info(f"[streaming] created task: {mytask.get_name()}")

            def on_task_done(t):
                logger.info(
                    f"[streaming] task done: {t.get_name()}, cancelled={t.cancelled()}, exception={t.exception() if not t.cancelled() else 'N/A'}"
                )
                if not t.cancelled() and t.exception() is not None:
                    logger.info(
                        f"Task {t.get_name()} failed with exception: {t.exception()}"
                    )
                # Only clear the bookkeeping if it still points at THIS task.
                # add_done_callback fires via loop.call_soon, so on a restart
                # (e.g. full recalc) the cancelled old task's callback can run
                # *after* the new task has already registered itself under the
                # same key - an unconditional pop would untrack the live new
                # task, orphaning it and breaking subsequent restarts.
                if background_tasks.get(task_key) is t:
                    background_tasks.pop(task_key, None)
                    task_key_to_sid_counts.pop(task_key, None)
                    task_key_to_task.pop(task_key, None)

            mytask.add_done_callback(on_task_done)
            return mytask
        else:
            logger.info(f"Reusing existing stream for task key '{task_key}'")
            return

    elif inspect.iscoroutinefunction(func):
        ret = await func(*args, **kwargs)
    else:
        ret = func(*args, **kwargs)

    ret = await convert(ret, ret_info, data)
    if caller_address and produces_handles:
        # Deterministically drop the object-handle entries that this cell's previous
        # invocation wrote. Only handle-producing functions are tracked - for any other
        # function this would be a pointless producer-map lookup on every call (a Redis
        # round trip in xlwings Server). The price: a cell whose formula changes from a
        # producing to a non-producing function keeps its last generation until LRU
        # eviction/expiry - the same backstop that covers deleted formulas (which never
        # trigger a call) and streaming functions (which carry no caller address).
        object_handles.evict_superseded(
            caller_address,
            ret,
            user_id=getattr(current_user, "id", None),
            session_id=data.get("session_id"),
            discriminator=producer_scope,
        )
    return ret


def custom_functions_code(
    module, custom_functions_call_path="/xlwings/custom-functions-call"
):
    js = (Path(__file__).parent / "custom_functions_code.js").read_text()
    # format string would require to double all curly braces
    js = js.replace("placeholder_xlwings_version", __version__).replace(
        "placeholder_custom_functions_call_path", custom_functions_call_path
    )
    for name, obj in inspect.getmembers(module):
        if hasattr(obj, "__xlfunc__"):
            xlfunc = obj.__xlfunc__
            func_name = xlfunc["name"]
            streaming = "true" if inspect.isasyncgenfunction(obj) else "false"
            js += dedent(
                f"""\
            async function {func_name}() {{
                let args = ["{func_name}", {streaming}]
                args.push.apply(args, arguments);
                return await base.apply(null, args);
            }}
            CustomFunctions.associate("{func_name.upper()}", {func_name});
            """
            )
    return js


def custom_functions_meta(module, typehinted_params_to_exclude=None):
    # Kept for backwards compatibility: injectable typehints are normally registered via
    # register_injectable_typehint() and are then already absent from xlfunc["args"].
    if typehinted_params_to_exclude is None:
        typehinted_params_to_exclude = []
    funcs = []
    enums = []
    for name, obj in inspect.getmembers(module):
        if hasattr(obj, "__xlfunc__"):
            xlfunc = obj.__xlfunc__
            func = {}
            func["description"] = xlfunc["ret"]["doc"]
            if xlfunc["help_url"]:
                func["helpUrl"] = xlfunc["help_url"]
            func["id"] = xlfunc["name"].upper()
            display_name = xlfunc.get("excel_name") or xlfunc["name"].upper()
            namespace = get_custom_function_namespace(obj, module)
            if namespace:
                func["name"] = f"{namespace.upper()}.{display_name}"
            else:
                func["name"] = display_name
            if inspect.isasyncgenfunction(obj):
                func["options"] = {
                    "stream": True,
                }
            else:
                func["options"] = {
                    "requiresAddress": True,
                    "requiresParameterAddresses": True,
                }
            if xlfunc["volatile"]:
                func["options"]["volatile"] = True
            func["result"] = {"dimensionality": "matrix", "type": "any"}

            type_hints = get_type_hints(obj)
            params = []
            for arg in xlfunc["args"]:
                if (
                    arg["name"] in type_hints
                    and type_hints[arg["name"]] in typehinted_params_to_exclude
                ):
                    continue
                param = {}
                param["description"] = arg["doc"]
                param["name"] = arg["name"]
                enum = arg["options"].get("enum")
                if enum is not None and enum["type"] == "string":
                    param["dimensionality"] = "scalar"
                    param["type"] = "string"
                    param["customEnumId"] = enum["id"]
                    enums.append(
                        {
                            "id": enum["id"],
                            "type": "string",
                            "values": [
                                _enum_value_entry("string", v, enum["tooltips"])
                                for v in enum["values"]
                            ],
                        }
                    )
                else:
                    param["dimensionality"] = "matrix"
                    param["type"] = "any"
                if "optional" in arg:
                    param["optional"] = True
                elif arg["vararg"]:
                    param["repeating"] = True
                params.append(param)
            func["parameters"] = params
            funcs.append(func)
    # With `name=` aliases, two functions with different ids can end up with the
    # same Excel-facing name, which would be ambiguous in Excel
    seen_names = set()
    for func in funcs:
        if func["name"].upper() in seen_names:
            raise XlwingsError(f"Duplicate custom function name: '{func['name']}'")
        seen_names.add(func["name"].upper())
    result = {
        "allowCustomDataForDataTypeAny": True,
        "allowErrorForDataTypeAny": True,
        "functions": funcs,
    }
    if enums:
        result["enums"] = enums
    return result


def _enum_value_entry(enum_type, value, tooltips):
    entry = {"name": str(value)}
    if enum_type == "string":
        entry["stringValue"] = value
    else:
        entry["numberValue"] = value
    if value in tooltips:
        entry["tooltip"] = tooltips[value]
    return entry


# Custom scripts
def _is_book_hint(hint) -> bool:
    """True if a type hint refers to the injected book (xw.Book or xw.BookAsync).

    xw.BookAsync is a Book subclass, so `hint == xw.Book` alone would miss it.
    Optional[X] is unwrapped so that a `book: xw.Book = None` parameter is still
    recognized as the book - the injection path accepts it, so every detection
    path has to agree (else the script runs but the book isn't found afterwards).
    """
    hint = _unwrap_optional_hint(hint)
    return hint is xw.Book or hint is xw.BookAsync


def _normalize_book_hint(hint):
    """Return (lookup_hint, is_async_book) for an injected-value lookup.

    BookAsync is a Book subclass and a type hint for the async API; the caller keys
    the injected book under xw.Book, so normalize before the lookup. Unhashable
    hints (e.g. some parametrized generics) are never injectable, so map them to
    None to keep the `in typehint_to_value` test from raising.

    Optional[X] is unwrapped first: a `= None` default makes get_type_hints() report
    Optional[X] rather than X (see is_injectable_typehint).
    """
    hint = _unwrap_optional_hint(hint)
    if hint is xw.BookAsync:
        return xw.Book, True
    try:
        hash(hint)
    except TypeError:
        return None, False
    return hint, False


def _inject_value(hint, typehint_to_value):
    """Return the framework-provided value for an injected parameter."""
    lookup_hint, book_is_async = _normalize_book_hint(hint)
    value = typehint_to_value[lookup_hint]
    # A BookAsync annotation makes the injected book lazy: no cell values were
    # pre-loaded, so sync `.value` reads raise until values are loaded (see
    # Range.raw_value in the remote backend). The book is constructed by the caller
    # (e.g. xw.Book(json=...) in xlwings Lite), which can't know the annotation, so
    # we set it here.
    if book_is_async:
        value.impl._lazy = True
    return value


def _book_param_hint(func):
    """Return the type hint of the injected book *parameter*, or None if absent.

    Only parameter annotations are considered — never the return annotation or
    unrelated parameter names. A script must have exactly one Book/BookAsync
    parameter (the injected workbook), so more than one is an error rather than
    an ambiguous guess. Zero returns None: the sync path stays the default and
    the existing runtime raises "No xlwings.Book found" at call time.
    """
    try:
        type_hints = get_type_hints(inspect.unwrap(func))
        params = inspect.signature(func).parameters
    except Exception:
        # A bad/forward-ref annotation shouldn't crash decoration; the caller
        # falls back to sync (BookAsync just won't be auto-detected).
        return None
    # Unwrapped, so that callers can compare the result against xw.Book/xw.BookAsync
    # by identity even when the parameter is annotated Optional[...] (= None).
    book_hints = [
        _unwrap_optional_hint(type_hints[pname])
        for pname in params
        if _is_book_hint(type_hints.get(pname))
    ]
    if len(book_hints) > 1:
        raise XlwingsError(
            "@script functions must have exactly one parameter annotated "
            "'xw.Book' or 'xw.BookAsync' (the injected workbook); found "
            f"{len(book_hints)}."
        )
    return book_hints[0] if book_hints else None


def _script_uses_book_async(func) -> bool:
    """True if the script's injected book parameter is annotated `xw.BookAsync`."""
    return _book_param_hint(func) is xw.BookAsync


@overload
def script(f: _F) -> _F:
    ...


@overload
def script(
    f: None = ...,
    name: str | None = ...,
    required_roles: list[str] | None = ...,
    include: str | list[str] | None = ...,
    exclude: str | list[str] | None = ...,
    button: str | None = ...,
    show_taskpane: bool | None = ...,
    **kwargs: Any,
) -> Callable[[_F], _F]:
    ...


def script(
    f: Callable[..., Any] | None = None,
    name: str | None = None,
    required_roles: list[str] | None = None,
    include: str | list[str] | None = None,
    exclude: str | list[str] | None = None,
    button: str | None = None,
    show_taskpane: bool | None = None,
    **kwargs: Any,
) -> Any:
    # Opt into the async, on-demand API by annotating the book parameter
    # `book: xw.BookAsync` (see the "Async API" docs in xlwings Lite). The choice
    # sits on the parameter it affects, and type checkers see it.
    #
    # `lazy=` is a deprecated alias that does the same
    # thing at the decorator level. Internally we always emit the `"lazy"`
    # metadata key the JS side keys off, so the wire protocol is unchanged.
    # args_lazy is None when `lazy=` wasn't passed, letting the BookAsync
    # annotation decide.
    if "lazy" in kwargs:
        warnings.warn(
            "The 'lazy' argument of @script is deprecated; annotate the book "
            "parameter with 'xw.BookAsync' instead (equivalent to the old "
            "lazy=True).",
            UserWarning,
            stacklevel=2,
        )
        lazy = kwargs.pop("lazy")
        if not isinstance(lazy, bool):
            raise XlwingsError(
                f"The 'lazy' argument of @script must be a boolean, not {lazy!r}."
            )
        args_lazy = lazy
    else:
        args_lazy = None  # nothing specified — the annotation decides

    def inner(func):
        # BookAsync on the book parameter is the way to opt into the async API.
        # It takes precedence, but must not contradict a deprecated lazy=False.
        annotation_lazy = _script_uses_book_async(func)
        if annotation_lazy and args_lazy is False:
            raise XlwingsError(
                "@script: the book parameter is annotated 'xw.BookAsync' (async "
                "API) but lazy=False was passed. Drop the conflicting argument."
            )
        # Precedence: deprecated lazy= if given, else the annotation, else sync.
        lazy_value = args_lazy if args_lazy is not None else annotation_lazy

        @wraps(func)
        async def wrapper(*args, **kwargs):
            # Remove the first arg and assign it to current_user
            current_user, *args = args
            if current_user:
                await check_user_roles(current_user, required_roles)
            if inspect.iscoroutinefunction(func):
                await func(*args, **kwargs)
            else:
                func(*args, **kwargs)

            type_hints = get_type_hints(func)
            sig = inspect.signature(func)

            # Keyword-only params (e.g. an injected value after *args) arrive via
            # kwargs, so check both when looking for the book to return.
            for param_name, arg_value in zip(sig.parameters.keys(), args):
                if param_name in type_hints and _is_book_hint(type_hints[param_name]):
                    return arg_value
            for param_name, arg_value in kwargs.items():
                if param_name in type_hints and _is_book_hint(type_hints[param_name]):
                    return arg_value

            raise XlwingsError("No xlwings.Book found in your function arguments!")

        wrapper.__xlscript__ = {
            "name": name,
            "required_roles": required_roles,
            "include": include,
            "exclude": exclude,
            # target_cell is deprecated
            "button": button or kwargs.get("target_cell"),
            "show_taskpane": show_taskpane,
            # Internal wire key consumed by the JS side; BookAsync -> lazy=True.
            "lazy": lazy_value,
        }
        wrapper.__xlscript__.update(kwargs)

        # For backward compatibility with deprecated 'config' parameter
        if "config" in kwargs and isinstance(kwargs["config"], dict):
            wrapper.__xlscript__.update(kwargs["config"])

        return wrapper

    if f is None:
        return inner
    else:
        return inner(f)


def _coerce_script_arg(value, hint, script_name, param_name):
    """Convert a JSON script argument to the type its hint asks for.

    JSON has no date type, so a `dt.date`/`dt.datetime` parameter would
    otherwise receive an ISO string while the same hint on a custom function
    yields a real date object. Only values that came over the wire pass through
    here; Python defaults are already the right type.
    """
    if hint is None or not isinstance(value, str):
        return value
    # An optional date still gets a date control in the UI, so coerce through
    # the union rather than leaving `Optional[date]` as a string.
    hint = _unwrap_optional_hint(hint)
    if hint is dt.datetime:
        parse, expected = dt.datetime.fromisoformat, "a date and time"
    elif hint is dt.date:
        parse, expected = dt.date.fromisoformat, "a date"
    else:
        return value
    try:
        return parse(value)
    except ValueError:
        raise XlwingsError(
            f"Script '{script_name}': argument '{param_name}' must be "
            f"{expected} in ISO format, got {value!r}"
        ) from None


async def custom_scripts_call(
    module, script_name, current_user=None, typehint_to_value: dict = None, args=None
):
    if typehint_to_value is None:
        typehint_to_value = {}
    if args is None:
        args = []
    func = getattr(module, script_name)

    # Get the function signature
    sig = inspect.signature(func)
    # Resolve type hints so `from __future__ import annotations` works
    resolved_hints = get_type_hints(inspect.unwrap(func))
    # Prepend current_user, which will be removed again by the script decorator
    call_args = [current_user]
    call_kwargs = {}

    # Iterate over the parameters and check their type hints
    arg_iter = iter(args)
    for param in sig.parameters.values():
        hint = resolved_hints.get(param.name)
        injectable = _normalize_book_hint(hint)[0] in typehint_to_value
        if param.kind is inspect.Parameter.KEYWORD_ONLY and injectable:
            # Args arrive positionally, so a keyword-only param can never be filled
            # by the caller - but a framework-injected one (e.g. CurrentUser) is
            # provided here and may therefore sit after *args.
            call_kwargs[param.name] = _inject_value(hint, typehint_to_value)
            continue
        if param.kind in (
            inspect.Parameter.KEYWORD_ONLY,
            inspect.Parameter.VAR_KEYWORD,
        ):
            raise XlwingsError(
                f"Script '{script_name}': keyword-only and **kwargs parameters "
                f"are not supported ('{param.name}')"
            )
        if injectable:
            call_args.append(_inject_value(hint, typehint_to_value))
        elif param.kind == inspect.Parameter.VAR_POSITIONAL:
            call_args.extend(
                _coerce_script_arg(value, hint, script_name, param.name)
                for value in arg_iter
            )
        else:
            try:
                call_args.append(
                    _coerce_script_arg(next(arg_iter), hint, script_name, param.name)
                )
            except StopIteration:
                if param.default is not inspect.Parameter.empty:
                    call_args.append(param.default)
                else:
                    raise XlwingsError(
                        f"Script '{script_name}': missing required argument "
                        f"'{param.name}'"
                    )
    leftover = list(arg_iter)
    if leftover:
        raise XlwingsError(
            f"Script '{script_name}' received {len(leftover)} extra argument(s)"
        )

    if inspect.iscoroutinefunction(func):
        book = await func(*call_args, **call_kwargs)
    else:
        book = func(*call_args, **call_kwargs)

    return book


def custom_scripts_meta(module):
    scripts_meta = []
    for name, func in inspect.getmembers(module, inspect.isfunction):
        meta = getattr(func, "__xlscript__", None)
        if meta:
            script_entry = {"function_name": name}
            if isinstance(meta, dict):
                # Allow include/exclude to be delivered as list
                meta_copy = meta.copy()
                for key in ["include", "exclude"]:
                    if key in meta_copy and isinstance(meta_copy[key], list):
                        meta_copy[key] = ",".join(meta_copy[key])
                script_entry.update(meta_copy)
            scripts_meta.append(script_entry)
    return scripts_meta


# Socket.io (sid is the session ID)
# task_key_to_sid_counts: task_key -> {sid: subscription_count}
task_key_to_sid_counts = {}
task_key_to_task = {}


async def sio_connect(sid, environ, auth, sio, authenticate=None):
    auth = auth if isinstance(auth, dict) else {}
    token = auth.get("token")
    if authenticate:
        try:
            if inspect.iscoroutinefunction(authenticate):
                current_user = await authenticate(token)
            else:
                current_user = authenticate(token)
            logger.info(f"Socket.io: connect {sid}")
            logger.info(f"Socket.io: User authenticated {current_user.name}")
        except Exception as e:
            logger.info(f"Socket.io: authentication failed for sid {sid}: {repr(e)}")
            await sio.disconnect(sid)
            return
    else:
        logger.info(f"Socket.io: connect {sid}")


async def sio_disconnect(sid):
    logger.info(f"disconnect {sid}")
    # Using list() to prevent the loop from changing the dict directly
    for task_key in list(task_key_to_sid_counts.keys()):
        sid_counts = task_key_to_sid_counts.get(task_key)
        if sid_counts is None:
            continue
        sid_counts.pop(sid, None)
        if not sid_counts:
            task = task_key_to_task.get(task_key)
            if task:
                task.cancel()
                logger.info(f"Cancelled task {task.get_name()}")
            task_key_to_sid_counts.pop(task_key, None)
            task_key_to_task.pop(task_key, None)
    await asyncio.sleep(0)  # Allow event loop to cancel the tasks
    active_tasks = [
        task.get_name()
        for task in asyncio.all_tasks()
        if task.get_name().startswith("xlwings")
    ]
    logger.info(f"Active xlwings tasks: {active_tasks}")


async def sio_custom_function_call(
    sid, data, custom_functions, current_user, sio, typehint_to_value: dict = None
):
    if typehint_to_value is None:
        typehint_to_value = {}
    task_key = data["task_key"]
    sid_counts = task_key_to_sid_counts.setdefault(task_key, {})
    sid_counts[sid] = sid_counts.get(sid, 0) + 1
    try:
        task = await custom_functions_call(
            data, custom_functions, current_user, sio, typehint_to_value
        )
    except Exception:
        sid_counts[sid] -= 1
        if sid_counts[sid] <= 0:
            sid_counts.pop(sid, None)
        if not sid_counts:
            task_key_to_sid_counts.pop(task_key, None)
        raise
    if task:
        task_key_to_task[task_key] = task


async def sio_cancel_task(sid, task_key):
    sid_counts = task_key_to_sid_counts.get(task_key)
    if sid_counts is None:
        return
    count = sid_counts.get(sid, 0)
    if count > 1:
        sid_counts[sid] = count - 1
    else:
        sid_counts.pop(sid, None)
    if not sid_counts:
        task = task_key_to_task.get(task_key)
        if task:
            task.cancel()
            logger.info(f"Cancelled task {task.get_name()}")
        task_key_to_sid_counts.pop(task_key, None)
        task_key_to_task.pop(task_key, None)
