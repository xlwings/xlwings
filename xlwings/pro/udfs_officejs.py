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
import inspect
import logging
import os
from functools import wraps
from pathlib import Path
from textwrap import dedent
from typing import (
    Annotated,
    Any,
    Callable,
    Literal,
    TypeVar,
    get_args,
    get_origin,
    get_type_hints,
    overload,
)

_F = TypeVar("_F", bound=Callable[..., Any])

import xlwings as xw

from .. import ObjectHandle, XlwingsError, __version__, conversion

logger = logging.getLogger(__name__)

# Tasks started by streaming functions
background_tasks = {}


def func_sig(f):
    sig = inspect.signature(f)
    vararg = None
    args = []
    defaults = []
    for param in sig.parameters.values():
        if param.kind is inspect.Parameter.POSITIONAL_OR_KEYWORD:
            args.append(param.name)
            if param.default is not inspect.Signature.empty:
                defaults.append(param.default)
        elif param.kind is inspect.Parameter.VAR_POSITIONAL:
            args.append(param.name)
            vararg = param.name
        else:
            raise XlwingsError("xlwings does not support UDFs with keyword arguments")
    return {"args": args, "defaults": defaults, "vararg": vararg}


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
        return top_level_type, annotations
    else:
        top_level_type = origin or type_hint
        return top_level_type, []


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


def to_scalar(arg):
    if isinstance(arg, (list, tuple)) and len(arg) == 1:
        if isinstance(arg[0], (list, tuple)) and len(arg[0]) == 1:
            arg = arg[0][0]
        else:
            arg = arg[0]
    return arg


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


def convert(result, ret_info, data):
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
    result = conversion.write(result, None, options, engine_name="officejs")
    return result


def provide_values_for_special_args(func, args, typehint_to_value: dict) -> tuple:
    if typehint_to_value is None:
        typehint_to_value = {}

    type_hints = get_type_hints(func)
    args_list = list(args)
    for param, hint in type_hints.items():
        if hint in typehint_to_value:
            param_index = list(func.__code__.co_varnames).index(param)
            args_list.insert(param_index, typehint_to_value[hint])
    args = tuple(args_list)
    return args


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
    args = provide_values_for_special_args(func, args, typehint_to_value)

    if inspect.isasyncgenfunction(func):
        # Streaming functions
        task_key = data["task_key"]

        async def task():
            ctx = streaming_context or contextlib.nullcontext()
            with ctx:
                try:
                    async for result in func(*args):
                        result = convert(result, ret_info, data)
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
                background_tasks.pop(task_key, None)
                task_key_to_sid_counts.pop(task_key, None)
                task_key_to_task.pop(task_key, None)

            mytask.add_done_callback(on_task_done)
            return mytask
        else:
            logger.info(f"Reusing existing stream for task key '{task_key}'")
            return

    elif inspect.iscoroutinefunction(func):
        ret = await func(*args)
    else:
        ret = func(*args)

    ret = convert(ret, ret_info, data)
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
            if xlfunc["namespace"]:
                func["name"] = f"{xlfunc['namespace'].upper()}.{xlfunc['name'].upper()}"
            else:
                func["name"] = xlfunc["name"].upper()
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
    lazy: bool = ...,
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
    lazy: bool = False,
    **kwargs: Any,
) -> Any:
    def inner(func):
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

            for param_name, arg_value in zip(sig.parameters.keys(), args):
                if param_name in type_hints and type_hints[param_name] == xw.Book:
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
            "lazy": lazy,
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

    # Iterate over the parameters and check their type hints
    arg_iter = iter(args)
    for param in sig.parameters.values():
        if param.kind in (
            inspect.Parameter.KEYWORD_ONLY,
            inspect.Parameter.VAR_KEYWORD,
        ):
            raise XlwingsError(
                f"Script '{script_name}': keyword-only and **kwargs parameters "
                f"are not supported ('{param.name}')"
            )
        hint = resolved_hints.get(param.name)
        if hint in typehint_to_value:
            call_args.append(typehint_to_value[hint])
        elif param.kind == inspect.Parameter.VAR_POSITIONAL:
            call_args.extend(arg_iter)
        else:
            try:
                call_args.append(next(arg_iter))
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
        book = await func(*call_args)
    else:
        book = func(*call_args)

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
