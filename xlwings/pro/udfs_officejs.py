"""
Required Notice: Copyright (C) Zoomer Analytics GmbH.

xlwings PRO is dual-licensed under one of the following licenses:

* PolyForm Noncommercial License 1.0.0 (for noncommercial use):
  https://polyformproject.org/licenses/noncommercial/1.0.0
* xlwings PRO License (for commercial use):
  https://github.com/xlwings/xlwings/blob/main/LICENSE_PRO.txt

Commercial licenses can be purchased at https://www.xlwings.org
"""

import asyncio
import inspect
import logging
import os
from functools import wraps
from pathlib import Path
from textwrap import dedent
from typing import Annotated, get_args, get_origin, get_type_hints

import xlwings as xw

from .. import XlwingsError, __version__, conversion

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
        return top_level_type, annotations
    else:
        top_level_type = origin or type_hint
        return top_level_type, []


def xlfunc(f=None, **kwargs):
    def inner(f):
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
                    type_hint, annotations = extract_type_and_annotations(
                        type_hints[var_name]
                    )
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
        return f

    if f is None:
        return inner
    else:
        return inner(f)


def xlret(convert=None, **kwargs):
    if convert is not None:
        kwargs["convert"] = convert

    def inner(f):
        xlf = xlfunc(f).__xlfunc__
        xlr = xlf["ret"]
        xlr["options"].update(kwargs)
        return f

    return inner


def xlarg(arg, convert=None, **kwargs):
    if convert is not None:
        kwargs["convert"] = convert

    def inner(f):
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


def convert(result, ret_info, data):
    if "date_format" not in ret_info["options"]:
        date_format = os.getenv("XLWINGS_DATE_FORMAT") or locale_to_shortdate.get(
            data["content_language"].lower()
        )
        if date_format is None:
            logger.warning(
                f"Locale {data['content_language'].lower()} not found, so custom "
                "functions won't format dates automatically. Please open an issue with "
                "this warning on https://github.com/xlwings/xlwings/issues. In the "
                "meantime, you can set the XLWINGS_DATE_FORMAT env var to fix that."
            )
        ret_info["options"]["date_format"] = date_format
    ret_info["options"]["runtime"] = data["runtime"]
    result = conversion.write(result, None, ret_info["options"], engine_name="officejs")
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


async def custom_functions_call(data, module, sio=None, typehint_to_value: dict = None):
    """
    sio : socketio.AsyncServer instance
    """
    func_name = data["func_name"]
    args = data["args"]
    func = getattr(module, func_name)
    func_info = func.__xlfunc__
    args_info = func_info["args"]
    ret_info = func_info["ret"]

    if data["version"] != __version__:
        raise XlwingsError(
            "xlwings version mismatch: please restart Excel or "
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
            try:
                async for result in func(*args):
                    result = convert(result, ret_info, data)
                    await sio.emit(
                        f"xlwings:set-result-{task_key}",
                        {"result": result},
                    )
            except Exception as e:  # noqa: E722
                await sio.emit(
                    f"xlwings:set-result-{task_key}",
                    {"result": [[f"ERROR: {repr(e)}"]]},
                )
                logger.exception(f"Error in custom function '{func_name}'")
                raise

        if task_key not in background_tasks:
            mytask = asyncio.create_task(task(), name=f"xlwings-{task_key}")
            background_tasks[task_key] = mytask

            def on_task_done(t):
                if not t.cancelled() and t.exception() is not None:
                    t.cancel()
                    logger.info(
                        f"Task {t.get_name()} cancelled as it failed with exception: {t.exception()}"
                    )
                del background_tasks[task_key]

            mytask.add_done_callback(on_task_done)
            return mytask
        else:
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
    funcs = []
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
                param["dimensionality"] = "matrix"
                param["type"] = "any"
                if "optional" in arg:
                    param["optional"] = True
                elif arg["vararg"]:
                    param["repeating"] = True
                params.append(param)
            func["parameters"] = params
            funcs.append(func)
    return {
        "allowCustomDataForDataTypeAny": True,
        "allowErrorForDataTypeAny": True,
        "functions": funcs,
    }


# Custom scripts
def script(func):
    @wraps(func)
    async def wrapper(*args, **kwargs):
        if inspect.iscoroutinefunction(func):
            await func(*args, **kwargs)
        else:
            func(*args, **kwargs)

        type_hints = get_type_hints(func)
        sig = inspect.signature(func)

        for param_name, arg_value in zip(sig.parameters.keys(), args):
            if param_name in type_hints and type_hints[param_name] == xw.Book:
                return arg_value

        raise XlwingsError("No xw.Book found in your function arguments!")

    return wrapper


async def custom_scripts_call(module, script_name, typehint_to_value: dict = None):
    # Currently, there are no arguments from the client accepted, only internally
    # provided args via type hints are allowed
    func = getattr(module, script_name)

    # Get the function signature
    sig = inspect.signature(func)
    args = []

    # Iterate over the parameters and check their type hints
    for param in sig.parameters.values():
        if param.annotation in typehint_to_value:
            args.append(typehint_to_value[param.annotation])
        else:
            raise XlwingsError(
                "Scripts currently only allow Book and CurrentUser as params"
            )

    if inspect.iscoroutinefunction(func):
        book = await func(*args)
    else:
        book = func(*args)

    return book


# Socket.io (sid is the session ID)
task_key_to_sids = {}
task_key_to_task = {}


async def sio_connect(sid, environ, auth, sio, authenticate=None):
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
    logger.info(f"Socket.io: connect {sid}")


async def sio_disconnect(sid):
    logger.info(f"disconnect {sid}")
    try:
        # Using list() to prevent the loop from changing the dict directly
        for task_key in list(task_key_to_sids.keys()):
            task = task_key_to_task[task_key]
            task_key_to_sids[task_key].discard(sid)
            if not task_key_to_sids[task_key]:
                task.cancel()
                logger.info(f"Cancelled task {task.get_name()}")
                del task_key_to_sids[task_key]
                del task_key_to_task[task_key]
    except KeyError:
        # Renaming functions during development can cause issues
        pass
    await asyncio.sleep(0)  # Allow event loop to cancel the tasks
    active_tasks = [
        task.get_name()
        for task in asyncio.all_tasks()
        if task.get_name().startswith("xlwings")
    ]
    logger.info(f"Active xlwings tasks:" f"{active_tasks}")


async def sio_custom_function_call(
    sid, data, custom_functions, sio, typehint_to_value: dict = None
):
    task_key = data["task_key"]
    task_key_to_sids[task_key] = task_key_to_sids.get(task_key, set()).union({sid})
    task = await custom_functions_call(data, custom_functions, sio, typehint_to_value)
    if task:
        task_key_to_task[task_key] = task


locale_to_shortdate = {
    # This is using the locales from https://github.com/OfficeDev/office-js/tree/release/dist
    # matched with values from https://stackoverflow.com/a/9893752/918626
    # TODO: https://metacpan.org/dist/DateTime-Locale seems to be much better than SO
    # Also, office-js doesn't have all locales supported by Office, such as en-ch or
    # en-ae are missing
    "af-za": "yyyy/mm/dd",
    "am-et": "d/m/yyyy",
    "ar-ae": "dd/mm/yyyy",
    "ar-bh": "dd/mm/yyyy",
    "ar-dz": "dd-mm-yyyy",
    "ar-eg": "dd/mm/yyyy",
    "ar-iq": "dd/mm/yyyy",
    "ar-jo": "dd/mm/yyyy",
    "ar-kw": "dd/mm/yyyy",
    "ar-lb": "dd/mm/yyyy",
    "ar-ly": "dd/mm/yyyy",
    "ar-ma": "dd-mm-yyyy",
    "ar-om": "dd/mm/yyyy",
    "ar-qa": "dd/mm/yyyy",
    "ar-sa": "dd/mm/yy",
    "ar-sy": "dd/mm/yyyy",
    "ar-tn": "dd-mm-yyyy",
    "ar-ye": "dd/mm/yyyy",
    "az-latn-az": "dd.mm.yyyy",
    "be-by": "dd.mm.yyyy",
    "bg-bg": "dd.m.yyyy",
    "bn-in": "dd-mm-yy",
    "bs-latn-ba": "d.m.yyyy",
    "ca-es": "dd/mm/yyyy",
    "cs-cz": "d.m.yyyy",
    "cy-gb": "dd/mm/yyyy",
    "da-dk": "dd-mm-yyyy",
    "de-at": "dd.mm.yyyy",
    "de-ch": "dd.mm.yyyy",
    "de-de": "dd.mm.yyyy",
    "de-li": "dd.mm.yyyy",
    "de-lu": "dd.mm.yyyy",
    "el-gr": "d/m/yyyy",
    "en-029": "mm/dd/yyyy",
    "en-au": "d/mm/yyyy",
    "en-bz": "dd/mm/yyyy",
    "en-ca": "dd/mm/yyyy",
    "en-ch": "dd.mm.yyyy",
    "en-gb": "dd/mm/yyyy",
    "en-ie": "dd/mm/yyyy",
    "en-in": "dd-mm-yyyy",
    "en-jm": "dd/mm/yyyy",
    "en-my": "d/m/yyyy",
    "en-nz": "d/mm/yyyy",
    "en-ph": "m/d/yyyy",
    "en-sg": "d/m/yyyy",
    "en-tt": "dd/mm/yyyy",
    "en-us": "m/d/yyyy",
    "en-za": "yyyy/mm/dd",
    "en-zw": "m/d/yyyy",
    "en-ae": "dd/mm/yyyy",
    "es-ar": "dd/mm/yyyy",
    "es-bo": "dd/mm/yyyy",
    "es-cl": "dd-mm-yyyy",
    "es-co": "dd/mm/yyyy",
    "es-cr": "dd/mm/yyyy",
    "es-do": "dd/mm/yyyy",
    "es-ec": "dd/mm/yyyy",
    "es-es": "dd/mm/yyyy",
    "es-gt": "dd/mm/yyyy",
    "es-hn": "dd/mm/yyyy",
    "es-mx": "dd/mm/yyyy",
    "es-ni": "dd/mm/yyyy",
    "es-pa": "mm/dd/yyyy",
    "es-pe": "dd/mm/yyyy",
    "es-pr": "dd/mm/yyyy",
    "es-py": "dd/mm/yyyy",
    "es-sv": "dd/mm/yyyy",
    "es-us": "m/d/yyyy",
    "es-uy": "dd/mm/yyyy",
    "es-ve": "dd/mm/yyyy",
    "et-ee": "d.mm.yyyy",
    "eu-es": "yyyy/mm/dd",
    "fa-ir": "mm/dd/yyyy",
    "fi-fi": "d.m.yyyy",
    "fil-ph": "m/d/yyyy",
    "fr-be": "d/mm/yyyy",
    "fr-ca": "yyyy-mm-dd",
    "fr-ch": "dd.mm.yyyy",
    "fr-fr": "dd/mm/yyyy",
    "fr-lu": "dd/mm/yyyy",
    "fr-mc": "dd/mm/yyyy",
    "ga-ie": "dd/mm/yyyy",
    "gl-es": "dd/mm/yy",
    "gu-in": "dd-mm-yy",
    "he-il": "dd/mm/yyyy",
    "hi-in": "dd-mm-yyyy",
    "hr-ba": "d.m.yyyy.",
    "hr-hr": "d.m.yyyy",
    "hu-hu": "yyyy. mm. dd.",
    "hy-am": "dd.mm.yyyy",
    "id-id": "dd/mm/yyyy",
    "is-is": "d.m.yyyy",
    "it-ch": "dd.mm.yyyy",
    "it-it": "dd/mm/yyyy",
    "ja-jp": "yyyy/mm/dd",
    "ka-ge": "dd.mm.yyyy",
    "kk-kz": "dd.mm.yyyy",
    "km-kh": "yyyy-mm-dd",
    "kn-in": "dd-mm-yy",
    "ko-kr": "yyyy. mm. dd",
    "lb-lu": "dd/mm/yyyy",
    "lo-la": "dd/mm/yyyy",
    "lt-lt": "yyyy.mm.dd",
    "lv-lv": "yyyy.mm.dd.",
    "mk-mk": "dd.mm.yyyy",
    "ml-in": "dd-mm-yy",
    "mn-mn": "yy.mm.dd",
    "mr-in": "dd-mm-yyyy",
    "ms-bn": "dd/mm/yyyy",
    "ms-my": "dd/mm/yyyy",
    "mt-mt": "dd/mm/yyyy",
    "nb-no": "dd.mm.yyyy",
    "ne-np": "m/d/yyyy",
    "nl-be": "d/mm/yyyy",
    "nl-nl": "d-m-yyyy",
    "nn-no": "dd.mm.yyyy",
    "pl-pl": "dd.mm.yyyy",
    "pt-br": "d/m/yyyy",
    "pt-pt": "dd-mm-yyyy",
    "ro-ro": "dd.mm.yyyy",
    "ru-ru": "dd.mm.yyyy",
    "si-lk": "yyyy-mm-dd",
    "sk-sk": "d. m. yyyy",
    "sl-si": "d.m.yyyy",
    "sq-al": "yyyy-mm-dd",
    "sr-cyrl-cs": "d.m.yyyy",
    "sr-cyrl-rs": "d.m.yyyy",
    "sr-latn-cs": "d.m.yyyy",
    "sr-latn-rs": "d.m.yyyy",
    "sv-fi": "d.m.yyyy",
    "sv-se": "yyyy-mm-dd",
    "sw-ke": "m/d/yyyy",
    "ta-in": "dd-mm-yyyy",
    "te-in": "dd-mm-yy",
    "th-th": "d/m/yyyy",
    "tr-tr": "dd.mm.yyyy",
    "uk-ua": "dd.mm.yyyy",
    "ur-pk": "dd/mm/yyyy",
    "vi-vn": "dd/mm/yyyy",
    "zh-cn": "yyyy/m/d",
    "zh-hk": "d/m/yyyy",
    "zh-mo": "d/m/yyyy",
    "zh-sg": "d/m/yyyy",
    "zh-tw": "yyyy/m/d",
}
