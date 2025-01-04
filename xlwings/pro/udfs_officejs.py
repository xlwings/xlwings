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
        f.__xlfunc__["required_roles"] = kwargs.get("required_roles")
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
    date_format = (
        ret_info["options"].get("date_format")  # @ret decorator
        or os.getenv("XLWINGS_DATE_FORMAT")  # env var
        or locale_to_shortdate.get(
            data["culture_info_name"]
        )  # Excel cultureInfo (default)
    )
    ret_info["options"].update({"date_format": date_format, "runtime": data["runtime"]})
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
    data, module, current_user=None, sio=None, typehint_to_value: dict = None
):
    """
    sio : socketio.AsyncServer instance
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

    if data["version"] != __version__:
        print(
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
    if typehinted_params_to_exclude is None:
        typehinted_params_to_exclude = []
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
def script(f=None, target_cell=None, config=None, required_roles=None):
    if config is None:
        config = {}

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

            raise XlwingsError("No xw.Book found in your function arguments!")

        wrapper.target_cell = target_cell
        wrapper.config = config
        return wrapper

    if f is None:
        return inner
    else:
        return inner(f)


async def custom_scripts_call(
    module, script_name, current_user=None, typehint_to_value: dict = None
):
    if typehint_to_value is None:
        typehint_to_value = {}
    # Currently, there are no arguments from the client accepted, only internally
    # provided args via type hints are allowed
    func = getattr(module, script_name)

    # Get the function signature
    sig = inspect.signature(func)
    # Prepend current_user, which will be removed again by the script decorator
    args = [current_user]

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
    sid, data, custom_functions, current_user, sio, typehint_to_value: dict = None
):
    if typehint_to_value is None:
        typehint_to_value = {}
    task_key = data["task_key"]
    task_key_to_sids[task_key] = task_key_to_sids.get(task_key, set()).union({sid})
    task = await custom_functions_call(
        data, custom_functions, current_user, sio, typehint_to_value
    )
    if task:
        task_key_to_task[task_key] = task


locale_to_shortdate = {
    # Generated with this PowerShell script
    #
    # $dict = @{}
    # [System.Globalization.CultureInfo]::GetCultures("AllCultures") | ForEach-Object {
    #     $dict[$_.Name] = $_.DateTimeFormat.ShortDatePattern
    # }
    # $sortedDict = $dict.GetEnumerator() | Sort-Object Key | ForEach-Object {
    #     "`"$($_.Key)`": `"$($_.Value)`""
    # }
    # "{`n" + ($sortedDict -join ",`n") + "`n}"
    "": "MM/dd/yyyy",
    "aa": "dd/MM/yyyy",
    "aa-DJ": "dd/MM/yyyy",
    "aa-ER": "dd/MM/yyyy",
    "aa-ET": "dd/MM/yyyy",
    "af": "yyyy-MM-dd",
    "af-NA": "yyyy-MM-dd",
    "af-ZA": "yyyy-MM-dd",
    "agq": "d/M/yyyy",
    "agq-CM": "d/M/yyyy",
    "ak": "yyyy/MM/dd",
    "ak-GH": "yyyy/MM/dd",
    "am": "dd/MM/yyyy",
    "am-ET": "dd/MM/yyyy",
    "ar": "dd/MM/yy",
    "ar-001": "d/M/yyyy",
    "ar-AE": "dd/MM/yyyy",
    "ar-BH": "dd/MM/yyyy",
    "ar-DJ": "d/M/yyyy",
    "ar-DZ": "dd-MM-yyyy",
    "ar-EG": "dd/MM/yyyy",
    "ar-ER": "d/M/yyyy",
    "ar-IL": "d/M/yyyy",
    "ar-IQ": "dd/MM/yyyy",
    "ar-JO": "dd/MM/yyyy",
    "ar-KM": "d/M/yyyy",
    "ar-KW": "dd/MM/yyyy",
    "ar-LB": "dd/MM/yyyy",
    "ar-LY": "dd/MM/yyyy",
    "ar-MA": "dd-MM-yyyy",
    "ar-MR": "d/M/yyyy",
    "arn": "dd-MM-yyyy",
    "arn-CL": "dd-MM-yyyy",
    "ar-OM": "dd/MM/yyyy",
    "ar-PS": "d/M/yyyy",
    "ar-QA": "dd/MM/yyyy",
    "ar-SA": "dd/MM/yy",
    "ar-SD": "d/M/yyyy",
    "ar-SO": "d/M/yyyy",
    "ar-SS": "d/M/yyyy",
    "ar-SY": "dd/MM/yyyy",
    "ar-TD": "d/M/yyyy",
    "ar-TN": "dd-MM-yyyy",
    "ar-YE": "dd/MM/yyyy",
    "as": "dd-MM-yyyy",
    "asa": "dd/MM/yyyy",
    "asa-TZ": "dd/MM/yyyy",
    "as-IN": "dd-MM-yyyy",
    "ast": "d/M/yyyy",
    "ast-ES": "d/M/yyyy",
    "az": "dd.MM.yyyy",
    "az-Cyrl": "dd.MM.yyyy",
    "az-Cyrl-AZ": "dd.MM.yyyy",
    "az-Latn": "dd.MM.yyyy",
    "az-Latn-AZ": "dd.MM.yyyy",
    "ba": "dd.MM.yy",
    "ba-RU": "dd.MM.yy",
    "bas": "d/M/yyyy",
    "bas-CM": "d/M/yyyy",
    "be": "dd.MM.yy",
    "be-BY": "dd.MM.yy",
    "bem": "dd/MM/yyyy",
    "bem-ZM": "dd/MM/yyyy",
    "bez": "dd/MM/yyyy",
    "bez-TZ": "dd/MM/yyyy",
    "bg": "d.M.yyyy 'г.'",
    "bg-BG": "d.M.yyyy 'г.'",
    "bgc": "yyyy-MM-dd",
    "bgc-Deva": "yyyy-MM-dd",
    "bgc-Deva-IN": "yyyy-MM-dd",
    "bho": "yyyy-MM-dd",
    "bho-Deva": "yyyy-MM-dd",
    "bho-Deva-IN": "yyyy-MM-dd",
    "bin": "d/M/yyyy",
    "bin-NG": "d/M/yyyy",
    "bm": "d/M/yyyy",
    "bm-Latn": "d/M/yyyy",
    "bm-Latn-ML": "d/M/yyyy",
    "bn": "d/M/yyyy",
    "bn-BD": "d/M/yyyy",
    "bn-IN": "dd-MM-yy",
    "bo": "yyyy/M/d",
    "bo-CN": "yyyy/M/d",
    "bo-IN": "yyyy-MM-dd",
    "br": "dd/MM/yyyy",
    "br-FR": "dd/MM/yyyy",
    "brx": "dd-MM-yyyy",
    "brx-IN": "dd-MM-yyyy",
    "bs": "d. M. yyyy.",
    "bs-Cyrl": "d.M.yyyy",
    "bs-Cyrl-BA": "d.M.yyyy",
    "bs-Latn": "d. M. yyyy.",
    "bs-Latn-BA": "d. M. yyyy.",
    "byn": "dd/MM/yyyy",
    "byn-ER": "dd/MM/yyyy",
    "ca": "d/M/yyyy",
    "ca-AD": "d/M/yyyy",
    "ca-ES": "d/M/yyyy",
    "ca-ES-valencia": "d/M/yyyy",
    "ca-FR": "d/M/yyyy",
    "ca-IT": "d/M/yyyy",
    "ccp": "d/M/yyyy",
    "ccp-Cakm": "d/M/yyyy",
    "ccp-Cakm-BD": "d/M/yyyy",
    "ccp-Cakm-IN": "d/M/yyyy",
    "ce": "yyyy-MM-dd",
    "ceb": "M/d/yyyy",
    "ceb-Latn": "M/d/yyyy",
    "ceb-Latn-PH": "M/d/yyyy",
    "ce-RU": "yyyy-MM-dd",
    "cgg": "dd/MM/yyyy",
    "cgg-UG": "dd/MM/yyyy",
    "chr": "M/d/yyyy",
    "chr-Cher": "M/d/yyyy",
    "chr-Cher-US": "M/d/yyyy",
    "co": "dd/MM/yyyy",
    "co-FR": "dd/MM/yyyy",
    "cs": "dd.MM.yyyy",
    "cs-CZ": "dd.MM.yyyy",
    "cu": "yyyy.MM.dd",
    "cu-RU": "yyyy.MM.dd",
    "cv": "dd.MM.yyyy",
    "cv-Cyrl": "dd.MM.yyyy",
    "cv-Cyrl-RU": "dd.MM.yyyy",
    "cy": "dd/MM/yyyy",
    "cy-GB": "dd/MM/yyyy",
    "da": "dd-MM-yyyy",
    "da-DK": "dd-MM-yyyy",
    "da-GL": "dd.MM.yyyy",
    "dav": "dd/MM/yyyy",
    "dav-KE": "dd/MM/yyyy",
    "de": "dd.MM.yyyy",
    "de-AT": "dd.MM.yyyy",
    "de-BE": "dd.MM.yyyy",
    "de-CH": "dd.MM.yyyy",
    "de-DE": "dd.MM.yyyy",
    "de-IT": "dd.MM.yyyy",
    "de-LI": "dd.MM.yyyy",
    "de-LU": "dd.MM.yyyy",
    "dje": "d/M/yyyy",
    "dje-NE": "d/M/yyyy",
    "doi": "d/M/yyyy",
    "doi-Deva": "d/M/yyyy",
    "doi-Deva-IN": "d/M/yyyy",
    "dsb": "d. M. yyyy",
    "dsb-DE": "d. M. yyyy",
    "dua": "d/M/yyyy",
    "dua-CM": "d/M/yyyy",
    "dv": "dd/MM/yy",
    "dv-MV": "dd/MM/yy",
    "dyo": "d/M/yyyy",
    "dyo-SN": "d/M/yyyy",
    "dz": "yyyy-MM-dd",
    "dz-BT": "yyyy-MM-dd",
    "ebu": "dd/MM/yyyy",
    "ebu-KE": "dd/MM/yyyy",
    "ee": "M/d/yyyy",
    "ee-GH": "M/d/yyyy",
    "ee-TG": "M/d/yyyy",
    "el": "d/M/yyyy",
    "el-CY": "d/M/yyyy",
    "el-GR": "d/M/yyyy",
    "en": "M/d/yyyy",
    "en-001": "dd/MM/yyyy",
    "en-029": "dd/MM/yyyy",
    "en-150": "dd/MM/yyyy",
    "en-AE": "dd/MM/yyyy",
    "en-AG": "dd/MM/yyyy",
    "en-AI": "dd/MM/yyyy",
    "en-AS": "M/d/yyyy",
    "en-AT": "dd/MM/yyyy",
    "en-AU": "d/MM/yyyy",
    "en-BB": "dd/MM/yyyy",
    "en-BE": "dd/MM/yyyy",
    "en-BI": "M/d/yyyy",
    "en-BM": "dd/MM/yyyy",
    "en-BS": "dd/MM/yyyy",
    "en-BW": "dd/MM/yyyy",
    "en-BZ": "dd/MM/yyyy",
    "en-CA": "yyyy-MM-dd",
    "en-CC": "dd/MM/yyyy",
    "en-CH": "dd.MM.yyyy",
    "en-CK": "dd/MM/yyyy",
    "en-CM": "dd/MM/yyyy",
    "en-CX": "dd/MM/yyyy",
    "en-CY": "dd/MM/yyyy",
    "en-DE": "dd/MM/yyyy",
    "en-DK": "dd/MM/yyyy",
    "en-DM": "dd/MM/yyyy",
    "en-ER": "dd/MM/yyyy",
    "en-FI": "dd/MM/yyyy",
    "en-FJ": "dd/MM/yyyy",
    "en-FK": "dd/MM/yyyy",
    "en-FM": "dd/MM/yyyy",
    "en-GB": "dd/MM/yyyy",
    "en-GD": "dd/MM/yyyy",
    "en-GG": "dd/MM/yyyy",
    "en-GH": "dd/MM/yyyy",
    "en-GI": "dd/MM/yyyy",
    "en-GM": "dd/MM/yyyy",
    "en-GU": "M/d/yyyy",
    "en-GY": "dd/MM/yyyy",
    "en-HK": "d/M/yyyy",
    "en-ID": "dd/MM/yyyy",
    "en-IE": "dd/MM/yyyy",
    "en-IL": "dd/MM/yyyy",
    "en-IM": "dd/MM/yyyy",
    "en-IN": "dd-MM-yyyy",
    "en-IO": "dd/MM/yyyy",
    "en-JE": "dd/MM/yyyy",
    "en-JM": "d/M/yyyy",
    "en-KE": "dd/MM/yyyy",
    "en-KI": "dd/MM/yyyy",
    "en-KN": "dd/MM/yyyy",
    "en-KY": "dd/MM/yyyy",
    "en-LC": "dd/MM/yyyy",
    "en-LR": "dd/MM/yyyy",
    "en-LS": "dd/MM/yyyy",
    "en-MG": "dd/MM/yyyy",
    "en-MH": "M/d/yyyy",
    "en-MO": "dd/MM/yyyy",
    "en-MP": "M/d/yyyy",
    "en-MS": "dd/MM/yyyy",
    "en-MT": "dd/MM/yyyy",
    "en-MU": "dd/MM/yyyy",
    "en-MV": "d-M-yyyy",
    "en-MW": "dd/MM/yyyy",
    "en-MY": "d/M/yyyy",
    "en-NA": "dd/MM/yyyy",
    "en-NF": "dd/MM/yyyy",
    "en-NG": "dd/MM/yyyy",
    "en-NL": "dd/MM/yyyy",
    "en-NR": "dd/MM/yyyy",
    "en-NU": "dd/MM/yyyy",
    "en-NZ": "d/MM/yyyy",
    "en-PG": "dd/MM/yyyy",
    "en-PH": "M/d/yyyy",
    "en-PK": "dd/MM/yyyy",
    "en-PN": "dd/MM/yyyy",
    "en-PR": "M/d/yyyy",
    "en-PW": "dd/MM/yyyy",
    "en-RW": "dd/MM/yyyy",
    "en-SB": "dd/MM/yyyy",
    "en-SC": "dd/MM/yyyy",
    "en-SD": "dd/MM/yyyy",
    "en-SE": "yyyy-MM-dd",
    "en-SG": "d/M/yyyy",
    "en-SH": "dd/MM/yyyy",
    "en-SI": "dd/MM/yyyy",
    "en-SL": "dd/MM/yyyy",
    "en-SS": "dd/MM/yyyy",
    "en-SX": "dd/MM/yyyy",
    "en-SZ": "dd/MM/yyyy",
    "en-TC": "dd/MM/yyyy",
    "en-TK": "dd/MM/yyyy",
    "en-TO": "dd/MM/yyyy",
    "en-TT": "dd/MM/yyyy",
    "en-TV": "dd/MM/yyyy",
    "en-TZ": "dd/MM/yyyy",
    "en-UG": "dd/MM/yyyy",
    "en-UM": "M/d/yyyy",
    "en-US": "M/d/yyyy",
    "en-VC": "dd/MM/yyyy",
    "en-VG": "dd/MM/yyyy",
    "en-VI": "M/d/yyyy",
    "en-VU": "dd/MM/yyyy",
    "en-WS": "dd/MM/yyyy",
    "en-ZA": "yyyy/MM/dd",
    "en-ZM": "dd/MM/yyyy",
    "en-ZW": "d/M/yyyy",
    "eo": "yyyy-MM-dd",
    "eo-001": "yyyy-MM-dd",
    "es": "dd/MM/yyyy",
    "es-419": "d/M/yyyy",
    "es-AR": "d/M/yyyy",
    "es-BO": "d/M/yyyy",
    "es-BR": "d/M/yyyy",
    "es-BZ": "d/M/yyyy",
    "es-CL": "dd-MM-yyyy",
    "es-CO": "d/MM/yyyy",
    "es-CR": "d/M/yyyy",
    "es-CU": "d/M/yyyy",
    "es-DO": "d/M/yyyy",
    "es-EC": "d/M/yyyy",
    "es-ES": "dd/MM/yyyy",
    "es-GQ": "d/M/yyyy",
    "es-GT": "d/MM/yyyy",
    "es-HN": "d/M/yyyy",
    "es-MX": "dd/MM/yyyy",
    "es-NI": "d/M/yyyy",
    "es-PA": "MM/dd/yyyy",
    "es-PE": "d/MM/yyyy",
    "es-PH": "d/M/yyyy",
    "es-PR": "MM/dd/yyyy",
    "es-PY": "d/M/yyyy",
    "es-SV": "d/M/yyyy",
    "es-US": "M/d/yyyy",
    "es-UY": "d/M/yyyy",
    "es-VE": "d/M/yyyy",
    "et": "dd.MM.yyyy",
    "et-EE": "dd.MM.yyyy",
    "eu": "yyyy/M/d",
    "eu-ES": "yyyy/M/d",
    "ewo": "d/M/yyyy",
    "ewo-CM": "d/M/yyyy",
    "fa": "dd/MM/yyyy",
    "fa-AF": "yyyy/M/d",
    "fa-IR": "dd/MM/yyyy",
    "ff": "dd/MM/yyyy",
    "ff-Adlm": "d-M-yyyy",
    "ff-Adlm-BF": "d-M-yyyy",
    "ff-Adlm-CM": "d-M-yyyy",
    "ff-Adlm-GH": "d-M-yyyy",
    "ff-Adlm-GM": "d-M-yyyy",
    "ff-Adlm-GN": "d-M-yyyy",
    "ff-Adlm-GW": "d-M-yyyy",
    "ff-Adlm-LR": "d-M-yyyy",
    "ff-Adlm-MR": "d-M-yyyy",
    "ff-Adlm-NE": "d-M-yyyy",
    "ff-Adlm-NG": "d-M-yyyy",
    "ff-Adlm-SL": "d-M-yyyy",
    "ff-Adlm-SN": "d-M-yyyy",
    "ff-Latn": "dd/MM/yyyy",
    "ff-Latn-BF": "d/M/yyyy",
    "ff-Latn-CM": "d/M/yyyy",
    "ff-Latn-GH": "d/M/yyyy",
    "ff-Latn-GM": "d/M/yyyy",
    "ff-Latn-GN": "d/M/yyyy",
    "ff-Latn-GW": "d/M/yyyy",
    "ff-Latn-LR": "d/M/yyyy",
    "ff-Latn-MR": "d/M/yyyy",
    "ff-Latn-NE": "d/M/yyyy",
    "ff-Latn-NG": "d/M/yyyy",
    "ff-Latn-SL": "d/M/yyyy",
    "ff-Latn-SN": "dd/MM/yyyy",
    "fi": "d.M.yyyy",
    "fi-FI": "d.M.yyyy",
    "fil": "M/d/yyyy",
    "fil-PH": "M/d/yyyy",
    "fo": "dd.MM.yyyy",
    "fo-DK": "dd.MM.yyyy",
    "fo-FO": "dd.MM.yyyy",
    "fr": "dd/MM/yyyy",
    "fr-029": "dd/MM/yyyy",
    "fr-BE": "dd-MM-yy",
    "fr-BF": "dd/MM/yyyy",
    "fr-BI": "dd/MM/yyyy",
    "fr-BJ": "dd/MM/yyyy",
    "fr-BL": "dd/MM/yyyy",
    "fr-CA": "yyyy-MM-dd",
    "fr-CD": "dd/MM/yyyy",
    "fr-CF": "dd/MM/yyyy",
    "fr-CG": "dd/MM/yyyy",
    "fr-CH": "dd.MM.yyyy",
    "fr-CI": "dd/MM/yyyy",
    "fr-CM": "dd/MM/yyyy",
    "fr-DJ": "dd/MM/yyyy",
    "fr-DZ": "dd/MM/yyyy",
    "fr-FR": "dd/MM/yyyy",
    "fr-GA": "dd/MM/yyyy",
    "fr-GF": "dd/MM/yyyy",
    "fr-GN": "dd/MM/yyyy",
    "fr-GP": "dd/MM/yyyy",
    "fr-GQ": "dd/MM/yyyy",
    "fr-HT": "dd/MM/yyyy",
    "fr-KM": "dd/MM/yyyy",
    "fr-LU": "dd/MM/yyyy",
    "fr-MA": "dd/MM/yyyy",
    "fr-MC": "dd/MM/yyyy",
    "fr-MF": "dd/MM/yyyy",
    "fr-MG": "dd/MM/yyyy",
    "fr-ML": "dd/MM/yyyy",
    "fr-MQ": "dd/MM/yyyy",
    "fr-MR": "dd/MM/yyyy",
    "fr-MU": "dd/MM/yyyy",
    "fr-NC": "dd/MM/yyyy",
    "fr-NE": "dd/MM/yyyy",
    "fr-PF": "dd/MM/yyyy",
    "fr-PM": "dd/MM/yyyy",
    "fr-RE": "dd/MM/yyyy",
    "fr-RW": "dd/MM/yyyy",
    "fr-SC": "dd/MM/yyyy",
    "fr-SN": "dd/MM/yyyy",
    "fr-SY": "dd/MM/yyyy",
    "fr-TD": "dd/MM/yyyy",
    "fr-TG": "dd/MM/yyyy",
    "fr-TN": "dd/MM/yyyy",
    "fr-VU": "dd/MM/yyyy",
    "fr-WF": "dd/MM/yyyy",
    "fr-YT": "dd/MM/yyyy",
    "fur": "dd/MM/yyyy",
    "fur-IT": "dd/MM/yyyy",
    "fy": "dd-MM-yyyy",
    "fy-NL": "dd-MM-yyyy",
    "ga": "dd/MM/yyyy",
    "ga-GB": "dd/MM/yyyy",
    "ga-IE": "dd/MM/yyyy",
    "gd": "dd/MM/yyyy",
    "gd-GB": "dd/MM/yyyy",
    "gl": "dd/MM/yyyy",
    "gl-ES": "dd/MM/yyyy",
    "gn": "dd/MM/yyyy",
    "gn-PY": "dd/MM/yyyy",
    "gsw": "dd.MM.yyyy",
    "gsw-CH": "dd.MM.yyyy",
    "gsw-FR": "dd/MM/yyyy",
    "gsw-LI": "dd.MM.yyyy",
    "gu": "dd-MM-yy",
    "gu-IN": "dd-MM-yy",
    "guz": "dd/MM/yyyy",
    "guz-KE": "dd/MM/yyyy",
    "gv": "dd/MM/yyyy",
    "gv-IM": "dd/MM/yyyy",
    "ha": "d/M/yyyy",
    "ha-Latn": "d/M/yyyy",
    "ha-Latn-GH": "d/M/yyyy",
    "ha-Latn-NE": "d/M/yyyy",
    "ha-Latn-NG": "d/M/yyyy",
    "haw": "d/M/yyyy",
    "haw-US": "d/M/yyyy",
    "he": "dd/MM/yyyy",
    "he-IL": "dd/MM/yyyy",
    "hi": "dd-MM-yyyy",
    "hi-IN": "dd-MM-yyyy",
    "hi-Latn": "dd/MM/yyyy",
    "hi-Latn-IN": "dd/MM/yyyy",
    "hr": "d.M.yyyy.",
    "hr-BA": "d. M. yyyy.",
    "hr-HR": "d.M.yyyy.",
    "hsb": "d.M.yyyy",
    "hsb-DE": "d.M.yyyy",
    "hu": "yyyy. MM. dd.",
    "hu-HU": "yyyy. MM. dd.",
    "hy": "dd.MM.yyyy",
    "hy-AM": "dd.MM.yyyy",
    "ia": "dd-MM-yyyy",
    "ia-001": "dd-MM-yyyy",
    "ibb": "d/M/yyyy",
    "ibb-NG": "d/M/yyyy",
    "id": "dd/MM/yyyy",
    "id-ID": "dd/MM/yyyy",
    "ig": "d/M/yyyy",
    "ig-NG": "d/M/yyyy",
    "ii": "yyyy/M/d",
    "ii-CN": "yyyy/M/d",
    "is": "d.M.yyyy",
    "is-IS": "d.M.yyyy",
    "it": "dd/MM/yyyy",
    "it-CH": "dd.MM.yyyy",
    "it-IT": "dd/MM/yyyy",
    "it-SM": "dd/MM/yyyy",
    "it-VA": "dd/MM/yyyy",
    "iu": "d/MM/yyyy",
    "iu-Cans": "d/M/yyyy",
    "iu-Cans-CA": "d/M/yyyy",
    "iu-Latn": "d/MM/yyyy",
    "iu-Latn-CA": "d/MM/yyyy",
    "ja": "yyyy/MM/dd",
    "ja-JP": "yyyy/MM/dd",
    "jgo": "yyyy-MM-dd",
    "jgo-CM": "yyyy-MM-dd",
    "jmc": "dd/MM/yyyy",
    "jmc-TZ": "dd/MM/yyyy",
    "jv": "dd/MM/yyyy",
    "jv-Java": "dd-MM-yyyy",
    "jv-Java-ID": "dd-MM-yyyy",
    "jv-Latn": "dd/MM/yyyy",
    "jv-Latn-ID": "dd/MM/yyyy",
    "ka": "dd.MM.yyyy",
    "kab": "d/M/yyyy",
    "kab-DZ": "d/M/yyyy",
    "ka-GE": "dd.MM.yyyy",
    "kam": "dd/MM/yyyy",
    "kam-KE": "dd/MM/yyyy",
    "kde": "dd/MM/yyyy",
    "kde-TZ": "dd/MM/yyyy",
    "kea": "dd/MM/yyyy",
    "kea-CV": "dd/MM/yyyy",
    "kgp": "dd/MM/yyyy",
    "kgp-Latn": "dd/MM/yyyy",
    "kgp-Latn-BR": "dd/MM/yyyy",
    "khq": "d/M/yyyy",
    "khq-ML": "d/M/yyyy",
    "ki": "dd/MM/yyyy",
    "ki-KE": "dd/MM/yyyy",
    "kk": "dd.MM.yyyy",
    "kkj": "dd/MM yyyy",
    "kkj-CM": "dd/MM yyyy",
    "kk-KZ": "dd.MM.yyyy",
    "kl": "dd-MM-yyyy",
    "kl-GL": "dd-MM-yyyy",
    "kln": "dd/MM/yyyy",
    "kln-KE": "dd/MM/yyyy",
    "km": "dd/MM/yy",
    "km-KH": "dd/MM/yy",
    "kn": "dd-MM-yy",
    "kn-IN": "dd-MM-yy",
    "ko": "yyyy-MM-dd",
    "kok": "dd-MM-yyyy",
    "kok-IN": "dd-MM-yyyy",
    "ko-KP": "yyyy. M. d.",
    "ko-KR": "yyyy-MM-dd",
    "kr": "d/M/yyyy",
    "kr-Latn": "d/M/yyyy",
    "kr-Latn-NG": "d/M/yyyy",
    "ks": "M/d/yyyy",
    "ks-Arab": "M/d/yyyy",
    "ks-Arab-IN": "M/d/yyyy",
    "ksb": "dd/MM/yyyy",
    "ksb-TZ": "dd/MM/yyyy",
    "ks-Deva": "d/M/yyyy",
    "ks-Deva-IN": "d/M/yyyy",
    "ksf": "d/M/yyyy",
    "ksf-CM": "d/M/yyyy",
    "ksh": "d. M. yyyy",
    "ksh-DE": "d. M. yyyy",
    "ku": "yyyy/MM/dd",
    "ku-Arab": "yyyy/MM/dd",
    "ku-Arab-IQ": "yyyy/MM/dd",
    "ku-Arab-IR": "dd/MM/yyyy",
    "kw": "dd/MM/yyyy",
    "kw-GB": "dd/MM/yyyy",
    "ky": "d/M/yyyy",
    "ky-KG": "d/M/yyyy",
    "la": "d M yyyy gg",
    "lag": "dd/MM/yyyy",
    "lag-TZ": "dd/MM/yyyy",
    "la-VA": "d M yyyy gg",
    "lb": "dd.MM.yy",
    "lb-LU": "dd.MM.yy",
    "lg": "dd/MM/yyyy",
    "lg-UG": "dd/MM/yyyy",
    "lkt": "M/d/yyyy",
    "lkt-US": "M/d/yyyy",
    "ln": "d/M/yyyy",
    "ln-AO": "d/M/yyyy",
    "ln-CD": "d/M/yyyy",
    "ln-CF": "d/M/yyyy",
    "ln-CG": "d/M/yyyy",
    "lo": "d/M/yyyy",
    "lo-LA": "d/M/yyyy",
    "lrc": "dd/MM/yyyy",
    "lrc-IQ": "yyyy-MM-dd",
    "lrc-IR": "dd/MM/yyyy",
    "lt": "yyyy-MM-dd",
    "lt-LT": "yyyy-MM-dd",
    "lu": "d/M/yyyy",
    "lu-CD": "d/M/yyyy",
    "luo": "dd/MM/yyyy",
    "luo-KE": "dd/MM/yyyy",
    "luy": "dd/MM/yyyy",
    "luy-KE": "dd/MM/yyyy",
    "lv": "dd.MM.yyyy",
    "lv-LV": "dd.MM.yyyy",
    "mai": "d/M/yyyy",
    "mai-IN": "d/M/yyyy",
    "mas": "dd/MM/yyyy",
    "mas-KE": "dd/MM/yyyy",
    "mas-TZ": "dd/MM/yyyy",
    "mer": "dd/MM/yyyy",
    "mer-KE": "dd/MM/yyyy",
    "mfe": "d/M/yyyy",
    "mfe-MU": "d/M/yyyy",
    "mg": "yyyy-MM-dd",
    "mgh": "dd/MM/yyyy",
    "mgh-MZ": "dd/MM/yyyy",
    "mg-MG": "yyyy-MM-dd",
    "mgo": "yyyy-MM-dd",
    "mgo-CM": "yyyy-MM-dd",
    "mi": "dd-MM-yyyy",
    "mi-NZ": "dd-MM-yyyy",
    "mk": "d.M.yyyy",
    "mk-MK": "d.M.yyyy",
    "ml": "d/M/yyyy",
    "ml-IN": "d/M/yyyy",
    "mn": "yyyy.MM.dd",
    "mn-Cyrl": "yyyy.MM.dd",
    "mni": "d/M/yyyy",
    "mni-Beng": "d/M/yyyy",
    "mni-IN": "d/M/yyyy",
    "mn-MN": "yyyy.MM.dd",
    "mn-Mong": "yyyy/M/d",
    "mn-Mong-CN": "yyyy/M/d",
    "mn-Mong-MN": "yyyy/M/d",
    "moh": "M/d/yyyy",
    "moh-CA": "M/d/yyyy",
    "mr": "dd-MM-yyyy",
    "mr-IN": "dd-MM-yyyy",
    "ms": "d/MM/yyyy",
    "ms-BN": "d/MM/yyyy",
    "ms-ID": "dd/MM/yyyy",
    "ms-MY": "d/MM/yyyy",
    "ms-SG": "d/MM/yyyy",
    "mt": "dd/MM/yyyy",
    "mt-MT": "dd/MM/yyyy",
    "mua": "d/M/yyyy",
    "mua-CM": "d/M/yyyy",
    "my": "d/M/yyyy",
    "my-MM": "d/M/yyyy",
    "mzn": "dd/MM/yyyy",
    "mzn-IR": "dd/MM/yyyy",
    "naq": "dd/MM/yyyy",
    "naq-NA": "dd/MM/yyyy",
    "nb": "dd.MM.yyyy",
    "nb-NO": "dd.MM.yyyy",
    "nb-SJ": "dd.MM.yyyy",
    "nd": "dd/MM/yyyy",
    "nds": "d.MM.yyyy",
    "nds-DE": "d.MM.yyyy",
    "nds-NL": "d.MM.yyyy",
    "nd-ZW": "dd/MM/yyyy",
    "ne": "M/d/yyyy",
    "ne-IN": "yyyy/M/d",
    "ne-NP": "M/d/yyyy",
    "nl": "d-M-yyyy",
    "nl-AW": "dd-MM-yyyy",
    "nl-BE": "d/MM/yyyy",
    "nl-BQ": "dd-MM-yyyy",
    "nl-CW": "dd-MM-yyyy",
    "nl-NL": "d-M-yyyy",
    "nl-SR": "dd-MM-yyyy",
    "nl-SX": "dd-MM-yyyy",
    "nmg": "d/M/yyyy",
    "nmg-CM": "d/M/yyyy",
    "nn": "dd.MM.yyyy",
    "nnh": "dd/MM/yyyy",
    "nnh-CM": "dd/MM/yyyy",
    "nn-NO": "dd.MM.yyyy",
    "no": "dd.MM.yyyy",
    "nqo": "dd/MM/yyyy",
    "nqo-GN": "dd/MM/yyyy",
    "nr": "yyyy-MM-dd",
    "nr-ZA": "yyyy-MM-dd",
    "nso": "yyyy-MM-dd",
    "nso-ZA": "yyyy-MM-dd",
    "nus": "d/MM/yyyy",
    "nus-SS": "d/MM/yyyy",
    "nyn": "dd/MM/yyyy",
    "nyn-UG": "dd/MM/yyyy",
    "oc": "d/MM/yyyy",
    "oc-ES": "d/MM/yyyy",
    "oc-FR": "d/MM/yyyy",
    "om": "dd/MM/yyyy",
    "om-ET": "dd/MM/yyyy",
    "om-KE": "dd/MM/yyyy",
    "or": "dd-MM-yy",
    "or-IN": "dd-MM-yy",
    "os": "dd.MM.yyyy",
    "os-GE": "dd.MM.yyyy",
    "os-RU": "dd.MM.yyyy",
    "pa": "dd-MM-yy",
    "pa-Arab": "dd-MM-yy",
    "pa-Arab-PK": "dd-MM-yy",
    "pa-Guru": "dd-MM-yy",
    "pa-IN": "dd-MM-yy",
    "pap": "d-M-yyyy",
    "pap-029": "d-M-yyyy",
    "pcm": "dd/MM/yyyy",
    "pcm-Latn": "dd/MM/yyyy",
    "pcm-Latn-NG": "dd/MM/yyyy",
    "pl": "d.MM.yyyy",
    "pl-PL": "d.MM.yyyy",
    "prg": "dd.MM.yyyy",
    "prg-001": "dd.MM.yyyy",
    "ps": "yyyy/M/d",
    "ps-AF": "yyyy/M/d",
    "ps-PK": "yyyy/M/d",
    "pt": "dd/MM/yyyy",
    "pt-AO": "dd/MM/yyyy",
    "pt-BR": "dd/MM/yyyy",
    "pt-CH": "dd/MM/yyyy",
    "pt-CV": "dd/MM/yyyy",
    "pt-GQ": "dd/MM/yyyy",
    "pt-GW": "dd/MM/yyyy",
    "pt-LU": "dd/MM/yyyy",
    "pt-MO": "dd/MM/yyyy",
    "pt-MZ": "dd/MM/yyyy",
    "pt-PT": "dd/MM/yyyy",
    "pt-ST": "dd/MM/yyyy",
    "pt-TL": "dd/MM/yyyy",
    "quc": "dd/MM/yyyy",
    "quc-Latn": "dd/MM/yyyy",
    "quc-Latn-GT": "dd/MM/yyyy",
    "quz": "dd/MM/yyyy",
    "quz-BO": "dd/MM/yyyy",
    "quz-EC": "dd/MM/yyyy",
    "quz-PE": "dd/MM/yyyy",
    "raj": "yyyy-MM-dd",
    "raj-Deva": "yyyy-MM-dd",
    "raj-Deva-IN": "yyyy-MM-dd",
    "rm": "dd-MM-yyyy",
    "rm-CH": "dd-MM-yyyy",
    "rn": "d/M/yyyy",
    "rn-BI": "d/M/yyyy",
    "ro": "dd.MM.yyyy",
    "rof": "dd/MM/yyyy",
    "rof-TZ": "dd/MM/yyyy",
    "ro-MD": "dd.MM.yyyy",
    "ro-RO": "dd.MM.yyyy",
    "ru": "dd.MM.yyyy",
    "ru-BY": "dd.MM.yyyy",
    "ru-KG": "dd.MM.yyyy",
    "ru-KZ": "dd.MM.yyyy",
    "ru-MD": "dd.MM.yyyy",
    "ru-RU": "dd.MM.yyyy",
    "ru-UA": "dd.MM.yyyy",
    "rw": "yyyy-MM-dd",
    "rwk": "dd/MM/yyyy",
    "rwk-TZ": "dd/MM/yyyy",
    "rw-RW": "yyyy-MM-dd",
    "sa": "d/M/yyyy",
    "sah": "dd.MM.yyyy",
    "sah-RU": "dd.MM.yyyy",
    "sa-IN": "d/M/yyyy",
    "saq": "dd/MM/yyyy",
    "saq-KE": "dd/MM/yyyy",
    "sat": "d/M/yyyy",
    "sat-Olck": "d/M/yyyy",
    "sat-Olck-IN": "d/M/yyyy",
    "sbp": "dd/MM/yyyy",
    "sbp-TZ": "dd/MM/yyyy",
    "sc": "dd/MM/yyyy",
    "sc-Latn": "dd/MM/yyyy",
    "sc-Latn-IT": "dd/MM/yyyy",
    "sd": "dd/MM/yyyy",
    "sd-Arab": "dd/MM/yyyy",
    "sd-Arab-PK": "dd/MM/yyyy",
    "sd-Deva": "M/d/yyyy",
    "sd-Deva-IN": "M/d/yyyy",
    "se": "yyyy-MM-dd",
    "se-FI": "d.M.yyyy",
    "seh": "d/M/yyyy",
    "seh-MZ": "d/M/yyyy",
    "se-NO": "yyyy-MM-dd",
    "ses": "d/M/yyyy",
    "se-SE": "yyyy-MM-dd",
    "ses-ML": "d/M/yyyy",
    "sg": "d/M/yyyy",
    "sg-CF": "d/M/yyyy",
    "shi": "d/M/yyyy",
    "shi-Latn": "d/M/yyyy",
    "shi-Latn-MA": "d/M/yyyy",
    "shi-Tfng": "d/M/yyyy",
    "shi-Tfng-MA": "d/M/yyyy",
    "si": "yyyy-MM-dd",
    "si-LK": "yyyy-MM-dd",
    "sk": "d. M. yyyy",
    "sk-SK": "d. M. yyyy",
    "sl": "d. MM. yyyy",
    "sl-SI": "d. MM. yyyy",
    "sma": "yyyy-MM-dd",
    "sma-NO": "dd.MM.yyyy",
    "sma-SE": "yyyy-MM-dd",
    "smj": "yyyy-MM-dd",
    "smj-NO": "dd.MM.yyyy",
    "smj-SE": "yyyy-MM-dd",
    "smn": "d.M.yyyy",
    "smn-FI": "d.M.yyyy",
    "sms": "d.M.yyyy",
    "sms-FI": "d.M.yyyy",
    "sn": "yyyy-MM-dd",
    "sn-Latn": "yyyy-MM-dd",
    "sn-Latn-ZW": "yyyy-MM-dd",
    "so": "dd/MM/yyyy",
    "so-DJ": "dd/MM/yyyy",
    "so-ET": "dd/MM/yyyy",
    "so-KE": "dd/MM/yyyy",
    "so-SO": "dd/MM/yyyy",
    "sq": "d.M.yyyy",
    "sq-AL": "d.M.yyyy",
    "sq-MK": "d.M.yyyy",
    "sq-XK": "d.M.yyyy",
    "sr": "d.M.yyyy.",
    "sr-Cyrl": "dd.MM.yyyy.",
    "sr-Cyrl-BA": "d.M.yyyy.",
    "sr-Cyrl-ME": "d.M.yyyy.",
    "sr-Cyrl-RS": "dd.MM.yyyy.",
    "sr-Cyrl-XK": "d.M.yyyy.",
    "sr-Latn": "d.M.yyyy.",
    "sr-Latn-BA": "d.M.yyyy.",
    "sr-Latn-ME": "d.M.yyyy.",
    "sr-Latn-RS": "d.M.yyyy.",
    "sr-Latn-XK": "d.M.yyyy.",
    "ss": "yyyy-MM-dd",
    "ss-SZ": "yyyy-MM-dd",
    "ssy": "dd/MM/yyyy",
    "ssy-ER": "dd/MM/yyyy",
    "ss-ZA": "yyyy-MM-dd",
    "st": "yyyy-MM-dd",
    "st-LS": "yyyy-MM-dd",
    "st-ZA": "yyyy-MM-dd",
    "su": "d/M/yyyy",
    "su-Latn": "d/M/yyyy",
    "su-Latn-ID": "d/M/yyyy",
    "sv": "yyyy-MM-dd",
    "sv-AX": "yyyy-MM-dd",
    "sv-FI": "yyyy-MM-dd",
    "sv-SE": "yyyy-MM-dd",
    "sw": "dd/MM/yyyy",
    "sw-CD": "dd/MM/yyyy",
    "sw-KE": "dd/MM/yyyy",
    "sw-TZ": "dd/MM/yyyy",
    "sw-UG": "dd/MM/yyyy",
    "syr": "dd/MM/yyyy",
    "syr-SY": "dd/MM/yyyy",
    "ta": "dd-MM-yyyy",
    "ta-IN": "dd-MM-yyyy",
    "ta-LK": "d/M/yyyy",
    "ta-MY": "d/M/yyyy",
    "ta-SG": "d/M/yyyy",
    "te": "dd-MM-yyyy",
    "te-IN": "dd-MM-yyyy",
    "teo": "dd/MM/yyyy",
    "teo-KE": "dd/MM/yyyy",
    "teo-UG": "dd/MM/yyyy",
    "tg": "dd.MM.yyyy",
    "tg-Cyrl": "dd.MM.yyyy",
    "tg-Cyrl-TJ": "dd.MM.yyyy",
    "th": "d/M/yyyy",
    "th-TH": "d/M/yyyy",
    "ti": "dd/MM/yyyy",
    "ti-ER": "dd/MM/yyyy",
    "ti-ET": "dd/MM/yyyy",
    "tig": "dd/MM/yyyy",
    "tig-ER": "dd/MM/yyyy",
    "tk": "dd.MM.yy 'ý.'",
    "tk-TM": "dd.MM.yy 'ý.'",
    "tn": "yyyy-MM-dd",
    "tn-BW": "yyyy-MM-dd",
    "tn-ZA": "yyyy-MM-dd",
    "to": "d/M/yyyy",
    "to-TO": "d/M/yyyy",
    "tr": "d.MM.yyyy",
    "tr-CY": "d.MM.yyyy",
    "tr-TR": "d.MM.yyyy",
    "ts": "yyyy-MM-dd",
    "ts-ZA": "yyyy-MM-dd",
    "tt": "dd.MM.yyyy",
    "tt-RU": "dd.MM.yyyy",
    "twq": "d/M/yyyy",
    "twq-NE": "d/M/yyyy",
    "tzm": "dd-MM-yyyy",
    "tzm-Arab": "d/M/yyyy",
    "tzm-Arab-MA": "d/M/yyyy",
    "tzm-Latn": "dd-MM-yyyy",
    "tzm-Latn-DZ": "dd-MM-yyyy",
    "tzm-Latn-MA": "dd/MM/yyyy",
    "tzm-Tfng": "dd-MM-yyyy",
    "tzm-Tfng-MA": "dd-MM-yyyy",
    "ug": "yyyy-M-d",
    "ug-CN": "yyyy-M-d",
    "uk": "dd.MM.yyyy",
    "uk-UA": "dd.MM.yyyy",
    "ur": "dd/MM/yyyy",
    "ur-IN": "d/M/yy",
    "ur-PK": "dd/MM/yyyy",
    "uz": "dd/MM/yyyy",
    "uz-Arab": "dd/MM yyyy",
    "uz-Arab-AF": "dd/MM yyyy",
    "uz-Cyrl": "dd/MM/yyyy",
    "uz-Cyrl-UZ": "dd/MM/yyyy",
    "uz-Latn": "dd/MM/yyyy",
    "uz-Latn-UZ": "dd/MM/yyyy",
    "vai": "dd/MM/yyyy",
    "vai-Latn": "dd/MM/yyyy",
    "vai-Latn-LR": "dd/MM/yyyy",
    "vai-Vaii": "dd/MM/yyyy",
    "vai-Vaii-LR": "dd/MM/yyyy",
    "ve": "yyyy-MM-dd",
    "ve-ZA": "yyyy-MM-dd",
    "vi": "dd/MM/yyyy",
    "vi-VN": "dd/MM/yyyy",
    "vo": "yyyy-MM-dd",
    "vo-001": "yyyy-MM-dd",
    "vun": "dd/MM/yyyy",
    "vun-TZ": "dd/MM/yyyy",
    "wae": "yyyy-MM-dd",
    "wae-CH": "yyyy-MM-dd",
    "wal": "dd/MM/yyyy",
    "wal-ET": "dd/MM/yyyy",
    "wo": "dd-MM-yyyy",
    "wo-SN": "dd-MM-yyyy",
    "xh": "M/d/yyyy",
    "xh-ZA": "M/d/yyyy",
    "xog": "dd/MM/yyyy",
    "xog-UG": "dd/MM/yyyy",
    "yav": "d/M/yyyy",
    "yav-CM": "d/M/yyyy",
    "yi": "dd/MM/yyyy",
    "yi-001": "dd/MM/yyyy",
    "yo": "d/M/yyyy",
    "yo-BJ": "d/M/yyyy",
    "yo-NG": "d/M/yyyy",
    "yrl": "dd/MM/yyyy",
    "yrl-Latn": "dd/MM/yyyy",
    "yrl-Latn-BR": "dd/MM/yyyy",
    "yrl-Latn-CO": "dd/MM/yyyy",
    "yrl-Latn-VE": "dd/MM/yyyy",
    "zgh": "d/M/yyyy",
    "zgh-Tfng": "d/M/yyyy",
    "zgh-Tfng-MA": "d/M/yyyy",
    "zh": "yyyy/M/d",
    "zh-CHS": "yyyy/M/d",
    "zh-CHT": "d/M/yyyy",
    "zh-CN": "yyyy/M/d",
    "zh-Hans": "yyyy/M/d",
    "zh-Hans-HK": "d/M/yyyy",
    "zh-Hans-MO": "d/M/yyyy",
    "zh-Hant": "d/M/yyyy",
    "zh-HK": "d/M/yyyy",
    "zh-MO": "d/M/yyyy",
    "zh-SG": "d/M/yyyy",
    "zh-TW": "yyyy/M/d",
    "zu": "M/d/yyyy",
    "zu-ZA": "M/d/yyyy",
}
