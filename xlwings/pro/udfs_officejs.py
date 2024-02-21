"""
Required Notice: Copyright (C) Zoomer Analytics GmbH.

xlwings PRO is dual-licensed under one of the following licenses:

* PolyForm Noncommercial License 1.0.0 (for noncommercial use):
  https://polyformproject.org/licenses/noncommercial/1.0.0
* xlwings PRO License (for commercial use):
  https://github.com/xlwings/xlwings/blob/main/LICENSE_PRO.txt

Commercial licenses can be purchased at https://www.xlwings.org
"""

import inspect
from textwrap import dedent

from .. import XlwingsError, __version__, conversion


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


def xlfunc(f=None, **kwargs):
    def inner(f):
        if not hasattr(f, "__xlfunc__"):
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


async def custom_functions_call(data, module):
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
    if inspect.iscoroutinefunction(func):
        ret = await func(*args)
    else:
        ret = func(*args)

    if "date_format" not in ret_info["options"]:
        ret_info["options"]["date_format"] = locale_to_shortdate[
            data["content_language"].lower()
        ]
    ret_info["options"]["runtime"] = data["runtime"]
    ret = conversion.write(ret, None, ret_info["options"], engine_name="officejs")
    return ret


def custom_functions_code(
    module, custom_functions_call_path="/xlwings/custom-functions-call"
):
    js = """\
         async function base() {
           // Turn arguments into an array, the last one is the invocation parameter
           let argsArr = Array.prototype.slice.call(arguments);
           let func_name = argsArr[0];
           let args = argsArr.slice(1, -1);
           let invocation = argsArr[argsArr.length - 1];
           // headers
           let headers = {};
           headers["Content-Type"] = "application/json";
           headers["Authorization"] = await globalThis.getAuth();
           let runtime;
           if (
             Office.context.requirements.isSetSupported("CustomFunctionsRuntime", "1.4")
           ) {
             runtime = "1.4";
           } else if (
             Office.context.requirements.isSetSupported("CustomFunctionsRuntime", "1.3")
           ) {
             runtime = "1.3";
           } else if (
             Office.context.requirements.isSetSupported("CustomFunctionsRuntime", "1.2")
           ) {
             runtime = "1.2";
           } else {
             runtime = "1.1";
           }
           let response = await fetch(
             window.location.origin + "custom_functions_call_path",
             {
               method: "POST",
               headers: headers,
               body: JSON.stringify({
                 func_name: func_name,
                 args: args,
                 caller_address: invocation.address,
                 formula_name: invocation.functionName,
                 content_language: Office.context.contentLanguage,
                 version: "xlwings_version",
                 runtime: runtime,
               }),
             }
           );
           if (response.status !== 200) {
             let errMsg = await response.text();
             // Error message only visible by hovering over the error flag!
             if (
               Office.context.requirements.isSetSupported(
                 "CustomFunctionsRuntime",
                 "1.2"
               )
             ) {
               let error = new CustomFunctions.Error(
                 CustomFunctions.ErrorCode.invalidValue,
                 errMsg
               );
               throw error;
             } else {
               return [[errMsg]];
             }
           } else {
             rawData = await response.json();
           }
           return rawData.result;
         }
    """.replace(
        "xlwings_version", __version__
    ).replace(
        "custom_functions_call_path", custom_functions_call_path
    )  # format string would require to double all curly braces
    js = dedent(js)
    for name, obj in inspect.getmembers(module):
        if hasattr(obj, "__xlfunc__"):
            xlfunc = obj.__xlfunc__
            func_name = xlfunc["name"]
            js += dedent(
                f"""\
            async function {func_name}() {{
                args = ["{func_name}"]
                args.push.apply(args, arguments);
                return await base.apply(null, args);
            }}
            CustomFunctions.associate("{func_name.upper()}", {func_name});
            """
            )
    return js


def custom_functions_meta(module):
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
            func["options"] = {
                "requiresAddress": True,
                "requiresParameterAddresses": True,
            }
            if xlfunc["volatile"]:
                func["options"]["volatile"] = True
            func["result"] = {"dimensionality": "matrix", "type": "any"}

            params = []
            for arg in xlfunc["args"]:
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


locale_to_shortdate = {
    # This is using the locales from https://github.com/OfficeDev/office-js/tree/release/dist
    # matched with values from https://stackoverflow.com/a/9893752/918626
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
