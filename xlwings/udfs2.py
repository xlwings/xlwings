import inspect
from importlib import import_module
from textwrap import dedent

from . import conversion


def func_sig(f):
    s = inspect.signature(f)
    vararg = None
    args = []
    defaults = []
    for p in s.parameters.values():
        if p.kind is inspect.Parameter.POSITIONAL_OR_KEYWORD:
            args.append(p.name)
            if p.default is not inspect.Signature.empty:
                defaults.append(p.default)
        elif p.kind is inspect.Parameter.VAR_POSITIONAL:
            args.append(p.name)
            vararg = p.name
        else:
            raise Exception("xlwings does not support UDFs with keyword arguments")
    return {"args": args, "defaults": defaults, "vararg": vararg}


def check_bool(kw, default, **func_kwargs):
    if kw in func_kwargs:
        check = func_kwargs.pop(kw)
        if isinstance(check, bool):
            return check
        raise Exception(
            '{0} only takes boolean values. ("{1}" provided).'.format(kw, check)
        )
    return default


def xlfunc(f=None, **kwargs):
    def inner(f):
        if not hasattr(f, "__xlfunc__"):
            xlf = f.__xlfunc__ = {}
            xlf["name"] = f.__name__
            xlargs = xlf["args"] = []
            xlargmap = xlf["argmap"] = {}
            sig = func_sig(f)
            nArgs = len(sig["args"])
            nDefaults = len(sig["defaults"])
            nRequiredArgs = nArgs - nDefaults
            if sig["vararg"] and nDefaults > 0:
                raise Exception(
                    "xlwings does not support UDFs "
                    "with both optional and variable length arguments"
                )
            for vpos, vname in enumerate(sig["args"]):
                arg_info = {
                    "name": vname,
                    "pos": vpos,
                    "doc": "Positional argument " + str(vpos + 1),
                    "vararg": vname == sig["vararg"],
                    "options": {},
                }
                if vpos >= nRequiredArgs:
                    arg_info["optional"] = sig["defaults"][vpos - nRequiredArgs]
                xlargs.append(arg_info)
                xlargmap[vname] = xlargs[-1]
            xlf["ret"] = {
                "doc": f.__doc__
                if f.__doc__ is not None
                else f"Python function '{f.__name__}'",
                "options": {},
            }
        f.__xlfunc__["volatile"] = check_bool("volatile", default=False, **kwargs)
        # If there's a global namespace defined in the manifest, this will be the
        # sub-namespace, i.e. NAMESPACE.SUBNAMESPACE.FUNCTIONNAME
        f.__xlfunc__["namespace"] = kwargs.get("namespace")
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
        if arg not in xlf["argmap"]:
            raise Exception("Invalid argument name '" + arg + "'.")
        xla = xlf["argmap"][arg]
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


async def call_udf(data):
    module_name = data["module_name"]
    func_name = data["func_name"]
    args = data["args"]
    module = import_module(module_name)
    func = getattr(module, func_name)
    func_info = func.__xlfunc__
    args_info = func_info["args"]
    ret_info = func_info["ret"]

    args = list(args)
    # Turn varargs into regular arguments (remove the outermost list)
    for i, arg in enumerate(args):
        arg_info = args_info[min(i, len(args_info) - 1)]
        if arg_info["vararg"]:
            del args[i]
            args[i:i] = arg

    for i, arg in enumerate(args):
        arg_info = args_info[min(i, len(args_info) - 1)]
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

    ret_info["options"]["date_format"] = locale_to_shortdate[
        data["content_language"].lower()
    ]
    ret = conversion.write(ret, None, ret_info["options"], engine_name="officejs")
    return ret


def generate_js_wrapper(module):
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
           let response = await fetch(window.location.origin + "/xlwings/udfs", {
             method: "POST",
             headers: headers,
             body: JSON.stringify({
               module_name: "functions",
               func_name: func_name,
               args: args,
               caller_address: invocation.address,
               formula_name: invocation.functionName,
               content_language: Office.context.contentLanguage,
             }),
           });
           if (response.status !== 200) {
             let errMsg = await response.text();
             // Error message only visible by hovering over the error flag, not by clicking it!
             let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, errMsg);
             throw error;
           } else {
             rawData = await response.json();
           }
           result = rawData.result;
           return result;
         }
    """
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


def generate_json_meta(module):
    funcs = []
    for name, obj in inspect.getmembers(module):
        if hasattr(obj, "__xlfunc__"):
            xlfunc = obj.__xlfunc__
            func = {}
            func["description"] = xlfunc["ret"]["doc"]
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
    # with the short date format from ChatGPT. Maybe replace with https://stackoverflow.com/a/9893752/918626
    "af-za": "yyyy/MM/dd",
    "am-et": "d/M/yyyy",
    "ar-ae": "dd/MM/yy",
    "ar-bh": "dd/MM/yy",
    "ar-dz": "dd/MM/yy",
    "ar-eg": "dd/MM/yy",
    "ar-iq": "dd/MM/yy",
    "ar-jo": "dd/MM/yy",
    "ar-kw": "dd/MM/yy",
    "ar-lb": "dd/MM/yy",
    "ar-ly": "dd/MM/yyyy",
    "ar-ma": "dd/MM/yyyy",
    "ar-om": "dd/MM/yyyy",
    "ar-qa": "dd/MM/yyyy",
    "ar-sa": "dd/MM/yy",
    "ar-sy": "dd/MM/yyyy",
    "ar-tn": "dd/MM/yyyy",
    "ar-ye": "dd/MM/yyyy",
    "az-latn-az": "dd.MM.yyyy",
    "be-by": "dd.MM.yyyy",
    "bg-bg": "d.M.yyyy",
    "bn-in": "dd/MM/yy",
    "bs-latn-ba": "d.M.yyyy.",
    "ca-es": "dd/MM/yy",
    "cs-cz": "d.M.yyyy",
    "cy-gb": "dd/MM/yy",
    "da-dk": "dd-MM-yy",
    "de-at": "dd.MM.yy",
    "de-ch": "dd.MM.yy",
    "de-de": "dd.MM.yy",
    "de-li": "dd.MM.yy",
    "de-lu": "dd.MM.yy",
    "el-gr": "d/M/yy",
    "en-029": "MM/dd/yyyy",
    "en-au": "d/MM/yyyy",
    "en-bz": "dd/MM/yyyy",
    "en-ca": "yyyy-MM-dd",
    "en-gb": "dd/MM/yyyy",
    "en-ie": "dd/MM/yyyy",
    "en-in": "dd-MM-yyyy",
    "en-jm": "dd/MM/yyyy",
    "en-my": "d/M/yyyy",
    "en-nz": "d/MM/yyyy",
    "en-ph": "M/d/yyyy",
    "en-sg": "d/M/yyyy",
    "en-tt": "dd/MM/yyyy",
    "en-us": "m/d/yy",
    "en-za": "yyyy/MM/dd",
    "en-zw": "M/d/yyyy",
    "es-ar": "dd/MM/yy",
    "es-bo": "dd/MM/yy",
    "es-cl": "dd-MM-yy",
    "es-co": "dd/MM/yy",
    "es-cr": "dd/MM/yy",
    "es-do": "dd/MM/yyyy",
    "es-ec": "dd/MM/yyyy",
    "es-es": "dd/MM/yy",
    "es-gt": "dd/MM/yyyy",
    "es-hn": "dd/MM/yyyy",
    "es-mx": "dd/MM/yyyy",
    "es-ni": "dd/MM/yyyy",
    "es-pa": "MM/dd/yyyy",
    "es-pe": "dd/MM/yyyy",
    "es-pr": "MM-dd-yy",
    "es-py": "dd/MM/yyyy",
    "es-sv": "dd/MM/yyyy",
    "es-us": "M/d/yyyy",
    "es-uy": "d/m/yyyy",
    "es-ve": "dd/mm/yyyy",
    "et-ee": "dd.mm.yyyy",
    "eu-es": "yyyy/mm/dd",
    "fa-ir": "mm/dd/yyyy",
    "fi-fi": "d.m.yyyy",
    "fil-ph": "m/d/yyyy",
    "fr-be": "d/m/yyyy",
    "fr-ca": "yyyy-mm-dd",
    "fr-ch": "dd.mm.yyyy",
    "fr-fr": "dd/mm/yyyy",
    "fr-lu": "d/m/yyyy",
    "fr-mc": "dd/mm/yyyy",
    "ga-ie": "dd/mm/yyyy",
    "gl-es": "dd/mm/yyyy",
    "gu-in": "dd/mm/yyyy",
    "he-il": "dd/mm/yyyy",
    "hi-in": "dd-mm-yyyy",
    "hr-ba": "d.m.yyyy.",
    "hr-hr": "d.m.yyyy.",
    "hu-hu": "yyyy.mm.dd.",
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
    "ko-kr": "yyyy-mm-dd",
    "lb-lu": "dd/mm/yyyy",
    "lo-la": "dd/mm/yyyy",
    "lt-lt": "yyyy.mm.dd",
    "lv-lv": "dd.mm.yyyy.",
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
    "pl-pl": "yyyy-mm-dd",
    "pt-br": "dd/mm/yyyy",
    "pt-pt": "dd-mm-yyyy",
    "ro-ro": "dd.mm.yyyy",
    "ru-ru": "dd.mm.yyyy",
    "si-lk": "yyyy-mm-dd",
    "sk-sk": "d.m.yyyy",
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
