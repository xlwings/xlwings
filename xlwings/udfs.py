import os
import re
import os.path
import tempfile
import inspect

from win32com.client import Dispatch

from . import conversion
from .utils import VBAWriter
from . import xlplatform
from . import Range

from . import PY3

if PY3:
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
        return {
            'args': args,
            'defaults': defaults,
            'vararg': vararg
        }

else:
    def func_sig(f):
        s = inspect.getargspec(f)
        if s.keywords:
            raise Exception("xlwings does not support UDFs with keyword arguments")
        return {
            'args': (s.args + [s.varargs]) if s.varargs else s.args,
            'defaults': s.defaults or [],
            'vararg': s.varargs
        }


def xlfunc(f=None, **kwargs):
    def inner(f):
        if not hasattr(f, "__xlfunc__"):
            xlf = f.__xlfunc__ = {}
            xlf["name"] = f.__name__
            xlf["sub"] = False
            xlargs = xlf["args"] = []
            xlargmap = xlf["argmap"] = {}
            sig = func_sig(f)
            nArgs = len(sig['args'])
            nDefaults = len(sig['defaults'])
            nRequiredArgs = nArgs - nDefaults
            if sig['vararg'] and nDefaults > 0:
                raise Exception("xlwings does not support UDFs with both optional and variable length arguments")
            for vpos, vname in enumerate(sig['args']):
                arg_info = {
                    "name": vname,
                    "pos": vpos,
                    "vba": None,
                    "doc": "Positional argument " + str(vpos+1),
                    "vararg": vname == sig['vararg'],
                    "options": {}
                }
                if vpos >= nRequiredArgs:
                    arg_info["optional"] = sig['defaults'][vpos - nRequiredArgs]
                xlargs.append(arg_info)
                xlargmap[vname] = xlargs[-1]
            xlf["ret"] = {
                "doc": f.__doc__ if f.__doc__ is not None else "Python function '" + f.__name__ + "' defined in '" + str(f.__code__.co_filename) + "'.",
                "options": {}
            }
        return f
    if f is None:
        return inner
    else:
        return inner(f)


def xlsub(f=None, **kwargs):
    def inner(f):
        f = xlfunc(**kwargs)(f)
        f.__xlfunc__["sub"] = True
        return f
    if f is None:
        return inner
    else:
        return inner(f)


def xlret(convert=None, **kwargs):
    if convert is not None:
        kwargs['convert'] = convert
    def inner(f):
        xlf = xlfunc(f).__xlfunc__
        xlr = xlf["ret"]
        xlr['options'].update(kwargs)
        return f
    return inner


def xlarg(arg, convert=None, **kwargs):
    if convert is not None:
        kwargs['convert'] = convert
    def inner(f):
        xlf = xlfunc(f).__xlfunc__
        if arg not in xlf["argmap"]:
            raise Exception("Invalid argument name '" + arg + "'.")
        xla = xlf["argmap"][arg]
        for special in ('vba', 'doc'):
            if special in kwargs:
                xla[special] = kwargs.pop(special)
        xla['options'].update(kwargs)
        return f
    return inner


udf_scripts = {}
def udf_script(filename):
    filename = os.path.normcase(filename.lower())
    mtime = os.path.getmtime(filename)
    if filename in udf_scripts:
        mtime2, vars = udf_scripts[filename]
        if mtime == mtime2:
            return vars
    vars = {}
    with open(filename, "r") as f:
        exec(compile(f.read(), filename, "exec"), vars)
    udf_scripts[filename] = (mtime, vars)
    return vars


def call_udf(module_names, func_name, args, this_workbook):
    module_vars = {}
    for module in module_names.split(';'):
        source_module = __import__(module)
        module_vars.update(udf_script(source_module.__file__))
    func = module_vars[func_name]

    func_info = func.__xlfunc__
    args_info = func_info['args']
    ret_info = func_info['ret']

    args = list(args)
    for i, arg in enumerate(args):
        arg_info = args_info[min(i, len(args_info)-1)]
        if type(arg) is int and arg == -2147352572:      # missing
            args[i] = arg_info.get('optional', None)
        elif xlplatform.is_range_instance(arg):
            args[i] = conversion.read(Range(arg), None, arg_info['options'])
        else:
            args[i] = conversion.read(None, arg, arg_info['options'])

    xlplatform.xl_workbook_current = Dispatch(this_workbook)
    ret = func(*args)

    return conversion.write(ret, None, ret_info['options'])


def generate_vba_wrapper(script_vars, f):

    vba = VBAWriter(f)

    vba.writeln('Attribute VB_Name = "xlwings_udfs"')

    vba.writeln("'Autogenerated code by xlwings - changes will be lost with next import!")

    for svar in script_vars.values():
        if hasattr(svar, '__xlfunc__'):
            xlfunc = svar.__xlfunc__
            xlret = xlfunc['ret']
            fname = xlfunc['name']

            ftype = 'Sub' if xlfunc['sub'] else 'Function'

            func_sig = ftype + " " + fname + "("

            first = True
            vararg = ''
            n_args = len(xlfunc['args'])
            for arg in xlfunc['args']:
                if not arg['vba']:
                    argname = arg['name']
                    if not first:
                        func_sig += ', '
                    if 'optional' in arg:
                        func_sig += 'Optional '
                    elif arg['vararg']:
                        func_sig += 'ParamArray '
                        vararg = argname
                    func_sig += argname
                    if arg['vararg']:
                        func_sig += '()'
                    first = False
            func_sig += ')'

            with vba.block(func_sig):

                if ftype == 'Function':
                    vba.write("If TypeOf Application.Caller Is Range Then On Error GoTo failed\n")

                if vararg != '':
                    vba.write("ReDim argsArray(1 to UBound(" + vararg + ") - LBound(" + vararg + ") + " + str(n_args) + ")\n")

                j = 1
                for arg in xlfunc['args']:
                    argname = arg['name']
                    if arg['vararg']:
                        vba.write("For k = LBound(" + vararg + ") To UBound(" + vararg + ")\n")
                        argname = vararg + "(k)"

                    if arg['vararg']:
                        vba.write("argsArray(" + str(j) + " + k - LBound(" + vararg + ")) = " + argname + "\n")
                        vba.write("Next k\n")
                    else:
                        if vararg != "":
                            vba.write("argsArray(" + str(j) + ") = " + argname + "\n")
                            j += 1

                if vararg != '':
                    args_vba = 'argsArray'
                else:
                    args_vba = 'Array(' + ', '.join(arg['vba'] or arg['name'] for arg in xlfunc['args']) + ')'

                if ftype == "Sub":
                    vba.write('Py.CallUDF GetUdfModules, "{fname}", {args_vba}, ThisWorkbook\n',
                        fname=fname,
                        args_vba=args_vba,
                    )
                else:
                    vba.write('{fname} = Py.CallUDF(GetUdfModules, "{fname}", {args_vba}, ThisWorkbook)\n',
                        fname=fname,
                        args_vba=args_vba,
                    )

                if ftype == "Function":
                    vba.write("Exit " + ftype + "\n")
                    vba.write_label("failed")
                    vba.write(fname + " = Err.Description\n")

            vba.write('End ' + ftype + "\n")
            vba.write("\n")


def import_udfs(module_names, xl_workbook):
    module_vars = {}
    for module in module_names.split(';'):
        source_module = __import__(module)
        module_vars.update(udf_script(source_module.__file__))

    tf = tempfile.NamedTemporaryFile(mode='w', delete=False)
    generate_vba_wrapper(module_vars, tf.file)
    tf.close()

    try:
        xl_workbook.VBProject.VBComponents.Remove(xl_workbook.VBProject.VBComponents("xlwings_udfs"))
    except:
        pass
    xl_workbook.VBProject.VBComponents.Import(tf.name)

    for mvar in module_vars.values():
        if hasattr(mvar, '__xlfunc__'):
            xlfunc = mvar.__xlfunc__
            xlret = xlfunc['ret']
            xlargs = xlfunc['args']
            fname = xlfunc['name']
            fdoc = xlret['doc'][:255]
            n_args = 0
            for arg in xlargs:
                if not arg['vba']:
                    n_args += 1

            excel_version = [int(x) for x in re.split("[,\\.]", xl_workbook.Application.Version)]
            if n_args > 0 and excel_version[0] >= 14:
                argdocs = []
                for arg in xlargs:
                    if not arg['vba']:
                        argdocs.append(arg['doc'][:255])
                xl_workbook.Application.MacroOptions("'" + xl_workbook.Name + "'!" + fname, Description=fdoc, ArgumentDescriptions=argdocs)
            else:
                xl_workbook.Application.MacroOptions("'" + xl_workbook.Name + "'!" + fname, Description=fdoc)

    # try to delete the temp file - doesn't matter too much if it fails
    try:
        os.unlink(tf.name)
    except:
        pass
