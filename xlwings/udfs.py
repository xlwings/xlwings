import os
import re
import os.path
import tempfile


from . import conversion
from .utils import VBAWriter
from . import xlplatform
from . import Range

def xlfunc(f=None, **kwargs):
    def inner(f):
        if not hasattr(f, "__xlfunc__"):
            xlf = f.__xlfunc__ = {}
            xlf["name"] = f.__name__
            xlf["sub"] = False
            xlargs = xlf["args"] = []
            xlargmap = xlf["argmap"] = {}
            nArgs = f.__code__.co_argcount
            if f.__code__.co_flags & 4:  # function has an '*args' argument
                nArgs += 1
            for vpos, vname in enumerate(f.__code__.co_varnames[:nArgs]):
                xlargs.append({
                    "name": vname,
                    "pos": vpos,
                    "marshal": "var",
                    "vba": None,
                    "range": False,
                    "dtype": None,
                    "ndim": None,
                    "doc": "Positional argument " + str(vpos+1),
                    "vararg": True if vpos == f.__code__.co_argcount else False
                })
                xlargmap[vname] = xlargs[-1]
            xlf["ret"] = {
                "marshal": "var",
                "lax": True,
                "doc": f.__doc__ if f.__doc__ is not None else "Python function '" + f.__name__ + "' defined in '" + str(f.__code__.co_filename) + "'."
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


def xlret(**kwargs):
    def inner(f):
        xlf = xlfunc(f).__xlfunc__
        xlr = xlf["ret"]
        xlr.update(kwargs)
        return f
    return inner


def xlarg(arg, **kwargs):
    def inner(f):
        xlf = xlfunc(f).__xlfunc__
        if arg not in xlf["argmap"]:
            raise Exception("Invalid argument name '" + arg + "'.")
        xla = xlf["argmap"][arg]
        xla.update(kwargs)
        return f
    return inner


udf_scripts = {}
def udf_script(filename):
    filename = filename.lower()
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


def call_udf(script_name, func_name, args, this_workbook):
    script = udf_script(script_name)
    func = script[func_name]

    func_info = func.__xlfunc__
    args_info = func_info['args']
    ret_info = func_info['ret']

    args = list(args)
    for i, arg in enumerate(args):
        arg_info = args_info[i]
        args[i] = conversion.read_from_range(Range(arg), arg_info)

    xlplatform.xl_workbook_current = this_workbook
    ret = func(*args)

    return conversion.write_to_range(ret, None, ret_info)


def generate_vba_wrapper(script_vars, f):

    vba = VBAWriter(f)

    vba.writeln('Attribute VB_Name = "xlwings_udfs"')

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
                    if arg['vararg']:
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
                    as_ = arg.get('as_', None)
                    converter = conversion.accessors.get(as_, as_)

                    argname = arg['name']
                    if arg['vararg']:
                        vba.write("For k = LBound(" + vararg + ") To UBound(" + vararg + ")\n")
                        argname = vararg + "(k)"

                    #converter.vba_read(vba, argname, arg)

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
                    vba.write('Py.CallMacro PyScriptPath, "{fname}", {args_vba}, ThisWorkbook\n',
                        fname=fname,
                        args_vba=args_vba,
                    )
                else:
                    vba.write('{fname} = Py.CallUDF(PyScriptPath, "{fname}", {args_vba}, ThisWorkbook)\n',
                        fname=fname,
                        args_vba=args_vba,
                    )

                if ftype == "Function":
                    vba.write("Exit " + ftype + "\n")
                    vba.write_label("failed")
                    vba.write(fname + " = Err.Description\n")

            vba.write('End ' + ftype + "\n")
            vba.write("\n")


def import_udfs(script_path, xl_workbook):

    script_vars = udf_script(script_path)

    tf = tempfile.NamedTemporaryFile(mode='w', delete=False)
    generate_vba_wrapper(script_vars, tf.file)
    tf.close()

    try:
        xl_workbook.VBProject.VBComponents.Remove(xl_workbook.VBProject.VBComponents("xlwings_udfs"))
    except:
        pass
    xl_workbook.VBProject.VBComponents.Import(tf.name)

    for svar in script_vars.values():
        if hasattr(svar, '__xlfunc__'):
            xlfunc = svar.__xlfunc__
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