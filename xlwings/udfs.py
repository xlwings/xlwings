def xlfunc(f = None, **kwargs):
    def inner(f):
        if not hasattr(f, "__xlfunc__"):
            xlf = f.__xlfunc__ = {}
            xlf["name"] = f.__name__
            xlf["sub"] = False
            xlargs = xlf["args"] = []
            xlargmap = xlf["argmap"] = {}
            nArgs = f.__code__.co_argcount
            if f.__code__.co_flags & 4:		# function has an '*args' argument
                nArgs += 1
            for vpos, vname in enumerate(f.__code__.co_varnames[:nArgs]):
                xlargs.append({
                    "name": vname,
                    "pos": vpos,
                    "marshal": "var",
                    "vba": None,
                    "range": False,
                    "dtype": None,
                    "dims": -1,
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

def xlsub(f = None, **kwargs):
    def inner(f):
        f = xlfunc(**kwargs)(f)
        f.__xlfunc__["sub"] = True
        return f
    if f is None:
        return inner
    else:
        return inner(f)

xlretparams = set(("marshal", "lax", "doc"))
def xlret(marshal=None, **kwargs):
    if marshal is not None:
        kwargs["marshal"] = marshal
    def inner(f):
        xlf = xlfunc(f).__xlfunc__
        xlr = xlf["ret"]
        for k, v in kwargs.items():
            if k in xlretparams:
                xlr[k] = v
            else:
                raise Exception("Invalid parameter '" + k + "'.")
        return f
    return inner

xlargparams = set(("marshal", "dims", "dtype", "range", "doc", "vba"))
def xlarg(arg, marshal=None, dims=None, **kwargs):
    if marshal is not None:
        kwargs["marshal"] = marshal
    if dims is not None:
        kwargs["dims"] = dims
    def inner(f):
        xlf = xlfunc(f).__xlfunc__
        if arg not in xlf["argmap"]:
            raise Exception("Invalid argument name '" + arg + "'.")
        xla = xlf["argmap"][arg]
        for k, v in kwargs.items():
            if k in xlargparams:
                xla[k] = v
            else:
                raise Exception("Invalid parameter '" + k + "'.")
        return f
    return inner

udf_scripts = {}
def udf_script(filename):
    import os.path
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


def generate_vba_udf_wrapper(script_path):

    tab = '\t'

    import tempfile
    tf = tempfile.NamedTemporaryFile(mode='w', delete=False)
    f = tf.file

    f.write('Attribute VB_Name = "xlwings_udfs"\n')

    script_vars = udf_script(script_path)

    for svar in script_vars.values():
        if hasattr(svar, '__xlfunc__'):
            xlfunc = svar.__xlfunc__
            xlret = xlfunc['ret']
            fname = xlfunc['name']

            ftype = 'Sub' if xlfunc['sub'] else 'Function'

            f.write(ftype + " " + fname + "(")

            first = True
            vararg = ''
            n_args = len(xlfunc['args'])
            for arg in xlfunc['args']:
                if not arg['vba']:
                    argname = arg['name']
                    if not first:
                        f.write(', ')
                    if arg['vararg']:
                        f.write('ParamArray ')
                        vararg = argname
                    f.write(argname)
                    if arg['vararg']:
                        f.write('()')
                    first = False
            f.write(')\n')
            if ftype == 'Function':
                f.write(tab + "If TypeOf Application.Caller Is Range Then On Error GoTo failed\n")

            if vararg != '':
                f.write(tab + "ReDim argsArray(1 to UBound(" + vararg + ") - LBound(" + vararg + ") + " + str(n_args) + ")\n")

            j = 1
            for arg in xlfunc['args']:
                if not arg['vba']:
                    argname = arg['name']
                    if arg['vararg']:
                        f.write(tab + "For k = LBound(" + vararg + ") To UBound(" + vararg + ")\n")
                        argname = vararg + "(k)"
                    if not arg['range']:
                        f.write(tab + "If TypeOf " + argname + " Is Range Then " + argname + " = " + argname + ".Value2\n")
                    dims = arg['dims']
                    marshal = arg['marshal']
                    if dims != -2 or marshal == "nparray" or marshal == "list":
                        f.write(tab + "If Not TypeOf " + argname + " Is Object Then\n")
                        if dims != -2:
                            f.write(tab + tab + argname + " = NDims(" + argname + ", " + str(dims) + ")\n")
                        if arg['marshal'] == "nparray":
                            dtype = arg['dtype']
                            if dtype is None:
                                f.write(tab + tab + 'Set ' + argname + ' = Py.Call(Py.Module("numpy"), "array", Py.Tuple(' + argname + '))\n')
                            else:
                                f.write(tab + tab + 'Set ' + argname + ' = Py.Call(Py.Module("numpy"), "array", Py.Tuple(' + argname + ', "' + dtype + '"))\n')
                        elif marshal == 'list':
                            f.write(tab + tab + 'Set ' + argname + ' = Py.Call(Py.Eval("lambda t: [ list(x) if isinstance(x, tuple) else x for x in t ] if isinstance(t, tuple) else t"), Py.Tuple(' + argname + '))\n')
                        f.write(tab + "End If\n")
                    if arg['vararg']:
                        f.write(tab + "argsArray(" + str(j) + " + k - LBound(" + vararg + ")) = " + argname + "\n")
                        f.write(tab + "Next k\n")
                    else:
                        if vararg != "":
                            f.write(tab + "argsArray(" + str(j) + ") = " + argname + "\n")
                            j += 1

            if vararg != '':
                f.write(tab + "Set args = Py.TupleFromArray(argsArray)\n")
            else:
                f.write(tab + "Set args = Py.Tuple(")
                first = True
                for arg in xlfunc['args']:
                    if not first:
                        f.write(", ")
                    if not arg['vba']:
                        f.write(str(arg['name']))
                    else:
                        f.write(str(arg['vba']))
                    first = False
                f.write(")\n")

            f.write(tab + 'Set xlpy = Py.Module("xlwings")\n')
            f.write(tab + 'Set script = Py.Call(xlpy, "udf_script", Py.Tuple(PyScriptPath))\n')
            f.write(tab + 'Set func = Py.GetItem(script, "' + fname + '")\n')
            if ftype == "Sub":
                f.write(tab + 'Py.SetAttr Py.Module("xlwings._xlwindows"), "xl_workbook_current", ThisWorkbook\n')
                f.write(tab + "Py.Call func, args\n")
            else:
                f.write(tab + "Set " + fname + " = Py.Call(func, args)\n")
                marshal = xlret["marshal"]
                if marshal == "auto":
                    f.write(tab + "If TypeOf Application.Caller Is Range Then " + fname + " = Py.Var(" + fname + ", " + str(xlret["lax"]) + ")\n")
                elif marshal == "var":
                    f.write(tab + fname + " = Py.Var(" + fname + ", " + str(xlret["lax"]) + ")\n")
                elif marshal == "str":
                    f.write(tab + fname + " = Py.Str(" + fname + ")\n")

            if ftype == "Function":
                f.write(tab + "Exit " + ftype + "\n")
                f.write("failed:\n")
                f.write(tab + fname + " = Err.Description\n")
            f.write("End " + ftype + "\n")
            f.write("\n")

    # for svar in script_vars.values():
    #     if hasattr(svar, '__xlfunc__'):
    #         xlfunc = svar.__xlfunc__
    #         xlret = xlfunc['ret']
    #         xlargs = xlfunc['args']
    #         fname = xlfunc['name']
    #         fdoc = xlret['doc']
    #         n_args = 0
    #         for arg in xlargs:
    #             if not arg['vba']:
    #                 n_args += 1
    #
    #         if n_args > 0: # and Val(Application.Version) >= 14
    #             argdocs = []
    #             for args in xlargs:
    #                 if not arg['vba']:
    #                     argdocs.append(arg['doc'][:255])
    #             fdoc = fdoc[:255]
    #             # XLPMacroOptions2010 "'" + wb.Name + "'!" + fname, Left$(fdoc, 255), argdocs
    #         else:
    #             pass # Application.MacroOptions "'" + wb.Name + "'!" + fname, Description:=Left$(fdoc, 255)

    return tf.name

