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