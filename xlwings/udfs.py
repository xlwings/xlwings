import os
import sys
import re
import os.path
import tempfile
import inspect
from importlib import import_module

from win32com.client import Dispatch

from . import conversion
from .utils import VBAWriter
from . import xlplatform
from . import Range

from . import PY3

if PY3:
    try:
        from imp import reload
    except:
        from importlib import reload

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


def get_category(**func_kwargs):
    if 'category' in func_kwargs:
        category = func_kwargs.pop('category')
        if isinstance(category, int):
            if 1 <= category <= 14:
                return category
            raise Exception(
                'There is only 14 build-in categories available in Excel. Please use a string value to specify a custom category.')
        if isinstance(category, str):
            return category[:255]
        raise Exception(
            'Category {0} should either be a predefined Excel category (int value) or a custom one (str value).'.format(
                category))
    return "xlwings"  # Default category


def should_call_in_wizard(**func_kwargs):
    if 'call_in_wizard' in func_kwargs:
        call_in_wizard = func_kwargs.pop('call_in_wizard')
        if isinstance(call_in_wizard, bool):
            return call_in_wizard
        raise Exception('call_in_wizard only takes boolean values ("{0}" provided).'.format(call_in_wizard))
    return True


def check_volatile(**func_kwargs):
    if 'volatile' in func_kwargs:
        volatile = func_kwargs.pop('volatile')
        if isinstance(volatile, bool):
            return volatile
        raise Exception('volatile only takes boolean values ("{0}" provided).'.format(volatile))
    return False


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
                    "doc": "Positional argument " + str(vpos + 1),
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
        f.__xlfunc__["category"] = get_category(**kwargs)
        f.__xlfunc__['call_in_wizard'] = should_call_in_wizard(**kwargs)
        f.__xlfunc__['volatile'] = check_volatile(**kwargs)
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


udf_modules = {}


class DelayWrite(object):
    def __init__(self, rng, options, value, caller):
        self.range = rng
        self.options = options
        self.value = value
        self.skip = (caller.Rows.Count, caller.Columns.Count)
        self.nb_remaining_call = 10

    def __call__(self, *args, **kwargs):
        self.nb_remaining_call -= 1
        conversion.write(
            self.value,
            self.range,
            conversion.Options(self.options)
            .override(_skip_tl_cells=self.skip)
        )


def get_udf_module(module_name):
    module_info = udf_modules.get(module_name, None)
    if module_info is not None:
        mtime = os.path.getmtime(module_info['filename'])
        module = module_info['module']
        if mtime == module_info['filetime']:
            return module
        else:
            module = reload(module)
            module_info['filetime'] = mtime
            module_info['module'] = module
            return module
    else:
        if sys.version_info[:2] < (2, 7):
            # For Python 2.6. we don't handle modules in subpackages
            module = __import__(module_name)
        else:
            module = import_module(module_name)
        filename = os.path.normcase(module.__file__.lower())
        mtime = os.path.getmtime(filename)
        udf_modules[module_name] = {
            'filename': filename,
            'filetime': mtime,
            'module': module
        }
        return module


def call_udf(module_name, func_name, args, this_workbook, caller):

    module = get_udf_module(module_name)

    func = getattr(module, func_name)

    func_info = func.__xlfunc__
    args_info = func_info['args']
    ret_info = func_info['ret']

    writing = func_info.get('writing', None)
    if writing and writing == caller.Address:
        return func_info['rval']

    output_param_indices = []

    args = list(args)
    for i, arg in enumerate(args):
        arg_info = args_info[min(i, len(args_info)-1)]
        if type(arg) is int and arg == -2147352572:      # missing
            args[i] = arg_info.get('optional', None)
        elif xlplatform.is_range_instance(arg):
            if arg_info.get('output', False):
                output_param_indices.append(i)
                args[i] = OutputParameter(Range(impl=xlplatform.Range(xl=arg)), arg_info['options'], func, caller)
            else:
                args[i] = conversion.read(Range(impl=xlplatform.Range(xl=arg)), None, arg_info['options'])
        else:
            args[i] = conversion.read(None, arg, arg_info['options'])

    xlplatform.BOOK_CALLER = Dispatch(this_workbook)
    ret = func(*args)

    if ret_info['options'].get('expand', None):
        from .server import add_idle_task
        add_idle_task(DelayWrite(Range(impl=xlplatform.Range(xl=caller)), ret_info['options'], ret, caller))

    return conversion.write(ret, None, ret_info['options'])


def generate_vba_wrapper(module_name, module, f):

    vba = VBAWriter(f)

    vba.writeln('Attribute VB_Name = "xlwings_udfs"')

    vba.writeln("'Autogenerated code by xlwings - changes will be lost with next import!")

    for svar in map(lambda attr: getattr(module, attr), dir(module)):
        if hasattr(svar, '__xlfunc__'):
            xlfunc = svar.__xlfunc__
            xlret = xlfunc['ret']
            fname = xlfunc['name']
            call_in_wizard = xlfunc['call_in_wizard']
            volatile = xlfunc['volatile']

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
                    if not call_in_wizard:
                        vba.writeln('If (Not Application.CommandBars("Standard").Controls(1).Enabled) Then Exit Function')
                    if volatile:
                        vba.writeln('Application.Volatile')
                    vba.writeln("If TypeOf Application.Caller Is Range Then On Error GoTo failed")

                if vararg != '':
                    vba.writeln("ReDim argsArray(1 to UBound(" + vararg + ") - LBound(" + vararg + ") + " + str(n_args) + ")")

                j = 1
                for arg in xlfunc['args']:
                    argname = arg['name']
                    if arg['vararg']:
                        vba.writeln("For k = LBound(" + vararg + ") To UBound(" + vararg + ")")
                        argname = vararg + "(k)"

                    if arg['vararg']:
                        vba.writeln("argsArray(" + str(j) + " + k - LBound(" + vararg + ")) = " + argname)
                        vba.writeln("Next k")
                    else:
                        if vararg != "":
                            vba.writeln("argsArray(" + str(j) + ") = " + argname)
                            j += 1

                if vararg != '':
                    args_vba = 'argsArray'
                else:
                    args_vba = 'Array(' + ', '.join(arg['vba'] or arg['name'] for arg in xlfunc['args']) + ')'

                if ftype == "Sub":
                    vba.writeln('Py.CallUDF "{module_name}", "{fname}", {args_vba}, ThisWorkbook, Application.Caller',
                        module_name=module_name,
                        fname=fname,
                        args_vba=args_vba,
                    )
                else:
                    vba.writeln('{fname} = Py.CallUDF("{module_name}", "{fname}", {args_vba}, ThisWorkbook, Application.Caller)',
                        module_name=module_name,
                        fname=fname,
                        args_vba=args_vba,
                    )

                if ftype == "Function":
                    vba.writeln("Exit " + ftype)
                    vba.write_label("failed")
                    vba.writeln(fname + " = Err.Description")

            vba.writeln('End ' + ftype)
            vba.writeln('')


def import_udfs(module_names, xl_workbook):
    module_names = module_names.split(';')

    tf = tempfile.NamedTemporaryFile(mode='w', delete=False)

    for module_name in module_names:
        module = get_udf_module(module_name)
        generate_vba_wrapper(module_name, module, tf.file)

    tf.close()

    try:
        xl_workbook.VBProject.VBComponents.Remove(xl_workbook.VBProject.VBComponents("xlwings_udfs"))
    except:
        pass
    xl_workbook.VBProject.VBComponents.Import(tf.name)

    for module_name in module_names:
        module = get_udf_module(module_name)
        for mvar in map(lambda attr: getattr(module, attr), dir(module)):
            if hasattr(mvar, '__xlfunc__'):
                xlfunc = mvar.__xlfunc__
                xlret = xlfunc['ret']
                xlargs = xlfunc['args']
                fname = xlfunc['name']
                fdoc = xlret['doc'][:255]
                fcategory = xlfunc['category']

                excel_version = [int(x) for x in re.split("[,\\.]", xl_workbook.Application.Version)]
                if excel_version[0] >= 14:
                    argdocs = [arg['doc'][:255] for arg in xlargs if not arg['vba']]
                    xl_workbook.Application.MacroOptions("'" + xl_workbook.Name + "'!" + fname,
                                                         Description=fdoc,
                                                         HasMenu=False,
                                                         MenuText=None,
                                                         HasShortcutKey=False,
                                                         ShortcutKey=None,
                                                         Category=fcategory,
                                                         StatusBar=None,
                                                         HelpContextID=None,
                                                         HelpFile=None,
                                                         ArgumentDescriptions=argdocs if argdocs else None)
                else:
                    xl_workbook.Application.MacroOptions("'" + xl_workbook.Name + "'!" + fname, Description=fdoc)

    # try to delete the temp file - doesn't matter too much if it fails
    try:
        os.unlink(tf.name)
    except:
        pass
