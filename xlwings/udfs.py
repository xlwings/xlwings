import os
import os.path
import sys
import re
import tempfile
import inspect
from importlib import import_module
from threading import Thread
from importlib import reload  # requires >= py 3.4

from win32com.client import Dispatch, CDispatch

from . import conversion, xlplatform, Range, apps
from .utils import VBAWriter


cache = {}


class AsyncThread(Thread):
    def __init__(self, pid, book_name, sheet_name, address, func, args, cache_key, expand):
        Thread.__init__(self)
        self.pid = pid
        self.book = book_name
        self.sheet = sheet_name
        self.address = address
        self.func = func
        self.args = args
        self.cache_key = cache_key
        self.expand = expand

    def run(self):
        cache[self.cache_key] = self.func(*self.args)

        if self.expand:
            apps[self.pid].books[self.book].sheets[self.sheet][self.address].formula_array = \
                apps[self.pid].books[self.book].sheets[self.sheet][self.address].formula_array
        else:
            apps[self.pid].books[self.book].sheets[self.sheet][self.address].formula = \
                apps[self.pid].books[self.book].sheets[self.sheet][self.address].formula


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


def get_async_mode(**func_kwargs):
    if 'async_mode' in func_kwargs:
        value = func_kwargs.pop('async_mode')
        if value in ['threading']:
            return value
        raise Exception('The only supported async_mode mode is currently "threading".')
    else:
        return None


def check_bool(kw, **func_kwargs):
    if kw in func_kwargs:
        check = func_kwargs.pop(kw)
        if isinstance(check, bool):
            return check
        raise Exception('{0} only takes boolean values. ("{1}" provided).'.format(kw, check))
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
        f.__xlfunc__['call_in_wizard'] = check_bool('call_in_wizard', **kwargs)
        f.__xlfunc__['volatile'] = check_bool('volatile', **kwargs)
        f.__xlfunc__['async_mode'] = get_async_mode(**kwargs)
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


class DelayedResizeDynamicArrayFormula:
    def __init__(self, target_range, caller, needs_clearing):
        self.target_range = target_range
        self.caller = caller
        self.needs_clearing = needs_clearing

    def __call__(self, *args, **kwargs):
        formula = self.caller.FormulaArray
        if self.needs_clearing:
            self.caller.ClearContents()
        self.target_range.api.FormulaArray = formula


def get_udf_module(module_name):
    module_info = udf_modules.get(module_name, None)
    if module_info is not None:
        module = module_info['module']
        # If filetime is None, it's not reloadable
        if module_info['filetime'] is not None:
            mtime = os.path.getmtime(module_info['filename'])
            if mtime != module_info['filetime']:
                module = reload(module)
                module_info['filetime'] = mtime
                module_info['module'] = module
    else:
        if sys.version_info[:2] < (2, 7):
            # For Python 2.6. we don't handle modules in subpackages
            module = __import__(module_name)
        else:
            module = import_module(module_name)

        filename = os.path.normcase(module.__file__.lower())

        try:  # getmtime fails for zip imports and frozen modules
            mtime = os.path.getmtime(filename)
        except OSError:
            mtime = None

        udf_modules[module_name] = {
            'filename': filename,
            'filetime': mtime,
            'module': module
        }

    return module


def get_cache_key(func, args, caller):
    """only use this if function is called from cells, not VBA"""
    xw_caller = Range(impl=xlplatform.Range(xl=caller))
    return (func.__name__ + str(args) + str(xw_caller.sheet.book.app.pid) +
            xw_caller.sheet.book.name + xw_caller.sheet.name + xw_caller.address.split(':')[0])


def call_udf(module_name, func_name, args, this_workbook=None, caller=None):

    module = get_udf_module(module_name)
    func = getattr(module, func_name)
    func_info = func.__xlfunc__
    args_info = func_info['args']
    ret_info = func_info['ret']
    is_dynamic_array = ret_info['options'].get('expand')

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
    if this_workbook:
        xlplatform.BOOK_CALLER = Dispatch(this_workbook)

    if func_info['async_mode'] and func_info['async_mode'] == 'threading':
        cache_key = get_cache_key(func, args, caller)
        cached_value = cache.get(cache_key)
        if cached_value is not None:  # test against None as np arrays don't have a truth value
            if not is_dynamic_array:  # for dynamic arrays, the cache is cleared below
                del cache[cache_key]
            ret = cached_value
        else:
            # You can't pass pywin32 objects directly to threads
            xw_caller = Range(impl=xlplatform.Range(xl=caller))
            thread = AsyncThread(xw_caller.sheet.book.app.pid,
                                 xw_caller.sheet.book.name,
                                 xw_caller.sheet.name,
                                 xw_caller.address,
                                 func,
                                 args,
                                 cache_key,
                                 is_dynamic_array)
            thread.start()
            return [["#N/A waiting..." * xw_caller.columns.count] * xw_caller.rows.count]
    else:
        if is_dynamic_array:
            cache_key = get_cache_key(func, args, caller)
            cached_value = cache.get(cache_key)
            if cached_value is not None:
                ret = cached_value
            else:
                ret = func(*args)
                cache[cache_key] = ret
        else:
            ret = func(*args)

    xl_result = conversion.write(ret, None, ret_info['options'])

    if is_dynamic_array:
        current_size = (caller.Rows.Count, caller.Columns.Count)
        result_size = (1, 1)
        if type(xl_result) is list:
            result_height = len(xl_result)
            result_width = result_height and len(xl_result[0])
            result_size = (max(1, result_height), max(1, result_width))
        if current_size != result_size:
            from .server import add_idle_task
            add_idle_task(DelayedResizeDynamicArrayFormula(
                Range(impl=xlplatform.Range(xl=caller)).resize(*result_size),
                caller,
                current_size[0] > result_size[0] or current_size[1] > result_size[1]
            ))
        else:
            del cache[cache_key]

    return xl_result


def generate_vba_wrapper(module_name, module, f):

    vba = VBAWriter(f)

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

                if vararg != '':
                    vba.writeln("Dim argsArray() As Variant")
                    non_varargs = [arg['vba'] or arg['name'] for arg in xlfunc['args'] if not arg['vararg']]
                    vba.writeln("argsArray = Array(%s)" % tuple({', '.join(non_varargs)}))

                    vba.writeln("ReDim Preserve argsArray(0 to UBound(" + vararg + ") - LBound(" + vararg + ") + " + str(len(non_varargs)) + ")")
                    vba.writeln("For k = LBound(" + vararg + ") To UBound(" + vararg + ")")
                    vba.writeln("argsArray(" + str(len(non_varargs)) + " + k - LBound(" + vararg + ")) = " + argname + "(k)")
                    vba.writeln("Next k")

                    args_vba = 'argsArray'
                else:
                    args_vba = 'Array(' + ', '.join(arg['vba'] or arg['name'] for arg in xlfunc['args']) + ')'

                if ftype == "Sub":
                    with vba.block('#If App = "Microsoft Excel" Then'):
                        vba.writeln('Py.CallUDF "{module_name}", "{fname}", {args_vba}, ThisWorkbook, Application.Caller',
                                    module_name=module_name,
                                    fname=fname,
                                    args_vba=args_vba,
                                    )
                    with vba.block("#Else"):
                        vba.writeln('Py.CallUDF "{module_name}", "{fname}", {args_vba}',
                                    module_name=module_name,
                                    fname=fname,
                                    args_vba=args_vba,
                                    )
                    vba.writeln("#End If")
                else:
                    with vba.block('#If App = "Microsoft Excel" Then'):
                        vba.writeln("If TypeOf Application.Caller Is Range Then On Error GoTo failed")
                        vba.writeln('{fname} = Py.CallUDF("{module_name}", "{fname}", {args_vba}, ThisWorkbook, Application.Caller)',
                                    module_name=module_name,
                                    fname=fname,
                                    args_vba=args_vba,
                                    )
                        vba.writeln("Exit " + ftype)
                    with vba.block("#Else"):
                        vba.writeln('{fname} = Py.CallUDF("{module_name}", "{fname}", {args_vba})',
                                module_name=module_name,
                                fname=fname,
                                args_vba=args_vba,
                                )
                        vba.writeln("Exit " + ftype)
                    vba.writeln("#End If")

                    vba.write_label("failed")
                    vba.writeln(fname + " = Err.Description")

            vba.writeln('End ' + ftype)
            vba.writeln('')


def import_udfs(module_names, xl_workbook):
    module_names = module_names.split(';')

    tf = tempfile.NamedTemporaryFile(mode='w', delete=False)

    vba = VBAWriter(tf.file)

    vba.writeln('Attribute VB_Name = "xlwings_udfs"')

    vba.writeln("'Autogenerated code by xlwings - changes will be lost with next import!")
    vba.writeln("""#Const App = "Microsoft Excel" 'Adjust when using outside of Excel""")

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
