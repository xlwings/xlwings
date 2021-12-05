# First of all see if we can load PyWin32
try:
    import _win32sysloader
except:
    raise Exception("Cannot import PyWin32. Are you sure it's installed?")
import logging
import sys
import os
import collections
import importlib
# Hack to find pythoncom.dll - needed for some distribution/setups
# E.g. if python is started with the full path outside of the python path, then it almost certainly fails
cwd = os.getcwd()
if not hasattr(sys, 'frozen'):
    # cx_Freeze etc. will fail here otherwise
    os.chdir(sys.exec_prefix)
import win32api
import win32event
os.chdir(cwd)

import pythoncom
import pywintypes
import win32com.client
import win32com.server.util as serverutil
import win32com.server.dispatcher
import win32com.server.policy

import asyncio
import threading

from .udfs import call_udf


# If no handler is configured, print is used to make the statements show up in the console that opens when using
# 'python' instead of 'pythonw' as the interpreter
logger = logging.getLogger(__name__)


class XLPythonOption:
    """ The XLPython class itself """
    def __init__(self, option, value):
        self.option = option
        self.value = value


class XLPythonObject:
    _public_methods_ = ['Item', 'Count']
    _public_attrs_ = ['_NewEnum']

    def __init__(self, obj):
        self.obj = obj

    def _NewEnum(self):
        return win32com.server.util.wrap(XLPythonEnumerator(self.obj), iid=pythoncom.IID_IEnumVARIANT)

    def Item(self, key):
        return ToVariant(self.obj[key])

    def Count(self):
        return len(self.obj)


class XLPythonEnumerator:
    _public_methods_ = ["Next", "Skip", "Reset", "Clone"]

    def __init__(self, gen):
        self.iter = gen.__iter__()

    def _query_interface_(self, iid):
        if iid == pythoncom.IID_IEnumVARIANT:
            return 1

    def Next(self, count):
        r = []
        try:
            r.append(ToVariant(next(self.iter)))
        except StopIteration:
            pass
        return r

    def Skip(self, count):
        raise win32com.server.exception.COMException(scode=0x80004001)  # E_NOTIMPL

    def Reset(self):
        raise win32com.server.exception.COMException(scode=0x80004001)  # E_NOTIMPL

    def Clone(self):
        raise win32com.server.exception.COMException(scode=0x80004001)  # E_NOTIMPL


PyIDispatch = pythoncom.TypeIIDs[pythoncom.IID_IDispatch]


def FromVariant(var):
    try:
        obj = win32com.server.util.unwrap(var).obj
    except:
        obj = var
    if type(obj) is PyIDispatch:
        obj = win32com.client.Dispatch(obj, userName=obj.GetTypeInfo().GetDocumentation(-1)[0])
    return obj


def ToVariant(obj):
    return win32com.server.util.wrap(XLPythonObject(obj))


class XLPython:
    _public_methods_ = ['Module', 'Tuple', 'TupleFromArray', 'Dict', 'DictFromArray', 'List', 'ListFromArray', 'Obj',
                        'Str', 'Var', 'Call', 'GetItem', 'SetItem', 'DelItem', 'Contains', 'GetAttr', 'SetAttr',
                        'DelAttr', 'HasAttr', 'Eval', 'Exec', 'ShowConsole', 'Builtin', 'Len', 'Bool',
                        'CallUDF']

    def ShowConsole(self):
        import ctypes
        import sys

        ctypes.windll.kernel32.AllocConsole()
        sys.stdout = open("CONOUT$", "a", 0)
        sys.stderr = open("CONOUT$", "a", 0)

    def Module(self, module, reload=False):
        vars = {}
        exec("import " + module + " as the_module", vars)
        m = vars["the_module"]
        if reload:
            m = importlib.reload(m)
        return ToVariant(m)

    def TupleFromArray(self, elements):
        return self.Tuple(*elements)

    def Tuple(self, *elements):
        return ToVariant(tuple((FromVariant(e) for e in elements)))

    def DictFromArray(self, kvpairs):
        return self.Dict(*kvpairs)

    def Dict(self, *kvpairs):
        if len(kvpairs) % 2 != 0:
            raise Exception("Arguments must be alternating keys and values.")
        n = int(len(kvpairs) / 2)
        d = {}
        for k in range(n):
            key = FromVariant(kvpairs[2 * k])
            value = FromVariant(kvpairs[2 * k + 1])
            d[key] = value
        return ToVariant(d)

    def ListFromArray(self, elements):
        return self.List(*elements)

    def List(self, *elements):
        return ToVariant(list((FromVariant(e) for e in elements)))

    def Obj(self, var, dispatch=True):
        return ToVariant(FromVariant(var, dispatch))

    def Str(self, obj):
        return str(FromVariant(obj))

    def Var(self, obj, lax=False):
        value = FromVariant(obj)
        if lax:
            t = type(value)
            if t is dict:
                value = tuple(value.items())
            elif t.__name__ == 'ndarray' and t.__module__ == 'numpy':
                value = value.tolist()
        if type(value) is tuple:
            return (value,)
        # elif isinstance(value, types.InstanceType) and value.__class__ is win32com.client.CDispatch:
        # return value._oleobj_
        else:
            return value

    def Call(self, obj, *args):
        obj = FromVariant(obj)
        method = None
        pargs = ()
        kwargs = {}
        for arg in args:
            arg = FromVariant(arg)
            if isinstance(arg, tuple):
                pargs = arg
            elif isinstance(arg, dict):
                kwargs = arg
            else:
                # assume string
                method = arg
        if method is None:
            return ToVariant(obj(*pargs, **kwargs))
        else:
            return ToVariant(getattr(obj, method)(*pargs, **kwargs))

    def CallUDF(self, script, fname, args, this_workbook=None, caller=None):
        args = tuple(FromVariant(arg) for arg in args)
        res = call_udf(script, fname, args, this_workbook, FromVariant(caller))
        if len(res) == 1 and len(res[0]) == 1:
            res = res[0][0]
        elif len(res) == 1 and len(res[0]) > 1:
            res = res[0]
        return res

    def Len(self, obj):
        obj = FromVariant(obj)
        return len(obj)

    def Bool(self, obj):
        obj = FromVariant(obj)
        if obj:
            return True
        else:
            return False

    def Builtin(self):
        import builtins
        return ToVariant(builtins)

    def GetItem(self, obj, key):
        obj = FromVariant(obj)
        key = FromVariant(key)
        return ToVariant(obj[key])

    def SetItem(self, obj, key, value):
        obj = FromVariant(obj)
        key = FromVariant(key)
        value = FromVariant(value)
        obj[key] = value

    def DelItem(self, obj, key):
        del obj[key]

    def Contains(self, obj, key):
        return key in obj

    def GetAttr(self, obj, attr):
        obj = FromVariant(obj)
        attr = FromVariant(attr)
        return ToVariant(getattr(obj, attr))

    def SetAttr(self, obj, attr, value):
        obj = FromVariant(obj)
        attr = FromVariant(attr)
        value = FromVariant(value)
        setattr(obj, attr, value)

    def HasAttr(self, obj, attr):
        obj = FromVariant(obj)
        attr = FromVariant(attr)
        return hasattr(obj, attr)

    def DelAttr(self, obj, attr):
        delattr(obj, attr)

    def Eval(self, expr, *args):
        globals = None
        locals = None
        for arg in args:
            arg = FromVariant(arg)
            if type(arg) is dict:
                if globals is None:
                    globals = arg
                elif locals is None:
                    locals = arg
                else:
                    raise Exception("Eval can be called with at most 2 dictionary arguments")
            else:
                pass
        return ToVariant(eval(expr, globals, locals))

    def Exec(self, stmt, *args):
        globals = None
        locals = None
        for arg in args:
            arg = FromVariant(arg)
            if type(arg) is dict:
                if globals is None:
                    globals = arg
                elif locals is None:
                    locals = arg
                else:
                    raise Exception("Exec can be called with at most 2 dictionary arguments")
            else:
                pass
        exec(stmt, globals, locals)


loop = asyncio.new_event_loop()


def _start_background_loop():
    pythoncom.CoInitialize()
    asyncio.set_event_loop(loop)
    loop.run_forever()


def serve(clsid="{506e67c3-55b5-48c3-a035-eed5deea7d6d}"):
    """Launch the COM server, clsid is the XLPython object class id """
    clsid = pywintypes.IID(clsid)

    # Ovveride CreateInstance in default policy to instantiate the XLPython object ---
    BaseDefaultPolicy = win32com.server.policy.DefaultPolicy

    class MyPolicy(BaseDefaultPolicy):
        def _CreateInstance_(self, reqClsid, reqIID):
            if reqClsid == clsid:
                return serverutil.wrap(XLPython(), reqIID)
            else:
                return BaseDefaultPolicy._CreateInstance_(self, clsid, reqIID)

    win32com.server.policy.DefaultPolicy = MyPolicy

    # Create the class factory and register it
    factory = pythoncom.MakePyFactory(clsid)

    clsctx = pythoncom.CLSCTX_LOCAL_SERVER
    flags = pythoncom.REGCLS_MULTIPLEUSE | pythoncom.REGCLS_SUSPENDED
    revokeId = pythoncom.CoRegisterClassObject(clsid, factory, clsctx, flags)

    pythoncom.EnableQuitMessage(win32api.GetCurrentThreadId())
    pythoncom.CoResumeClassObjects()

    if not loop.is_running():
        t = threading.Thread(
            target=_start_background_loop,
            daemon=True)
        t.start()
        tid = t.ident
    else:
        tid = None

    msg = 'xlwings server running, clsid=%s, event loop on %s'
    logger.info(msg, clsid, tid) if logger.hasHandlers() else print(msg % (clsid, tid))

    while True:
        rc = win32event.MsgWaitForMultipleObjects(
            (),
            0,
            win32event.INFINITE,
            win32event.QS_ALLEVENTS
        )
        if rc == win32event.WAIT_OBJECT_0:
            if pythoncom.PumpWaitingMessages():
                break  # wm_quit

    pythoncom.CoRevokeClassObject(revokeId)
    pythoncom.CoUninitialize()
