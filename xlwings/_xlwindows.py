import atexit
import locale
import os
import subprocess
import sys
from shlex import split

# Hack to find pythoncom.dll - needed for some distribution/setups (includes seemingly
# unused import win32api) E.g. if python is started with the full path outside of the
# python path, then it almost certainly fails
cwd = os.getcwd()
if not hasattr(sys, "frozen"):
    # cx_Freeze etc. will fail here otherwise
    os.chdir(sys.exec_prefix)
# Since Python 3.8, pywintypes needs to be imported before win32api or you get
# ImportError: DLL load failed while importing win32api: The specified module could not
# be found.
# See: https://stackoverflow.com/questions/58805040/pywin32-226-and-virtual-environments
# Seems to be required even with pywin32 227
import pywintypes
import win32api
import win32con

os.chdir(cwd)

import ctypes
import datetime as dt
import numbers
import types
from ctypes import PyDLL, byref, oledll, py_object, windll
from pathlib import Path
from warnings import warn

import pythoncom

# Patching CoClassBaseClass, see https://github.com/xlwings/xlwings/issues/1789
import win32com.client

from ._win32patch import CoClassBaseClass

win32com.client.CoClassBaseClass = CoClassBaseClass
# End Patch

import win32gui
import win32process
import win32timezone
from win32com.client import (
    CDispatch,
    CoClassBaseClass,
    Dispatch,
    DispatchBaseClass,
    DispatchEx,
)

import xlwings

from . import base_classes, constants, utils
from .constants import (
    AxisGroup,
    AxisType,
    ColorIndex,
    ConsolidationFunction,
    DeleteShiftDirection,
    FileFormat,
    FixedFormatType,
    HAlign,
    HtmlType,
    InsertFormatOrigin,
    InsertShiftDirection,
    LayoutFormType,
    LayoutRowType,
    LegendPosition,
    ListObjectSourceType,
    PivotFieldOrientation,
    PivotTableSourceType,
    ReferenceStyle,
    RowCol,
    SourceType,
    UpdateLinks,
    VAlign,
)
from .utils import (
    col_name,
    fullname_url_to_local_path,
    hex_to_rgb,
    int_to_rgb,
    np_datetime_to_datetime,
    read_config_sheet,
    rgb_to_int,
)

# Optional imports
try:
    import pandas as pd
except ImportError:
    pd = None
try:
    import numpy as np
except ImportError:
    np = None
try:
    from PIL import ImageGrab
except ImportError:
    PIL = None


time_types = (dt.date, dt.datetime, pywintypes.TimeType)
if np:
    time_types = time_types + (np.datetime64,)


N_COM_ATTEMPTS = 0  # 0 means try indefinitely
BOOK_CALLER = None
missing = object()


@atexit.register
def cleanup():
    """Clear up any zombie processes"""
    try:
        Apps.cleanup()
    except:  # noqa: E722
        pass


class COMRetryMethodWrapper:
    def __init__(self, method):
        self.__method = method

    def __call__(self, *args, **kwargs):
        n_attempt = 1
        while True:
            try:
                v = self.__method(*args, **kwargs)
                if isinstance(v, (CDispatch, CoClassBaseClass, DispatchBaseClass)):
                    return COMRetryObjectWrapper(v)
                elif isinstance(v, types.MethodType):
                    return COMRetryMethodWrapper(v)
                else:
                    return v
            except pywintypes.com_error as e:
                if (
                    not N_COM_ATTEMPTS or n_attempt < N_COM_ATTEMPTS
                ) and e.hresult == -2147418111:
                    n_attempt += 1
                    continue
                else:
                    raise
            except AttributeError:
                if not N_COM_ATTEMPTS or n_attempt < N_COM_ATTEMPTS:
                    n_attempt += 1
                    continue
                else:
                    raise


class ExcelBusyError(Exception):
    def __init__(self):
        super(ExcelBusyError, self).__init__("Excel application is not responding")


class COMRetryObjectWrapper:
    def __init__(self, inner):
        object.__setattr__(self, "_inner", inner)

    def __repr__(self):
        return repr(self._inner)

    def __setattr__(self, key, value):
        n_attempt = 1
        while True:
            try:
                return setattr(self._inner, key, value)
            except pywintypes.com_error as e:
                hresult, msg, exc, arg = e.args
                if exc:
                    wcode, source, text, help_file, help_id, scode = exc
                else:
                    wcode, source, text, help_file, help_id, scode = (  # noqa: F841
                        None,
                        None,
                        None,
                        None,
                        None,
                        None,
                    )
                # -2147352567 is the error you get when clicking into cells. If we
                # wouldn't check for scode, actions like renaming a sheet with >31
                # characters would be tried forever, causing xlwings to hang (they
                # also have hresult -2147352567).
                if (
                    (not N_COM_ATTEMPTS or n_attempt < N_COM_ATTEMPTS)
                    and e.hresult in [-2147418111, -2147352567]
                    and scode in [None, -2146777998]
                ):
                    n_attempt += 1
                    continue
                else:
                    raise
            except AttributeError:
                if not N_COM_ATTEMPTS or n_attempt < N_COM_ATTEMPTS:
                    n_attempt += 1
                    continue
                else:
                    raise

    def __getattr__(self, item):
        n_attempt = 1
        while True:
            try:
                v = getattr(self._inner, item)
                if isinstance(v, (CDispatch, CoClassBaseClass, DispatchBaseClass)):
                    return COMRetryObjectWrapper(v)
                elif isinstance(v, types.MethodType):
                    return COMRetryMethodWrapper(v)
                else:
                    return v
            except pywintypes.com_error as e:
                if (
                    not N_COM_ATTEMPTS or n_attempt < N_COM_ATTEMPTS
                ) and e.hresult == -2147418111:
                    n_attempt += 1
                    continue
                else:
                    raise
            except AttributeError:
                # pywin32 reacts incorrectly to RPC_E_CALL_REJECTED (i.e. assumes
                # attribute doesn't exist, thus not allowing to distinguish between
                # cases where attribute really doesn't exist or error is only being
                # thrown because the COM RPC server is busy). Here we try to test to
                # see what's going on really
                try:
                    self._oleobj_.GetIDsOfNames(0, item)
                except pythoncom.ole_error as e:
                    if e.hresult != -2147418111:  # RPC_E_CALL_REJECTED
                        # attribute probably really doesn't exist
                        raise
                if not N_COM_ATTEMPTS or n_attempt < N_COM_ATTEMPTS:
                    n_attempt += 1
                    continue
                else:
                    raise ExcelBusyError()

    def __call__(self, *args, **kwargs):
        n_attempt = 1
        for i in range(N_COM_ATTEMPTS + 1):
            try:
                v = self._inner(*args, **kwargs)
                if isinstance(v, (CDispatch, CoClassBaseClass, DispatchBaseClass)):
                    return COMRetryObjectWrapper(v)
                elif isinstance(v, types.MethodType):
                    return COMRetryMethodWrapper(v)
                else:
                    return v
            except pywintypes.com_error as e:
                if (
                    not N_COM_ATTEMPTS or n_attempt < N_COM_ATTEMPTS
                ) and e.hresult == -2147418111:
                    n_attempt += 1
                    continue
                else:
                    raise
            except AttributeError:
                if not N_COM_ATTEMPTS or n_attempt < N_COM_ATTEMPTS:
                    n_attempt += 1
                    continue
                else:
                    raise

    def __iter__(self):
        for v in self._inner:
            if isinstance(v, (CDispatch, CoClassBaseClass, DispatchBaseClass)):
                yield COMRetryObjectWrapper(v)
            else:
                yield v


# Constants
OBJID_NATIVEOM = -16


class _GUID(ctypes.Structure):
    # https://docs.microsoft.com/en-us/openspecs/windows_protocols/
    #  ms-dtyp/49e490b8-f972-45d6-a3a4-99f924998d97
    _fields_ = [
        ("Data1", ctypes.c_ulong),
        ("Data2", ctypes.c_ushort),
        ("Data3", ctypes.c_ushort),
        ("Data4", ctypes.c_byte * 8),
    ]


_IDISPATCH_GUID = _GUID()
oledll.ole32.CLSIDFromString(
    "{00020400-0000-0000-C000-000000000046}", byref(_IDISPATCH_GUID)
)


def accessible_object_from_window(hwnd):
    # ptr is a pointer to an IDispatch:
    # https://docs.microsoft.com/en-us/windows/win32/api/oaidl/nn-oaidl-idispatch
    # We don't bother using ctypes.POINTER(comtypes.automation.IDispatch)()
    # because we won't dereference the pointer except through pywin32's
    # pythoncom.PyCom_PyObjectFromIUnknown below in get_xl_app_from_hwnd().
    ptr = ctypes.c_void_p()
    res = oledll.oleacc.AccessibleObjectFromWindow(  # noqa: F841
        hwnd, OBJID_NATIVEOM, byref(_IDISPATCH_GUID), byref(ptr)
    )
    return ptr


def is_hwnd_xl_app(hwnd):
    try:
        child_hwnd = win32gui.FindWindowEx(hwnd, 0, "XLDESK", None)
        child_hwnd = win32gui.FindWindowEx(child_hwnd, 0, "EXCEL7", None)
        ptr = accessible_object_from_window(child_hwnd)  # noqa: F841
        return True
    except WindowsError:
        return False
    except pywintypes.error:
        return False


_PyCom_PyObjectFromIUnknown = PyDLL(pythoncom.__file__).PyCom_PyObjectFromIUnknown
_PyCom_PyObjectFromIUnknown.restype = py_object


def get_xl_app_from_hwnd(hwnd):
    pythoncom.CoInitialize()
    child_hwnd = win32gui.FindWindowEx(hwnd, 0, "XLDESK", None)
    child_hwnd = win32gui.FindWindowEx(child_hwnd, 0, "EXCEL7", None)

    ptr = accessible_object_from_window(child_hwnd)
    p = _PyCom_PyObjectFromIUnknown(ptr, byref(_IDISPATCH_GUID), True)
    disp = COMRetryObjectWrapper(Dispatch(p))
    return disp.Application


def get_excel_hwnds():
    pythoncom.CoInitialize()
    hwnd = windll.user32.GetTopWindow(None)
    pids = set()
    while hwnd:
        try:
            # Apparently, this fails on some systems when Excel is closed
            child_hwnd = win32gui.FindWindowEx(hwnd, 0, "XLDESK", None)
            if child_hwnd:
                child_hwnd = win32gui.FindWindowEx(child_hwnd, 0, "EXCEL7", None)
            if child_hwnd:
                pid = win32process.GetWindowThreadProcessId(hwnd)[1]
                if pid not in pids:
                    pids.add(pid)
                    yield hwnd
        except pywintypes.error:
            pass

        hwnd = windll.user32.GetWindow(hwnd, 2)  # 2 = next window according to Z-order


def get_xl_apps():
    for hwnd in get_excel_hwnds():
        try:
            yield get_xl_app_from_hwnd(hwnd)
        except ExcelBusyError:
            pass
        except WindowsError:
            # This happens if the bare Excel Application is open without Workbook, i.e.,
            # there's no 'EXCEL7' child hwnd that would be necessary for a connection
            pass


def is_range_instance(xl_range):
    pyid = getattr(xl_range, "_oleobj_", None)
    if pyid is None:
        return False
    return xl_range._oleobj_.GetTypeInfo().GetTypeAttr().iid == pywintypes.IID(
        "{00020846-0000-0000-C000-000000000046}"
    )
    # return pyid.GetTypeInfo().GetDocumentation(-1)[0] == 'Range'


def _com_time_to_datetime(com_time, datetime_builder):
    return datetime_builder(
        month=com_time.month,
        day=com_time.day,
        year=com_time.year,
        hour=com_time.hour,
        minute=com_time.minute,
        second=com_time.second,
        microsecond=com_time.microsecond,
        tzinfo=None,
    )


def _datetime_to_com_time(dt_time):
    """
    This function is a modified version from Pyvot (https://pypi.python.org/pypi/Pyvot)
    and subject to the following copyright:

    Copyright (c) Microsoft Corporation.

    This source code is subject to terms and conditions of the Apache License,
    Version 2.0. A copy of the license can be found in the LICENSE.txt file at the root
    of this distribution. If you cannot locate the Apache License, Version 2.0, please
    send an email to vspython@microsoft.com. By using this source code in any fashion,
    you are agreeing to be bound by the terms of the Apache License, Version 2.0.

    You must not remove this notice, or any other, from this software.

    """
    # Convert date to datetime
    if pd and isinstance(dt_time, type(pd.NaT)):
        return ""
    if np:
        if type(dt_time) is np.datetime64:
            dt_time = np_datetime_to_datetime(dt_time)

    if type(dt_time) is dt.date:
        dt_time = dt.datetime(
            dt_time.year,
            dt_time.month,
            dt_time.day,
            tzinfo=win32timezone.TimeZoneInfo.utc(),
        )

    # pywintypes has its time type inherit from datetime.
    # For some reason, though it accepts plain datetimes, they must have a timezone set.
    # See http://docs.activestate.com/activepython/2.7/pywin32/html/win32/help/py3k.html
    # We replace no timezone -> UTC to allow round-trips in the naive case
    if pd and isinstance(dt_time, pd.Timestamp):
        # Otherwise pandas prints ignored exceptions on Python 3
        dt_time = dt_time.to_pydatetime()
    # We don't use pytz.utc to get rid of additional dependency
    # Don't do any timezone transformation: simply cutoff the tz info
    # If we don't reset it first, it gets transformed into UTC before sending to Excel
    dt_time = dt_time.replace(tzinfo=None)
    dt_time = dt_time.replace(tzinfo=win32timezone.TimeZoneInfo.utc())

    return dt_time


cell_errors = {
    -2146826281: "#DIV/0!",
    -2146826246: "#N/A",
    -2146826259: "#NAME?",
    -2146826288: "#NULL!",
    -2146826252: "#NUM!",
    -2146826265: "#REF!",
    -2146826273: "#VALUE!",
}


def _clean_value_data_element(
    value, datetime_builder, empty_as, number_builder, err_to_str
):
    if value in ("", None):
        return empty_as
    elif isinstance(value, time_types):
        return _com_time_to_datetime(value, datetime_builder)
    elif number_builder is not None and isinstance(value, float):
        value = number_builder(value)
    elif isinstance(value, int) and value in cell_errors:
        if err_to_str:
            return cell_errors[value]
        else:
            return None
    return value


class Engine:
    @property
    def apps(self):
        return Apps()

    @property
    def name(self):
        return "excel"

    @property
    def type(self):
        return "desktop"

    @staticmethod
    def prepare_xl_data_element(x, date_format):
        if isinstance(x, time_types):
            return _datetime_to_com_time(x)
        elif pd and pd.isna(x):
            return ""
        elif np and isinstance(x, (np.floating, float)) and np.isnan(x):
            return ""
        elif np and isinstance(x, np.number):
            return float(x)
        elif x is None:
            return ""
        else:
            return x

    @staticmethod
    def clean_value_data(data, datetime_builder, empty_as, number_builder, err_to_str):
        return [
            [
                _clean_value_data_element(
                    c, datetime_builder, empty_as, number_builder, err_to_str
                )
                for c in row
            ]
            for row in data
        ]


engine = Engine()


class Apps(base_classes.Apps):
    def keys(self):
        k = []
        for hwnd in get_excel_hwnds():
            k.append(App(xl=hwnd).pid)
        return k

    def add(self, spec=None, add_book=None, xl=None, visible=None):
        return App(spec=spec, add_book=add_book, xl=xl, visible=visible)

    @staticmethod
    def cleanup():
        res = subprocess.run(
            split('tasklist /FI "IMAGENAME eq EXCEL.exe"'),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            creationflags=subprocess.CREATE_NO_WINDOW,
            encoding=locale.getpreferredencoding(),
        )

        all_pids = set()
        for line in res.stdout.splitlines()[3:]:
            # Ignored if there's no processes as it prints only 1 line
            _, pid, _, _, _, _ = line.split()
            all_pids.add(int(pid))

        active_pids = {app.pid for app in xlwings.apps}
        zombie_pids = all_pids - active_pids

        for pid in zombie_pids:
            subprocess.run(
                split(f"taskkill /PID {pid} /F"),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                creationflags=subprocess.CREATE_NO_WINDOW,
                encoding=locale.getpreferredencoding(),
            )

    def __iter__(self):
        for hwnd in get_excel_hwnds():
            yield App(xl=hwnd)

    def __len__(self):
        return len(list(get_excel_hwnds()))

    def __getitem__(self, pid):
        for hwnd in get_excel_hwnds():
            app = App(xl=hwnd)
            if app.pid == pid:
                return app
        raise KeyError("Could not find an Excel instance with this PID.")


class App(base_classes.App):
    def __init__(self, spec=None, add_book=True, xl=None, visible=None):
        # visible is only required on mac
        pythoncom.CoInitialize()
        if spec is not None:
            warn("spec is ignored on Windows.")
        if xl is None:
            # new instance
            self._xl = COMRetryObjectWrapper(DispatchEx("Excel.Application"))
            if add_book:
                self._xl.Workbooks.Add()
            self._hwnd = None
        elif isinstance(xl, int):
            self._xl = None
            self._hwnd = xl
        else:
            self._xl = xl
            self._hwnd = None
        self._pid = self.pid

    @property
    def xl(self):
        if self._xl is None:
            self._xl = get_xl_app_from_hwnd(self._hwnd)
        return self._xl

    @xl.setter
    def xl(self, value):
        self._xl = value

    api = xl

    @property
    def engine(self):
        return engine

    @property
    def selection(self):
        try:
            _ = (
                self.xl.Selection.Address
            )  # Force exception outside of the retry wrapper e.g., if chart is selected
            return Range(xl=self.xl.Selection)
        except pywintypes.com_error:
            return None

    def activate(self, steal_focus=False):
        # makes the Excel instance the foreground Excel instance,
        # but not the foreground desktop app if the current foreground
        # app isn't already an Excel instance
        hwnd = windll.user32.GetForegroundWindow()
        if steal_focus or is_hwnd_xl_app(hwnd):
            windll.user32.SetForegroundWindow(self.xl.Hwnd)
        else:
            windll.user32.SetWindowPos(self.xl.Hwnd, hwnd, 0, 0, 0, 0, 0x1 | 0x2 | 0x10)

    @property
    def visible(self):
        return self.xl.Visible

    @visible.setter
    def visible(self, visible):
        self.xl.Visible = visible

    def quit(self):
        self.xl.DisplayAlerts = False
        self.xl.Quit()
        self.xl = None
        try:
            Apps.cleanup()
        except:  # noqa: E722
            pass

    def kill(self):
        PROCESS_TERMINATE = 1
        handle = win32api.OpenProcess(PROCESS_TERMINATE, False, self._pid)
        win32api.TerminateProcess(handle, -1)
        win32api.CloseHandle(handle)
        try:
            Apps.cleanup()
        except:  # noqa: E722
            pass

    @property
    def screen_updating(self):
        return self.xl.ScreenUpdating

    @screen_updating.setter
    def screen_updating(self, value):
        self.xl.ScreenUpdating = value

    @property
    def display_alerts(self):
        return self.xl.DisplayAlerts

    @display_alerts.setter
    def display_alerts(self, value):
        self.xl.DisplayAlerts = value

    @property
    def enable_events(self):
        return self.xl.EnableEvents

    @enable_events.setter
    def enable_events(self, value):
        self.xl.EnableEvents = value

    @property
    def interactive(self):
        return self.xl.Interactive

    @interactive.setter
    def interactive(self, value):
        self.xl.Interactive = value

    @property
    def startup_path(self):
        return self.xl.StartupPath

    @property
    def calculation(self):
        return calculation_i2s[self.xl.Calculation]

    @calculation.setter
    def calculation(self, value):
        self.xl.Calculation = calculation_s2i[value]

    def calculate(self):
        self.xl.Calculate()

    @property
    def version(self):
        return self.xl.Version

    @property
    def books(self):
        return Books(xl=self.xl.Workbooks, app=self)

    @property
    def hwnd(self):
        if self._hwnd is None:
            self._hwnd = self._xl.Hwnd
        return self._hwnd

    @property
    def path(self):
        return self.xl.Path

    @property
    def pid(self):
        return win32process.GetWindowThreadProcessId(self.hwnd)[1]

    def run(self, macro, args):
        return self.xl.Run(macro, *args)

    @property
    def status_bar(self):
        return self.xl.StatusBar

    @status_bar.setter
    def status_bar(self, value):
        self.xl.StatusBar = value

    @property
    def cut_copy_mode(self):
        modes = {2: "cut", 1: "copy"}
        return modes.get(self.xl.CutCopyMode)

    @cut_copy_mode.setter
    def cut_copy_mode(self, value):
        self.xl.CutCopyMode = value

    def alert(self, prompt, title, buttons, mode, callback):
        buttons_dict = {
            None: win32con.MB_OK,
            "ok": win32con.MB_OK,
            "ok_cancel": win32con.MB_OKCANCEL,
            "yes_no": win32con.MB_YESNO,
            "yes_no_cancel": win32con.MB_YESNOCANCEL,
        }
        modes = {
            "info": win32con.MB_ICONINFORMATION,
            "critical": win32con.MB_ICONWARNING,
        }
        style = buttons_dict[buttons]
        if mode:
            style += modes[mode]
        rv = win32api.MessageBox(
            self.hwnd,
            "" if prompt is None else prompt,
            "" if title is None else title,
            style,
        )
        return_values = {1: "ok", 2: "cancel", 6: "yes", 7: "no"}
        return return_values[rv]


class Books(base_classes.Books):
    def __init__(self, xl, app):
        self.xl = xl
        self.app = app

    @property
    def api(self):
        return self.xl

    @property
    def active(self):
        return Book(self.xl.Application.ActiveWorkbook)

    def __call__(self, name_or_index):
        try:
            return Book(xl=self.xl(name_or_index))
        except pywintypes.com_error:
            raise KeyError(name_or_index)

    def __len__(self):
        return self.xl.Count

    def add(self):
        return Book(xl=self.xl.Add())

    def open(
        self,
        fullname,
        update_links=None,
        read_only=None,
        format=None,
        password=None,
        write_res_password=None,
        ignore_read_only_recommended=None,
        origin=None,
        delimiter=None,
        editable=None,
        notify=None,
        converter=None,
        add_to_mru=None,
        local=None,
        corrupt_load=None,
    ):
        # update_links: According to VBA docs, only constants 0 and 3 are supported
        if update_links:
            update_links = UpdateLinks.xlUpdateLinksAlways
        # Workbooks.Open params are position only on pywin32
        return Book(
            xl=self.xl.Open(
                fullname,
                update_links,
                read_only,
                format,
                password,
                write_res_password,
                ignore_read_only_recommended,
                origin,
                delimiter,
                editable,
                notify,
                converter,
                add_to_mru,
                local,
                corrupt_load,
            )
        )

    def __iter__(self):
        for xl in self.xl:
            yield Book(xl=xl)


class Book(base_classes.Book):
    def __init__(self, xl):
        self.xl = xl

    @property
    def api(self):
        return self.xl

    def json(self):
        raise NotImplementedError()

    @property
    def name(self):
        return self.xl.Name

    @property
    def sheets(self):
        return Sheets(xl=self.xl.Worksheets)

    @property
    def app(self):
        return App(xl=self.xl.Application)

    def close(self):
        self.xl.Close(SaveChanges=False)

    def save(self, path=None, password=None):
        saved_path = self.xl.Path
        source_ext = os.path.splitext(self.name)[1] if saved_path else None
        target_ext = os.path.splitext(path)[1] if path else ".xlsx"
        if saved_path and source_ext == target_ext:
            file_format = self.xl.FileFormat
        else:
            ext_to_file_format = {
                ".xlsx": FileFormat.xlOpenXMLWorkbook,
                ".xlsm": FileFormat.xlOpenXMLWorkbookMacroEnabled,
                ".xlsb": FileFormat.xlExcel12,
                ".xltm": FileFormat.xlOpenXMLTemplateMacroEnabled,
                ".xltx": FileFormat.xlOpenXMLTemplateMacroEnabled,
                ".xlam": FileFormat.xlOpenXMLAddIn,
                ".xls": FileFormat.xlWorkbookNormal,
                ".xlt": FileFormat.xlTemplate,
                ".xla": FileFormat.xlAddIn,
                ".html": FileFormat.xlHtml,
            }
            file_format = ext_to_file_format[target_ext]
        if (saved_path != "") and (path is None):
            # Previously saved: Save under existing name
            self.xl.Save()
        elif (
            (saved_path != "") and (path is not None) and (os.path.split(path)[0] == "")
        ):
            # Save existing book under new name in cwd if no path has been provided
            path = os.path.join(os.getcwd(), path)
            self.xl.SaveAs(
                os.path.realpath(path), FileFormat=file_format, Password=password
            )
        elif (saved_path == "") and (path is None):
            # Previously unsaved: Save under current name in current working directory
            path = os.path.join(os.getcwd(), self.xl.Name + ".xlsx")
            alerts_state = self.xl.Application.DisplayAlerts
            self.xl.Application.DisplayAlerts = False
            self.xl.SaveAs(
                os.path.realpath(path), FileFormat=file_format, Password=password
            )
            self.xl.Application.DisplayAlerts = alerts_state
        elif path:
            # Save under new name/location
            alerts_state = self.xl.Application.DisplayAlerts
            self.xl.Application.DisplayAlerts = False
            self.xl.SaveAs(
                os.path.realpath(path), FileFormat=file_format, Password=password
            )
            self.xl.Application.DisplayAlerts = alerts_state

    @property
    def fullname(self):
        if "://" in self.xl.FullName:
            config = read_config_sheet(xlwings.Book(impl=self))
            return fullname_url_to_local_path(
                url=self.xl.FullName,
                sheet_onedrive_consumer_config=config.get("ONEDRIVE_CONSUMER_WIN"),
                sheet_onedrive_commercial_config=config.get("ONEDRIVE_COMMERCIAL_WIN"),
                sheet_sharepoint_config=config.get("SHAREPOINT_WIN"),
            )
        else:
            return self.xl.FullName

    @property
    def names(self):
        return Names(xl=self.xl.Names)

    def activate(self):
        self.xl.Activate()

    def to_pdf(self, path, quality):
        self.xl.ExportAsFixedFormat(
            Type=FixedFormatType.xlTypePDF,
            Filename=path,
            Quality=quality_types[quality],
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False,
        )


class Sheets(base_classes.Sheets):
    def __init__(self, xl):
        self.xl = xl

    @property
    def api(self):
        return self.xl

    @property
    def active(self):
        return Sheet(self.xl.Parent.ActiveSheet)

    def __call__(self, name_or_index):
        return Sheet(xl=self.xl(name_or_index))

    def __len__(self):
        return self.xl.Count

    def __iter__(self):
        for xl in self.xl:
            yield Sheet(xl=xl)

    def add(self, before=None, after=None, name=None):
        if before:
            sheet = Sheet(xl=self.xl.Add(Before=before.xl))
            if name is not None:
                sheet.name = name
            return sheet
        elif after:
            # Hack, since "After" is broken in certain environments
            # see: http://code.activestate.com/lists/python-win32/11554/
            count = self.xl.Count
            new_sheet_index = after.xl.Index + 1
            if new_sheet_index > count:
                xl_sheet = self.xl.Add(Before=after.xl)
                self.xl(self.xl.Count).Move(Before=self.xl(self.xl.Count - 1))
                self.xl(self.xl.Count).Activate()
            else:
                xl_sheet = self.xl.Add(Before=self.xl(after.xl.Index + 1))
            sheet = Sheet(xl=xl_sheet)
            if name is not None:
                sheet.name = name
            return sheet
        else:
            sheet = Sheet(xl=self.xl.Add())
            if name is not None:
                sheet.name = name
            return sheet


class Sheet(base_classes.Sheet):
    def __init__(self, xl):
        self.xl = xl

    @property
    def api(self):
        return self.xl

    @property
    def name(self):
        return self.xl.Name

    @name.setter
    def name(self, value):
        self.xl.Name = value

    @property
    def names(self):
        return Names(xl=self.xl.Names)

    @property
    def book(self):
        return Book(xl=self.xl.Parent)

    @property
    def notes(self):
        return [Note(xl=comment) for comment in self.xl.Comments]

    @property
    def index(self):
        return self.xl.Index

    def range(self, arg1, arg2=None):
        if isinstance(arg1, Range):
            xl1 = arg1.xl
        elif isinstance(arg1, tuple):
            if len(arg1) == 4:
                row, col, nrows, ncols = arg1
                return Range(xl=(self.xl, row, col, nrows, ncols))
            if 0 in arg1:
                raise IndexError(
                    "Attempted to access 0-based Range. "
                    "xlwings/Excel Ranges are 1-based."
                )
            xl1 = self.xl.Cells(arg1[0], arg1[1])
        elif isinstance(arg1, numbers.Number) and isinstance(arg2, numbers.Number):
            xl1 = self.xl.Cells(arg1, arg2)
            arg2 = None
        else:
            xl1 = self.xl.Range(arg1)

        if arg2 is None:
            return Range(xl=xl1)

        if isinstance(arg2, Range):
            xl2 = arg2.xl
        elif isinstance(arg2, tuple):
            if 0 in arg2:
                raise IndexError(
                    "Attempted to access 0-based Range. "
                    "xlwings/Excel Ranges are 1-based."
                )
            xl2 = self.xl.Cells(arg2[0], arg2[1])
        else:
            xl2 = self.xl.Range(arg2)

        return Range(xl=self.xl.Range(xl1, xl2))

    @property
    def cells(self):
        return Range(xl=self.xl.Cells)

    def activate(self):
        return self.xl.Activate()

    def select(self):
        return self.xl.Select()

    def clear_contents(self):
        self.xl.Cells.ClearContents()

    def clear_formats(self):
        self.xl.Cells.ClearFormats()

    def clear(self):
        self.xl.Cells.Clear()

    def autofit(self, axis=None):
        if axis == "rows" or axis == "r":
            self.xl.Rows.AutoFit()
        elif axis == "columns" or axis == "c":
            self.xl.Columns.AutoFit()
        elif axis is None:
            self.xl.Rows.AutoFit()
            self.xl.Columns.AutoFit()

    def delete(self):
        app = self.xl.Parent.Application
        alerts_state = app.DisplayAlerts
        app.DisplayAlerts = False
        self.xl.Delete()
        app.DisplayAlerts = alerts_state

    def copy(self, before, after):
        if before:
            before = before.xl
        if after:
            after = after.xl
        self.xl.Copy(Before=before, After=after)

    def move(self, before, after):
        if before:
            before = before.xl
        if after:
            after = after.xl
        self.xl.Move(Before=before, After=after)

    @property
    def charts(self):
        return Charts(xl=self.xl.ChartObjects())

    @property
    def shapes(self):
        return Shapes(xl=self.xl.Shapes)

    @property
    def tables(self):
        return Tables(xl=self.xl.ListObjects)

    @property
    def pivot_tables(self):
        return PivotTables(xl=self.xl.PivotTables())

    @property
    def pictures(self):
        return Pictures(xl=self.xl.Pictures())

    @property
    def used_range(self):
        return Range(xl=self.xl.UsedRange)

    @property
    def visible(self):
        return self.xl.Visible

    @visible.setter
    def visible(self, value):
        self.xl.Visible = value

    def _window_property(self, name, *value):
        """Get (no `value`) or set (one `value`) a property of the book's window.

        Gridlines and the like are window properties in COM, applying to the
        sheet that the window currently shows. A sheet that isn't active is
        therefore activated for the duration of the call and the previously
        active sheet restored afterwards, with screen updating off.
        """
        book = self.xl.Parent
        window = book.Windows(1)
        previous_book_sheet = book.ActiveSheet
        if previous_book_sheet.Name != self.xl.Name:
            if self.xl.Visible != constants.SheetVisibility.xlSheetVisible:
                raise ValueError(
                    f"Sheet.{name}: hidden sheets can't be activated. Set "
                    "sheet.visible = True first."
                )
            app = self.xl.Application
            previous_sheet = app.ActiveSheet
            previous_screen_updating = app.ScreenUpdating
            app.ScreenUpdating = False
            try:
                self.xl.Activate()
                if value:
                    setattr(window, name, value[0])
                    return None
                return getattr(window, name)
            finally:
                try:
                    previous_book_sheet.Activate()
                finally:
                    try:
                        previous_sheet.Activate()
                    finally:
                        app.ScreenUpdating = previous_screen_updating
        if value:
            setattr(window, name, value[0])
            return None
        return getattr(window, name)

    @property
    def show_gridlines(self):
        return bool(self._window_property("DisplayGridlines"))

    @show_gridlines.setter
    def show_gridlines(self, value):
        self._window_property("DisplayGridlines", bool(value))

    @property
    def page_setup(self):
        return PageSetup(self.xl.PageSetup)

    def to_html(self, path):
        if not Path(path).is_absolute():
            path = Path(".").resolve() / path
        source_cell2 = self.used_range.address.split(":")
        if len(source_cell2) == 2:
            source = f"A1:{source_cell2[1]}"
        else:
            source = f"A1:{source_cell2[0]}"
        self.book.xl.PublishObjects.Add(
            SourceType=SourceType.xlSourceRange,
            Filename=path,
            Sheet=self.name,
            Source=source,
            HtmlType=HtmlType.xlHtmlStatic,
        ).Publish(True)
        html_file = Path(path)
        content = html_file.read_text()
        html_file.write_text(
            content.replace(
                "align=center x:publishsource=", "align=left x:publishsource="
            )
        )


_CONDITIONAL_FORMAT_TYPE_FROM_XL = {
    constants.FormatConditionType.xlCellValue: "cell_value",
    constants.FormatConditionType.xlExpression: "custom",
    constants.FormatConditionType.xlColorScale: "color_scale",
    constants.FormatConditionType.xlDatabar: "data_bar",
    constants.FormatConditionType.xlIconSets: "icon_set",
}
_CONDITIONAL_FORMAT_OPERATOR_TO_XL = {
    "between": constants.FormatConditionOperator.xlBetween,
    "not_between": constants.FormatConditionOperator.xlNotBetween,
    "equal_to": constants.FormatConditionOperator.xlEqual,
    "not_equal_to": constants.FormatConditionOperator.xlNotEqual,
    "greater_than": constants.FormatConditionOperator.xlGreater,
    "less_than": constants.FormatConditionOperator.xlLess,
    "greater_than_or_equal": constants.FormatConditionOperator.xlGreaterEqual,
    "less_than_or_equal": constants.FormatConditionOperator.xlLessEqual,
}
_CONDITIONAL_FORMAT_OPERATOR_FROM_XL = {
    value: key for key, value in _CONDITIONAL_FORMAT_OPERATOR_TO_XL.items()
}
_CONDITIONAL_FORMAT_THRESHOLD_TO_XL = {
    "lowest_value": constants.ConditionValueTypes.xlConditionValueLowestValue,
    "highest_value": constants.ConditionValueTypes.xlConditionValueHighestValue,
    "number": constants.ConditionValueTypes.xlConditionValueNumber,
    "percent": constants.ConditionValueTypes.xlConditionValuePercent,
    "percentile": constants.ConditionValueTypes.xlConditionValuePercentile,
    "formula": constants.ConditionValueTypes.xlConditionValueFormula,
}
_CONDITIONAL_FORMAT_THRESHOLD_FROM_XL = {
    **{value: key for key, value in _CONDITIONAL_FORMAT_THRESHOLD_TO_XL.items()},
    constants.ConditionValueTypes.xlConditionValueAutomaticMin: "automatic",
    constants.ConditionValueTypes.xlConditionValueAutomaticMax: "automatic",
}
_CONDITIONAL_FORMAT_ICON_SET_TO_XL = {
    "3_arrows": constants.IconSet.xl3Arrows,
    "3_arrows_gray": constants.IconSet.xl3ArrowsGray,
    "3_flags": constants.IconSet.xl3Flags,
    "3_traffic_lights_1": constants.IconSet.xl3TrafficLights1,
    "3_traffic_lights_2": constants.IconSet.xl3TrafficLights2,
    "3_signs": constants.IconSet.xl3Signs,
    "3_symbols": constants.IconSet.xl3Symbols,
    "3_symbols_2": constants.IconSet.xl3Symbols2,
    "4_arrows": constants.IconSet.xl4Arrows,
    "4_arrows_gray": constants.IconSet.xl4ArrowsGray,
    "4_red_to_black": constants.IconSet.xl4RedToBlack,
    "4_rating": constants.IconSet.xl4CRV,
    "4_traffic_lights": constants.IconSet.xl4TrafficLights,
    "5_arrows": constants.IconSet.xl5Arrows,
    "5_arrows_gray": constants.IconSet.xl5ArrowsGray,
    "5_rating": constants.IconSet.xl5CRV,
    "5_quarters": constants.IconSet.xl5Quarters,
    "3_stars": constants.IconSet.xl3Stars,
    "3_triangles": constants.IconSet.xl3Triangles,
    "5_boxes": constants.IconSet.xl5Boxes,
}
_CONDITIONAL_FORMAT_ICON_SET_FROM_XL = {
    value: key for key, value in _CONDITIONAL_FORMAT_ICON_SET_TO_XL.items()
}


class Range(base_classes.Range):
    def __init__(self, xl):
        if isinstance(xl, tuple):
            self._coords = xl
            self._xl = missing
        else:
            self._coords = missing
            self._xl = xl

    @property
    def xl(self):
        if self._xl is missing:
            xl_sheet, row, col, nrows, ncols = self._coords
            if nrows and ncols:
                self._xl = xl_sheet.Range(
                    xl_sheet.Cells(row, col),
                    xl_sheet.Cells(row + nrows - 1, col + ncols - 1),
                )
            else:
                self._xl = None
        return self._xl

    @property
    def coords(self):
        if self._coords is missing:
            self._coords = (
                self.xl.Worksheet,
                self.xl.Row,
                self.xl.Column,
                self.xl.Rows.Count,
                self.xl.Columns.Count,
            )
        return self._coords

    @property
    def api(self):
        return self.xl

    @property
    def autofilter(self):
        return AutoFilter(self)

    @property
    def sheet(self):
        return Sheet(xl=self.coords[0])

    def __len__(self):
        return (self.xl and self.xl.Count) or 0

    @property
    def row(self):
        return self.coords[1]

    @property
    def column(self):
        return self.coords[2]

    @property
    def shape(self):
        return self.coords[3], self.coords[4]

    @property
    def raw_value(self):
        if self.xl is not None:
            return self.xl.Value
        else:
            return None

    @raw_value.setter
    def raw_value(self, data):
        if self.xl is not None:
            self.xl.Value = data

    def clear_contents(self):
        if self.xl is not None:
            self.xl.ClearContents()

    def clear_formats(self):
        self.xl.ClearFormats()

    def clear(self):
        if self.xl is not None:
            self.xl.Clear()

    @property
    def formula(self):
        if self.xl is not None:
            return self.xl.Formula
        else:
            return None

    @formula.setter
    def formula(self, value):
        if self.xl is not None:
            self.xl.Formula = value

    @property
    def formula2(self):
        if self.xl is not None:
            return self.xl.Formula2
        else:
            return None

    @formula2.setter
    def formula2(self, value):
        if self.xl is not None:
            self.xl.Formula2 = value

    def end(self, direction):
        direction = directions_s2i.get(direction, direction)
        return Range(xl=self.xl.End(direction))

    @property
    def formula_array(self):
        if self.xl is not None:
            return self.xl.FormulaArray
        else:
            return None

    @formula_array.setter
    def formula_array(self, value):
        if self.xl is not None:
            self.xl.FormulaArray = value

    @property
    def font(self):
        return Font(self, self.xl.Font)

    @property
    def borders(self):
        return Borders(self, self.xl)

    @property
    def data_validation(self):
        return DataValidation(self)

    @property
    def column_width(self):
        if self.xl is not None:
            return self.xl.ColumnWidth
        else:
            return 0

    @column_width.setter
    def column_width(self, value):
        if self.xl is not None:
            self.xl.ColumnWidth = value

    @property
    def row_height(self):
        if self.xl is not None:
            return self.xl.RowHeight
        else:
            return 0

    @row_height.setter
    def row_height(self, value):
        if self.xl is not None:
            self.xl.RowHeight = value

    @property
    def width(self):
        if self.xl is not None:
            return self.xl.Width
        else:
            return 0

    @property
    def height(self):
        if self.xl is not None:
            return self.xl.Height
        else:
            return 0

    @property
    def left(self):
        if self.xl is not None:
            return self.xl.Left
        else:
            return 0

    @property
    def top(self):
        if self.xl is not None:
            return self.xl.Top
        else:
            return 0

    @property
    def number_format(self):
        if self.xl is not None:
            return self.xl.NumberFormat
        else:
            return ""

    @number_format.setter
    def number_format(self, value):
        if self.xl is not None:
            self.xl.NumberFormat = value

    def get_address(self, row_absolute, col_absolute, external):
        if self.xl is not None:
            return self.xl.GetAddress(row_absolute, col_absolute, 1, external)
        else:
            raise NotImplementedError()

    @property
    def address(self):
        if self.xl is not None:
            return self.xl.Address
        else:
            _, row, col, nrows, ncols = self.coords
            return "$%s$%s{%sx%s}" % (col_name(col), str(row), nrows, ncols)

    @property
    def current_region(self):
        if self.xl is not None:
            return Range(xl=self.xl.CurrentRegion)
        else:
            return self

    def autofit(self, axis=None):
        if self.xl is not None:
            if axis == "rows" or axis == "r":
                self.xl.Rows.AutoFit()
            elif axis == "columns" or axis == "c":
                self.xl.Columns.AutoFit()
            elif axis is None:
                self.xl.Columns.AutoFit()
                self.xl.Rows.AutoFit()

    def insert(self, shift=None, copy_origin=None):
        shifts = {
            "down": InsertShiftDirection.xlShiftDown,
            "right": InsertShiftDirection.xlShiftToRight,
            None: None,
        }
        copy_origins = {
            "format_from_left_or_above": InsertFormatOrigin.xlFormatFromLeftOrAbove,
            "format_from_right_or_below": InsertFormatOrigin.xlFormatFromRightOrBelow,
        }
        self.xl.Insert(Shift=shifts[shift], CopyOrigin=copy_origins[copy_origin])

    def delete(self, shift=None):
        shifts = {
            "up": DeleteShiftDirection.xlShiftUp,
            "left": DeleteShiftDirection.xlShiftToLeft,
            None: None,
        }
        self.xl.Delete(Shift=shifts[shift])

    def copy(self, destination=None):
        self.xl.Copy(Destination=destination.api if destination else None)

    def paste(self, paste=None, operation=None, skip_blanks=False, transpose=False):
        pastes = {
            "all": -4104,
            None: -4104,
            "all_except_borders": 7,
            "all_merging_conditional_formats": 14,
            "all_using_source_theme": 13,
            "column_widths": 8,
            "comments": -4144,
            "formats": -4122,
            "formulas": -4123,
            "formulas_and_number_formats": 11,
            "validation": 6,
            "values": -4163,
            "values_and_number_formats": 12,
        }

        operations = {
            "add": 2,
            "divide": 5,
            "multiply": 4,
            None: -4142,
            "subtract": 3,
        }

        self.xl.PasteSpecial(
            Paste=pastes[paste],
            Operation=operations[operation],
            SkipBlanks=skip_blanks,
            Transpose=transpose,
        )

    @property
    def hyperlink(self):
        if self.xl is not None:
            try:
                return self.xl.Hyperlinks(1).Address
            except pywintypes.com_error:
                raise Exception("The cell doesn't seem to contain a hyperlink!")
        else:
            return ""

    def add_hyperlink(self, address, text_to_display, screen_tip):
        if self.xl is not None:
            # Another one of these pywin32 bugs that only materialize under certain
            # circumstances: https://stackoverflow.com/questions/
            #  6284227/hyperlink-will-not-show-display-proper-text
            link = self.xl.Hyperlinks.Add(Anchor=self.xl, Address=address)
            link.TextToDisplay = text_to_display
            link.ScreenTip = screen_tip

    @property
    def color(self):
        if self.xl is not None:
            if self.xl.Interior.ColorIndex == ColorIndex.xlColorIndexNone:
                return None
            else:
                return int_to_rgb(self.xl.Interior.Color)
        else:
            return None

    @color.setter
    def color(self, color_or_rgb):
        if isinstance(color_or_rgb, str):
            color_or_rgb = hex_to_rgb(color_or_rgb)
        if self.xl is not None:
            if color_or_rgb is None:
                self.xl.Interior.ColorIndex = ColorIndex.xlColorIndexNone
            elif isinstance(color_or_rgb, int):
                self.xl.Interior.Color = color_or_rgb
            else:
                self.xl.Interior.Color = rgb_to_int(color_or_rgb)

    def set_colors(self, colors):
        for row_index, row in enumerate(colors):
            for column_index, color in enumerate(row):
                if color is not ...:
                    self(row_index + 1, column_index + 1).color = color

    @property
    def name(self):
        if self.xl is not None:
            try:
                name = Name(xl=self.xl.Name)
            except pywintypes.com_error:
                name = None
            return name
        else:
            return None

    @property
    def has_array(self):
        if self.xl is not None:
            try:
                return self.xl.HasArray
            except pywintypes.com_error:
                return False
        else:
            return False

    @name.setter
    def name(self, value):
        if self.xl is not None:
            self.xl.Name = value

    def __call__(self, *args):
        if self.xl is not None:
            if len(args) == 0:
                raise ValueError("Invalid arguments")
            return Range(xl=self.xl(*args))
        else:
            raise NotImplementedError()

    @property
    def rows(self):
        return Range(xl=self.xl.Rows)

    @property
    def columns(self):
        return Range(xl=self.xl.Columns)

    def select(self):
        return self.xl.Select()

    @property
    def merge_area(self):
        return Range(xl=self.xl.MergeArea)

    @property
    def merge_cells(self):
        return self.xl.MergeCells

    def merge(self, across):
        self.xl.Merge(across)

    def unmerge(self):
        self.xl.UnMerge()

    @property
    def table(self):
        if self.xl.ListObject:
            return Table(self.xl.ListObject)

    @property
    def characters(self):
        return Characters(parent=self, xl=self.xl.GetCharacters)

    @property
    def wrap_text(self):
        return self.xl.WrapText

    @wrap_text.setter
    def wrap_text(self, value):
        self.xl.WrapText = value

    @property
    def horizontal_alignment(self):
        # COM returns None for a range whose cells disagree.
        return horizontal_alignments_i2s.get(self.xl.HorizontalAlignment)

    @horizontal_alignment.setter
    def horizontal_alignment(self, value):
        self.xl.HorizontalAlignment = horizontal_alignments_s2i[value]

    @property
    def vertical_alignment(self):
        return vertical_alignments_i2s.get(self.xl.VerticalAlignment)

    @vertical_alignment.setter
    def vertical_alignment(self, value):
        self.xl.VerticalAlignment = vertical_alignments_s2i[value]

    @property
    def note(self):
        return Note(xl=self.xl.Comment) if self.xl.Comment else None

    def add_note(self, text):
        return Note(xl=self.xl.AddComment(text))

    @property
    def conditional_formats(self):
        return ConditionalFormats(xl=self.xl.FormatConditions)

    def copy_picture(self, appearance, format):
        _appearance = {"screen": 1, "printer": 2}
        _format = {"picture": -4147, "bitmap": 2}
        self.xl.CopyPicture(Appearance=_appearance[appearance], Format=_format[format])

    def to_png(self, path):
        max_retries = 10
        for retry in range(max_retries):
            # https://stackoverflow.com/questions/
            #  24740062/copypicture-method-of-range-class-failed-sometimes
            try:
                # appearance="printer" fails here, not sure why
                self.copy_picture(appearance="screen", format="bitmap")
                im = ImageGrab.grabclipboard()
                im.save(path)
                break
            except (pywintypes.com_error, AttributeError):
                if retry == max_retries - 1:
                    raise

    def to_pdf(self, path, quality):
        self.xl.ExportAsFixedFormat(
            Type=FixedFormatType.xlTypePDF,
            Filename=path,
            Quality=quality_types[quality],
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False,
        )

    def autofill(self, destination, type_):
        types = {
            "fill_copy": constants.AutoFillType.xlFillCopy,
            "fill_days": constants.AutoFillType.xlFillDays,
            "fill_default": constants.AutoFillType.xlFillDefault,
            "fill_formats": constants.AutoFillType.xlFillFormats,
            "fill_months": constants.AutoFillType.xlFillMonths,
            "fill_series": constants.AutoFillType.xlFillSeries,
            "fill_values": constants.AutoFillType.xlFillValues,
            "fill_weekdays": constants.AutoFillType.xlFillWeekdays,
            "fill_years": constants.AutoFillType.xlFillYears,
            "growth_trend": constants.AutoFillType.xlGrowthTrend,
            "linear_trend": constants.AutoFillType.xlLinearTrend,
            "flash_fill": constants.AutoFillType.xlFlashFill,
        }
        self.xl.AutoFill(Destination=destination.api, Type=types[type_])


class Shape(base_classes.Shape):
    def __init__(self, xl):
        self.xl = xl

    @property
    def api(self):
        return self.xl

    @property
    def name(self):
        return self.xl.Name

    @property
    def parent(self):
        return Sheet(xl=self.xl.Parent)

    @property
    def type(self):
        return shape_types_i2s[self.xl.Type]

    @property
    def left(self):
        return self.xl.Left

    @left.setter
    def left(self, value):
        self.xl.Left = value

    @property
    def top(self):
        return self.xl.Top

    @top.setter
    def top(self, value):
        self.xl.Top = value

    @property
    def width(self):
        return self.xl.Width

    @width.setter
    def width(self, value):
        self.xl.Width = value

    @property
    def height(self):
        return self.xl.Height

    @height.setter
    def height(self, value):
        self.xl.Height = value

    def delete(self):
        self.xl.Delete()

    @name.setter
    def name(self, value):
        self.xl.Name = value

    @property
    def index(self):
        return self.xl.Index

    def activate(self):
        self.xl.Activate()

    def scale_height(self, factor, relative_to_original_size, scale):
        self.xl.ScaleHeight(
            Scale=scaling[scale],
            RelativeToOriginalSize=relative_to_original_size,
            Factor=factor,
        )

    def scale_width(self, factor, relative_to_original_size, scale):
        self.xl.ScaleWidth(
            Scale=scaling[scale],
            RelativeToOriginalSize=relative_to_original_size,
            Factor=factor,
        )

    @property
    def text(self):
        if self.xl.TextFrame2.HasText:
            return self.xl.TextFrame2.TextRange.Text

    @text.setter
    def text(self, value):
        self.xl.TextFrame2.TextRange.Text = value

    @property
    def font(self):
        return Font(self, self.xl.TextFrame2.TextRange.Font)

    @property
    def characters(self):
        return Characters(parent=self, xl=self.xl.TextFrame2.TextRange.GetCharacters)


_BORDER_SIDE_TO_XL = {
    "edge_top": constants.BordersIndex.xlEdgeTop,
    "edge_bottom": constants.BordersIndex.xlEdgeBottom,
    "edge_left": constants.BordersIndex.xlEdgeLeft,
    "edge_right": constants.BordersIndex.xlEdgeRight,
    "inside_vertical": constants.BordersIndex.xlInsideVertical,
    "inside_horizontal": constants.BordersIndex.xlInsideHorizontal,
    "diagonal_down": constants.BordersIndex.xlDiagonalDown,
    "diagonal_up": constants.BordersIndex.xlDiagonalUp,
}
# None is the normalized form of "none", see main.Borders
_BORDER_LINE_STYLE_TO_XL = {
    "continuous": constants.LineStyle.xlContinuous,
    "dash": constants.LineStyle.xlDash,
    "dash_dot": constants.LineStyle.xlDashDot,
    "dash_dot_dot": constants.LineStyle.xlDashDotDot,
    "dot": constants.LineStyle.xlDot,
    "double": constants.LineStyle.xlDouble,
    "slant_dash_dot": constants.LineStyle.xlSlantDashDot,
    None: constants.LineStyle.xlLineStyleNone,
}
_BORDER_LINE_STYLE_FROM_XL = {
    xl: name or "none" for name, xl in _BORDER_LINE_STYLE_TO_XL.items()
}
_BORDER_WEIGHT_TO_XL = {
    "hairline": constants.BorderWeight.xlHairline,
    "thin": constants.BorderWeight.xlThin,
    "medium": constants.BorderWeight.xlMedium,
    "thick": constants.BorderWeight.xlThick,
}
_BORDER_WEIGHT_FROM_XL = {xl: name for name, xl in _BORDER_WEIGHT_TO_XL.items()}


class Border(base_classes.Border):
    def __init__(self, parent, side, xl):
        self.parent = parent
        self.side = side
        self.xl = xl

    @property
    def api(self):
        return self.xl

    @property
    def line_style(self):
        """The line style Excel reports for the range as a whole.

        Excel doesn't flag a range whose cells disagree, so a mixed range
        reports one of its values rather than None. Read a single cell to
        get an unambiguous answer.
        """
        if self.xl is not None:
            value = self.xl.LineStyle
            return None if value is None else _BORDER_LINE_STYLE_FROM_XL.get(value)

    @line_style.setter
    def line_style(self, value):
        if self.xl is not None:
            self.xl.LineStyle = _BORDER_LINE_STYLE_TO_XL[value]

    @property
    def weight(self):
        # Like line_style, a mixed range reports one of its values, not None
        if self.xl is not None:
            value = self.xl.Weight
            return None if value is None else _BORDER_WEIGHT_FROM_XL.get(value)

    @weight.setter
    def weight(self, value):
        if self.xl is not None:
            self.xl.Weight = _BORDER_WEIGHT_TO_XL[value]

    @property
    def color(self):
        """The color Excel reports for the range as a whole.

        Like line_style, a mixed range isn't flagged. Diagonals are a
        further special case: on a multi-cell range Excel reports no color
        for them even right after one was set, so this returns None. Read
        the diagonal of a single cell to get its color.
        """
        if self.xl is not None:
            if self.xl.ColorIndex == ColorIndex.xlColorIndexNone:
                return None
            color = self.xl.Color
            return None if color is None else int_to_rgb(color)

    @color.setter
    def color(self, color_or_rgb):
        if isinstance(color_or_rgb, str):
            color_or_rgb = hex_to_rgb(color_or_rgb)
        if self.xl is not None:
            if isinstance(color_or_rgb, int):
                self.xl.Color = color_or_rgb
            else:
                self.xl.Color = rgb_to_int(color_or_rgb)


class DataValidation(base_classes.DataValidation):
    _TYPE_FROM_XL = {
        constants.DVType.xlValidateWholeNumber: "whole_number",
        constants.DVType.xlValidateDecimal: "decimal",
        constants.DVType.xlValidateList: "list",
        constants.DVType.xlValidateDate: "date",
        constants.DVType.xlValidateTime: "time",
        constants.DVType.xlValidateTextLength: "text_length",
        constants.DVType.xlValidateCustom: "custom",
    }
    _TYPE_TO_XL = {value: key for key, value in _TYPE_FROM_XL.items()}
    _OPERATOR_FROM_XL = {
        constants.FormatConditionOperator.xlBetween: "between",
        constants.FormatConditionOperator.xlNotBetween: "not_between",
        constants.FormatConditionOperator.xlEqual: "equal_to",
        constants.FormatConditionOperator.xlNotEqual: "not_equal_to",
        constants.FormatConditionOperator.xlGreater: "greater_than",
        constants.FormatConditionOperator.xlLess: "less_than",
        constants.FormatConditionOperator.xlGreaterEqual: "greater_than_or_equal",
        constants.FormatConditionOperator.xlLessEqual: "less_than_or_equal",
    }
    _OPERATOR_TO_XL = {value: key for key, value in _OPERATOR_FROM_XL.items()}
    _ALERT_STYLE_FROM_XL = {
        constants.DVAlertStyle.xlValidAlertStop: "stop",
        constants.DVAlertStyle.xlValidAlertWarning: "warning",
        constants.DVAlertStyle.xlValidAlertInformation: "information",
    }

    def __init__(self, parent):
        self.parent = parent

    @property
    def api(self):
        return self.parent.xl.Validation

    def _nonuniform_type(self):
        try:
            validation_cells = self.parent.xl.SpecialCells(
                constants.CellType.xlCellTypeAllValidation
            )
            intersection = self.parent.xl.Application.Intersect(
                self.parent.xl, validation_cells
            )
        except pywintypes.com_error:
            return "none"
        if intersection is None:
            return "none"
        try:
            validated_count = int(intersection.Cells.CountLarge)
            target_count = int(self.parent.xl.Cells.CountLarge)
        except (AttributeError, TypeError, ValueError, pywintypes.com_error):
            validated_count = int(intersection.Cells.Count)
            target_count = int(self.parent.xl.Cells.Count)
        return "mixed_criteria" if validated_count < target_count else "inconsistent"

    @property
    def type(self):
        try:
            native_type = self.parent.xl.Validation.Type
        except pywintypes.com_error:
            return self._nonuniform_type()
        if native_type is None:
            return self._nonuniform_type()
        return self._TYPE_FROM_XL.get(native_type, "unknown")

    def _property(self, name):
        if self.type in {"none", "mixed_criteria", "inconsistent"}:
            return None
        try:
            value = getattr(self.parent.xl.Validation, name)
        except pywintypes.com_error:
            return None
        return value

    @property
    def operator(self):
        if self.type not in {
            "whole_number",
            "decimal",
            "date",
            "time",
            "text_length",
        }:
            return None
        return self._OPERATOR_FROM_XL.get(self._property("Operator"))

    @property
    def formula1(self):
        if self.type not in {
            "whole_number",
            "decimal",
            "date",
            "time",
            "text_length",
        }:
            return None
        value = self._property("Formula1")
        return None if value is None else str(value)

    @property
    def formula2(self):
        if self.operator not in {"between", "not_between"}:
            return None
        value = self._property("Formula2")
        return None if value is None else str(value)

    @property
    def formula(self):
        if self.type != "custom":
            return None
        value = self._property("Formula1")
        return None if value is None else str(value)

    @property
    def source(self):
        if self.type != "list":
            return None
        value = self._property("Formula1")
        return None if value is None else str(value)

    @property
    def in_cell_dropdown(self):
        value = self._property("InCellDropdown") if self.type == "list" else None
        return None if value is None else bool(value)

    @property
    def ignore_blank(self):
        value = self._property("IgnoreBlank")
        return None if value is None else bool(value)

    @property
    def input_title(self):
        return self._property("InputTitle")

    @property
    def input_message(self):
        return self._property("InputMessage")

    @property
    def show_input(self):
        value = self._property("ShowInput")
        return None if value is None else bool(value)

    @property
    def error_title(self):
        return self._property("ErrorTitle")

    @property
    def error_message(self):
        return self._property("ErrorMessage")

    @property
    def show_error(self):
        value = self._property("ShowError")
        return None if value is None else bool(value)

    @property
    def alert_style(self):
        return self._ALERT_STYLE_FROM_XL.get(self._property("AlertStyle"))

    def _formula(self, source):
        if isinstance(source, base_classes.Range):
            formula = f"={source.get_address(True, True, True)}"
        elif isinstance(source, base_classes.Name):
            formula = f"={source.name}"
        else:
            separator = self.parent.xl.Application.International[
                constants.ApplicationInternational.xlListSeparator
            ]
            formula = separator.join(source)
        if len(formula) > 255:
            raise ValueError(
                "the Excel data validation source cannot exceed 255 characters"
            )
        return formula

    def _set(self, rule_type, operator=None, formula1=None, formula2=None):
        current_type = self.type
        if current_type in {"mixed_criteria", "inconsistent"}:
            raise xlwings.XlwingsError(
                "Cannot update data validation because the target cells have "
                "different validation rules."
            )
        kwargs = {"Type": self._TYPE_TO_XL[rule_type], "Formula1": formula1}
        if operator is not None:
            kwargs["Operator"] = self._OPERATOR_TO_XL[operator]
        if formula2 is not None:
            kwargs["Formula2"] = formula2
        validation = self.parent.xl.Validation
        if current_type == "none":
            validation.Add(**kwargs)
        else:
            validation.Modify(**kwargs)

    def set_list(self, source, in_cell_dropdown):
        formula = self._formula(source)
        self._set("list", formula1=formula)
        self.parent.xl.Validation.InCellDropdown = in_cell_dropdown

    def set_rule(self, rule_type, operator, formula1, formula2):
        self._set(rule_type, operator, formula1, formula2)

    def delete(self):
        self.parent.xl.Validation.Delete()


class Borders(base_classes.Borders):
    def __init__(self, parent, xl):
        # xl is the Range object: the Borders collection is looked up per side
        self.parent = parent
        self.xl = xl

    @property
    def api(self):
        if self.xl is not None:
            return self.xl.Borders

    def __getitem__(self, side):
        if self.xl is not None:
            return Border(self.parent, side, self.xl.Borders(_BORDER_SIDE_TO_XL[side]))
        return Border(self.parent, side, None)

    def _common_value(self, attribute):
        """The value the existing grid sides share, or None if they differ.

        This compares the sides against each other. It can't detect cells
        that disagree within one side, because the per-side getters don't
        report that on Windows.
        """
        if self.xl is None:
            return None
        values = {getattr(self[side], attribute) for side in self._grid_sides()}
        return values.pop() if len(values) == 1 else None

    @property
    def line_style(self):
        return self._common_value("line_style")

    @line_style.setter
    def line_style(self, value):
        self.set(base_classes.BORDER_GRID_SIDES, line_style=value)

    @property
    def weight(self):
        return self._common_value("weight")

    @weight.setter
    def weight(self, value):
        self.set(base_classes.BORDER_GRID_SIDES, weight=value)

    @property
    def color(self):
        return self._common_value("color")

    @color.setter
    def color(self, color_or_rgb):
        self.set(base_classes.BORDER_GRID_SIDES, color=color_or_rgb)

    def set(
        self,
        which,
        *,
        line_style=base_classes._UNSET,
        weight=base_classes._UNSET,
        color=base_classes._UNSET,
    ):
        # `which` arrives validated and expanded by main.Borders. The fixed
        # order color, weight, line style is documented: Excel's border
        # attributes interfere, and this makes the line style win.
        for side in which:
            border = self[side]
            if color is not base_classes._UNSET:
                border.color = color
            if weight is not base_classes._UNSET:
                border.weight = weight
            if line_style is not base_classes._UNSET:
                border.line_style = line_style

    def clear(self, which):
        self.set(which, line_style=None)


class Font(base_classes.Font):
    def __init__(self, parent, xl):
        self.parent = parent
        self.xl = xl

    @property
    def api(self):
        return self.xl

    @property
    def bold(self):
        if isinstance(self.parent, Range):
            return self.xl.Bold
        elif isinstance(self.parent, Shape):
            return True if self.xl.Bold == -1 else False
        elif isinstance(self.parent.parent, Range):
            return self.xl.Bold
        elif isinstance(self.parent.parent, Shape):
            return True if self.xl.Bold == -1 else False
        elif isinstance(self.parent.parent.parent, Range):
            return self.xl.Bold
        elif isinstance(self.parent.parent.parent, Shape):
            return True if self.xl.Bold == -1 else False

    @bold.setter
    def bold(self, value):
        self.xl.Bold = value

    @property
    def italic(self):
        if isinstance(self.parent, Range):
            return self.xl.Italic
        elif isinstance(self.parent, Shape):
            return True if self.xl.Italic == -1 else False
        elif isinstance(self.parent.parent, Range):
            return self.xl.Italic
        elif isinstance(self.parent.parent, Shape):
            return True if self.xl.Italic == -1 else False
        elif isinstance(self.parent.parent.parent, Range):
            return self.xl.Italic
        elif isinstance(self.parent.parent.parent, Shape):
            return True if self.xl.Italic == -1 else False

    @italic.setter
    def italic(self, value):
        self.xl.Italic = value

    @property
    def size(self):
        return self.xl.Size

    @size.setter
    def size(self, value):
        self.xl.Size = value

    @property
    def color(self):
        # self.parent is used for direct access, self.parent.parent via characters
        if isinstance(self.parent, Shape):
            return int_to_rgb(self.xl.Fill.ForeColor.RGB)
        elif isinstance(self.parent, Range):
            return int_to_rgb(self.xl.Color)
        elif isinstance(self.parent.parent, Shape):
            return int_to_rgb(self.xl.Fill.ForeColor.RGB)
        elif isinstance(self.parent.parent, Range):
            return int_to_rgb(self.xl.Color)
        elif isinstance(self.parent.parent.parent, Shape):
            return int_to_rgb(self.xl.Fill.ForeColor.RGB)
        elif isinstance(self.parent.parent.parent, Range):
            return int_to_rgb(self.xl.Color)

    @color.setter
    def color(self, color_or_rgb):
        if isinstance(color_or_rgb, str):
            color_or_rgb = utils.hex_to_rgb(color_or_rgb)
        # TODO: refactor
        if self.xl is not None:
            if isinstance(self.parent, Shape):
                if isinstance(color_or_rgb, int):
                    self.xl.Fill.ForeColor.RGB = color_or_rgb
                else:
                    self.xl.Fill.ForeColor.RGB = rgb_to_int(color_or_rgb)
            elif isinstance(self.parent, Range):
                if isinstance(color_or_rgb, int):
                    self.xl.Color = color_or_rgb
                else:
                    self.xl.Color = rgb_to_int(color_or_rgb)

            elif isinstance(self.parent.parent, Shape):
                if isinstance(color_or_rgb, int):
                    self.xl.Fill.ForeColor.RGB = color_or_rgb
                else:
                    self.xl.Fill.ForeColor.RGB = rgb_to_int(color_or_rgb)
            elif isinstance(self.parent.parent, Range):
                if isinstance(color_or_rgb, int):
                    self.xl.Color = color_or_rgb
                else:
                    self.xl.Color = rgb_to_int(color_or_rgb)

            elif isinstance(self.parent.parent.parent, Shape):
                if isinstance(color_or_rgb, int):
                    self.xl.Fill.ForeColor.RGB = color_or_rgb
                else:
                    self.xl.Fill.ForeColor.RGB = rgb_to_int(color_or_rgb)
            elif isinstance(self.parent.parent.parent, Range):
                if isinstance(color_or_rgb, int):
                    self.xl.Color = color_or_rgb
                else:
                    self.xl.Color = rgb_to_int(color_or_rgb)

    @property
    def name(self):
        return self.xl.Name

    @name.setter
    def name(self, value):
        self.xl.Name = value


class Characters(base_classes.Characters):
    def __init__(self, parent, xl, start=None, length=None):
        self.parent = parent
        self.xl = xl
        self.start = start if start else 1
        self.length = length if length else xl().Count

    @property
    def api(self):
        return self.xl(self.start, self.length)

    @property
    def text(self):
        return self.xl(self.start, self.length).Text

    @property
    def font(self):
        return Font(self, self.xl(self.start, self.length).Font)

    def __getitem__(self, item):
        if isinstance(item, slice):
            if (item.start and item.start < 0) or (item.stop and item.stop < 0):
                raise ValueError(
                    self.__class__.__name__
                    + " object does not support slicing with negative indexes"
                )
            start = item.start + 1 if item.start else 1
            length = item.stop + 1 - start if item.stop else self.length + 1 - start
            return Characters(parent=self, xl=self.xl, start=start, length=length)
        else:
            if item >= 0:
                return Characters(parent=self, xl=self.xl, start=item + 1, length=1)
            else:
                return Characters(
                    parent=self, xl=self.xl, start=len(self.text) + 1 + item, length=1
                )


class Collection(base_classes.Collection):
    def __init__(self, xl):
        self.xl = xl

    @property
    def api(self):
        return self.xl

    def __call__(self, key):
        try:
            return self._wrap(xl=self.xl.Item(key))
        except pywintypes.com_error:
            raise KeyError(key)

    def __len__(self):
        return self.xl.Count

    def __iter__(self):
        for xl in self.xl:
            yield self._wrap(xl=xl)

    def __contains__(self, key):
        try:
            self.xl.Item(key)
            return True
        except pywintypes.com_error:
            return False


class PageSetup(base_classes.PageSetup):
    def __init__(self, xl):
        self.xl = xl

    @property
    def api(self):
        return self.xl

    @property
    def print_area(self):
        value = self.xl.PrintArea
        return None if value == "" else value

    @print_area.setter
    def print_area(self, value):
        self.xl.PrintArea = value


class Note(base_classes.Note):
    def __init__(self, xl):
        self.xl = xl

    @property
    def api(self):
        return self.xl

    @property
    def text(self):
        return self.xl.Text()

    @property
    def author(self):
        return self.xl.Author

    @property
    def location(self):
        return Range(xl=self.xl.Parent)

    @text.setter
    def text(self, value):
        self.xl.Text(value)

    def delete(self):
        self.xl.Delete()


class ConditionalFormat(base_classes.ConditionalFormat):
    def __init__(self, xl):
        self.xl = xl

    @property
    def api(self):
        return self.xl

    @property
    def type(self):
        return _CONDITIONAL_FORMAT_TYPE_FROM_XL.get(self.xl.Type, "unknown")

    @property
    def stop_if_true(self):
        if self.type in {"color_scale", "data_bar", "icon_set"}:
            return None
        return bool(self.xl.StopIfTrue)

    @property
    def operator(self):
        if self.type != "cell_value":
            return None
        return _CONDITIONAL_FORMAT_OPERATOR_FROM_XL.get(self.xl.Operator)

    @property
    def formula1(self):
        return self.xl.Formula1 if self.type == "cell_value" else None

    @property
    def formula2(self):
        if self.type != "cell_value" or self.operator not in {
            "between",
            "not_between",
        }:
            return None
        return self.xl.Formula2

    @property
    def formula(self):
        return self.xl.Formula1 if self.type == "custom" else None

    @staticmethod
    def _color(obj):
        color_index = obj.ColorIndex
        if color_index in {
            ColorIndex.xlColorIndexNone,
            ColorIndex.xlColorIndexAutomatic,
        }:
            return None
        color = obj.Color
        return None if color is None else int_to_rgb(color)

    @property
    def fill_color(self):
        if self.type not in {"cell_value", "custom"}:
            return None
        return self._color(self.xl.Interior)

    @property
    def font_color(self):
        if self.type not in {"cell_value", "custom"}:
            return None
        return self._color(self.xl.Font)

    @property
    def font_bold(self):
        if self.type not in {"cell_value", "custom"}:
            return None
        return self.xl.Font.Bold

    @property
    def font_italic(self):
        if self.type not in {"cell_value", "custom"}:
            return None
        return self.xl.Font.Italic

    @staticmethod
    def _threshold(criterion):
        criterion_type = _CONDITIONAL_FORMAT_THRESHOLD_FROM_XL.get(
            criterion.Type, "unknown"
        )
        value = (
            None
            if criterion_type
            in {"automatic", "lowest_value", "highest_value", "unknown"}
            else criterion.Value
        )
        return criterion_type, value

    @property
    def colors(self):
        if self.type != "color_scale":
            return None
        criteria = self.xl.ColorScaleCriteria
        return tuple(
            self._color(criteria(index).FormatColor)
            for index in range(1, criteria.Count + 1)
        )

    @property
    def bar_color(self):
        return self._color(self.xl.BarColor) if self.type == "data_bar" else None

    @property
    def gradient(self):
        if self.type != "data_bar":
            return None
        return self.xl.BarFillType == constants.DataBarFillType.xlDataBarFillGradient

    @property
    def show_value(self):
        if self.type == "data_bar":
            return bool(self.xl.ShowValue)
        if self.type == "icon_set":
            return not bool(self.xl.ShowIconOnly)
        return None

    @property
    def icon_set(self):
        if self.type != "icon_set":
            return None
        return _CONDITIONAL_FORMAT_ICON_SET_FROM_XL.get(self.xl.IconSet.ID)

    @property
    def reverse_order(self):
        return bool(self.xl.ReverseOrder) if self.type == "icon_set" else None

    def _threshold_pairs(self):
        if self.type == "color_scale":
            criteria = self.xl.ColorScaleCriteria
            return tuple(
                self._threshold(criteria(index))
                for index in range(1, criteria.Count + 1)
            )
        if self.type == "data_bar":
            return (
                self._threshold(self.xl.MinPoint),
                self._threshold(self.xl.MaxPoint),
            )
        if self.type == "icon_set":
            criteria = self.xl.IconCriteria
            return tuple(
                self._threshold(criteria(index))
                for index in range(2, criteria.Count + 1)
            )
        return None

    @property
    def threshold_types(self):
        pairs = self._threshold_pairs()
        return None if pairs is None else tuple(pair[0] for pair in pairs)

    @property
    def thresholds(self):
        pairs = self._threshold_pairs()
        return None if pairs is None else tuple(pair[1] for pair in pairs)

    def set(self, changes):
        criteria = {"operator", "formula1", "formula2", "formula"} & changes.keys()
        if criteria:
            if self.type == "cell_value":
                kwargs = {
                    "Type": constants.FormatConditionType.xlCellValue,
                    "Operator": _CONDITIONAL_FORMAT_OPERATOR_TO_XL[
                        changes.get("operator", self.operator)
                    ],
                    "Formula1": changes.get("formula1", self.formula1),
                }
                formula2 = changes.get("formula2", self.formula2)
                if formula2 is not None:
                    kwargs["Formula2"] = formula2
                self.xl.Modify(**kwargs)
            else:
                self.xl.Modify(
                    Type=constants.FormatConditionType.xlExpression,
                    Formula1=changes.get("formula", self.formula),
                )
        if "fill_color" in changes:
            self.xl.Interior.Color = rgb_to_int(changes["fill_color"])
        if "font_color" in changes:
            self.xl.Font.Color = rgb_to_int(changes["font_color"])
        if "font_bold" in changes:
            self.xl.Font.Bold = changes["font_bold"]
        if "font_italic" in changes:
            self.xl.Font.Italic = changes["font_italic"]
        if "stop_if_true" in changes:
            self.xl.StopIfTrue = changes["stop_if_true"]

    def delete(self):
        self.xl.Delete()


class ConditionalFormats(Collection, base_classes.ConditionalFormats):
    _wrap = ConditionalFormat

    def _finish_add(self, rule, spec):
        wrapped = ConditionalFormat(rule)
        wrapped.set(
            {
                key: value
                for key, value in spec.items()
                if key
                in {
                    "fill_color",
                    "font_color",
                    "font_bold",
                    "font_italic",
                    "stop_if_true",
                }
            }
        )
        rule.SetFirstPriority()
        return ConditionalFormat(self.xl(1))

    def add_cell_value(self, spec):
        kwargs = {
            "Type": constants.FormatConditionType.xlCellValue,
            "Operator": _CONDITIONAL_FORMAT_OPERATOR_TO_XL[spec["operator"]],
            "Formula1": spec["formula1"],
        }
        if spec["formula2"] is not None:
            kwargs["Formula2"] = spec["formula2"]
        return self._finish_add(self.xl.Add(**kwargs), spec)

    def add_custom(self, spec):
        return self._finish_add(
            self.xl.Add(
                Type=constants.FormatConditionType.xlExpression,
                Formula1=spec["formula"],
            ),
            spec,
        )

    @staticmethod
    def _set_threshold(
        criterion, criterion_type, value, *, automatic_type=None, modify=False
    ):
        xl_type = (
            automatic_type
            if criterion_type == "automatic"
            else _CONDITIONAL_FORMAT_THRESHOLD_TO_XL[criterion_type]
        )
        if modify:
            if value is None:
                criterion.Modify(xl_type)
            else:
                criterion.Modify(xl_type, value)
        else:
            criterion.Type = xl_type
            if value is not None:
                criterion.Value = value

    def _finish_visual_add(self, rule):
        rule.SetFirstPriority()
        return ConditionalFormat(self.xl(1))

    def add_color_scale(self, spec):
        rule = self.xl.AddColorScale(ColorScaleType=len(spec["colors"]))
        for index, (color, criterion_type, value) in enumerate(
            zip(spec["colors"], spec["threshold_types"], spec["thresholds"]),
            start=1,
        ):
            criterion = rule.ColorScaleCriteria(index)
            self._set_threshold(criterion, criterion_type, value)
            criterion.FormatColor.Color = rgb_to_int(color)
        return self._finish_visual_add(rule)

    def add_data_bar(self, spec):
        rule = self.xl.AddDatabar()
        rule.BarColor.Color = rgb_to_int(spec["bar_color"])
        rule.BarFillType = (
            constants.DataBarFillType.xlDataBarFillGradient
            if spec["gradient"]
            else constants.DataBarFillType.xlDataBarFillSolid
        )
        rule.ShowValue = spec["show_value"]
        self._set_threshold(
            rule.MinPoint,
            spec["threshold_types"][0],
            spec["thresholds"][0],
            automatic_type=constants.ConditionValueTypes.xlConditionValueAutomaticMin,
            modify=True,
        )
        self._set_threshold(
            rule.MaxPoint,
            spec["threshold_types"][1],
            spec["thresholds"][1],
            automatic_type=constants.ConditionValueTypes.xlConditionValueAutomaticMax,
            modify=True,
        )
        return self._finish_visual_add(rule)

    def add_icon_set(self, spec):
        rule = self.xl.AddIconSetCondition()
        workbook = rule.AppliesTo.Parent.Parent
        rule.IconSet = workbook.IconSets(
            _CONDITIONAL_FORMAT_ICON_SET_TO_XL[spec["icon_set"]]
        )
        rule.ShowIconOnly = not spec["show_value"]
        rule.ReverseOrder = spec["reverse_order"]
        for index, (criterion_type, value) in enumerate(
            zip(spec["threshold_types"], spec["thresholds"]), start=2
        ):
            criterion = rule.IconCriteria(index)
            self._set_threshold(criterion, criterion_type, value)
            criterion.Operator = constants.FormatConditionOperator.xlGreaterEqual
        return self._finish_visual_add(rule)

    def clear(self):
        self.xl.Delete()


class Shapes(Collection):
    _wrap = Shape


class AutoFilter(base_classes.AutoFilter):
    _COMPARISON_PREFIXES = {
        "equal_to": "=",
        "not_equal_to": "<>",
        "greater_than": ">",
        "less_than": "<",
        "greater_than_or_equal": ">=",
        "less_than_or_equal": "<=",
    }

    def __init__(self, parent, is_table=False):
        self.parent = parent
        self.is_table = is_table

    @staticmethod
    def _escape(value):
        return value.replace("~", "~~").replace("*", "~*").replace("?", "~?")

    @staticmethod
    def _criterion_value(value):
        if not isinstance(value, dict):
            return value
        if value.get("type") == "date":
            parsed = dt.date.fromisoformat(value["value"])
            return f"{parsed.month}/{parsed.day}/{parsed.year}"
        if value.get("type") == "datetime":
            parsed = dt.datetime.fromisoformat(value["value"])
            seconds = f"{parsed.second + parsed.microsecond / 1_000_000:g}"
            return (
                f"{parsed.month}/{parsed.day}/{parsed.year} "
                f"{parsed.hour}:{parsed.minute}:{seconds}"
            )
        raise ValueError("Unknown AutoFilter comparison value type")

    def _criteria(self, operator, value1, value2):
        if value1 is None:
            return ("=" if operator == "equal_to" else "<>"), None, None
        value1 = self._escape(self._criterion_value(value1))
        if operator == "between":
            return (
                f">={value1}",
                constants.AutoFilterOperator.xlAnd,
                f"<={self._escape(self._criterion_value(value2))}",
            )
        if operator == "not_between":
            return (
                f"<{value1}",
                constants.AutoFilterOperator.xlOr,
                f">{self._escape(self._criterion_value(value2))}",
            )
        return f"{self._COMPARISON_PREFIXES[operator]}{value1}", None, None

    @property
    def _range(self):
        return self.parent.xl.Range if self.is_table else self.parent.xl

    @property
    def _column_count(self):
        return self.parent.range.shape[1] if self.is_table else self.parent.shape[1]

    def _worksheet_filter_range(self):
        sheet = self.parent.sheet.xl
        if not sheet.AutoFilterMode:
            return None
        autofilter = sheet.AutoFilter
        return autofilter.Range if autofilter is not None else None

    def _ensure_target(self):
        if self.is_table:
            return
        existing = self._worksheet_filter_range()
        if existing is not None and existing.Address != self.parent.xl.Address:
            raise ValueError(
                "This worksheet already has an AutoFilter on a different range"
            )

    def _native_autofilter(self):
        if self.is_table:
            return self.parent.xl.AutoFilter
        existing = self._worksheet_filter_range()
        if existing is None or existing.Address != self.parent.xl.Address:
            return None
        return self.parent.sheet.xl.AutoFilter

    @property
    def criteria(self):
        autofilter = self._native_autofilter()
        if autofilter is None:
            return [
                base_classes.empty_autofilter_criteria(field)
                for field in range(1, self._column_count + 1)
            ]
        snapshots = []
        operator_types = {
            constants.AutoFilterOperator.xlFilterValues: "values",
            constants.AutoFilterOperator.xlTop10Items: "top_items",
            constants.AutoFilterOperator.xlBottom10Items: "bottom_items",
            constants.AutoFilterOperator.xlTop10Percent: "top_percent",
            constants.AutoFilterOperator.xlBottom10Percent: "bottom_percent",
        }
        for field in range(1, self._column_count + 1):
            native_filter = autofilter.Filters(field)
            if not native_filter.On:
                snapshots.append(base_classes.empty_autofilter_criteria(field))
                continue
            try:
                operator = native_filter.Operator
            except Exception:
                operator = None
            type_ = operator_types.get(operator, "comparison")
            if operator not in (
                0,
                constants.AutoFilterOperator.xlAnd,
                constants.AutoFilterOperator.xlOr,
                *operator_types,
            ):
                type_ = "unknown"
            try:
                criteria2 = native_filter.Criteria2
            except Exception:
                criteria2 = None
            try:
                criteria1 = native_filter.Criteria1
            except Exception:
                criteria1 = None
            snapshots.append(
                base_classes.autofilter_criteria_snapshot(
                    field,
                    type_,
                    criteria1,
                    criteria2,
                    {
                        constants.AutoFilterOperator.xlAnd: "and",
                        constants.AutoFilterOperator.xlOr: "or",
                    }.get(operator),
                )
            )
        return snapshots

    def apply_values(self, field, values):
        self._ensure_target()
        self._range.AutoFilter(
            Field=field,
            Criteria1=values,
            Operator=constants.AutoFilterOperator.xlFilterValues,
        )

    def apply_comparison(self, field, operator, value1, value2):
        self._ensure_target()
        criteria1, native_operator, criteria2 = self._criteria(operator, value1, value2)
        kwargs = {"Field": field, "Criteria1": criteria1}
        if native_operator is not None:
            kwargs["Operator"] = native_operator
            kwargs["Criteria2"] = criteria2
        self._range.AutoFilter(**kwargs)

    def _apply_top_bottom(self, field, value, operator):
        self._ensure_target()
        self._range.AutoFilter(Field=field, Criteria1=str(value), Operator=operator)

    def apply_top_items(self, field, count):
        self._apply_top_bottom(field, count, constants.AutoFilterOperator.xlTop10Items)

    def apply_bottom_items(self, field, count):
        self._apply_top_bottom(
            field, count, constants.AutoFilterOperator.xlBottom10Items
        )

    def apply_top_percent(self, field, percent):
        self._apply_top_bottom(
            field, percent, constants.AutoFilterOperator.xlTop10Percent
        )

    def apply_bottom_percent(self, field, percent):
        self._apply_top_bottom(
            field, percent, constants.AutoFilterOperator.xlBottom10Percent
        )

    def clear(self, field):
        if not self.is_table:
            existing = self._worksheet_filter_range()
            if existing is None or existing.Address != self.parent.xl.Address:
                return
        fields = [field] if field is not None else range(1, self._column_count + 1)
        for field_index in fields:
            self._range.AutoFilter(Field=field_index)


class Table(base_classes.Table):
    def __init__(self, xl):
        self.xl = xl

    @property
    def api(self):
        return self.xl

    @property
    def name(self):
        return self.xl.Name

    @name.setter
    def name(self, value):
        self.xl.Name = value

    @property
    def data_body_range(self):
        return Range(xl=self.xl.DataBodyRange) if self.xl.DataBodyRange else None

    @property
    def display_name(self):
        return self.xl.DisplayName

    @display_name.setter
    def display_name(self, value):
        self.xl.DisplayName = value

    @property
    def header_row_range(self):
        return Range(xl=self.xl.HeaderRowRange)

    @property
    def insert_row_range(self):
        return Range(xl=self.xl.InsertRowRange)

    @property
    def parent(self):
        return Sheet(xl=self.xl.Parent)

    @property
    def range(self):
        return Range(xl=self.xl.Range)

    @property
    def autofilter(self):
        return AutoFilter(self, is_table=True)

    @property
    def show_autofilter(self):
        return self.xl.ShowAutoFilter

    @show_autofilter.setter
    def show_autofilter(self, value):
        self.xl.ShowAutoFilter = value

    @property
    def show_headers(self):
        return self.xl.ShowHeaders

    @show_headers.setter
    def show_headers(self, value):
        self.xl.ShowHeaders = value

    @property
    def show_table_style_column_stripes(self):
        return self.xl.ShowTableStyleColumnStripes

    @show_table_style_column_stripes.setter
    def show_table_style_column_stripes(self, value):
        self.xl.ShowTableStyleColumnStripes = value

    @property
    def show_table_style_first_column(self):
        return self.xl.ShowTableStyleFirstColumn

    @show_table_style_first_column.setter
    def show_table_style_first_column(self, value):
        self.xl.ShowTableStyleFirstColumn = value

    @property
    def show_table_style_last_column(self):
        return self.xl.ShowTableStyleLastColumn

    @show_table_style_last_column.setter
    def show_table_style_last_column(self, value):
        self.xl.ShowTableStyleLastColumn = value

    @property
    def show_table_style_row_stripes(self):
        return self.xl.ShowTableStyleRowStripes

    @show_table_style_row_stripes.setter
    def show_table_style_row_stripes(self, value):
        self.xl.ShowTableStyleRowStripes = value

    @property
    def show_totals(self):
        return self.xl.ShowTotals

    @show_totals.setter
    def show_totals(self, value):
        self.xl.ShowTotals = value

    @property
    def table_style(self):
        return self.xl.TableStyle.Name

    @table_style.setter
    def table_style(self, value):
        self.xl.TableStyle = value

    @property
    def totals_row_range(self):
        return Range(xl=self.xl.TotalsRowRange)

    def resize(self, range):
        self.xl.Resize(range.api)


class Tables(Collection, base_classes.Tables):
    _wrap = Table

    def add(
        self,
        source_type=None,
        source=None,
        link_source=None,
        has_headers=None,
        destination=None,
        table_style_name=None,
        name=None,
    ):
        table = Table(
            xl=self.xl.Add(
                SourceType=ListObjectSourceType.xlSrcRange,
                Source=source.api,
                LinkSource=link_source,
                XlListObjectHasHeaders=True,
                Destination=destination,
                TableStyleName=table_style_name,
            )
        )
        if name is not None:
            table.name = name
        return table


class Chart(base_classes.Chart):
    def __init__(self, xl_obj=None, xl=None):
        self.xl = xl_obj.Chart if xl is None else xl
        self.xl_obj = xl_obj

    @property
    def api(self):
        return self.xl_obj, self.xl

    @property
    def name(self):
        if self.xl_obj is None:
            return self.xl.Name
        else:
            return self.xl_obj.Name

    @name.setter
    def name(self, value):
        if self.xl_obj is None:
            self.xl.Name = value
        else:
            self.xl_obj.Name = value

    @property
    def parent(self):
        if self.xl_obj is None:
            return Book(xl=self.xl.Parent)
        else:
            return Sheet(xl=self.xl_obj.Parent)

    def set_source_data(self, rng, plot_by=None):
        if plot_by is None:
            self.xl.SetSourceData(rng.xl)
        else:
            self.xl.SetSourceData(rng.xl, plot_by_s2i[plot_by])

    def set_x_axis_values(self, rng):
        series_collection = self.xl.SeriesCollection()
        for index in range(1, series_collection.Count + 1):
            series_collection(index).XValues = rng.xl

    @property
    def chart_type(self):
        return chart_types_i2s[self.xl.ChartType]

    @chart_type.setter
    def chart_type(self, chart_type):
        self.xl.ChartType = chart_types_s2i[chart_type]

    @property
    def title(self):
        return self.xl.ChartTitle.Text if self.xl.HasTitle else None

    @title.setter
    def title(self, value):
        if value is None:
            self.xl.HasTitle = False
        else:
            # ChartTitle is only reachable once HasTitle is True
            self.xl.HasTitle = True
            self.xl.ChartTitle.Text = value

    @property
    def legend(self):
        return ChartLegend(self)

    @property
    def category_axis(self):
        return ChartAxis(self, "category")

    @property
    def value_axis(self):
        return ChartAxis(self, "value")

    @property
    def series(self):
        return ChartSeriesCollection(self.xl.SeriesCollection())

    @property
    def plot_by(self):
        return plot_by_i2s[self.xl.PlotBy]

    @plot_by.setter
    def plot_by(self, value):
        self.xl.PlotBy = plot_by_s2i[value]

    @property
    def style(self):
        return self.xl.ChartStyle

    @style.setter
    def style(self, value):
        self.xl.ChartStyle = value

    @property
    def left(self):
        if self.xl_obj is None:
            raise Exception("This chart is not embedded.")
        return self.xl_obj.Left

    @left.setter
    def left(self, value):
        if self.xl_obj is None:
            raise Exception("This chart is not embedded.")
        self.xl_obj.Left = value

    @property
    def top(self):
        if self.xl_obj is None:
            raise Exception("This chart is not embedded.")
        return self.xl_obj.Top

    @top.setter
    def top(self, value):
        if self.xl_obj is None:
            raise Exception("This chart is not embedded.")
        self.xl_obj.Top = value

    @property
    def width(self):
        if self.xl_obj is None:
            raise Exception("This chart is not embedded.")
        return self.xl_obj.Width

    @width.setter
    def width(self, value):
        if self.xl_obj is None:
            raise Exception("This chart is not embedded.")
        self.xl_obj.Width = value

    @property
    def height(self):
        if self.xl_obj is None:
            raise Exception("This chart is not embedded.")
        return self.xl_obj.Height

    @height.setter
    def height(self, value):
        if self.xl_obj is None:
            raise Exception("This chart is not embedded.")
        self.xl_obj.Height = value

    def delete(self):
        if self.xl_obj is None:
            # chart sheet
            self.xl.Delete()
        else:
            self.xl_obj.Delete()

    def to_png(self, path):
        self.xl.Export(path)

    def to_pdf(self, path, quality):
        self.xl_obj.Select()
        self.xl.ExportAsFixedFormat(
            Type=FixedFormatType.xlTypePDF,
            Filename=path,
            Quality=quality_types[quality],
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False,
        )
        try:
            self.parent.range("A1").select()
        except:  # noqa: E722
            pass


class ChartSeries(base_classes.ChartSeries):
    def __init__(self, xl):
        self.xl = xl

    @property
    def api(self):
        return self.xl

    @property
    def name(self):
        return self.xl.Name

    @name.setter
    def name(self, value):
        self.xl.Name = value

    @property
    def marker_style(self):
        return marker_styles_i2s[self.xl.MarkerStyle]

    @marker_style.setter
    def marker_style(self, value):
        self.xl.MarkerStyle = marker_styles_s2i[value]

    @property
    def marker_size(self):
        return int(self.xl.MarkerSize)

    @marker_size.setter
    def marker_size(self, value):
        self.xl.MarkerSize = value

    @staticmethod
    def _color(value):
        return None if value is None or value < 0 else int_to_rgb(value)

    @property
    def marker_foreground_color(self):
        return self._color(self.xl.MarkerForegroundColor)

    @marker_foreground_color.setter
    def marker_foreground_color(self, value):
        self.xl.MarkerForegroundColor = rgb_to_int(value)

    @property
    def marker_background_color(self):
        return self._color(self.xl.MarkerBackgroundColor)

    @marker_background_color.setter
    def marker_background_color(self, value):
        self.xl.MarkerBackgroundColor = rgb_to_int(value)

    @property
    def line_color(self):
        line = self.xl.Format.Line
        return None if not line.Visible else self._color(line.ForeColor.RGB)

    @line_color.setter
    def line_color(self, value):
        line = self.xl.Format.Line
        line.Visible = True
        line.ForeColor.RGB = rgb_to_int(value)

    @property
    def fill_color(self):
        fill = self.xl.Format.Fill
        return None if not fill.Visible else self._color(fill.ForeColor.RGB)

    @fill_color.setter
    def fill_color(self, value):
        fill = self.xl.Format.Fill
        fill.Visible = True
        fill.Solid()
        fill.ForeColor.RGB = rgb_to_int(value)

    def set(
        self,
        *,
        name=base_classes._UNSET,
        marker_style=base_classes._UNSET,
        marker_size=base_classes._UNSET,
        marker_foreground_color=base_classes._UNSET,
        marker_background_color=base_classes._UNSET,
        line_color=base_classes._UNSET,
        fill_color=base_classes._UNSET,
    ):
        for attribute, value in (
            ("name", name),
            # Excel may propagate series line/fill formatting to markers. Apply
            # explicit marker overrides afterwards so one bulk set preserves
            # independently requested colors.
            ("line_color", line_color),
            ("fill_color", fill_color),
            ("marker_style", marker_style),
            ("marker_size", marker_size),
            ("marker_foreground_color", marker_foreground_color),
            ("marker_background_color", marker_background_color),
        ):
            if value is not base_classes._UNSET:
                setattr(self, attribute, value)


class ChartSeriesCollection(Collection, base_classes.ChartSeriesCollection):
    _wrap = ChartSeries

    def __call__(self, key):
        if not isinstance(key, numbers.Integral) or isinstance(key, bool):
            raise KeyError(key)
        return super().__call__(key)


class ChartAxis(base_classes.ChartAxis):
    _axis_types = {
        "category": AxisType.xlCategory,
        "value": AxisType.xlValue,
    }

    def __init__(self, parent, axis_type):
        self.parent = parent
        self.axis_type = axis_type

    @property
    def _xl_axis_type(self):
        return self._axis_types[self.axis_type]

    @property
    def visible(self):
        # HasAxis is an indexed COM property. The generated pywin32 wrapper
        # treats attribute access as a zero-argument property read, even when
        # called with the two indexes, so invoke the property directly. Wrap
        # both low-level calls to retain the normal retry-on-busy behavior.
        oleobj = self.parent.xl._oleobj_
        dispid = COMRetryMethodWrapper(oleobj.GetIDsOfNames)(0, "HasAxis")
        return bool(
            COMRetryMethodWrapper(oleobj.Invoke)(
                dispid,
                0,
                pythoncom.DISPATCH_PROPERTYGET,
                1,
                self._xl_axis_type,
                AxisGroup.xlPrimary,
            )
        )

    @visible.setter
    def visible(self, value):
        # Python assignment can't express the two indexes, so invoke the
        # property-put directly with its runtime-resolved DISPID. Wrap both
        # low-level calls to retain the normal retry-on-busy behavior.
        oleobj = self.parent.xl._oleobj_
        dispid = COMRetryMethodWrapper(oleobj.GetIDsOfNames)(0, "HasAxis")
        COMRetryMethodWrapper(oleobj.Invoke)(
            dispid,
            0,
            pythoncom.DISPATCH_PROPERTYPUT,
            0,
            self._xl_axis_type,
            AxisGroup.xlPrimary,
            bool(value),
        )

    def _axis(self):
        if not self.visible:
            raise xlwings.XlwingsError(
                f"The chart has no visible primary {self.axis_type} axis."
            )
        return self.parent.xl.Axes(self._xl_axis_type, AxisGroup.xlPrimary)

    @property
    def api(self):
        return self._axis() if self.visible else None

    @property
    def title(self):
        if not self.visible:
            return None
        axis = self._axis()
        return axis.AxisTitle.Text if axis.HasTitle else None

    @title.setter
    def title(self, value):
        if value is None:
            if self.visible:
                self._axis().HasTitle = False
            return
        self.visible = True
        axis = self._axis()
        axis.HasTitle = True
        axis.AxisTitle.Text = value

    def _get_scale(self, attribute):
        return float(getattr(self._axis(), attribute))

    def _set_scale(self, attribute, auto_attribute, value):
        axis = self._axis()
        if value is None:
            setattr(axis, auto_attribute, True)
        else:
            setattr(axis, attribute, value)

    @property
    def minimum_scale(self):
        return self._get_scale("MinimumScale")

    @minimum_scale.setter
    def minimum_scale(self, value):
        self._set_scale("MinimumScale", "MinimumScaleIsAuto", value)

    @property
    def maximum_scale(self):
        return self._get_scale("MaximumScale")

    @maximum_scale.setter
    def maximum_scale(self, value):
        self._set_scale("MaximumScale", "MaximumScaleIsAuto", value)

    @property
    def major_unit(self):
        return self._get_scale("MajorUnit")

    @major_unit.setter
    def major_unit(self, value):
        self._set_scale("MajorUnit", "MajorUnitIsAuto", value)

    @property
    def number_format(self):
        return self._axis().TickLabels.NumberFormat

    @number_format.setter
    def number_format(self, value):
        self._axis().TickLabels.NumberFormat = value

    def set(
        self,
        *,
        title=base_classes._UNSET,
        minimum_scale=base_classes._UNSET,
        maximum_scale=base_classes._UNSET,
        major_unit=base_classes._UNSET,
        number_format=base_classes._UNSET,
        visible=base_classes._UNSET,
    ):
        attributes = (title, minimum_scale, maximum_scale, major_unit, number_format)
        if visible is True or (
            visible is False
            and any(value is not base_classes._UNSET for value in attributes)
        ):
            self.visible = True
        if title is not base_classes._UNSET:
            self.title = title
        if minimum_scale is not base_classes._UNSET:
            self.minimum_scale = minimum_scale
        if maximum_scale is not base_classes._UNSET:
            self.maximum_scale = maximum_scale
        if major_unit is not base_classes._UNSET:
            self.major_unit = major_unit
        if number_format is not base_classes._UNSET:
            self.number_format = number_format
        if visible is False:
            self.visible = False


class ChartLegend(base_classes.ChartLegend):
    def __init__(self, parent):
        self.parent = parent

    @property
    def xl(self):
        return self.parent.xl

    @property
    def api(self):
        return self.xl.Legend if self.xl.HasLegend else None

    @property
    def visible(self):
        return bool(self.xl.HasLegend)

    @visible.setter
    def visible(self, value):
        self.xl.HasLegend = value

    @property
    def position(self):
        if not self.xl.HasLegend:
            # Legend.Position raises on a hidden legend
            return None
        return legend_positions_i2s.get(self.xl.Legend.Position)

    @position.setter
    def position(self, value):
        self.xl.HasLegend = True
        self.xl.Legend.Position = legend_positions_s2i[value]


class Charts(Collection, base_classes.Charts):
    @property
    def parent(self):
        return Sheet(xl=self.xl.Parent)

    def _wrap(self, xl):
        return Chart(xl_obj=xl)

    def add(
        self,
        left,
        top,
        width,
        height,
        chart_type=None,
        source=None,
        plot_by=None,
        name=None,
        anchor=None,
        style=227,
    ):
        if anchor:
            top, left = anchor.top, anchor.left
        chart = Chart(xl_obj=self.xl.Add(left, top, width, height))
        # data before type: stock/xy types need series to exist
        if source is not None:
            chart.set_source_data(source, plot_by)
        if chart_type is not None:
            chart.chart_type = chart_type
        if style is not None:
            chart.style = style
        if name is not None:
            chart.name = name
        return chart


class PivotTable(base_classes.PivotTable):
    def __init__(self, xl):
        self.xl = xl

    @property
    def api(self):
        return self.xl

    @property
    def parent(self):
        return Sheet(xl=self.xl.Parent)

    @property
    def name(self):
        return self.xl.Name

    @name.setter
    def name(self, value):
        self.xl.Name = value

    def _values_pseudo_field_name(self):
        """The name of the "Values" pseudo field that Excel places in the
        columns (or rows) area once there are two or more value fields."""
        try:
            return self.xl.DataPivotField.Name
        except pywintypes.com_error:
            return None

    @property
    def field_names(self):
        # PivotFields lists the source fields plus the "Values" pseudo field;
        # value fields themselves have the data orientation.
        values_name = self._values_pseudo_field_name()
        return [
            field.Name
            for field in self.xl.PivotFields()
            if field.Name != values_name
            and field.Orientation != PivotFieldOrientation.xlDataField
        ]

    @property
    def rows(self):
        return PivotFields(pivot=self, area="rows")

    @property
    def columns(self):
        return PivotFields(pivot=self, area="columns")

    @property
    def filters(self):
        return PivotFields(pivot=self, area="filters")

    @property
    def values(self):
        return PivotValueFields(pivot=self)

    @property
    def layout(self):
        # LayoutRowDefault only applies to fields added later, so read the
        # actual layout off the row fields; None when they disagree.
        row_fields = list(self.xl.RowFields)
        if not row_fields:
            return layout_row_types_i2s.get(self.xl.LayoutRowDefault)
        layouts = set()
        for field in row_fields:
            if field.LayoutCompactRow:
                layouts.add("compact")
            elif field.LayoutForm == LayoutFormType.xlTabular:
                layouts.add("tabular")
            else:
                layouts.add("outline")
        return layouts.pop() if len(layouts) == 1 else None

    @layout.setter
    def layout(self, value):
        self.xl.RowAxisLayout(layout_row_types_s2i[value])
        self.xl.LayoutRowDefault = layout_row_types_s2i[value]

    @property
    def show_row_grand_totals(self):
        return bool(self.xl.RowGrand)

    @show_row_grand_totals.setter
    def show_row_grand_totals(self, value):
        self.xl.RowGrand = value

    @property
    def show_column_grand_totals(self):
        return bool(self.xl.ColumnGrand)

    @show_column_grand_totals.setter
    def show_column_grand_totals(self, value):
        self.xl.ColumnGrand = value

    @property
    def range(self):
        return Range(xl=self.xl.TableRange1)

    @property
    def data_body_range(self):
        if self.xl.DataFields.Count == 0:
            return None
        return Range(xl=self.xl.DataBodyRange)

    def refresh(self):
        self.xl.RefreshTable()

    def delete(self):
        # There is no PivotTable.Delete; clearing the full report range
        # (incl. the filters area) removes it.
        self.xl.TableRange2.Clear()


class PivotField(base_classes.PivotField):
    def __init__(self, xl, pivot):
        # xl is the source field (PivotFields(name)), so the wrapper follows
        # the field when it is moved to another area
        self.xl = xl
        self._pivot = pivot

    @property
    def api(self):
        return self.xl

    @property
    def parent(self):
        return self._pivot

    @property
    def name(self):
        return self.xl.Name

    def remove(self):
        self.xl.Orientation = PivotFieldOrientation.xlHidden


class PivotFields(base_classes.PivotFields):
    def __init__(self, pivot, area):
        self._pivot = pivot
        self._area = area

    @property
    def xl(self):
        # RowFields/ColumnFields/PageFields are parameterized properties, which
        # pywin32's early binding exposes as plain attributes (calling the
        # returned collection fails). They're snapshots, so fetch them on
        # every access.
        return getattr(self._pivot.xl, pivot_area_collections[self._area])

    @property
    def api(self):
        return self.xl

    @property
    def parent(self):
        return self._pivot

    @property
    def area(self):
        return self._area

    def _fields(self):
        # Excel's "Values" pseudo field isn't a source field: hide it, like
        # field_names does and like Office.js
        values_name = self._pivot._values_pseudo_field_name()
        return [xl for xl in self.xl if xl.Name != values_name]

    def __call__(self, key):
        fields = self._fields()
        if isinstance(key, numbers.Number):
            if key < 1 or key > len(fields):
                raise KeyError(key)
            return PivotField(xl=fields[key - 1], pivot=self._pivot)
        for xl in fields:
            if xl.Name == key:
                return PivotField(xl=xl, pivot=self._pivot)
        raise KeyError(key)

    def __len__(self):
        return len(self._fields())

    def __iter__(self):
        for xl in self._fields():
            yield PivotField(xl=xl, pivot=self._pivot)

    def __contains__(self, key):
        if isinstance(key, numbers.Number):
            return 1 <= key <= len(self)
        return any(xl.Name == key for xl in self._fields())

    def add(self, name):
        try:
            field = self._pivot.xl.PivotFields(name)
        except pywintypes.com_error:
            raise KeyError(name)
        orientation = pivot_area_orientations[self._area]
        # setting the orientation appends the field to the area; leave a
        # field that is already here where it is
        if field.Orientation != orientation:
            field.Orientation = orientation
        return PivotField(xl=field, pivot=self._pivot)


class PivotValueField(base_classes.PivotValueField):
    def __init__(self, xl, pivot):
        self.xl = xl
        self._pivot = pivot

    @property
    def api(self):
        return self.xl

    @property
    def parent(self):
        return self._pivot

    @property
    def name(self):
        return self.xl.Name

    @name.setter
    def name(self, value):
        self.xl.Name = value

    @property
    def source_field(self):
        return self.xl.SourceName

    @property
    def function(self):
        return pivot_functions_i2s.get(self.xl.Function)

    @function.setter
    def function(self, value):
        self.xl.Function = pivot_functions_s2i[value]

    @property
    def number_format(self):
        return self.xl.NumberFormat

    @number_format.setter
    def number_format(self, value):
        self.xl.NumberFormat = value

    def remove(self):
        self.xl.Orientation = PivotFieldOrientation.xlHidden


class PivotValueFields(base_classes.PivotValueFields):
    def __init__(self, pivot):
        self._pivot = pivot

    @property
    def xl(self):
        # a parameterized property, see PivotFields.xl
        return self._pivot.xl.DataFields

    @property
    def api(self):
        return self.xl

    @property
    def parent(self):
        return self._pivot

    def __call__(self, key):
        try:
            return PivotValueField(xl=self.xl.Item(key), pivot=self._pivot)
        except pywintypes.com_error:
            raise KeyError(key)

    def __len__(self):
        return self.xl.Count

    def __iter__(self):
        for xl in self.xl:
            yield PivotValueField(xl=xl, pivot=self._pivot)

    def __contains__(self, key):
        try:
            self.xl.Item(key)
            return True
        except pywintypes.com_error:
            return False

    def add(self, field, function=None, name=None, number_format=None):
        try:
            source = self._pivot.xl.PivotFields(field)
        except pywintypes.com_error:
            raise KeyError(field)
        kwargs = {}
        if name is not None:
            kwargs["Caption"] = name
        if function is not None:
            kwargs["Function"] = pivot_functions_s2i[function]
        xl = self._pivot.xl.AddDataField(source, **kwargs)
        if number_format is not None:
            xl.NumberFormat = number_format
        return PivotValueField(xl=xl, pivot=self._pivot)


class PivotTables(Collection, base_classes.PivotTables):
    _wrap = PivotTable

    @property
    def parent(self):
        return Sheet(xl=self.xl.Parent)

    def add(self, source, destination, name=None):
        if isinstance(source, Table):
            source_data = source.xl.Name
        else:
            # Microsoft's docs warn that passing a Range object as SourceData
            # can raise a type mismatch, so pass an external R1C1 address
            # (GetAddress: pywin32's form of the parameterized Address property)
            source_data = source.xl.GetAddress(True, True, ReferenceStyle.xlR1C1, True)
        book = self.xl.Parent.Parent
        cache = book.PivotCaches().Create(
            SourceType=PivotTableSourceType.xlDatabase, SourceData=source_data
        )
        kwargs = {"TableDestination": destination.xl.Cells(1, 1)}
        if name:
            kwargs["TableName"] = name
        return PivotTable(xl=cache.CreatePivotTable(**kwargs))


class Picture(base_classes.Picture):
    def __init__(self, xl):
        self.xl = xl

    @property
    def api(self):
        return self.xl

    @property
    def name(self):
        return self.xl.Name

    @name.setter
    def name(self, value):
        self.xl.Name = value

    @property
    def parent(self):
        return Sheet(xl=self.xl.Parent)

    @property
    def left(self):
        return self.xl.Left

    @left.setter
    def left(self, value):
        self.xl.Left = value

    @property
    def top(self):
        return self.xl.Top

    @top.setter
    def top(self, value):
        self.xl.Top = value

    @property
    def width(self):
        return self.xl.Width

    @width.setter
    def width(self, value):
        self.xl.Width = value

    @property
    def height(self):
        return self.xl.Height

    @height.setter
    def height(self, value):
        self.xl.Height = value

    def delete(self):
        self.xl.Delete()

    @property
    def lock_aspect_ratio(self):
        return self.xl.ShapeRange.LockAspectRatio

    @lock_aspect_ratio.setter
    def lock_aspect_ratio(self, value):
        self.xl.ShapeRange.LockAspectRatio = value

    def update(self, filename):
        return utils.excel_update_picture(self, filename)


class Pictures(Collection, base_classes.Pictures):
    _wrap = Picture

    @property
    def parent(self):
        return Sheet(xl=self.xl.Parent)

    def add(
        self,
        filename,
        link_to_file,
        save_with_document,
        left,
        top,
        width,
        height,
        anchor,
    ):
        if anchor:
            top, left = anchor.top, anchor.left
        else:
            top = top if top else 0
            left = left if left else 0

        return Picture(
            xl=self.xl.Parent.Shapes.AddPicture(
                Filename=filename,
                LinkToFile=link_to_file,
                SaveWithDocument=save_with_document,
                Left=left,
                Top=top,
                Width=width,
                Height=height,
            ).DrawingObject
        )


class Names(base_classes.Names):
    def __init__(self, xl):
        self.xl = xl

    @property
    def api(self):
        return self.xl

    def __call__(self, name_or_index):
        return Name(xl=self.xl(name_or_index))

    def contains(self, name_or_index):
        try:
            self.xl(name_or_index)
        except pywintypes.com_error as e:
            if e.hresult == -2147352567:
                return False
            else:
                raise
        return True

    def __len__(self):
        return self.xl.Count

    def add(self, name, refers_to):
        return Name(xl=self.xl.Add(name, refers_to))


class Name(base_classes.Name):
    def __init__(self, xl):
        self.xl = xl

    @property
    def api(self):
        return self.xl

    def delete(self):
        self.xl.Delete()

    @property
    def name(self):
        return self.xl.Name

    @name.setter
    def name(self, value):
        self.xl.Name = value

    @property
    def refers_to(self):
        return self.xl.RefersTo

    @refers_to.setter
    def refers_to(self, value):
        self.xl.RefersTo = value

    @property
    def refers_to_range(self):
        return Range(xl=self.xl.RefersToRange)


# --- constants ---
quality_types = {"minimum": 1, "standard": 0}

chart_types_s2i = {
    "3d_area": -4098,
    "3d_area_stacked": 78,
    "3d_area_stacked_100": 79,
    "3d_bar_clustered": 60,
    "3d_bar_stacked": 61,
    "3d_bar_stacked_100": 62,
    "3d_column": -4100,
    "3d_column_clustered": 54,
    "3d_column_stacked": 55,
    "3d_column_stacked_100": 56,
    "3d_line": -4101,
    "3d_pie": -4102,
    "3d_pie_exploded": 70,
    "area": 1,
    "area_stacked": 76,
    "area_stacked_100": 77,
    "bar_clustered": 57,
    "bar_of_pie": 71,
    "bar_stacked": 58,
    "bar_stacked_100": 59,
    "bubble": 15,
    "bubble_3d_effect": 87,
    "column_clustered": 51,
    "column_stacked": 52,
    "column_stacked_100": 53,
    "cone_bar_clustered": 102,
    "cone_bar_stacked": 103,
    "cone_bar_stacked_100": 104,
    "cone_col": 105,
    "cone_col_clustered": 99,
    "cone_col_stacked": 100,
    "cone_col_stacked_100": 101,
    "cylinder_bar_clustered": 95,
    "cylinder_bar_stacked": 96,
    "cylinder_bar_stacked_100": 97,
    "cylinder_col": 98,
    "cylinder_col_clustered": 92,
    "cylinder_col_stacked": 93,
    "cylinder_col_stacked_100": 94,
    "doughnut": -4120,
    "doughnut_exploded": 80,
    "line": 4,
    "line_markers": 65,
    "line_markers_stacked": 66,
    "line_markers_stacked_100": 67,
    "line_stacked": 63,
    "line_stacked_100": 64,
    "pie": 5,
    "pie_exploded": 69,
    "pie_of_pie": 68,
    "pyramid_bar_clustered": 109,
    "pyramid_bar_stacked": 110,
    "pyramid_bar_stacked_100": 111,
    "pyramid_col": 112,
    "pyramid_col_clustered": 106,
    "pyramid_col_stacked": 107,
    "pyramid_col_stacked_100": 108,
    "radar": -4151,
    "radar_filled": 82,
    "radar_markers": 81,
    "stock_hlc": 88,
    "stock_ohlc": 89,
    "stock_vhlc": 90,
    "stock_vohlc": 91,
    "surface": 83,
    "surface_top_view": 85,
    "surface_top_view_wireframe": 86,
    "surface_wireframe": 84,
    "xy_scatter": -4169,
    "xy_scatter_lines": 74,
    "xy_scatter_lines_no_markers": 75,
    "xy_scatter_smooth": 72,
    "xy_scatter_smooth_no_markers": 73,
}

chart_types_i2s = {v: k for k, v in chart_types_s2i.items()}

plot_by_s2i = {"rows": RowCol.xlRows, "columns": RowCol.xlColumns}
plot_by_i2s = {v: k for k, v in plot_by_s2i.items()}

legend_positions_s2i = {
    "top": LegendPosition.xlLegendPositionTop,
    "bottom": LegendPosition.xlLegendPositionBottom,
    "left": LegendPosition.xlLegendPositionLeft,
    "right": LegendPosition.xlLegendPositionRight,
    "corner": LegendPosition.xlLegendPositionCorner,
}
legend_positions_i2s = {v: k for k, v in legend_positions_s2i.items()}

marker_styles_s2i = {
    "automatic": constants.MarkerStyle.xlMarkerStyleAutomatic,
    "none": constants.MarkerStyle.xlMarkerStyleNone,
    "square": constants.MarkerStyle.xlMarkerStyleSquare,
    "diamond": constants.MarkerStyle.xlMarkerStyleDiamond,
    "triangle": constants.MarkerStyle.xlMarkerStyleTriangle,
    "x": constants.MarkerStyle.xlMarkerStyleX,
    "star": constants.MarkerStyle.xlMarkerStyleStar,
    "dot": constants.MarkerStyle.xlMarkerStyleDot,
    "dash": constants.MarkerStyle.xlMarkerStyleDash,
    "circle": constants.MarkerStyle.xlMarkerStyleCircle,
    "plus": constants.MarkerStyle.xlMarkerStylePlus,
}
marker_styles_i2s = {v: k for k, v in marker_styles_s2i.items()}
# only ever read back, e.g. after a user dragged the legend
legend_positions_i2s[LegendPosition.xlLegendPositionCustom] = "custom"

horizontal_alignments_s2i = {
    "general": HAlign.xlHAlignGeneral,
    "left": HAlign.xlHAlignLeft,
    "center": HAlign.xlHAlignCenter,
    "right": HAlign.xlHAlignRight,
    "fill": HAlign.xlHAlignFill,
    "justify": HAlign.xlHAlignJustify,
    "center_across_selection": HAlign.xlHAlignCenterAcrossSelection,
    "distributed": HAlign.xlHAlignDistributed,
}
horizontal_alignments_i2s = {v: k for k, v in horizontal_alignments_s2i.items()}

vertical_alignments_s2i = {
    "top": VAlign.xlVAlignTop,
    "center": VAlign.xlVAlignCenter,
    "bottom": VAlign.xlVAlignBottom,
    "justify": VAlign.xlVAlignJustify,
    "distributed": VAlign.xlVAlignDistributed,
}
vertical_alignments_i2s = {v: k for k, v in vertical_alignments_s2i.items()}

pivot_functions_s2i = {
    "sum": ConsolidationFunction.xlSum,
    "count": ConsolidationFunction.xlCount,
    "average": ConsolidationFunction.xlAverage,
    "max": ConsolidationFunction.xlMax,
    "min": ConsolidationFunction.xlMin,
    "product": ConsolidationFunction.xlProduct,
    "count_numbers": ConsolidationFunction.xlCountNums,
    "stdev": ConsolidationFunction.xlStDev,
    "stdevp": ConsolidationFunction.xlStDevP,
    "var": ConsolidationFunction.xlVar,
    "varp": ConsolidationFunction.xlVarP,
}
pivot_functions_i2s = {v: k for k, v in pivot_functions_s2i.items()}

layout_row_types_s2i = {
    "compact": LayoutRowType.xlCompactRow,
    "outline": LayoutRowType.xlOutlineRow,
    "tabular": LayoutRowType.xlTabularRow,
}
layout_row_types_i2s = {v: k for k, v in layout_row_types_s2i.items()}

# xlwings' field areas -> the PivotTable collection method / the orientation
pivot_area_collections = {
    "rows": "RowFields",
    "columns": "ColumnFields",
    "filters": "PageFields",
}
pivot_area_orientations = {
    "rows": PivotFieldOrientation.xlRowField,
    "columns": PivotFieldOrientation.xlColumnField,
    "filters": PivotFieldOrientation.xlPageField,
}


directions_s2i = {
    "d": -4121,
    "down": -4121,
    "l": -4159,
    "left": -4159,
    "r": -4161,
    "right": -4161,
    "u": -4162,
    "up": -4162,
}

directions_i2s = {-4121: "down", -4159: "left", -4161: "right", -4162: "up"}

calculation_s2i = {"automatic": -4105, "manual": -4135, "semiautomatic": 2}

calculation_i2s = {v: k for k, v in calculation_s2i.items()}

shape_types_s2i = {
    "3d_model": 30,
    "auto_shape": 1,
    "callout": 2,
    "canvas": 20,
    "chart": 3,
    "comment": 4,
    "content_app": 27,
    "diagram": 21,
    "embedded_ole_object": 7,
    "form_control": 8,
    "free_form": 5,
    "graphic": 28,
    "group": 6,
    "igx_graphic": 24,
    "ink": 22,
    "ink_comment": 23,
    "line": 9,
    "linked_3d_model": 31,
    "linked_graphic": 29,
    "linked_ole_object": 10,
    "linked_picture": 11,
    "media": 16,
    "ole_control_object": 12,
    "picture": 13,
    "placeholder": 14,
    "script_anchor": 18,
    "shape_type_mixed": -2,
    "slicer": 25,
    "table": 19,
    "text_box": 17,
    "text_effect": 15,
    "web_video": 26,
}

scaling = {
    "scale_from_top_left": 0,
    "scale_from_bottom_right": 2,
    "scale_from_middle": 1,
}

shape_types_i2s = {v: k for k, v in shape_types_s2i.items()}
