import os
import sys

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

os.chdir(cwd)

from warnings import warn
import datetime as dt
import numbers
import types
import ctypes
from ctypes import oledll, PyDLL, py_object, byref, windll

import pythoncom
from win32com.client import (
    Dispatch,
    CoClassBaseClass,
    CDispatch,
    DispatchEx,
    DispatchBaseClass,
)
import win32timezone
import win32gui
import win32process

from .constants import (
    ColorIndex,
    UpdateLinks,
    InsertShiftDirection,
    InsertFormatOrigin,
    DeleteShiftDirection,
    ListObjectSourceType,
    FixedFormatType,
    FileFormat,
)
from .utils import (
    rgb_to_int,
    int_to_rgb,
    np_datetime_to_datetime,
    col_name,
    fullname_url_to_local_path,
    read_config_sheet,
    hex_to_rgb,
)
import xlwings

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
                elif type(v) is types.MethodType:
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
            except AttributeError as e:
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
                    wcode, source, text, help_file, help_id, scode = (
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
            except AttributeError as e:
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
                elif type(v) is types.MethodType:
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
            except AttributeError as e:
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
                elif type(v) is types.MethodType:
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
            except AttributeError as e:
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
    res = oledll.oleacc.AccessibleObjectFromWindow(
        hwnd, OBJID_NATIVEOM, byref(_IDISPATCH_GUID), byref(ptr)
    )
    return ptr


def is_hwnd_xl_app(hwnd):
    try:
        child_hwnd = win32gui.FindWindowEx(hwnd, 0, "XLDESK", None)
        child_hwnd = win32gui.FindWindowEx(child_hwnd, 0, "EXCEL7", None)
        ptr = accessible_object_from_window(child_hwnd)
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


class Engine:
    @property
    def apps(self):
        return Apps()

    @property
    def name(self):
        return "excel"

    @staticmethod
    def prepare_xl_data_element(x):
        if isinstance(x, time_types):
            return _datetime_to_com_time(x)
        elif np and isinstance(x, (np.floating, float)) and np.isnan(x):
            return ""
        elif np and isinstance(x, np.number):
            return float(x)
        elif x is None:
            return ""
        else:
            return x

    @staticmethod
    def clean_value_data(data, datetime_builder, empty_as, number_builder):
        if number_builder is not None:
            return [
                [
                    _com_time_to_datetime(c, datetime_builder)
                    if isinstance(c, time_types)
                    else number_builder(c)
                    if type(c) == float
                    else empty_as
                    # #DIV/0!, #N/A, #NAME?, #NULL!, #NUM!, #REF!, #VALUE!
                    if c is None
                    or (
                        isinstance(c, int)
                        and c
                        in [
                            -2146826281,
                            -2146826246,
                            -2146826259,
                            -2146826288,
                            -2146826252,
                            -2146826265,
                            -2146826273,
                        ]
                    )
                    else c
                    for c in row
                ]
                for row in data
            ]
        else:
            return [
                [
                    _com_time_to_datetime(c, datetime_builder)
                    if isinstance(c, time_types)
                    else empty_as
                    if c is None
                    or (
                        isinstance(c, int)
                        and c
                        in [
                            -2146826281,
                            -2146826246,
                            -2146826259,
                            -2146826288,
                            -2146826252,
                            -2146826265,
                            -2146826273,
                        ]
                    )
                    else c
                    for c in row
                ]
                for row in data
            ]


engine = Engine()


class Apps:
    def keys(self):
        k = []
        for hwnd in get_excel_hwnds():
            k.append(App(xl=hwnd).pid)
        return k

    def add(self, spec=None, add_book=None, xl=None, visible=None):
        return App(spec=spec, add_book=add_book, xl=xl, visible=visible)

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


class App:
    def __init__(self, spec=None, add_book=True, xl=None, visible=None):
        # visible is only required on mac
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

    def kill(self):
        import win32api

        PROCESS_TERMINATE = 1
        handle = win32api.OpenProcess(PROCESS_TERMINATE, False, self._pid)
        win32api.TerminateProcess(handle, -1)
        win32api.CloseHandle(handle)

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
        return Books(xl=self.xl.Workbooks)

    @property
    def hwnd(self):
        if self._hwnd is None:
            self._hwnd = self._xl.Hwnd
        return self._hwnd

    @property
    def pid(self):
        return win32process.GetWindowThreadProcessId(self.hwnd)[1]

    def range(self, arg1, arg2=None):
        if isinstance(arg1, Range):
            xl1 = arg1.xl
        else:
            xl1 = self.xl.Range(arg1)

        if arg2 is None:
            return Range(xl=xl1)

        if isinstance(arg2, Range):
            xl2 = arg2.xl
        else:
            xl2 = self.xl.Range(arg2)

        return Range(xl=self.xl.Range(xl1, xl2))

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


class Books:
    def __init__(self, xl):
        self.xl = xl

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


class Book:
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


class Sheets:
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

    def add(self, before=None, after=None):
        if before:
            return Sheet(xl=self.xl.Add(Before=before.xl))
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
            return Sheet(xl=xl_sheet)
        else:
            return Sheet(xl=self.xl.Add())


class Sheet:
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

    @property
    def page_setup(self):
        return PageSetup(self.xl.PageSetup)


class Range:
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
            raise NotImplemented()

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
            raise NotImplemented()

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
    def note(self):
        return Note(xl=self.xl.Comment) if self.xl.Comment else None

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


class Shape:
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


class Font:
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


class Characters:
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


class Collection:
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


class PageSetup:
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


class Note:
    def __init__(self, xl):
        self.xl = xl

    @property
    def api(self):
        return self.xl

    @property
    def text(self):
        return self.xl.Text()

    @text.setter
    def text(self, value):
        self.xl.Text(value)

    def delete(self):
        self.xl.Delete()


class Shapes(Collection):

    _wrap = Shape


class Table:
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
        self.xl.Resize(range)


class Tables(Collection):

    _wrap = Table

    def add(
        self,
        source_type=None,
        source=None,
        link_source=None,
        has_headers=None,
        destination=None,
        table_style_name=None,
    ):
        return Table(
            xl=self.xl.Add(
                SourceType=ListObjectSourceType.xlSrcRange,
                Source=source.api,
                LinkSource=link_source,
                XlListObjectHasHeaders=True,
                Destination=destination,
                TableStyleName=table_style_name,
            )
        )


class Chart:
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

    def set_source_data(self, rng):
        self.xl.SetSourceData(rng.xl)

    @property
    def chart_type(self):
        return chart_types_i2s[self.xl.ChartType]

    @chart_type.setter
    def chart_type(self, chart_type):
        self.xl.ChartType = chart_types_s2i[chart_type]

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
        # todo: what about chart sheets?
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
        except:
            pass


class Charts(Collection):
    def _wrap(self, xl):
        return Chart(xl_obj=xl)

    def add(self, left, top, width, height):
        return Chart(xl_obj=self.xl.Add(left, top, width, height))


class Picture:
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


class Pictures(Collection):

    _wrap = Picture

    @property
    def parent(self):
        return Sheet(xl=self.xl.Parent)

    def add(self, filename, link_to_file, save_with_document, left, top, width, height):
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


class Names:
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


class Name:
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
