import os
import sys

# Hack to find pythoncom.dll - needed for some distribution/setups (includes seemingly unused import win32api)
# E.g. if python is started with the full path outside of the python path, then it almost certainly fails
cwd = os.getcwd()
if not hasattr(sys, 'frozen'):
    # cx_Freeze etc. will fail here otherwise
    os.chdir(sys.exec_prefix)
import win32api
os.chdir(cwd)

from warnings import warn
import datetime as dt
import numbers
import types
from ctypes import oledll, PyDLL, py_object, byref, POINTER, windll

import pywintypes
import pythoncom
from win32com.client import Dispatch, CDispatch, DispatchEx
import win32timezone
import win32gui
import win32process
from comtypes import IUnknown
from comtypes.automation import IDispatch

from .constants import ColorIndex
from .utils import rgb_to_int, int_to_rgb, get_duplicates, np_datetime_to_datetime, col_name

# Optional imports
try:
    import pandas as pd
except ImportError:
    pd = None
try:
    import numpy as np
except ImportError:
    np = None

from . import PY3

time_types = (dt.date, dt.datetime, pywintypes.TimeType)
if np:
    time_types = time_types + (np.datetime64,)


N_COM_ATTEMPTS = 0      # 0 means try indefinitely
BOOK_CALLER = None

missing = object()


class COMRetryMethodWrapper(object):

    def __init__(self, method):
        self.__method = method

    def __call__(self, *args, **kwargs):
        n_attempt = 1
        while True:
            try:
                v = self.__method(*args, **kwargs)
                t = type(v)
                if t is CDispatch:
                    return COMRetryObjectWrapper(v)
                elif t is types.MethodType:
                    return COMRetryMethodWrapper(v)
                else:
                    return v
            except pywintypes.com_error as e:
                if (not N_COM_ATTEMPTS or n_attempt < N_COM_ATTEMPTS) and e.hresult == -2147418111:
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


class COMRetryObjectWrapper(object):
    def __init__(self, inner):
        object.__setattr__(self, '_inner', inner)

    def __setattr__(self, key, value):
        n_attempt = 1
        while True:
            try:
                return setattr(self._inner, key, value)
            except pywintypes.com_error as e:
                if (not N_COM_ATTEMPTS or n_attempt < N_COM_ATTEMPTS) and e.hresult == -2147418111:
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
                t = type(v)
                if t is CDispatch:
                    return COMRetryObjectWrapper(v)
                elif t is types.MethodType:
                    return COMRetryMethodWrapper(v)
                else:
                    return v
            except pywintypes.com_error as e:
                if (not N_COM_ATTEMPTS or n_attempt < N_COM_ATTEMPTS) and e.hresult == -2147418111:
                    n_attempt += 1
                    continue
                else:
                    raise
            except AttributeError as e:
                # Pywin32 reacts incorrectly to RPC_E_CALL_REJECTED (i.e. assumes attribute doesn't
                # exist, thus not allowing to destinguish between cases where attribute really doesn't
                # exist or error is only being thrown because the COM RPC server is busy). Here
                # we try to test to see what's going on really
                try:
                    self._oleobj_.GetIDsOfNames(0, item)
                except pythoncom.ole_error as e:
                    if e.hresult != -2147418111:   # RPC_E_CALL_REJECTED
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
                t = type(v)
                if t is CDispatch:
                    return COMRetryObjectWrapper(v)
                elif t is types.MethodType:
                    return COMRetryMethodWrapper(v)
                else:
                    return v
            except pywintypes.com_error as e:
                    if (not N_COM_ATTEMPTS or n_attempt < N_COM_ATTEMPTS) and e.hresult == -2147418111:
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
            t = type(v)
            if t is CDispatch:
                yield COMRetryObjectWrapper(v)
            else:
                yield v


# Constants
OBJID_NATIVEOM = -16


def accessible_object_from_window(hwnd):
    ptr = POINTER(IDispatch)()
    res = oledll.oleacc.AccessibleObjectFromWindow(
        hwnd, OBJID_NATIVEOM,
        byref(IDispatch._iid_), byref(ptr))
    return ptr


def comtypes_to_pywin(ptr, interface=None):
    _PyCom_PyObjectFromIUnknown = PyDLL(pythoncom.__file__).PyCom_PyObjectFromIUnknown
    _PyCom_PyObjectFromIUnknown.restype = py_object

    if interface is None:
        interface = IUnknown
    return _PyCom_PyObjectFromIUnknown(ptr, byref(interface._iid_), True)


def is_hwnd_xl_app(hwnd):
    try:
        child_hwnd = win32gui.FindWindowEx(hwnd, 0, 'XLDESK', None)
        child_hwnd = win32gui.FindWindowEx(child_hwnd, 0, 'EXCEL7', None)
        ptr = accessible_object_from_window(child_hwnd)
        return True
    except WindowsError:
        return False
    except pywintypes.error:
        return False


def get_xl_app_from_hwnd(hwnd):
    child_hwnd = win32gui.FindWindowEx(hwnd, 0, 'XLDESK', None)
    child_hwnd = win32gui.FindWindowEx(child_hwnd, 0, 'EXCEL7', None)

    ptr = accessible_object_from_window(child_hwnd)
    p = comtypes_to_pywin(ptr, interface=IDispatch)
    disp = COMRetryObjectWrapper(Dispatch(p))
    return disp.Application


def get_excel_hwnds():
    #win32gui.EnumWindows(lambda hwnd, result_list: result_list.append(hwnd), hwnds)
    hwnd = windll.user32.GetTopWindow(None)
    pids = set()
    while hwnd:
        try:
            # Apparently, this fails on some systems when Excel is closed
            child_hwnd = win32gui.FindWindowEx(hwnd, 0, 'XLDESK', None)
            if child_hwnd:
                child_hwnd = win32gui.FindWindowEx(child_hwnd, 0, 'EXCEL7', None)
            if child_hwnd:
                pid = win32process.GetWindowThreadProcessId(hwnd)[1]
                if pid not in pids:
                    pids.add(pid)
                    yield hwnd
            #if win32gui.FindWindowEx(hwnd, 0, 'XLDESK', None):
            #    yield hwnd
        except pywintypes.error:
            pass

        hwnd = windll.user32.GetWindow(hwnd, 2)   # 2 = next window according to Z-order


def get_xl_apps():
    for hwnd in get_excel_hwnds():
        try:
            yield get_xl_app_from_hwnd(hwnd)
        except ExcelBusyError:
            pass
        except WindowsError:
            # This happens if the bare Excel Application is open without Workbook
            # i.e. there is no 'EXCEL7' child hwnd that would be necessary to make a connection
            pass


def is_range_instance(xl_range):
    pyid = getattr(xl_range, '_oleobj_', None)
    if pyid is None:
        return False
    return xl_range._oleobj_.GetTypeInfo().GetTypeAttr().iid == pywintypes.IID('{00020846-0000-0000-C000-000000000046}')
    # return pyid.GetTypeInfo().GetDocumentation(-1)[0] == 'Range'


class Apps(object):

    def __iter__(self):
        for hwnd in get_excel_hwnds():
            yield App(xl=hwnd)

    def __len__(self):
        return len(list(get_excel_hwnds()))

    def __getitem__(self, index):
        hwnds = list(get_excel_hwnds())
        return App(xl=hwnds[index])


class App(object):

    def __init__(self, spec=None, add_book=True, xl=None):
        if spec is not None:
            warn('spec is ignored on Windows.')
        if xl is None:
            # new instance
            self._xl = COMRetryObjectWrapper(DispatchEx('Excel.Application'))
            if add_book:
                self._xl.Workbooks.Add()
            self._hwnd = None
        elif isinstance(xl, int):
            self._xl = None
            self._hwnd = xl
        else:
            self._xl = xl
            self._hwnd = None

    @property
    def xl(self):
        if self._xl is None:
            self._xl = get_xl_app_from_hwnd(self._hwnd)
        return self._xl

    api = xl

    @property
    def selection(self):
        # TODO: selection isn't always a range
        return Range(xl=self.xl.Selection)

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

    def kill(self):
        import win32api
        PROCESS_TERMINATE = 1
        handle = win32api.OpenProcess(PROCESS_TERMINATE, False, self.pid)
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


class Books(object):

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

    def open(self, fullname):
        return Book(xl=self.xl.Open(fullname))

    def __iter__(self):
        for xl in self.xl:
            yield Book(xl=xl)


class Book(object):

    def __init__(self, xl):
        self.xl = xl

    @property
    def api(self):
        return self.xl

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

    def save(self, path=None):
        saved_path = self.xl.Path
        if (saved_path != '') and (path is None):
            # Previously saved: Save under existing name
            self.xl.Save()
        elif (saved_path != '') and (path is not None) and (os.path.split(path)[0] == ''):
            # Save existing book under new name in cwd if no path has been provided
            path = os.path.join(os.getcwd(), path)
            self.xl.SaveAs(os.path.realpath(path))
        elif (saved_path == '') and (path is None):
            # Previously unsaved: Save under current name in current working directory
            path = os.path.join(os.getcwd(), self.xl.Name + '.xlsx')
            alerts_state = self.xl.Application.DisplayAlerts
            self.xl.Application.DisplayAlerts = False
            self.xl.SaveAs(os.path.realpath(path))
            self.xl.Application.DisplayAlerts = alerts_state
        elif path:
            # Save under new name/location
            alerts_state = self.xl.Application.DisplayAlerts
            self.xl.Application.DisplayAlerts = False
            self.xl.SaveAs(os.path.realpath(path))
            self.xl.Application.DisplayAlerts = alerts_state

    @property
    def fullname(self):
        return self.xl.FullName

    @property
    def names(self):
        return Names(xl=self.xl.Names)

    def activate(self):
        self.xl.Activate()


class Sheets(object):
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


class Sheet(object):

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
                raise IndexError("Attempted to access 0-based Range. xlwings/Excel Ranges are 1-based.")
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
                raise IndexError("Attempted to access 0-based Range. xlwings/Excel Ranges are 1-based.")
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

    def clear(self):
        self.xl.Cells.Clear()

    def autofit(self, axis=None):
        if axis == 'rows' or axis == 'r':
            self.xl.Rows.AutoFit()
        elif axis == 'columns' or axis == 'c':
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

    @property
    def charts(self):
        return Charts(xl=self.xl.ChartObjects())

    @property
    def shapes(self):
        return Shapes(xl=self.xl.Shapes)

    @property
    def pictures(self):
        return Pictures(xl=self.xl.Pictures())


class Range(object):

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
                self._xl = xl_sheet.Range(xl_sheet.Cells(row, col), xl_sheet.Cells(row+nrows-1, col+ncols-1))
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
                self.xl.Columns.Count
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
            return ''

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
            if axis == 'rows' or axis == 'r':
                self.xl.Rows.AutoFit()
            elif axis == 'columns' or axis == 'c':
                self.xl.Columns.AutoFit()
            elif axis is None:
                self.xl.Columns.AutoFit()
                self.xl.Rows.AutoFit()

    @property
    def hyperlink(self):
        if self.xl is not None:
            try:
                return self.xl.Hyperlinks(1).Address
            except pywintypes.com_error:
                raise Exception("The cell doesn't seem to contain a hyperlink!")
        else:
            return ''

    def add_hyperlink(self, address, text_to_display, screen_tip):
        if self.xl is not None:
            # Another one of these pywin32 bugs that only materialize under certain circumstances:
            # http://stackoverflow.com/questions/6284227/hyperlink-will-not-show-display-proper-text
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


def clean_value_data(data, datetime_builder, empty_as, number_builder):
    if number_builder is not None:
        return [
            [
                _com_time_to_datetime(c, datetime_builder)
                if isinstance(c, time_types) else
                number_builder(c)
                if type(c) == float else
                empty_as
                if c is None else
                c
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
                else c
                for c in row
            ]
            for row in data
        ]


def _com_time_to_datetime(com_time, datetime_builder):
    """
    This function is a modified version from Pyvot (https://pypi.python.org/pypi/Pyvot)
    and subject to the following copyright:

    Copyright (c) Microsoft Corporation.

    This source code is subject to terms and conditions of the Apache License, Version 2.0. A
    copy of the license can be found in the LICENSE.txt file at the root of this distribution. If
    you cannot locate the Apache License, Version 2.0, please send an email to
    vspython@microsoft.com. By using this source code in any fashion, you are agreeing to be bound
    by the terms of the Apache License, Version 2.0.

    You must not remove this notice, or any other, from this software.

    """

    if PY3:
        # The py3 version of pywintypes has its time type inherit from datetime.
        # We copy to a new datetime so that the returned type is the same between 2/3
        # Changed: We make the datetime object timezone naive as Excel doesn't provide info
        return datetime_builder(month=com_time.month, day=com_time.day, year=com_time.year,
                                hour=com_time.hour, minute=com_time.minute, second=com_time.second,
                                microsecond=com_time.microsecond, tzinfo=None)
    else:
        assert com_time.msec == 0, "fractional seconds not yet handled"
        return datetime_builder(month=com_time.month, day=com_time.day, year=com_time.year,
                                hour=com_time.hour, minute=com_time.minute, second=com_time.second)


def _datetime_to_com_time(dt_time):
    """
    This function is a modified version from Pyvot (https://pypi.python.org/pypi/Pyvot)
    and subject to the following copyright:

    Copyright (c) Microsoft Corporation.

    This source code is subject to terms and conditions of the Apache License, Version 2.0. A
    copy of the license can be found in the LICENSE.txt file at the root of this distribution. If
    you cannot locate the Apache License, Version 2.0, please send an email to
    vspython@microsoft.com. By using this source code in any fashion, you are agreeing to be bound
    by the terms of the Apache License, Version 2.0.

    You must not remove this notice, or any other, from this software.

    """
    # Convert date to datetime
    if pd and isinstance(dt_time, pd.tslib.NaTType):
        return None
    if np:
        if type(dt_time) is np.datetime64:
            dt_time = np_datetime_to_datetime(dt_time)

    if type(dt_time) is dt.date:
        dt_time = dt.datetime(dt_time.year, dt_time.month, dt_time.day,
                              tzinfo=win32timezone.TimeZoneInfo.utc())

    if PY3:
        # The py3 version of pywintypes has its time type inherit from datetime.
        # For some reason, though it accepts plain datetimes, they must have a timezone set.
        # See http://docs.activestate.com/activepython/2.7/pywin32/html/win32/help/py3k.html
        # We replace no timezone -> UTC to allow round-trips in the naive case
        if pd and isinstance(dt_time, pd.tslib.Timestamp):
            # Otherwise pandas prints ignored exceptions on Python 3
            dt_time = dt_time.to_datetime()
        # We don't use pytz.utc to get rid of additional dependency
        # Don't do any timezone transformation: simply cutoff the tz info
        # If we don't reset it first, it gets transformed into UTC before transferred to Excel
        dt_time = dt_time.replace(tzinfo=None)
        dt_time = dt_time.replace(tzinfo=win32timezone.TimeZoneInfo.utc())

        return dt_time
    else:
        assert dt_time.microsecond == 0, "fractional seconds not yet handled"
        return pywintypes.Time(dt_time.timetuple())


def prepare_xl_data_element(x):
    if isinstance(x, time_types):
        return _datetime_to_com_time(x)
    elif np and isinstance(x, np.generic):
        return float(x)
    elif x is None:
        return ""
    elif np and isinstance(x, float) and np.isnan(x):
        return ""
    else:
        return x


# TODO: move somewhere better, same on mac
def open_template(fullpath):
    os.startfile(fullpath)


class Shape(object):

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


class Collection(object):

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


class Shapes(Collection):

    _wrap = Shape


class Chart(object):

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


class Charts(Collection):

    def _wrap(self, xl):
        return Chart(xl_obj=xl)

    def add(self, left, top, width, height):
        return Chart(xl_obj=self.xl.Add(
            left, top, width, height
        ))


class Picture(object):

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


class Pictures(Collection):

    _wrap = Picture

    @property
    def parent(self):
        return Sheet(xl=self.xl.Parent)

    def add(self, filename, link_to_file, save_with_document, left, top, width, height):
        return Picture(xl=self.xl.Parent.Shapes.AddPicture(
            Filename=filename,
            LinkToFile=link_to_file,
            SaveWithDocument=save_with_document,
            Left=left,
            Top=top,
            Width=width,
            Height=height
        ).DrawingObject)


class Names(object):
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


class Name(object):
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

chart_types_s2i = {
    '3d_area': -4098,
    '3d_area_stacked': 78,
    '3d_area_stacked_100': 79,
    '3d_bar_clustered': 60,
    '3d_bar_stacked': 61,
    '3d_bar_stacked_100': 62,
    '3d_column': -4100,
    '3d_column_clustered': 54,
    '3d_column_stacked': 55,
    '3d_column_stacked_100': 56,
    '3d_line': -4101,
    '3d_pie': -4102,
    '3d_pie_exploded': 70,
    'area': 1,
    'area_stacked': 76,
    'area_stacked_100': 77,
    'bar_clustered': 57,
    'bar_of_pie': 71,
    'bar_stacked': 58,
    'bar_stacked_100': 59,
    'bubble': 15,
    'bubble_3d_effect': 87,
    'column_clustered': 51,
    'column_stacked': 52,
    'column_stacked_100': 53,
    'cone_bar_clustered': 102,
    'cone_bar_stacked': 103,
    'cone_bar_stacked_100': 104,
    'cone_col': 105,
    'cone_col_clustered': 99,
    'cone_col_stacked': 100,
    'cone_col_stacked_100': 101,
    'cylinder_bar_clustered': 95,
    'cylinder_bar_stacked': 96,
    'cylinder_bar_stacked_100': 97,
    'cylinder_col': 98,
    'cylinder_col_clustered': 92,
    'cylinder_col_stacked': 93,
    'cylinder_col_stacked_100': 94,
    'doughnut': -4120,
    'doughnut_exploded': 80,
    'line': 4,
    'line_markers': 65,
    'line_markers_stacked': 66,
    'line_markers_stacked_100': 67,
    'line_stacked': 63,
    'line_stacked_100': 64,
    'pie': 5,
    'pie_exploded': 69,
    'pie_of_pie': 68,
    'pyramid_bar_clustered': 109,
    'pyramid_bar_stacked': 110,
    'pyramid_bar_stacked_100': 111,
    'pyramid_col': 112,
    'pyramid_col_clustered': 106,
    'pyramid_col_stacked': 107,
    'pyramid_col_stacked_100': 108,
    'radar': -4151,
    'radar_filled': 82,
    'radar_markers': 81,
    'stock_hlc': 88,
    'stock_ohlc': 89,
    'stock_vhlc': 90,
    'stock_vohlc': 91,
    'surface': 83,
    'surface_top_view': 85,
    'surface_top_view_wireframe': 86,
    'surface_wireframe': 84,
    'xy_scatter': -4169,
    'xy_scatter_lines': 74,
    'xy_scatter_lines_no_markers': 75,
    'xy_scatter_smooth': 72,
    'xy_scatter_smooth_no_markers': 73
}

chart_types_i2s = {v: k for k, v in chart_types_s2i.items()}

directions_s2i = {
    'd': -4121,
    'down': -4121,
    'l': -4159,
    'left': -4159,
    'r': -4161,
    'right': -4161,
    'u': -4162,
    'up': -4162
}

directions_i2s = {
    -4121: 'down',
    -4159: 'left',
    -4161: 'right',
    -4162: 'up'
}

calculation_s2i = {
    "automatic": -4105,
    "manual": -4135,
    "semiautomatic": 2
}

calculation_i2s = {v: k for k, v in calculation_s2i.items()}

shape_types_s2i = {
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
    "group": 6,
    "igx_graphic": 24,
    "ink": 22,
    "ink_comment": 23,
    "line": 9,
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
    "web_video": 26
}

shape_types_i2s = {v: k for k, v in shape_types_s2i.items()}
