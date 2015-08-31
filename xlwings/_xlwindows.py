# TODO: align clean_xl_data and prepare_xl_data (should work on same dimensions of data)

import os
import sys

# Hack to find pythoncom.dll - needed for some distribution/setups
# E.g. if python is started with the full path outside of the python path, then it almost certainly fails
cwd = os.getcwd()
if not hasattr(sys, 'frozen'):
    # cx_Freeze etc. will fail here otherwise
    os.chdir(sys.exec_prefix)
import win32api
os.chdir(cwd)

import pywintypes
import pythoncom
import win32pdh
from win32com.client import GetObject, GetActiveObject, dynamic
import win32timezone
import datetime as dt
from .constants import Direction, ColorIndex
from .utils import rgb_to_int, int_to_rgb

# Optional imports
try:
    import pandas as pd
except ImportError:
    pd = None

from xlwings import PY3

# Time types: pywintypes.timetype doesn't work on Python 3
time_types = (dt.date, dt.datetime, type(pywintypes.Time(0)))


def get_number_of_instances():
    _, instances = win32pdh.EnumObjectItems(None, None, 'process', win32pdh.PERF_DETAIL_WIZARD)
    num_instances = 0
    for instance in instances:
        if instance == 'EXCEL':
            num_instances += 1
    return num_instances


def is_file_open(fullname):
    """
    Checks the Running Object Table (ROT) for the fully qualified filename
    """
    if not PY3:
        if isinstance(fullname, str):
            fullname = unicode(fullname, 'mbcs')
    context = pythoncom.CreateBindCtx()
    for moniker in pythoncom.GetRunningObjectTable():
        name = moniker.GetDisplayName(context, None)
        if name.lower() == fullname.lower():
            return True
    return False


def get_workbook(fullname, app_target=None):
    """
    Returns the COM Application and Workbook objects of an open Workbook.
    GetObject() returns the correct Excel instance if there are > 1
    """
    if app_target is not None:
        raise NotImplementedError('app_target is only available on Mac.')
    xl_workbook = GetObject(fullname)
    xl_app = xl_workbook.Application
    return xl_app, xl_workbook


def get_workbook_name(xl_workbook):
    return xl_workbook.Name


def get_worksheet_name(xl_sheet):
    return xl_sheet.Name


def get_xl_sheet(xl_workbook, sheet_name_or_index):
    return xl_workbook.Sheets(sheet_name_or_index)


def set_worksheet_name(xl_sheet, value):
    xl_sheet.Name = value


def get_worksheet_index(xl_sheet):
    return xl_sheet.Index


def get_app(xl_workbook, app_target):
    if app_target is not None:
        raise NotImplementedError('app_target is only available on Mac.')
    return xl_workbook.Application


def _get_latest_app():
    """
    Only dispatch Excel if there isn't an existing application - this allows us to run open_workbook() and
    new_workbook() in the correct Excel instance, i.e. in the one that was instantiated last. Otherwise it would pick
    the application that appears first in the Running Object Table (ROT).
    """
    try:
        return xl_workbook_current.Application
    except (NameError, AttributeError, pywintypes.com_error):
        return dynamic.Dispatch('Excel.Application')


def open_workbook(fullname, app_target):
    if app_target is not None:
        raise NotImplementedError('app_target is only available on Mac.')
    xl_app = _get_latest_app()
    xl_workbook = xl_app.Workbooks.Open(fullname)
    return xl_app, xl_workbook


def close_workbook(xl_workbook):
    xl_workbook.Close(SaveChanges=False)


def new_workbook(app_target):
    if app_target is not None:
        raise NotImplementedError('app_target is only available on Mac.')
    xl_app = _get_latest_app()
    xl_workbook = xl_app.Workbooks.Add()
    return xl_app, xl_workbook


def get_active_sheet(xl_workbook):
    return xl_workbook.ActiveSheet


def activate_sheet(xl_workbook, sheet):
    return xl_workbook.Sheets(sheet).Activate()


def get_worksheet(xl_workbook, sheet):
    return xl_workbook.Sheets(sheet)


def get_first_row(xl_sheet, range_address):
    return xl_sheet.Range(range_address).Row


def get_first_column(xl_sheet, range_address):
    return xl_sheet.Range(range_address).Column


def count_rows(xl_sheet, range_address):
    return xl_sheet.Range(range_address).Rows.Count


def count_columns(xl_sheet, range_address):
    return xl_sheet.Range(range_address).Columns.Count


def get_range_from_indices(xl_sheet, first_row, first_column, last_row, last_column):
    return xl_sheet.Range(xl_sheet.Cells(first_row, first_column),
                          xl_sheet.Cells(last_row, last_column))


def get_value_from_range(xl_range):
    return xl_range.Value


def get_value_from_index(xl_sheet, row_index, column_index):
    return xl_sheet.Cells(row_index, column_index).Value


def set_value(xl_range, data):
    xl_range.Value = data


def clean_xl_data(data):
    """
    Brings data from tuples of tuples into list of list and
    transforms pywintypes Time objects into Python datetime objects.

    Parameters
    ----------
    data : tuple of tuple
        raw data as returned from Excel through pywin32

    Returns
    -------
    list of list with native Python datetime objects

    """
    # Turn into list of list (e.g. makes it easier to create Pandas DataFrame) and handle dates
    data = [[_com_time_to_datetime(c) if isinstance(c, time_types) else c for c in row] for row in data]
    return data


def prepare_xl_data(data):
    if isinstance(data, time_types):
        return _datetime_to_com_time(data)
    else:
        return data

def _com_time_to_datetime(com_time):
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
        return dt.datetime(month=com_time.month, day=com_time.day, year=com_time.year,
                           hour=com_time.hour, minute=com_time.minute, second=com_time.second,
                           microsecond=com_time.microsecond, tzinfo=None)
    else:
        assert com_time.msec == 0, "fractional seconds not yet handled"
        return dt.datetime(month=com_time.month, day=com_time.day, year=com_time.year,
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
    if type(dt_time) is dt.date:
        dt_time = dt.datetime(dt_time.year, dt_time.month, dt_time.day,
                              tzinfo=win32timezone.TimeZoneInfo.utc())

    if PY3:
        # The py3 version of pywintypes has its time type inherit from datetime.
        # For some reason, though it accepts plain datetimes, they must have a timezone set.
        # See http://docs.activestate.com/activepython/2.7/pywin32/html/win32/help/py3k.html
        # We replace no timezone -> UTC to allow round-trips in the naive case
        if dt_time.tzinfo is None:
            if hasattr(pd, 'tslib') and isinstance(dt_time, pd.tslib.Timestamp):
                # Otherwise pandas prints ignored exceptions on Python 3
                dt_time = dt_time.to_datetime()
            # We don't use pytz.utc to get rid of additional dependency
            dt_time = dt_time.replace(tzinfo=win32timezone.TimeZoneInfo.utc())

        return dt_time
    else:
        assert dt_time.microsecond == 0, "fractional seconds not yet handled"
        return pywintypes.Time(dt_time.timetuple())


def get_selection_address(xl_app):
    return str(xl_app.Selection.Address)


def clear_contents_worksheet(xl_workbook, sheet_name_or_index):
    xl_workbook.Sheets(sheet_name_or_index).Cells.ClearContents()


def clear_worksheet(xl_workbook, sheet_name_or_index):
    xl_workbook.Sheets(sheet_name_or_index).Cells.Clear()


def clear_contents_range(xl_range):
    xl_range.ClearContents()


def clear_range(xl_range):
    xl_range.Clear()


def get_formula(xl_range):
    return xl_range.Formula


def set_formula(xl_range, value):
    xl_range.Formula = value


def get_row_index_end_down(xl_sheet, row_index, column_index):
    return xl_sheet.Cells(row_index, column_index).End(Direction.xlDown).Row


def get_column_index_end_right(xl_sheet, row_index, column_index):
    return xl_sheet.Cells(row_index, column_index).End(Direction.xlToRight).Column


def get_current_region_address(xl_sheet, row_index, column_index):
    return str(xl_sheet.Cells(row_index, column_index).CurrentRegion.Address)


def get_chart_object(xl_workbook, sheet_name_or_index, chart_name_or_index):
    return xl_workbook.Sheets(sheet_name_or_index).ChartObjects(chart_name_or_index)


def get_chart_index(xl_chart):
    return xl_chart.Index


def get_chart_name(xl_chart):
    return xl_chart.Name


def add_chart(xl_workbook, sheet_name_or_index, left, top, width, height):
    return xl_workbook.Sheets(sheet_name_or_index).ChartObjects().Add(left, top, width, height)


def set_chart_name(xl_chart, name):
    xl_chart.Name = name


def set_source_data_chart(xl_chart, xl_range):
    xl_chart.Chart.SetSourceData(xl_range)


def get_chart_type(xl_chart):
    return xl_chart.Chart.ChartType


def set_chart_type(xl_chart, chart_type):
    xl_chart.Chart.ChartType = chart_type


def activate_chart(xl_chart):
    xl_chart.Activate()


def get_column_width(xl_range):
    return xl_range.ColumnWidth


def set_column_width(xl_range, value):
    xl_range.ColumnWidth = value


def get_row_height(xl_range):
    return xl_range.RowHeight


def set_row_height(xl_range, value):
    xl_range.RowHeight = value


def get_width(xl_range):
    return xl_range.Width


def get_height(xl_range):
    return xl_range.Height


def autofit(range_, axis):
    if axis == 'rows' or axis == 'r':
        range_.xl_range.Rows.AutoFit()
    elif axis == 'columns' or axis == 'c':
        range_.xl_range.Columns.AutoFit()
    elif axis is None:
        range_.xl_range.Columns.AutoFit()
        range_.xl_range.Rows.AutoFit()


def autofit_sheet(sheet, axis):
    if axis == 'rows' or axis == 'r':
        sheet.xl_sheet.Rows.AutoFit()
    elif axis == 'columns' or axis == 'c':
        sheet.xl_sheet.Columns.AutoFit()
    elif axis is None:
        sheet.xl_sheet.Rows.AutoFit()
        sheet.xl_sheet.Columns.AutoFit()


def set_xl_workbook_current(xl_workbook):
    global xl_workbook_current
    xl_workbook_current = xl_workbook


def get_xl_workbook_current():
    try:
        return xl_workbook_current
    except NameError:
        return None


def get_number_format(range_):
    return range_.xl_range.NumberFormat


def set_number_format(range_, value):
    range_.xl_range.NumberFormat = value


def get_address(xl_range, row_absolute, col_absolute, external):
    return xl_range.GetAddress(row_absolute, col_absolute, 1, external)


def add_sheet(xl_workbook, before, after):
    if before:
        return xl_workbook.Worksheets.Add(Before=before.xl_sheet)
    else:
        # Hack, since "After" is broken in certain environments
        # see: http://code.activestate.com/lists/python-win32/11554/
        count = xl_workbook.Worksheets.Count
        new_sheet_index = after.xl_sheet.Index + 1
        if new_sheet_index > count:
            xl_sheet = xl_workbook.Worksheets.Add(Before=xl_workbook.Sheets(after.xl_sheet.Index))
            xl_workbook.Worksheets(xl_workbook.Worksheets.Count
                                   ).Move(Before=xl_workbook.Sheets(xl_workbook.Worksheets.Count - 1))
            xl_workbook.Worksheets(xl_workbook.Worksheets.Count).Activate()
        else:
            xl_sheet = xl_workbook.Worksheets.Add(Before=xl_workbook.Sheets(after.xl_sheet.Index + 1))
        return xl_sheet


def count_worksheets(xl_workbook):
    return xl_workbook.Worksheets.Count


def get_hyperlink_address(xl_range):
    try:
        return xl_range.Hyperlinks(1).Address
    except pywintypes.com_error:
        raise Exception("The cell doesn't seem to contain a hyperlink!")



def set_hyperlink(xl_range, address, text_to_display=None, screen_tip=None):
    # Another one of these pywin32 bugs that only materialize under certain circumstances:
    # http://stackoverflow.com/questions/6284227/hyperlink-will-not-show-display-proper-text
    link = xl_range.Hyperlinks.Add(Anchor=xl_range, Address=address)
    link.TextToDisplay = text_to_display
    link.ScreenTip = screen_tip


def set_color(xl_range, color_or_rgb):
    if color_or_rgb is None:
        xl_range.Interior.ColorIndex = ColorIndex.xlColorIndexNone
    elif isinstance(color_or_rgb, int):
        xl_range.Interior.Color = color_or_rgb
    else:
        xl_range.Interior.Color = rgb_to_int(color_or_rgb)


def get_color(xl_range):
    if xl_range.Interior.ColorIndex == ColorIndex.xlColorIndexNone:
        return None
    else:
        return int_to_rgb(xl_range.Interior.Color)


def get_xl_workbook_from_xl(fullname, app_target=None, hwnd=None):
    """
    Use GetActiveObject whenever possible, GetObject with a file path will only work if the file
    has been registered in the RunningObjectTable (ROT).
    Sometimes, e.g. if the files opens from an untrusted location, it doesn't appear in the ROT.
    app_target is only used on Mac.
    """
    num_of_instances = get_number_of_instances()

    if num_of_instances < 2:
        xl_app = GetActiveObject('Excel.Application')
        xl_workbook = xl_app.ActiveWorkbook
    else:
        if not is_file_open(fullname):
            # This means that the file doesn't appear in the ROT. If it's in the first instance of
            # Excel, we can still get it with GetActiveObject
            xl_app = GetActiveObject('Excel.Application')
            xl_workbook = xl_app.ActiveWorkbook
        else:
            xl_workbook = GetObject(fullname)
            xl_app = xl_workbook.Application
    if str(xl_app.hwnd) != hwnd:
        # The check of the window handle also works when the same file is opened
        # in two instances, whereas the comparison of fullpath would fail
        raise Exception("Can't establish connection! "
                        "Try to open the file in the first instance of Excel or "
                        "change your trusted location/document settings or "
                        "set OPTIMIZED_CONNECTION = True.")
    return xl_workbook


def save_workbook(xl_workbook, path):
    saved_path = xl_workbook.Path
    if (saved_path != '') and (path is None):
        # Previously saved: Save under existing name
        xl_workbook.Save()
    elif (saved_path == '') and (path is None):
        # Previously unsaved: Save under current name in current working directory
        path = os.path.join(os.getcwd(), xl_workbook.Name + '.xlsx')
        xl_workbook.Application.DisplayAlerts = False
        xl_workbook.SaveAs(path)
        xl_workbook.Application.DisplayAlerts = True
    elif path:
        # Save under new name/location
        xl_workbook.Application.DisplayAlerts = False
        xl_workbook.SaveAs(path)
        xl_workbook.Application.DisplayAlerts = True


def open_template(fullpath):
    os.startfile(fullpath)


def set_visible(xl_app, visible):
    xl_app.Visible = visible


def get_visible(xl_app):
    return xl_app.Visible


def get_fullname(xl_workbook):
    return xl_workbook.FullName


def quit_app(xl_app):
    xl_app.DisplayAlerts = False
    xl_app.Quit()
    xl_app.DisplayAlerts = True


def get_screen_updating(xl_app):
    return xl_app.ScreenUpdating


def set_screen_updating(xl_app, value):
    xl_app.ScreenUpdating = value


def get_calculation(xl_app):
    return xl_app.Calculation


def set_calculation(xl_app, value):
    xl_app.Calculation = value


def calculate(xl_app):
    xl_app.Calculate()


def get_named_range(range_):
    return range_.xl_range.Name.Name


def set_named_range(range_, value):
    range_.xl_range.Name = value
