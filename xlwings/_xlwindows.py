# TODO: create classes
# TODO: align clean_xl_data and prepare_xl_data (should work on same dimensions of data)

import datetime as dt
import win32api  # needed as first import to find all dlls
import pywintypes
import pythoncom
from win32com.client import GetObject, dynamic
import win32timezone
from .constants import Direction

# Optional imports
try:
    import pandas as pd
except ImportError:
    pd = None

from xlwings import PY3

# Time types: pywintypes.timetype doesn't work on Python 3
time_types = (dt.date, dt.datetime, type(pywintypes.Time(0)))


def is_file_open(fullname):
    """
    Checks the Running Object Table (ROT) for the fully qualified filename
    """
    context = pythoncom.CreateBindCtx()
    for moniker in pythoncom.GetRunningObjectTable():
        name = moniker.GetDisplayName(context, None)
        if name.lower() == fullname.lower():
            return True
    return False


def get_workbook(fullname):
    """
    Returns the COM Application and Workbook objects of an open Workbook.
    GetObject() returns the correct Excel instance if there are > 1
    """
    xl_workbook = GetObject(fullname)
    xl_app = xl_workbook.Application
    return xl_app, xl_workbook


def get_workbook_name(xl_workbook):
    return xl_workbook.Name


def get_worksheet_name(xl_sheet):
    return xl_sheet.Name


def get_worksheet_index(xl_sheet):
    return xl_sheet.Index


def open_workbook(fullname):
    xl_app = dynamic.Dispatch('Excel.Application')
    xl_workbook = xl_app.Workbooks.Open(fullname)
    xl_app.Visible = True
    return xl_app, xl_workbook


def close_workbook(xl_workbook):
    xl_workbook.Close(SaveChanges=False)


def new_workbook():
    xl_app = dynamic.Dispatch('Excel.Application')
    xl_app.Visible = True
    xl_workbook = xl_app.Workbooks.Add()
    return xl_app, xl_workbook

def new_worksheet(xl_workbook,worksheet_name):
    sheet_count = xl_workbook.Worksheets.Count    
    if worksheet_name == '':        
        xl_sheet = xl_workbook.Worksheets.Add (After=xl_workbook.Worksheets(sheet_count))
    else:
        xl_sheet = xl_workbook.Worksheets.Add (After=xl_workbook.Worksheets(sheet_count))
        xl_sheet.Name = worksheet_name
    return xl_sheet

def delete_worksheet(xl_workbook, sheetname):
    xl_workbook.Application.DisplayAlerts = False    
    xl_workbook.Sheets(sheetname).Delete()

def fit_columns(xl_workbook, sheetname):
    if sheetname == None:
        xl_workbook.ActiveSheet.Columns.AutoFit()
    else:    
        xl_workbook.Sheets(sheetname).Columns.AutoFit()


def get_active_sheet(xl_workbook):
    return xl_workbook.ActiveSheet

def activate_sheet(xl_workbook, sheet):
    return xl_workbook.Sheets(sheet).Activate()

def get_worksheet(xl_workbook, sheet):
    return xl_workbook.Sheets(sheet)

def get_first_row(xl_sheet, range_address):
    return xl_sheet.Range(range_address).Row


def sheet_list(xl_workbook):
    z = []
    sheet_count = xl_workbook.Worksheets.Count
    for i in range(1,sheet_count+1):
        z.append(xl_workbook.Worksheets(i).Name)
    return z

def hiden_rows(xl_workbook,rows, status):
    xl_workbook.ActiveSheet.Rows(str(rows)).EntireRow.Hidden = status
    
def hiden_columns(xl_workbook, cols, status):
    xl_workbook.ActiveSheet.Columns(str(cols)).EntireColumn.Hidden = status
    
def is_row_hidden(xl_workbook, row):
    return xl_workbook.ActiveSheet.Rows(str(row)).EntireRow.Hidden
    
def is_col_hidden(xl_workbook, col):
    return xl_workbook.ActiveSheet.Columns(str(col)).EntireColumn.Hidden
    


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
    
def set_color(xl_range,color_name_or_number):
    all_colors = {'Light Orange': 45.0, 'Ivory': 19.0, 'Teal+': 31.0,
                  'Teal': 14.0, 'Green': 10.0, 'Pale Blue': 37.0,
                  'Ice Blue': 24.0, 'Gray-80%': 56.0, 'Aqua': 42.0,
                  'Coral': 22.0, 'Dark Blue+': 25.0, 'Light Blue': 41.0,
                  'Plum+': 18.0, 'Sea Green': 50.0, 'Ocean Blue': 23.0,
                  'Dark Red+': 30.0, 'Violet+': 29.0, 'Brown': 53.0,
                  'Dark Teal': 49.0, 'Gray-25%': 15.0, 'Sky Blue': 33.0,
                  'Pink': 7.0, 'Light Yellow': 36.0, 'Periwinkle': 17.0,
                  'Turquoise': 8.0, 'Yellow+': 27.0, 'Lite Turquoise': 20.0,
                  'Red': 3.0, 'White': 2.0, 'Dark Green': 51.0, 'Orange': 46.0,
                  'Dark Purple': 21.0, 'Dark Yellow': 12.0, 'Black': 1.0,
                  'Light Turquoise': 34.0, 'Olive Green': 52.0, 'Rose': 38.0,
                  'Blue': 5.0, 'Blue-Gray': 47.0, 'Lime': 43.0, 'Tan': 40.0,
                  'Bright Green': 4.0, 'Light Green': 35.0, 'Violet': 13.0,
                  'Blue+': 32.0, 'Dark Blue': 11.0, 'Yellow': 6.0,
                  'Gray-50%': 16.0, 'Lavender': 39.0, 'Dark Red': 9.0,
                  'Gold': 44.0, 'Plum': 54.0, 'Indigo': 55.0, 'Pink+': 26.0,
                  'Gray-40%': 48.0, 'Turquoise+': 28.0, '':-4142}
    try:
        xl_range.Interior.ColorIndex = all_colors[color_name_or_number]
    except KeyError:
        xl_range.Interior.ColorIndex = color_name_or_number

def get_color(xl_range):
    color_codes = {1.0: 'Black', 2.0: 'White', 3.0: 'Red', 4.0: 'Bright Green',
                   5.0: 'Blue', 6.0: 'Yellow', 7.0: 'Pink', 8.0: 'Turquoise',
                   9.0: 'Dark Red', 10.0: 'Green', 11.0: 'Dark Blue',
                   12.0: 'Dark Yellow', 13.0: 'Violet', 14.0: 'Teal',
                   15.0: 'Gray-25%', 16.0: 'Gray-50%', 17.0: 'Periwinkle',
                   18.0: 'Plum+', 19.0: 'Ivory', 20.0: 'Lite Turquoise',
                   21.0: 'Dark Purple', 22.0: 'Coral', 23.0: 'Ocean Blue',
                   24.0: 'Ice Blue', 25.0: 'Dark Blue+', 26.0: 'Pink+',
                   27.0: 'Yellow+', 28.0: 'Turquoise+', 29.0: 'Violet+',
                   30.0: 'Dark Red+', 31.0: 'Teal+', 32.0: 'Blue+',
                   33.0: 'Sky Blue', 34.0: 'Light Turquoise', 35.0: 'Light Green',
                   36.0: 'Light Yellow', 37.0: 'Pale Blue', 38.0: 'Rose',
                   39.0: 'Lavender', 40.0: 'Tan', 41.0: 'Light Blue',
                   42.0: 'Aqua', 43.0: 'Lime', 44.0: 'Gold', 45.0: 'Light Orange',
                   46.0: 'Orange', 47.0: 'Blue-Gray', 48.0: 'Gray-40%',
                   49.0: 'Dark Teal', 50.0: 'Sea Green', 51.0: 'Dark Green',
                   52.0: 'Olive Green', 53.0: 'Brown', 54.0: 'Plum',
                   55.0: 'Indigo', 56.0: 'Gray-80%', -4142: ''}

    try:
        return color_codes[xl_range.Interior.ColorIndex]
    except KeyError:
        return xl_range.Interior.ColorIndex