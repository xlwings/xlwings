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