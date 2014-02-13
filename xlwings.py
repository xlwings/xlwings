"""
xlwings is the easiest way to deploy your Python powered Excel tools on Windows.
Homepage and documentation: http://xlwings.org/

Copyright (c) 2014, Zoomer Analytics LLC.
License: BSD (see LICENSE.txt for details)

"""

import sys
from win32com.client import GetObject
import win32com.client.dynamic
import pywintypes
import pythoncom
import numbers
import datetime as dt
import pytz

# Optional imports
try:
    import numpy as np
except ImportError:
    np = None
try:
    from pandas import MultiIndex
except ImportError:
    MultiIndex = None
try:
    import pandas as pd
except ImportError:
    pd = None


__version__ = '0.1.0-dev'


# Python 2 and 3 compatibility
PY3 = sys.version_info.major >= 3
if PY3:
    string_types = str
    time_types = (dt.date, dt.datetime, type(pywintypes.Time(0)))
else:
    string_types = basestring
    time_types = (dt.date, dt.datetime, pywintypes.TimeType)

# Excel constants: We can't use 'from win32com.client import constants' as we're dynamically dispatching
xlDown = -4121
xlToRight = -4161


def clean_com_data(data):
    """
    Brings data from tuples of tuples into list of list and
    transforms pywintypes Time objects into Python datetime objects.

    Parameters
    ----------
    data : tuple of tuple
        raw data as returned from Excel through pywin32

    Returns
    -------
    data : list of list
        data is a list of list with native Python datetime objects

    """
    # Turn into list of list for easier handling (e.g. for Pandas DataFrame)
    data = [list(row) for row in data]

    # Handle dates
    data = [[_com_time_to_datetime(c) if isinstance(c, time_types) else c for c in row] for row in data]

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
        # pywintypes promises to only return instances set to UTC; see doc link in _datetime_to_com_time
        assert com_time.tzinfo is not None
        return dt.datetime(month=com_time.month, day=com_time.day, year=com_time.year,
                           hour=com_time.hour, minute=com_time.minute, second=com_time.second,
                           microsecond=com_time.microsecond, tzinfo=com_time.tzinfo)
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
            dt_time = dt_time.replace(tzinfo=pytz.utc)
            
        return dt_time
    else:
        assert dt_time.microsecond == 0, "fractional seconds not yet handled"
        return pywintypes.Time(dt_time.timetuple())


def _is_file_open(fullname):
    """
    Checks the Running Object Table (ROT) for the fully qualified filename
    """
    context = pythoncom.CreateBindCtx(0)
    for moniker in pythoncom.GetRunningObjectTable():
        name = moniker.GetDisplayName(context, None)
        if name.lower() == fullname.lower():
            return True
    return False


class Workbook(object):
    """
    Workbook connects an Excel Workbook with Python. You can create a new connection from Python with
    - a new workbook: wb = Workbook()
    - an existing workbook: wb = Workbook(r'C:\path\to\file.xlsx')

    If you want to create the connection from Excel through the xlwings VBA module, use:
        wb = Workbook()

    Parameters
    ----------
    fullname : string, default None
        For debugging/interactive use from within Python, provide the fully qualified name, e.g: 'C:\path\to\file.xlsx'
        No arguments must be provided if called from Excel through the xlwings VBA module.

    Returns
    -------
    workbook : xlwings Workbook object

    """
    def __init__(self, fullname=None):
        if fullname:
            self.fullname = fullname.lower()
            if _is_file_open(self.fullname):
                # GetObject() returns the correct Excel instance if there are > 1
                self.Workbook = GetObject(self.fullname)
                self.App = self.Workbook.Application
            else:
                self.App = win32com.client.dynamic.Dispatch('Excel.Application')
                self.Workbook = self.App.Workbooks.Open(self.fullname)
                self.App.Visible = True
        elif len(sys.argv) >= 2 and sys.argv[2] == 'from_xl':
            self.fullname = sys.argv[1].lower()
            self.Workbook = GetObject(self.fullname)
            self.App = self.Workbook.Application
        else:
            self.App = win32com.client.dynamic.Dispatch('Excel.Application')
            self.App.Visible = True
            self.Workbook = self.App.Workbooks.Add()

        self.name = self.Workbook.Name

        # Make the most recently created Workbook the default when creating Range objects directly
        global wb
        wb = self.Workbook

    def get_selection(self, asarray=False):
        """
        Returns the currently selected Range from Excel as xlwing Range object.

        Parameters
        ----------
        asarray : boolean, default False
            returns a NumPy array where empty cells are shown as nan

        Returns
        -------
        Range : xlwings Range object
        """
        return self.range(str(self.Workbook.Application.Selection.Address), asarray=asarray)

    def range(self, *args, **kwargs):
        """
        The range method gets and sets the Range object with the following arguments:

        range('A1')          range('Sheet1', 'A1')          range(1, 'A1')
        range('A1:C3')       range('Sheet1', 'A1:C3')       range(1, 'A1:C3')
        range((1,2))         range('Sheet1, (1,2))          range(1, (1,2))
        range((1,1), (3,3))  range('Sheet1', (1,1), (3,3))  range(1, (1,1), (3,3))
        range('NamedRange')  range('Sheet1', 'NamedRange')  range(1, 'NamedRange')

        If no worksheet name is provided as first argument (as name or index),
        it will take the range from the active sheet.

        You usually want to go for something like wb.range('A1').value to get the values as list of lists.

        Parameters
        ----------
        asarray : boolean, default False
            returns a NumPy array where empty cells are shown as nan

        index : boolean, default True
            Includes the index when setting a Pandas DataFrame

        header : boolean, default True
            Includes the column headers when setting a Pandas DataFrame

        Returns
        -------
        Range : xlwings Range object
        """
        return Range(*args, workbook=self.Workbook, **kwargs)

    def __repr__(self):
        return "<xlwings.Workbook '{0}'>".format(self.name)

class Range(object):
    """
    A Range object can be created with the following arguments:

    Range('A1')          Range('Sheet1', 'A1')          Range(1, 'A1')
    Range('A1:C3')       Range('Sheet1', 'A1:C3')       Range(1, 'A1:C3')
    Range((1,2))         Range('Sheet1, (1,2))          Range(1, (1,2))
    Range((1,1), (3,3))  Range('Sheet1', (1,1), (3,3))  Range(1, (1,1), (3,3))
    Range('NamedRange')  Range('Sheet1', 'NamedRange')  Range(1, 'NamedRange')

    If no worksheet name is provided as first argument (as name or index),
    it will take the Range from the active sheet.

    You usually want to go for Range(...).value to get the values as list of lists.

    Parameters
    ----------
    asarray : boolean, default False
        returns a NumPy array where empty cells are shown as nan

    index : boolean, default True
        Includes the index when setting a Pandas DataFrame

    header : boolean, default True
        Includes the column headers when setting a Pandas DataFrame
    """
    def __init__(self, *args, **kwargs):
        # Arguments
        if len(args) == 1 and isinstance(args[0], string_types):
            sheet = None
            cell_range = args[0]
        elif len(args) == 1 and isinstance(args[0], tuple):
            sheet = None
            cell_range = None
            self.row1 = args[0][0]
            self.col1 = args[0][1]
            self.row2 = self.row1
            self.col2 = self.col1
        elif (len(args) == 2
              and isinstance(args[0], (numbers.Number, string_types))
              and isinstance(args[1], string_types)):
            sheet = args[0]
            cell_range = args[1]
        elif (len(args) == 2
              and isinstance(args[0], (numbers.Number, string_types))
              and isinstance(args[1], tuple)):
            sheet = args[0]
            cell_range = None
            self.row1 = args[1][0]
            self.col1 = args[1][1]
            self.row2 = self.row1
            self.col2 = self.col1
        elif len(args) == 2 and isinstance(args[0], tuple):
            sheet = None
            cell_range = None
            self.row1 = args[0][0]
            self.col1 = args[0][1]
            self.row2 = args[1][0]
            self.col2 = args[1][1]
        elif len(args) == 3:
            sheet = args[0]
            cell_range = None
            self.row1 = args[1][0]
            self.col1 = args[1][1]
            self.row2 = args[2][0]
            self.col2 = args[2][1]

        # Keyword Arguments
        self.kwargs = kwargs
        self.index = kwargs.get('index', True)  # Set DataFrame with index
        self.header = kwargs.get('header', True)  # Set DataFrame with header
        self.asarray = kwargs.get('asarray', False)  # Return Data as NumPy Array
        self.strict = kwargs.get('strict', False)  # Stop table/horizontal/vertical at empty cells that contain formulas
        self.workbook = kwargs.get('workbook', wb)

        # Get sheet
        if sheet:
            self.sheet = self.workbook.Worksheets(sheet)
        else:
            self.sheet = self.workbook.ActiveSheet

        # Get row1, col1, row2, col2 out of Range object
        if cell_range:
            self.row1 = self.sheet.Range(cell_range).Row
            self.col1 = self.sheet.Range(cell_range).Column
            self.row2 = self.row1 + self.sheet.Range(cell_range).Rows.Count - 1
            self.col2 = self.col1 + self.sheet.Range(cell_range).Columns.Count - 1

        self.cell_range = self.sheet.Range(self.sheet.Cells(self.row1, self.col1),
                                           self.sheet.Cells(self.row2, self.col2))

    @property
    def value(self):
        """
        Gets or sets the values for the given Range.

        Returns
        -------
        data : list of list or NumPy array
        """
        if self.row1 == self.row2 and self.col1 == self.col2:
            # Single cell - clean_com_data requires and returns a list of list
            data = clean_com_data([[self.cell_range.Value]])[0][0]
        else:
            # At least 2 cells
            data = clean_com_data(self.cell_range.Value)

        # Return as NumPy Array
        if self.asarray:
            # replace None (empty cells) with nan as None produces arrays with dtype=object
            if data is None:
                data = np.nan
            elif not isinstance(data, (numbers.Number, string_types)):
                data = [[np.nan if x is None else x for x in i] for i in data]
            return np.array(data)
        return data

    @value.setter
    def value(self, data):
        # Pandas DataFrame: Turn into NumPy object array with or without Index and Headers
        if hasattr(pd, 'DataFrame') and isinstance(data, pd.DataFrame):
            if self.index:
                data = data.reset_index()

            if self.header:
                if isinstance(data.columns, MultiIndex):
                    columns = np.array(zip(*data.columns.tolist()))
                else:
                    columns = np.array([data.columns.tolist()])
                data = np.vstack((columns, data.values))
            else:
                data = data.values

        # NumPy array: Handle NaN values and turn into list of list (Python 3 can't handle arrays directly)
        if hasattr(np, 'ndarray') and isinstance(data, np.ndarray):
            try:
                # nan have to be transformed to None, otherwise Excel shows them as 65535
                data = np.where(np.isnan(data), None, data)
            except TypeError:
                # isnan doesn't work on arrays of dtype=object
                data[pd.isnull(data)] = None
            data = data.tolist()

        # Simple Lists: Turn into list of lists
        if isinstance(data, list) and isinstance(data[0], (numbers.Number, string_types, time_types)):
            data = [data]

        # Get dimensions and handle date values
        if isinstance(data, (numbers.Number, string_types, time_types)):
            # Single cells
            row2 = self.row2
            col2 = self.col2
            if isinstance(data, time_types):
                data = _datetime_to_com_time(data)
        else:
            # List of List
            row2 = self.row1 + len(data) - 1
            col2 = self.col1 + len(data[0]) - 1
            data = [[_datetime_to_com_time(c) if isinstance(c, time_types) else c for c in row] for row in data]

        self.sheet.Range(self.sheet.Cells(self.row1, self.col1), self.sheet.Cells(row2, col2)).Value = data

    @property
    def table(self):
        """
        Returns a contiguous Range starting with the indicated cell as top-left corner and going down and right as long
        as no empty cell is hit. For example, to get the values of a contiguous range or clear its contents use:

            Range('A1').table.value
            Range('A1').table.clear_contents()

        Parameters
        ----------
        strict : boolean, default False
            strict stops the table at empty cells even if they contain a formula. Less efficient than if set to False.

        Returns
        -------
        range : Range
            xlwings Range object

        """
        row2 = Range(self.sheet.Name, (self.row1, self.col1), **self.kwargs).vertical.row2
        col2 = Range(self.sheet.Name, (self.row1, self.col1), **self.kwargs).horizontal.col2

        return Range(self.sheet.Name, (self.row1, self.col1), (row2, col2), **self.kwargs)

    @property
    def vertical(self):
        """
        Returns a contiguous Range starting with the indicated cell and going down as long as no empty cell is hit. For
        example, to get the values of a contiguous range or clear its contents use:

            Range('A1').vertical.value
            Range('A1').vertical.clear_contents()

        Parameters
        ----------
        strict : bool, default False
            strict stops the table at empty cells even if they contain a formula. Less efficient than if set to False.

        Returns
        -------
        range : Range
            xlwings Range object

        """
        # A single cell is a special case as End(xlDown) jumps over adjacent empty cells
        if self.sheet.Cells(self.row1 + 1, self.col1).Value in [None, ""]:
            row2 = self.row1
        else:
            row2 = self.sheet.Cells(self.row1, self.col1).End(xlDown).Row

        # Strict stops at cells that contain a formula but show an empty value
        if self.strict:
            row2 = self.row1
            while self.sheet.Cells(row2 + 1, self.col1).Value not in [None, ""]:
                row2 += 1

        col2 = self.col2

        return Range(self.sheet.Name, (self.row1, self.col1), (row2, col2), **self.kwargs)

    @property
    def horizontal(self):
        """
        Returns a contiguous Range starting with the indicated cell and going right as long as no empty cell is hit. For
        example, to get the values of a contiguous range or clear its contents use:

            Range('A1').horizontal.value
            Range('A1').horizontal.clear_contents()

        Parameters
        ----------
        strict : bool, default False
            strict stops the table at empty cells even if they contain a formula. Less efficient than if set to False.

        Returns
        -------
        range : Range
            xlwings Range object

        """
        # A single cell is a special case as End(xlDown) jumps over adjacent empty cells
        if self.sheet.Cells(self.row1, self.col1 + 1).Value in [None, ""]:
            col2 = self.col1
        else:
            col2 = self.sheet.Cells(self.row1, self.col1).End(xlToRight).Column

        # Strict: stops at cells that contain a formula but show an empty value
        if self.strict:
            col2 = self.col1
            while self.sheet.Cells(self.row1, col2 + 1).Value not in [None, ""]:
                col2 += 1

        row2 = self.row2

        return Range(self.sheet.Name, (self.row1, self.col1), (row2, col2), **self.kwargs)

    @property
    def current_region(self):
        """
        The current_region property returns a Range object representing a range bounded by (but not including) any
        combination of blank rows and blank columns or the edges of the worksheet
        VBA equivalent: CurrentRegion property of Range object

        Returns
        -------
        range : Range
            xlwings Range object

        """
        current_region = self.sheet.Cells(self.row1, self.col1).CurrentRegion
        row2 = self.row1 + current_region.Rows.Count - 1
        col2 = self.col1 + current_region.Columns.Count - 1
        return Range(self.sheet.Name, (self.row1, self.col1), (row2, col2), **self.kwargs)

    def clear(self):
        """
        Clears the content and the formatting of a Range.
        """
        self.cell_range.Clear()

    def clear_contents(self):
        """
        Clears the content of a Range but leaves the formatting.
        """
        self.cell_range.ClearContents()



