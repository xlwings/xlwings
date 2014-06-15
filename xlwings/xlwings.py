"""
Make Excel fly!

Homepage and documentation: http://xlwings.org
See also: http://zoomeranalytics.com

Copyright (C) 2014, Zoomer Analytics LLC.
All rights reserved.

License: BSD 3-clause (see LICENSE.txt for details)
"""

import sys
import numbers
import datetime as dt
from win32com.client import GetObject, dynamic
import win32timezone
import pywintypes
import pythoncom

# Optional imports
try:
    import numpy as np
except ImportError:
    np = None
try:
    import pandas as pd
except ImportError:
    pd = None

__version__ = '0.1.1dev'

# Python 2 and 3 compatibility
PY3 = sys.version_info[0] >= 3
if PY3:
    string_types = str
    time_types = (dt.date, dt.datetime, type(pywintypes.Time(0)))
else:
    string_types = basestring
    time_types = (dt.date, dt.datetime, pywintypes.TimeType)

# Excel constants: We can't use 'from win32com.client import constants' as we're dynamically dispatching
xlDown, xlToRight = -4121, -4161


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
    list
        list of list with native Python datetime objects

    """
    # Turn into list of list (e.g. for Pandas DataFrame) and handle dates
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
            dt_time = dt_time.replace(tzinfo=win32timezone.TimeZoneInfo.utc())
            
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

    * a new workbook: ``wb = Workbook()``
    * an existing workbook: ``wb = Workbook(r'C:\\path\\to\\file.xlsx')``

    If you want to create the connection from Excel through the xlwings VBA module, use:

    ``wb = Workbook()``

    Parameters
    ----------
    fullname : string, default None
        If you want to connect to an existing Excel file from Python, use the fullname, e.g:
        ``r'C:\\path\\to\\file.xlsx'``

    Returns
    -------
    Workbook
        xlwings Workbook object

    """
    def __init__(self, fullname=None):
        if fullname:
            # Use/open an existing workbook
            self.fullname = fullname.lower()
            if _is_file_open(self.fullname):
                # GetObject() returns the correct Excel instance if there are > 1
                self.com_workbook = GetObject(self.fullname)
                self.com_app = self.com_workbook.Application
            else:
                self.com_app = dynamic.Dispatch('Excel.Application')
                self.com_workbook = self.com_app.Workbooks.Open(self.fullname)
                self.com_app.Visible = True
        elif len(sys.argv) >= 2 and sys.argv[2] == 'from_xl':
            # Connect to the workbook from which this code has been invoked
            self.fullname = sys.argv[1].lower()
            self.com_workbook = GetObject(self.fullname)
            self.com_app = self.com_workbook.Application
        else:
            # Open Excel if necessary and create a new workbook
            self.com_app = dynamic.Dispatch('Excel.Application')
            self.com_app.Visible = True
            self.com_workbook = self.com_app.Workbooks.Add()

        self.name = self.com_workbook.Name
        self.active_sheet = ActiveSheet(workbook=self.com_workbook)

        # Make the most recently created Workbook the default when creating Range objects directly
        global wb  # TODO: rename into com_wb
        wb = self.com_workbook

    def activate(self, sheet):
        """
        Activates the given sheet.

        Parameters
        ----------
        sheet : string or integer
            Sheet name or index.
        """
        self.com_workbook.Sheets(sheet).Activate()

    def get_selection(self, asarray=False):
        """
        Returns the currently selected Range from Excel as xlwings Range object.

        Parameters
        ----------
        asarray : boolean, default False
            returns a NumPy array where empty cells are shown as nan

        Returns
        -------
        Range
            xlwings Range object
        """
        return self.range(str(self.com_workbook.Application.Selection.Address), asarray=asarray)

    def range(self, *args, **kwargs):
        """
        The range method gets and sets the Range object with the following arguments::

            range('A1')          range('Sheet1', 'A1')          range(1, 'A1')
            range('A1:C3')       range('Sheet1', 'A1:C3')       range(1, 'A1:C3')
            range((1,2))         range('Sheet1, (1,2))          range(1, (1,2))
            range((1,1), (3,3))  range('Sheet1', (1,1), (3,3))  range(1, (1,1), (3,3))
            range('NamedRange')  range('Sheet1', 'NamedRange')  range(1, 'NamedRange')

        If no worksheet name is provided as first argument (as name or index),
        it will take the range from the active sheet.

        Please check the available methods/properties directly under the Range object.

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
        Range
            xlwings Range object
        """
        return Range(*args, workbook=self.com_workbook, **kwargs)

    def chart(self, *args, **kwargs):
        """
        The chart method gives access to the Chart object and can be called with the following arguments::

            chart(1)            chart('Sheet1', 1)              chart(1, 1)
            chart('Chart 1')    chart('Sheet1', 'Chart 1')      chart(1, 'Chart 1')

        If no worksheet name is provided as first argument (as name or index),
        it will take the Chart from the active sheet.

        Please check the available methods/properties directly under the Chart object.

        Parameters
        ----------
        *args :
            Definition of Sheet (optional) and Chart in the above described combinations.



        """
        return Chart(*args, workbook=self.com_workbook, **kwargs)

    def clear_contents(self, sheet):
        """
        Clears the content of a whole Sheet but leaves the formatting.

        Parameters
        ----------
        sheet : string or integer
            Sheet name or index.
        """
        self.com_workbook.Sheets(sheet).Cells.ClearContents()

    def clear(self, sheet):
        """
        Clears the content and formatting of a whole Sheet.

        Parameters
        ----------
        sheet : string or integer
            Sheet name or index.
        """
        self.com_workbook.Sheets(sheet).Cells.Clear()

    def close(self):
        """Closes the Workbook without saving it"""
        self.com_workbook.Close(SaveChanges=False)

    def __repr__(self):
        return "<xlwings.Workbook '{0}'>".format(self.name)


class ActiveSheet(object):
    """

    """
    def __init__(self, workbook=None):
        if workbook is None:
            workbook = wb
        self.com_active_sheet = workbook.ActiveSheet
        self.name = self.com_active_sheet.Name


class Range(object):
    """
    A Range object can be created with the following arguments::

        Range('A1')          Range('Sheet1', 'A1')          Range(1, 'A1')
        Range('A1:C3')       Range('Sheet1', 'A1:C3')       Range(1, 'A1:C3')
        Range((1,2))         Range('Sheet1, (1,2))          Range(1, (1,2))
        Range((1,1), (3,3))  Range('Sheet1', (1,1), (3,3))  Range(1, (1,1), (3,3))
        Range('NamedRange')  Range('Sheet1', 'NamedRange')  Range(1, 'NamedRange')

    If no worksheet name is provided as first argument (as name or index),
    it will take the Range from the active sheet.

    You usually want to go for ``Range(...).value`` to get the values (as list of lists).

    Parameters
    ----------
    *args :
        Definition of Sheet (optional) and Range in the above described combinations.
    asarray : boolean, default False
        Returns a NumPy array where empty cells are transformed into nan.

    index : boolean, default True
        Includes the index when setting a Pandas DataFrame.

    header : boolean, default True
        Includes the column headers when setting a Pandas DataFrame.
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

        # Get COM Range object
        if cell_range:
            self.row1 = self.sheet.Range(cell_range).Row
            self.col1 = self.sheet.Range(cell_range).Column
            self.row2 = self.row1 + self.sheet.Range(cell_range).Rows.Count - 1
            self.col2 = self.col1 + self.sheet.Range(cell_range).Columns.Count - 1

        self.com_range = self.sheet.Range(self.sheet.Cells(self.row1, self.col1),
                                          self.sheet.Cells(self.row2, self.col2))

    @property
    def value(self):
        """
        Gets or sets the values for the given Range.

        Returns
        -------
        list
            Empty cells are set to None. If ``asarray=True``, a numpy array is returned where empty cells are set to nan.
        """
        if self.row1 == self.row2 and self.col1 == self.col2:
            # Single cell - clean_com_data requires and returns a list of list
            data = clean_com_data([[self.com_range.Value]])[0][0]
        else:
            # At least 2 cells
            data = clean_com_data(self.com_range.Value)

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
                if isinstance(data.columns, pd.MultiIndex):
                    columns = np.array(zip(*data.columns.tolist()))
                else:
                    columns = np.array([data.columns.tolist()])
                data = np.vstack((columns, data.values))
            else:
                data = data.values

        # NumPy array: nan have to be transformed to None, otherwise Excel shows them as 65535
        # Also, turn into list (Python 3 can't handle arrays directly)
        if hasattr(np, 'ndarray') and isinstance(data, np.ndarray):
            try:
                data = np.where(np.isnan(data), None, data)
                data = data.tolist()
            except TypeError:
                # isnan doesn't work on arrays of dtype=object
                if hasattr(pd, 'isnull'):
                    data[pd.isnull(data)] = None
                    data = data.tolist()
                else:
                    # expensive way of replacing nan with None in object arrays in case Pandas is not available
                    data = [[None if isinstance(c, float) and np.isnan(c) else c for c in row] for row in data]

        # Simple Lists: Turn into list of lists (np.nan is part of numbers.Number)
        if isinstance(data, list) and (isinstance(data[0],
                                                 (numbers.Number, string_types, time_types)) or data[0] is None):
            data = [data]

        # Get dimensions and handle date values
        if isinstance(data, (numbers.Number, string_types, time_types)) or data is None:
            # Single cells
            row2 = self.row2
            col2 = self.col2
            if isinstance(data, time_types):
                data = _datetime_to_com_time(data)
            try:
                # scalar np.nan need to be turned into None, otherwise Excel shows it as 65535 (same as for NumPy array)
                if hasattr(np, 'ndarray') and np.isnan(data):
                    data = None
            except TypeError:
                pass

        else:
            # List of List
            row2 = self.row1 + len(data) - 1
            col2 = self.col1 + len(data[0]) - 1
            data = [[_datetime_to_com_time(c) if isinstance(c, time_types) else c for c in row] for row in data]

        self.sheet.Range(self.sheet.Cells(self.row1, self.col1), self.sheet.Cells(row2, col2)).Value = data

    @property
    def formula(self):
        """
        Gets or sets the formula for the given Range.
        """
        return self.com_range.Formula

    @formula.setter
    def formula(self, value):
        self.com_range.Formula = value

    @property
    def table(self):
        """
        Returns a contiguous Range starting with the indicated cell as top-left corner and going down and right as
        long as no empty cell is hit.

        Parameters
        ----------
        strict : boolean, default False
            strict stops the table at empty cells even if they contain a formula. Less efficient than if set to False.

        Returns
        -------
        Range
            xlwings Range object

        Examples
        --------
        To get the values of a contiguous range or clear its contents use::

            Range('A1').table.value
            Range('A1').table.clear_contents()

        """
        row2 = Range(self.sheet.Name, (self.row1, self.col1), **self.kwargs).vertical.row2
        col2 = Range(self.sheet.Name, (self.row1, self.col1), **self.kwargs).horizontal.col2

        return Range(self.sheet.Name, (self.row1, self.col1), (row2, col2), **self.kwargs)

    @property
    def vertical(self):
        """
        Returns a contiguous Range starting with the indicated cell and going down as long as no empty cell is hit.
        This corresponds to ``Ctrl + Shift + Down Arrow`` in Excel.

        Parameters
        ----------
        strict : bool, default False
            strict stops the table at empty cells even if they contain a formula. Less efficient than if set to False.

        Returns
        -------
        Range
            xlwings Range object

        Examples
        --------
        To get the values of a contiguous range or clear its contents use::

            Range('A1').vertical.value
            Range('A1').vertical.clear_contents()

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
        Returns a contiguous Range starting with the indicated cell and going right as long as no empty cell is hit.

        Parameters
        ----------
        strict : bool, default False
            strict stops the table at empty cells even if they contain a formula. Less efficient than if set to False.

        Returns
        -------
        Range
            xlwings Range object

        Examples
        --------
        To get the values of a contiguous range or clear its contents use::

            Range('A1').horizontal.value
            Range('A1').horizontal.clear_contents()

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
        combination of blank rows and blank columns or the edges of the worksheet. It corresponds to ``Ctrl + *``.

        Returns
        -------
        Range
            xlwings Range object

        """
        address = str(self.sheet.Cells(self.row1, self.col1).CurrentRegion.Address)
        return Range(self.sheet.Name, address, **self.kwargs)

    def clear(self):
        """
        Clears the content and the formatting of a Range.
        """
        self.com_range.Clear()

    def clear_contents(self):
        """
        Clears the content of a Range but leaves the formatting.
        """
        self.com_range.ClearContents()

    def __repr__(self):
        return "<xlwings.Range of Workbook '{0}'>".format(self.workbook.name)


class Chart(object):
    """
    A Chart object can be created with the following arguments::

        Chart(1)            Chart('Sheet1', 1)              Chart(1, 1)
        Chart('Chart 1')    Chart('Sheet1', 'Chart 1')      Chart(1, 'Chart 1')

    If no worksheet name is provided as first argument (as name or index),
    it will take the Chart from the active sheet.

    Parameters
    ----------
    *args :
        Definition of Sheet (optional) and Chart in the above described combinations.

    """
    def __init__(self, *args, **kwargs):
        # Keyword Arguments
        self.workbook = kwargs.get('workbook', wb)

        # Arguments
        if len(args) == 0:
            pass
        elif len(args) > 0:
            if len(args) == 1:
                sheet = self.workbook.ActiveSheet.Name
                name_or_index = args[0]
            elif len(args) == 2:
                sheet = args[0]
                name_or_index = args[1]

            # Get Chart COM object
            self.com_chart = wb.Sheets(sheet).ChartObjects(name_or_index)
            self.index = self.com_chart.Index

    def __repr__(self):
        return "<xlwings.Chart '{0}'>".format(self.name)

    @property
    def name(self):
        """
        Gets and sets the name of a Chart
        """
        return self.com_chart.Name

    @name.setter
    def name(self, value):
        self.com_chart.Name = value

    def activate(self):
        self.com_chart.Activate()

    def set_source_data(self, source):
        """
        Sets the source for the chart

        Arguments
        ---------
        source : Range
            xlwings Range object, e.g. ``Range('A1')``
        """
        self.com_chart.Chart.SetSourceData(source.com_range)

    def add(self, sheet=None, left=168, top=217, width=355, height=211):
        """
        Adds a new Chart

        Arguments
        ---------
        sheet : string or integer, default None
            Name or Index of the sheet, defaults to the active sheet
        left : float, default 100
            left position in points
        top : float, default 75
            top position in points
        width : float, default 375
            width in points
        height : float, default 225
            height in points

        """
        if sheet is None:
            sheet = wb.ActiveSheet.Name

        com_chart = wb.Sheets(sheet).ChartObjects().Add(left, top, width, height)
        return Chart(sheet, com_chart.Name)

if __name__ == '__main__':
    wb = Workbook()
    Range('A1').value