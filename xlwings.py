"""
xlwings makes it easy to deploy your Python powered Excel tools on Windows.
Homepage and documentation: http://xlwings.org/

Copyright (c) 2014, Zoomer Analytics.
License: MIT (see LICENSE.txt for details)

"""

import sys
import os
from win32com.client import GetObject
import win32com.client.dynamic
import adodbapi
from pywintypes import TimeType
import numpy as np
from pandas import MultiIndex
import pandas as pd


__version__ = '0.1-dev'

_is_python3 = sys.version_info.major > 2


def xlwings_connect(fullname=None):
    """
    Establishes a connection between an Excel file and Python


    Parameters
    ----------
    fullname : string, default None
        For debugging/interactive use from within Python, provide the fully qualified name, e.g: 'C:\path\to\file.xlsx'
        No arguments must be provided if called from Excel through the xlwings VBA module.
    """
    if fullname:
        fullname = fullname.lower()
    else:
        fullname = sys.argv[1].lower()
    global Workbook
    global xlApp
    xlApp = win32com.client.dynamic.Dispatch('Excel.Application')
    Workbook = GetObject(fullname)  # GetObject() returns the correct Excel instance if there are > 1


def clean_com_data(data):
    """
    Transforms data from tuples of tuples into list of list and
    on Python 2, transforms PyTime Objects from COM into datetime objects.

    Parameters
    ----------
    data : raw data as returned from Excel (tuple of tuple)

    """
    # Turn into list of list for easier handling (e.g. for Pandas DataFrame)
    data = [list(row) for row in data]

    # Check which columns contain COM dates
    # TODO: replace with datetime transformations from pyvot -> python3?
    # TODO: simplify like this: [[tc.DateObjectFromCOMDate(c) for c in row] for row in data]
    if _is_python3 is True:
        return data
    else:
        tc = adodbapi.pythonDateTimeConverter()
        for i in range(len(data[0])):
            if any([type(row[i]) is TimeType for row in data]):
                # Transform PyTime into datetime
                for j, cell in enumerate([row[i] for row in data]):
                    if type(cell) is TimeType:
                        data[j][i] = tc.DateObjectFromCOMDate(cell)
    return data


class Range(object):
    """
    A Range object can be created with the following arguments:

    Range('A1')
    TODO: Range((1,2))
    Range('A1:C3')
    Range('NamedRange')
    Range((1,1), (3,3))

    If no worksheet name is provided as first argument, it will take the range from the active sheet. To get
    the range from a specific sheet, provide the worksheet name as first argument like so:

    Range('Sheet1','A1')
    """
    def __init__(self, *args, **kwargs):
        self.index = kwargs.get('index', True)
        self.header = kwargs.get('header', True)
        # Parse arguments
        if len(args) == 1:
            sheet = None
            cell_range = args[0]
        if len(args) == 2 and type(args[0]) is str:  # TODO: change to isinstance // elif
            sheet = args[0]
            cell_range = args[1]
        if len(args) == 2 and type(args[0]) is not str:
            sheet = None
            cell_range = None
            self.row1 = args[0][0]
            self.col1 = args[0][1]
            self.row2 = args[1][0]
            self.col2 = args[1][1]
        if len(args) == 3:
            sheet = args[0]
            cell_range = None
            self.row1 = args[1][0]
            self.col1 = args[1][1]
            self.row2 = args[2][0]
            self.col2 = args[2][1]

        # Get cells
        if sheet:
            self.sheet = Workbook.Worksheets(sheet)
        else:
            self.sheet = Workbook.ActiveSheet

        if cell_range:
            self.row1 = self.sheet.Range(cell_range.split(':')[0]).Row
            self.col1 = self.sheet.Range(cell_range.split(':')[0]).Column

            if len(cell_range.split(':')) == 2:
                self.row2 = self.sheet.Range(cell_range.split(':')[1]).Row
                self.col2 = self.sheet.Range(cell_range.split(':')[1]).Column
            else:
                self.row2 = self.row1
                self.col2 = self.col1

        self.cell_range = self.sheet.Range(self.sheet.Cells(self.row1, self.col1),
                                           self.sheet.Cells(self.row2, self.col2))

    @property
    def value(self):
            if self.row1 == self.row2 and self.col1 == self.col2:
                return clean_com_data([[self.cell_range.Value]])[0][0]  # TODO: introduce as_matrix method?
            else:
                return clean_com_data(self.cell_range.Value)

    @value.setter
    def value(self, data):
        # TODO: Range('A1:C3').value = 5 should apply 5 to the whole range
        if isinstance(data, np.ndarray):
            data = data.tolist()  # Python 3 can't handle arrays directly
        elif isinstance(data, pd.DataFrame):
            df = data
            if self.index:
                df = data.reset_index()

            if self.header:
                if type(df.columns) is MultiIndex:
                    columns = np.array(zip(*df.columns.tolist()))
                else:
                    columns = np.array([df.columns.tolist()])
                    data = np.vstack((columns, df.values))
            else:
                data = df.values  # TODO: simplify

        row2 = self.row1 + len(data) - 1
        col2 = self.col1 + len(data[0]) - 1

        self.sheet.Range(self.sheet.Cells(self.row1, self.col1), self.sheet.Cells(row2, col2)).Value = data

    @property
    def table(self):
        """
        TODO:
        """
        row2 = self.row1
        while self.sheet.Cells(row2 + 1, self.col1).Value not in [None, ""]:
            row2 += 1

        # Right column
        col2 = self.col1
        while self.sheet.Cells(self.row1, col2 + 1).Value not in [None, ""]:
            col2 += 1

        return Range(self.sheet.Name, (self.row1, self.col1), (row2, col2))

    @property
    def current_region(self):
        """
        The current_region property returns a CellRange object representing a range bounded by (but not including) any
        combination of blank rows and blank columns or the edges of the worksheet
        VBA equivalent: CurrentRegion property of Range object
        """
        current_region = self.sheet.Cells(self.row1, self.col1).CurrentRegion
        row2 = self.row1 + current_region.Rows.Count - 1
        col2 = self.col1 + current_region.Columns.Count - 1
        return Range(self.sheet.Name, (self.row1, self.col1), (row2, col2))

    def clear(self):
        self.cell_range.Clear()

    def clear_contents(self):
        self.cell_range.ClearContents()

if __name__ == "__main__":
    xlwings_connect(r'C:\DEV\Git\xlwings\example.xlsm')

    # Assumes Sheet3 to be the active one

    # Cell
    print Range('B1').value
    print Range('Sheet3', 'A1').value
    Range('G2').value = [[23]]  # TODO: accept numbers or strings without having to do [[]]

    # CellRange
    print Range('A1').value
    print np.array(Range('A1:C3').value)
    print np.array(Range('test_range').value)
    print np.array(Range((1,1), (3,3)).value)

    print Range('Sheet3', 'A1').value
    print np.array(Range('Sheet3', 'A1:C3').value)
    print np.array(Range('Sheet3', 'test_range').value)
    print np.array(Range('Sheet3', (1,1), (3,3)).value)

    Range('A25').value = [[23]]
    Range('A1:C3').value = [[11,22,33], [44,55,66], [77,88,99]]
    Range('single_cell').value = np.eye(4)
    Range((5,5), (8,8)).value = [[1,2,3], [4,5,6], [7,8,9]]

    print np.array(Range('Sheet3', 'G23', table=True).value)
    print np.array(Range('Sheet3', 'G23', current_region=True).value)
    print np.array(Range('Sheet3', 'G23').current_region.value)
    print np.array(Range('Sheet3', 'G23').table.value)

    # DataFrame
    data = Range('Sheet2', 'A1:E7').value
    df = pd.DataFrame(data[1:], columns=data[0])
    df.set_index('test 1', inplace=True)
    df.index = pd.to_datetime(df.index)
    print(df)
    print(df.info())

    Range('Sheet2', 'H1', index=False, header=False).value = df


class Xl:
    """
    TODO: Deprecated

    """

    def __init__(self, fullname=None):
        # TODO: check if fullname exists
        if fullname:
            self.fullname = fullname.lower()
        else:
            # TODO: catch AttributeError in case called from Python without argument or
            # TODO: pass 'from_excel' arg when called from VBA
            self.fullname = sys.argv[1].lower()
        self.App = win32com.client.dynamic.Dispatch('Excel.Application')
        self.Workbook = GetObject(self.fullname)  # GetObject() gives us the correct Excel instance if there are > 1
        self.filename = os.path.split(self.fullname)[1]

    def save(self, newfilename=None):
        if newfilename:
            self.filename = newfilename
            self.Workbook.SaveAs(newfilename)
        else:
            self.Workbook.Save()