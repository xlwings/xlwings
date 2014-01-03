"""
xlwings makes it easy to deploy your Python powered Excel tools on Windows.
Homepage and documentation: http://xlwings.org/

Copyright (c) 2014, Felix Zumstein.
License: MIT (see LICENSE.txt for details)

"""

import sys
import os
from win32com.client import Dispatch, GetObject
import win32com.client.dynamic
import adodbapi
from pywintypes import TimeType
import numpy as np
from pandas import MultiIndex

__version__ = '0.1-dev'
__license__ = 'MIT'

_is_python3 = sys.version_info.major > 2


def xlwings_connect(fullname=None):
    """
    Establishes a connection between a specific Excel file and Python


    Parameters
    ----------
    fullname : string, default None
        For debugging/interactive use from within Python, provide the fully qualified name, e.g: 'C:\path\to\file.xlsx'
        Leave empty if called from Excel.
    """
    if fullname:
        fullname = fullname.lower()
    else:
        fullname = sys.argv[1].lower()
    global Workbook
    Workbook = GetObject(fullname)  # GetObject() returns the correct Excel instance if there are > 1


class Xl:
    """
    Xl provides an easy interface to the Excel file from which this code is being called
    
    Parameters
    ----------
    fullname : string, default None
        For debugging/interactive use from within Python, provide the fully qualified filename.
        Leave empty if called from Excel.

    """

    def __init__(self, fullname=None):
        # TODO: check if fullname exists
        if fullname:
            self.fullname = fullname.lower()
        else:
            # TODO: catch AttributeError in case called from Python without argument
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
        
    def get_cell(self, sheet, row, col):
        """ Get value of one cell """
        sht = self.Workbook.Worksheets(sheet)
        cell = sht.Cells(row, col).Value
        if type(cell) is TimeType:
            return clean_com_data([[cell]])[0][0]  # TODO: introduce as_array method?
        return cell

    def get_range(self, sheet, row1, col1, row2, col2):
        """ Returns a list of lists """
        if row1 == row2 and col1 == col2:
            return self.get_cell(sheet, row1, row1)
        else:
            sht = self.Workbook.Worksheets(sheet)
            data = sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value
            return clean_com_data(data)

    def get_contiguous_range(self, sheet, row, col):
        # TODO: shortcut/option to ignore "" cells with xlup/xlright or CurrentRegion
        """
        Tracks down and across from top left cell until it
        encounters blank cells; returns the non-blank range.
        Looks at first row and column; blanks at bottom or right
        are OK and return None within the array

        """
        sht = self.Workbook.Worksheets(sheet)

        # Find the bottom row
        bottom = row
        while sht.Cells(bottom + 1, col).Value not in [None, ""]:
            bottom += 1

        # Right column
        right = col
        while sht.Cells(row, right + 1).Value not in [None, ""]:
            right += 1

        data = sht.Range(sht.Cells(row, col), sht.Cells(bottom, right)).Value
        return clean_com_data(data)

    def get_current_range(self, sheet, row, col):
        """
        Equivalent to CurrentRange in Excel: Takes all surrounding cells into account

        """
        data = self.Workbook.Worksheets(sheet).Cells(row, col).CurrentRegion.Value
        data = [list(row) for row in data]
        return clean_com_data(data)

    def set_cell(self, sheet, row, col, value):
        """ Set value of one cell """
        sht = self.Workbook.Worksheets(sheet)
        sht.Cells(row, col).Value = value

    def set_range(self, sheet, top_row, left_col, data):
        """
        Insert a 2d array starting at given location.
        Works out the size needed for itself
        
        """
        bottom_row = top_row + len(data) - 1
        right_col = left_col + len(data[0]) - 1
        sht = self.Workbook.Worksheets(sheet)
        if type(data) is np.ndarray:
            data = data.tolist()  # Python 3 cant handle arrays directly
        sht.Range(sht.Cells(top_row, left_col), sht.Cells(bottom_row, right_col)).Value = data

    def set_dataframe(self, sheet, top_row, left_col, dataframe, index=True, header=True):
        """
        Writes out a Pandas DataFrame

        Parameters
        ----------
        TODO:

        """
        if index:
            dataframe = dataframe.reset_index()

        if header:
            if type(dataframe.columns) is MultiIndex:
                columns = zip(*dataframe.columns.tolist())
            else:
                columns = [dataframe.columns.tolist()]
            self.set_range(sheet, top_row, left_col, columns)
            top_row += len(columns)

        self.set_range(sheet, top_row, left_col, dataframe.values)


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


class Cell(object):
    """
    A Cell object can be created with the following arguments:

    Cell('A1')
    Cell('NamedRange')
    Cell(1,1)

    If no worksheet name is provided as first argument, it will take the cells from the active sheet. To get
    the cells from a specific sheet, provide the worksheet name as first argument like so:

    Cell('Sheet1','A1')
    """
    def __init__(self, *args):
        # Parse arguments
        if len(args) == 1:
            sheet = None
            self.range = args[0]
        if len(args) == 2 and type(args[0]) is str:
            sheet = args[0]
            self.range = args[1]
        if len(args) == 2 and type(args[0]) is int:
            sheet = None
            self.range = None
            self.row = args[0]
            self.col = args[1]
        if len(args) == 3:
            sheet = args[0]
            self.range = None
            self.row = args[1]
            self.col = args[2]

        # Get cell
        if sheet:
            self.sheet = Workbook.Worksheets(sheet)
        else:
            self.sheet = Workbook.ActiveSheet
        if self.range:
            self.cell = self.sheet.Range(self.range)
            self.row = self.cell.Row
            self.col = self.cell.Column
        else:
            self.cell = self.sheet.Cells(self.row, self.col)

    @property
    def value(self):
        if type(self.cell.Value) is TimeType:
            return clean_com_data([[self.cell.Value]])[0][0]  # TODO: introduce as_matrix method?
        return self.cell.Value

    @value.setter
    def value(self, data):
        self.cell.Value = data

    def clear(self):
        self.cell.Clear()

    def clear_contents(self):
        self.cell.ClearContents()

def get_table(sheet, row, col):
    """
    Returns  down_row and right_col from table starting at Cell(row, col)
    """
    bottom = row
    while sheet.Cells(bottom + 1, col).Value not in [None, ""]:
        bottom += 1

    # Right column
    right = col
    while sheet.Cells(row, right + 1).Value not in [None, ""]:
        right += 1

    return bottom, right


class CellRange(object):
    """
    A CellRange object can be created with the following arguments:

    CellRange('A1', table=True)
    CellRange('A1:C3')
    CellRange('NamedRange')
    CellRange((1,1), (3,3))

    If no worksheet name is provided as first argument, it will take the range from the active sheet. To get
    the range from a specific sheet, provide the worksheet name as first argument like so:

    CellRange('Sheet1','A1')
    """
    def __init__(self, *args, **kwargs):
        # Parse arguments
        self.table = kwargs.get('table')
        if len(args) == 1:
            sheet = None
            cell_range = args[0]
        if len(args) == 2 and type(args[0]) is str:
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
            elif self.table:
                self.row2, self.col2 = get_table(self.sheet, self.row1, self.col1)
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
        bottom_row = self.row1 + len(data) - 1
        right_col = self.col1 + len(data[0]) - 1
        if type(data) is np.ndarray:
            data = data.tolist()  # Python 3 can't handle arrays directly
        self.sheet.Range(self.sheet.Cells(self.row1, self.col1), self.sheet.Cells(bottom_row, right_col)).Value = data

    @property
    def current_region(self):
        """
        The current_region property returns a CellRange object representing a range bounded by (but not including) any
        combination of blank rows and blank columns or the edges of the worksheet
        VBA equivalent: CurrentRegion property of Range object
        """
        current_region = self.sheet.Cells(self.row1, self.col1).CurrentRegion
        self.row2 = self.row1 + current_region.Rows.Count - 1
        self.col2 = self.col1 + current_region.Columns.Count - 1
        return CellRange(self.sheet.Name, (self.row1, self.col1), (self.row2, self.col2))

    def clear(self):
        self.cell_range.Clear()

    def clear_contents(self):
        self.cell_range.ClearContents()

if __name__ == "__main__":
    xlwings_connect(r'C:\DEV\Git\xlwings\example.xlsm')
    # Cell
    # print Cell('B1').value
    # print Cell('Sheet3', 'A1').value
    # Cell('G2').value = 23
    #
    # # CellRange
    # print CellRange('A1').value
    # print np.array(CellRange('A1:C3').value)
    # print np.array(CellRange('test_range').value)
    # print np.array(CellRange((1,1), (3,3)).value)
    #
    # print CellRange('Sheet3', 'A1').value
    # print np.array(CellRange('Sheet3', 'A1:C3').value)
    # print np.array(CellRange('Sheet3', 'test_range').value)
    # print np.array(CellRange('Sheet3', (1,1), (3,3)).value)
    #
    # CellRange('A25').value = [[23]]
    # CellRange('A1:C3').value = [[11,22,33], [44,55,66], [77,88,99]]
    # CellRange('single_cell').value = np.eye(4)
    # CellRange((5,5), (8,8)).value = [[1,2,3], [4,5,6], [7,8,9]]
    #
    # print np.array(CellRange('Sheet3', 'G23', table=True).value)
    # print np.array(CellRange('Sheet3', 'G23', current_region=True).value)
    print np.array(CellRange('Sheet3', 'G23').current_region.value)