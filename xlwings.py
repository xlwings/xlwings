"""
xlwings makes it easy to deploy your Python powered Excel tools on Windows.
Homepage and documentation: http://xlwings.org/

Copyright (c) 2013, Felix Zumstein.
License: MIT (see LICENSE.txt for details)

This module partly based on the easyExcel class as shown in the book
"Python Programming on Win32". It can be downloaded from
http://starship.python.net/crew/mhammond/ppw32
Copyright (c) 2000, Mark Hammond and Andy Robinson
"""

import sys
import os
from win32com.client import GetObject
import adodbapi
from pywintypes import TimeType
import numpy as np
from pandas import MultiIndex

__version__ = '0.1-dev'
__license__ = 'MIT'

_is_python3 = sys.version_info.major > 2

class Xl:
    """
    Xl provides an easy interface to the Excel file from which this code is being called
    
    Parameters
    ----------
    filepath : string, default None
        For debugging/interactive use from within Python, provide the full filepath. Leave empty if called from Excel.

    """

    def __init__(self, fullname=None):
        if fullname:
            # GetObject() gives us the correct Excel instance if there are > 1
            self.Workbook = GetObject(fullname)
            self.fullname = fullname
            self.filename = os.path.split(self.fullname)[1]
        else:
            # TODO: catch AttributeError in case called from Python without argument
            self.fullname = sys.argv[1]
            self.Workbook = GetObject(self.fullname)
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
            return self.clean_data([[cell]])[0][0]
        return cell

    def get_range(self, sheet, row1, col1, row2, col2):
        """ Returns a list of lists """
        if row1 == row2 and col1 == col2:
            return self.get_cell(sheet, row1, row1)
        else:
            sht = self.Workbook.Worksheets(sheet)
            data = sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value
            return self.clean_data(data)

    def get_contiguous_range(self, sheet, row, col):
        # TODO: shortcut/option to ignore "" cells with xlup/xlright or CurrentRegion
        """
        Tracks down and across from top left cell until it
        encounters blank cells; returns the non-blank range.
        Looks at first row and column; blanks at bottom or right
        are OK and return None within the array

        """
        sht = self.Workbook.Worksheets(sheet)

        # find the bottom row
        bottom = row
        while sht.Cells(bottom + 1, col).Value not in [None, ""]:
            bottom += 1

        # right column
        right = col
        while sht.Cells(row, right + 1).Value not in [None, ""]:
            right += 1

        data = sht.Range(sht.Cells(row, col), sht.Cells(bottom, right)).Value
        return self.clean_data(data)

    def get_current_range(self, sheet, row, col):
        """
        Equivalent to CurrentRange in Excel: Takes all surrounding cells into account

        """
        data = self.Workbook.Worksheets(sheet).Cells(row, col).CurrentRegion.Value
        data = [list(row) for row in data]
        return self.clean_data(data)

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

    @staticmethod
    def clean_data(data):
        """
        Transforms data from tuples of tuples into list of list and
        on Python 2, transforms PyTime Objects from COM into datetime objects.

        Parameters
        ----------
        data : raw data as returned from Excel (tuple of tuple)

        """
        # turn into list of list for easier handling (e.g. for Pandas DataFrame)
        data = [list(row) for row in data]

        # Check which columns contain COM dates
        # TODO: simplify
        if _is_python3 is True:
            return data
        else:
            cols_with_dates = []
            for i in range(len(data[0])):
                if any([type(row[i]) is TimeType for row in data]):
                    cols_with_dates.append(i)

            # Transform PyTime into datetime
            tc = adodbapi.pythonDateTimeConverter()
            for i in cols_with_dates:
                for j, cell in enumerate([row[i] for row in data]):
                    if type(cell) is TimeType:
                        data[j][i] = tc.DateObjectFromCOMDate(cell)
            return data
