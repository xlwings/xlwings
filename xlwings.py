"""
xlwings is an easy way to connect your Excel tools with Python (Windows only).
The aim is to make it as easy as possible to distribute the Excel files.

Homepage and documentation: http://xlwings.org/

Copyright (c) 2013, Felix Zumstein.
License: MIT (see LICENSE.txt for details)

This module is largely based on the EasyExcel class as described in the book
"Python Programming on Win32". It can be downloadable from
http://starship.python.net/crew/mhammond/ppw32
Copyright (c) 2000, Mark Hammond and Andy Robinson
"""

import sys
import os
import inspect
from win32com.client import GetObject
import adodbapi
from pywintypes import UnicodeType, TimeType
from pandas import DataFrame
import pandas as pd

__version__ = '0.1-dev'
__license__ = 'MIT'

class Xl:
    """
    Xl provides an easy interface to the Excel file from which this code is being called
    
    Parameters
    ----------
    filepath : string, default None
        For debugging/interactive use from within Python, provide the full filepath. Leave empty if called from Excel.
    """

    def __init__(self, filepath=None):
        if filepath:
            # GetObject() gives us the correct Excel instance if there are > 1
            self.Workbook = GetObject(filepath)
            self.filepath = filepath
        else:
            filename = sys.argv[1]
            # Get filepath of calling function in case this module is somewhere else
            _dirpath = os.path.dirname(inspect.getmodule(inspect.stack()[1][0]).__file__)
            self.filepath = os.path.abspath(os.path.join(_dirpath, filename))
            self.Workbook = GetObject(self.filepath)

    def save(self, newfilename=None):
        if newfilename:
            self.filename = newfilename
            self.Workbook.SaveAs(newfilename)
        else:
            self.Workbook.Save()
        
    def get_cell(self, sheet, row, col):
        """Get value of one cell"""
        sht = self.Workbook.Worksheets(sheet)
        cell = sht.Cells(row, col).Value
        if type(cell) is TimeType:
            tc = adodbapi.pythonDateTimeConverter()
            return tc.DateObjectFromCOMDate(cell)
        else:
            return cell

    def set_cell(self, sheet, row, col, value):
        """set value of one cell"""
        sht = self.Workbook.Worksheets(sheet)
        sht.Cells(row, col).Value = value

    def get_range(self, sheet, row1, col1, row2, col2):
        """return a 2d array (i.e. tuple of tuples)"""
        sht = self.Workbook.Worksheets(sheet)
        return list(sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value)
        
    def set_range(self, sheet, top_row, left_col, data):
        """
        Insert a 2d array starting at given location.
        Works out the size needed for itself
        
        """
    
        bottom_row = top_row + len(data) - 1
        right_col = left_col + len(data[0]) - 1
        sht = self.Workbook.Worksheets(sheet)
        sht.Range(sht.Cells(top_row, left_col), sht.Cells(bottom_row, right_col)).Value = data

    def get_contiguous_range(self, sheet, row, col):
        # TODO: shortcut/option to ignore "" cells with xlup/xlright
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
    
        return list(sht.Range(sht.Cells(row, col), sht.Cells(bottom, right)).Value)

    def fix_dates(self, matrix):
        # TODO: make it work for single cols or rows
        """converts all unicode strings and times"""
        tc = adodbapi.pythonDateTimeConverter()
        newmatrix = []
        for row in matrix:
            newrow = []
            for cell in row:
                if type(cell) is TimeType:
                    newrow.append(tc.DateObjectFromCOMDate(cell))
                else:
                    newrow.append(cell)
            newmatrix.append(tuple(newrow))
        return newmatrix
        
    def create_dataframe(self, data):
        """Create a pandas DataFrame"""
        
        data_fixed = self.fix_dates(data)
        df = DataFrame(data_fixed[1:], columns=data_fixed[0])
        df['test 1'] = pd.to_datetime(df['test 1'])
        return df
        
        
        
        
        
        
        
        
        
        
        
        
        