import sys
import os
import inspect
from win32com.client import GetObject
from pywintypes import UnicodeType, TimeType


class Xl:
    """For debugging, provide the full filepath, otherwise leave blank
    
    
    """

    def __init__(self, filepath=None):
        if filepath:
            # GetObject() gives us the correct Excel instance if there are > 1
            self.Workbook = GetObject(filepath)
        else:
            filename = sys.argv[1]
            # Get filepath of calling function in case this module is somewhere else
            _dirpath = os.path.dirname(inspect.getmodule(inspect.stack()[1][0]).__file__)
            filepath = os.path.abspath(os.path.join(_dirpath, filename))
            self.Workbook = GetObject(filepath)

    def save(self, newfilename=None):
        if newfilename:
            self.filename = newfilename
            self.Workbook.SaveAs(newfilename)
        else:
            self.Workbook.Save()
        
    def get_cell(self, sheet, row, col):
        """Get value of one cell"""
        sht = self.Workbook.Worksheets(sheet)
        return sht.Cells(row, col).Value

    def set_cell(self, sheet, row, col, value):
        """set value of one cell"""
        sht = self.Workbook.Worksheets(sheet)
        sht.Cells(row, col).Value = value

    def get_range(self, sheet, row1, col1, row2, col2):
        """return a 2d array (i.e. tuple of tuples)"""
        sht = self.Workbook.Worksheets(sheet)
        return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value
        
    def set_range(self, sheet, left_col, top_row, data):
        """insert a 2d array starting at given location.
        Works out the size needed for itself
        
        """
    
        bottom_row = top_row + len(data) - 1
        right_col = left_col + len(data[0]) - 1
        sht = self.Workbook.Worksheets(sheet)
        sht.Range(sht.Cells(top_row, left_col), sht.Cells(bottom_row, right_col)).Value = data

    def get_contiguous_range(self, sheet, row, col):
        """Tracks down and across from top left cell until it
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
    
        return sht.Range(sht.Cells(row, col), sht.Cells(bottom, right)).Value
        
    def fix_strings_and_dates(self, aMatrix):
        """converts all unicode strings and times"""
        newmatrix = []
        for row in aMatrix:
            newrow = []
            for cell in row:
                if type(cell) is UnicodeType:
                    newrow.append(str(cell))
                elif type(cell) is TimeType:
                    newrow.append(int(cell))
                else:
                    newrow.append(cell)
            newmatrix.append(tuple(newrow))
        return newmatrix