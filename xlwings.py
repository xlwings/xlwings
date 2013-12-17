import os
import sys
from win32com.client import GetObject
import win32com
from pywintypes import UnicodeType, TimeType
import logging
from datetime import datetime


class XlWings:
    """TODO: Description """

    def __init__(self):
        logging.info('{0} - start class __init__'.format(datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
#        filename = sys.argv[1]
#        #TODO: provide filepath of calling function in case installed in python dir
#        _dirpath = os.path.dirname(os.path.abspath(__file__))
#        _filepath = r'{0}\{1}'.format(_dirpath, filename)
#        self.xl_app = GetObject(_filepath)
        self.xl_app = win32com.client.Dispatch("Excel.Application")
        logging.info('{0} - end class __init__'.format(datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
        
    def save(self, newfilename=None):
        if newfilename:
            self.filename = newfilename
            self.xlBook.SaveAs(newfilename)
        else:
            self.xlBook.Save()

    def close(self):
        self.xlBook.Close(SaveChanges=0)
        del self.xl_app
        
    def getCell(self, sheet, row, col):
        "Get value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        return sht.Cells(row, col).Value

    def setCell(self, sheet, row, col, value):
        "set value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        sht.Cells(row, col).Value = value

    def getRange(self, sheet, row1, col1, row2, col2):
        "return a 2d array (i.e. tuple of tuples)"
        sht = self.xlBook.Worksheet(sheet)
        return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value
        
    def setRange(self, sheet, leftCol, topRow, data):
        """insert a 2d array starting at given location.
        Works out the size needed for itself"""
    
        bottomRow = topRow + len(data) - 1
    
        rightCol = leftCol + len(data[0]) - 1
        sht = self.xlBook.Worksheets(sheet)
        sht.Range(
            sht.Cells(topRow, leftCol),
            sht.Cells(bottomRow, rightCol)
            ).Value = data
            
    def getContiguousRange(self, sheet, row, col):
        """Tracks down and across from top left cell until it
        encounters blank cells; returns the non-blank range.
        Looks at first row and column; blanks at bottom or right
        are OK and return None within the array"""
    
        sht = self.xlBook.Worksheets(sheet)
    
        # find the bottom row
        bottom = row
        while sht.Cells(bottom + 1, col).Value not in [None, ""]:
            bottom = bottom + 1
    
        # right column
        right = col
        while sht.Cells(row, right + 1).Value not in [None, ""]:
            right = right + 1
    
        return sht.Range(sht.Cells(row, col), sht.Cells(bottom, right)).Value
        
    def fixStringsAndDates(self, aMatrix):
        # converts all unicode strings and times
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