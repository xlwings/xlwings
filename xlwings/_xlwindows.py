import datetime as dt
import win32api
import pywintypes
import pythoncom
from win32com.client import GetObject, dynamic

# Time types: pywintypes.timetype doesn't work on Python 3
time_types = (dt.date, dt.datetime, type(pywintypes.Time(0)))


def is_file_open(fullname):
    """
    Checks the Running Object Table (ROT) for the fully qualified filename
    """
    context = pythoncom.CreateBindCtx(0)
    for moniker in pythoncom.GetRunningObjectTable():
        name = moniker.GetDisplayName(context, None)
        if name.lower() == fullname.lower():
            return True
    return False


def get_xl_workbook(fullname):
    """
    Returns the COM Application and Workbook objects of an open Workbook.
    GetObject() returns the correct Excel instance if there are > 1
    """
    xl_workbook = GetObject(fullname)
    xl_app = xl_workbook.Application
    return xl_app, xl_workbook


def get_workbook_name(xl_workbook):
    return xl_workbook.Name


def get_workbook_index(xl_workbook):
    return xl_workbook.Index


def open_xl_workbook(fullname):
    """

    """
    xl_app = dynamic.Dispatch('Excel.Application')
    xl_workbook = xl_app.Workbooks.Open(fullname)
    xl_app.Visible = True
    return xl_app, xl_workbook


def new_xl_workbook():
    xl_app = dynamic.Dispatch('Excel.Application')
    xl_app.Visible = True
    xl_workbook = xl_app.Workbooks.Add()
    return xl_app, xl_workbook


def get_active_sheet(xl_workbook):
    return xl_workbook.ActiveSheet


def get_worksheet(xl_workbook, sheet):
    return xl_workbook.Sheets[sheet]


def get_first_row(xl_sheet, cell_range):
    return xl_sheet.Range(cell_range).Row


def get_first_column(xl_sheet, cell_range):
    return xl_sheet.Range(cell_range).Column


def count_rows(xl_sheet, cell_range):
    return xl_sheet.Range(cell_range).Rows.Count


def count_columns(xl_sheet, cell_range):
    return xl_sheet.Range(cell_range).Columns.Count


def get_range_from_indices(xl_sheet, first_row, first_column, last_row, last_column):
    return xl_sheet.Range(xl_sheet.Cells(first_row, first_column),
                          xl_sheet.Cells(last_row, last_column))