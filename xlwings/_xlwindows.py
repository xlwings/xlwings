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