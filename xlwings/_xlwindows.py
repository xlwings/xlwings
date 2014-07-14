import datetime as dt
import win32api
import pywintypes
import pythoncom
from win32com.client import GetObject

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
    Get the COM Workbook object.
    GetObject() returns the correct Excel instance if there are > 1
    """
    #
    return GetObject(fullname)


def get_xl_application(xl_workbook):
    """
    Get the COM Application object.
    """
    return xl_workbook.Application