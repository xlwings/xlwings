import datetime as dt
import win32api
import pywintypes
import pythoncom

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