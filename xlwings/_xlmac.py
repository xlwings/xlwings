import os
import datetime as dt
from appscript import app
from appscript import k as kw
import psutil


# Time types: pywintypes.timetype doesn't work on Python 3
time_types = (dt.date, dt.datetime)


def is_file_open(fullname):
    """
    Checks if the file is already open
    """
    for proc in psutil.process_iter():
        if proc.name() == 'Microsoft Excel':
            for i in proc.get_open_files():
                if i.path.lower() == fullname.lower():
                    return True
            else:
                return False


def is_excel_running():
    for proc in psutil.process_iter():
        if proc.name() == 'Microsoft Excel':
            return True
    return False


def get_xl_workbook(fullname):
    """
    Get the appscript Workbook object.
    On Mac, it seems that we don't have to deal with >1 instances of Excel,
    as each spreadsheet opens in a separate window anyway.
    """
    filename = os.path.basename(fullname)
    xl_workbook = app('Microsoft Excel').workbooks[filename]
    xl_app = app('Microsoft Excel')
    return xl_app, xl_workbook


def get_workbook_name(xl_workbook):
    return xl_workbook.name.get()


def get_workbook_index(xl_workbook):
    return xl_workbook.entry_index.get()


def open_xl_workbook(fullname):
    filename = os.path.basename(fullname)
    xl_app = app('Microsoft Excel')
    xl_app.open(fullname)
    xl_workbook = xl_app.workbooks[filename]
    return xl_app, xl_workbook


def new_xl_workbook():
    """

    """
    is_running = is_excel_running()

    xl_app = app('Microsoft Excel')
    xl_app.activate()

    if is_running:
        # If Excel is being fired up, a "Workbook1" is automatically added
        # If its already running, we create an new one that is called "Sheet1".
        # That's a feature: See p.14 on Excel 2004 AppleScript Reference
        xl_workbook = xl_app.make(new=kw.workbook)
    else:
        xl_workbook = xl_app.active_workbook

    return xl_app, xl_workbook


def get_active_sheet(xl_workbook):
    return xl_workbook.active_sheet