import win32api
import win32con
import xlwings as xw


@xw.event('WorkbookOpen')
def on_workbook_open(xl_workbook):
    win32api.MessageBox(
        xw.apps.active.api.Hwnd, 'All your workbooks are belong to us', 'Info', win32con.MB_ICONINFORMATION)


@xw.event('SheetActivate')
def on_sheet_activate(xl_sheet):
    win32api.MessageBox(
        xw.apps.active.api.Hwnd, 'All your sheets are belong to us', 'Info', win32con.MB_ICONINFORMATION)


@xw.event('SheetSelectionChange')
def on_sheet_selection_change(xl_sheet, xl_range):
    xl_range.Value = 'X'


@xw.event('SheetChange')
def on_sheet_change(xl_sheet, xl_range):
    xl_range.Value += 1
