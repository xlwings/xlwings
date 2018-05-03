import win32api
import win32con
import xlwings as xw


@xw.event('WorkbookOpen')
def on_workbook_open(book):
    win32api.MessageBox(
        xw.apps.active.api.Hwnd, 'All your workbooks are belong to us', 'Info', win32con.MB_ICONINFORMATION)


@xw.event('SheetActivate')
def on_sheet_activate(sheet):
    win32api.MessageBox(
        xw.apps.active.api.Hwnd, 'All your sheets are belong to us', 'Info', win32con.MB_ICONINFORMATION)


@xw.event('SheetSelectionChange')
def on_sheet_selection_change(sheet, rng):
    rng.value = 'X'


@xw.event('SheetChange')
def on_sheet_change(sheet, rng):
    rng.value += 1

