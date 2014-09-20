# TODO: create classes
# TODO: align clean_xl_data and prepare_xl_data (should work on same dimensions of data)

import os
import datetime as dt
from appscript import app, reference
from appscript import k as kw
import psutil
import atexit
try:
    import pandas as pd
except ImportError:
    pd = None

# Time types
time_types = (dt.date, dt.datetime)


def clean_up():
    """
    Since AppleScript cannot access Excel while a Macro is running, we have to run the Python call in a
    background process which makes the call return immediately: we rely on the StatusBar to give the user
    feedback.
    This function is triggered when the interpreter exits and runs the CleanUp Macro in VBA to show any
    errors and to reset the StatusBar.
    """
    if is_excel_running():
        app('Microsoft Excel').run_VB_macro('CleanUp')

atexit.register(clean_up)


def is_file_open(fullname):
    """
    Checks if the file is already open
    """
    for proc in psutil.process_iter():
        if proc.name() == 'Microsoft Excel':
            for i in proc.open_files():
                if i.path.lower() == fullname.lower():
                    return True
    return False


def is_excel_running():
    for proc in psutil.process_iter():
        if proc.name() == 'Microsoft Excel':
            return True
    return False


def get_workbook(fullname):
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


def get_worksheet_name(xl_sheet):
    return xl_sheet.name.get()


def get_xl_sheet(xl_workbook, sheet_name_or_index):
    return xl_workbook.sheets[sheet_name_or_index]


def set_worksheet_name(xl_sheet, value):
    return xl_sheet.name.set(value)


def get_worksheet_index(xl_sheet):
    return xl_sheet.entry_index.get()


def get_app(xl_workbook):
    # Since we can't have multiple instances of Excel on Mac (?), we ignore xl_workbook
    return app('Microsoft Excel')


def open_workbook(fullname):
    filename = os.path.basename(fullname)
    xl_app = app('Microsoft Excel')
    xl_app.activate()
    xl_app.open(fullname)
    xl_workbook = xl_app.workbooks[filename]
    return xl_app, xl_workbook


def close_workbook(xl_workbook):
    xl_workbook.close(saving=kw.no)


def new_workbook():
    is_running = is_excel_running()

    xl_app = app('Microsoft Excel')
    xl_app.activate()

    if is_running:
        # If Excel is being fired up, a "Workbook1" is automatically added
        # If its already running, we create an new one that Excel unfortunately calls "Sheet1".
        # It's a feature though: See p.14 on Excel 2004 AppleScript Reference
        xl_workbook = xl_app.make(new=kw.workbook)
    else:
        xl_workbook = xl_app.workbooks[1]

    return xl_app, xl_workbook


def get_active_sheet(xl_workbook):
    return xl_workbook.active_sheet


def activate_sheet(xl_workbook, sheet_name_or_index):
    return xl_workbook.sheets[sheet_name_or_index].activate_object()


def get_worksheet(xl_workbook, sheet_name_or_index):
    return xl_workbook.sheets[sheet_name_or_index]


def get_first_row(xl_sheet, range_address):
    return xl_sheet.cells[range_address].first_row_index.get()


def get_first_column(xl_sheet, range_address):
    return xl_sheet.cells[range_address].first_column_index.get()


def count_rows(xl_sheet, range_address):
    return xl_sheet.cells[range_address].count(each=kw.row)


def count_columns(xl_sheet, range_address):
    return xl_sheet.cells[range_address].count(each=kw.column)


def get_range_from_indices(xl_sheet, first_row, first_column, last_row, last_column):
    first_address = xl_sheet.columns[first_column].rows[first_row].get_address()
    last_address = xl_sheet.columns[last_column].rows[last_row].get_address()
    return xl_sheet.cells['{0}:{1}'.format(first_address, last_address)]


def get_value_from_range(xl_range):
    return xl_range.value.get()


def get_value_from_index(xl_sheet, row_index, column_index):
    return xl_sheet.columns[column_index].rows[row_index].value.get()


def clean_xl_data(data):
    """
    appscript returns empty cells as ''. So we replace those with None to be in line with pywin32
    """
    return [[None if c == '' else c for c in row] for row in data]


def prepare_xl_data(data):
    # This transformation seems to be only needed on Python 2.6 (?)
    if hasattr(pd, 'tslib') and isinstance(data, pd.tslib.Timestamp):
        data = data.to_datetime()
    return data


def set_value(xl_range, data):
    xl_range.value.set(data)


def get_selection_address(xl_app):
    return str(xl_app.selection.get_address())


def clear_contents_worksheet(xl_workbook, sheets_name_or_index):
    xl_workbook.sheets[sheets_name_or_index].used_range.clear_contents()


def clear_worksheet(xl_workbook, sheet_name_or_index):
    xl_workbook.sheets[sheet_name_or_index].used_range.clear_range()


def clear_contents_range(xl_range):
    xl_range.clear_contents()


def clear_range(xl_range):
    xl_range.clear_range()


def get_formula(xl_range):
    return xl_range.formula.get()


def set_formula(xl_range, value):
    xl_range.formula.set(value)


def get_row_index_end_down(xl_sheet, row_index, column_index):
    ix = xl_sheet.columns[column_index].rows[row_index].get_end(direction=kw.toward_the_bottom).first_row_index.get()
    return ix


def get_column_index_end_right(xl_sheet, row_index, column_index):
    ix = xl_sheet.columns[column_index].rows[row_index].get_end(direction=kw.toward_the_right).first_column_index.get()
    return ix


def get_current_region_address(xl_sheet, row_index, column_index):
    return str(xl_sheet.columns[column_index].rows[row_index].current_region.get_address())


def get_chart_object(xl_workbook, sheet_name_or_index, chart_name_or_index):
    return xl_workbook.sheets[sheet_name_or_index].chart_objects[chart_name_or_index]


def get_chart_index(xl_chart):
    return xl_chart.entry_index.get()


def get_chart_name(xl_chart):
    return xl_chart.name.get()


def add_chart(xl_workbook, sheet_name_or_index, left, top, width, height):
    # With the sheet name it won't find the chart later, so we go with the index (no idea why)
    sheet_index = xl_workbook.sheets[sheet_name_or_index].entry_index.get()
    return xl_workbook.make(at=xl_workbook.sheets[sheet_index],
                            new=kw.chart_object,
                            with_properties={kw.width: width,
                                             kw.top: top,
                                             kw.left_position: left,
                                             kw.height: height})


def set_chart_name(xl_chart, name):
    xl_chart.name.set(name)


def set_source_data_chart(xl_chart, xl_range):
    xl_chart.chart.set_source_data(source=xl_range)


def get_chart_type(xl_chart):
    return xl_chart.chart.chart_type.get()


def set_chart_type(xl_chart, chart_type):
    xl_chart.chart.chart_type.set(chart_type)


def activate_chart(xl_chart):
    """
    activate() doesn't seem to do anything so resolving to select() for now
    """
    xl_chart.select()


def autofit(range_, axis):
    if (axis == 0 or axis == 'rows' or axis == 'r') and not range_.is_column():
        range_.xl_range.rows.autofit()
    elif (axis == 1 or axis == 'columns' or axis == 'c') and not range_.is_row():
        range_.xl_range.columns.autofit()
    elif axis is None:
        if not range_.is_row():
            range_.xl_range.columns.autofit()
        if not range_.is_column():
            range_.xl_range.rows.autofit()


def set_xl_workbook_latest(xl_workbook):
    global xl_workbook_latest
    xl_workbook_latest = xl_workbook


def get_xl_workbook_latest():
    try:
        return xl_workbook_latest
    except NameError:
        return False