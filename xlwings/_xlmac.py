import os
import datetime as dt
from appscript import app
from appscript import k as kw
import psutil


# Time types
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


def get_worksheet_name(xl_sheet):
    return xl_sheet.name.get()


def get_workbook_index(xl_workbook):
    return xl_workbook.entry_index.get()


def open_xl_workbook(fullname):
    filename = os.path.basename(fullname)
    xl_app = app('Microsoft Excel')
    xl_app.open(fullname)
    xl_workbook = xl_app.workbooks[filename]
    return xl_app, xl_workbook


def close_workbook(xl_workbook):
    xl_workbook.close(saving=kw.no)


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


def activate_sheet(xl_workbook, sheet):
    return xl_workbook.sheets[sheet].activate_object()


def get_worksheet(xl_workbook, sheet):
    return xl_workbook.sheets[sheet]


def get_first_row(xl_sheet, cell_range):
    return xl_sheet.cells[cell_range].first_row_index.get()


def get_first_column(xl_sheet, cell_range):
    return xl_sheet.cells[cell_range].first_column_index.get()


def count_rows(xl_sheet, cell_range):
    return xl_sheet.cells[cell_range].count(each=kw.row)


def count_columns(xl_sheet, cell_range):
    return xl_sheet.cells[cell_range].count(each=kw.column)


def get_range_from_indices(xl_sheet, first_row, first_column, last_row, last_column):
    first_address = xl_sheet.columns[first_column].rows[first_row].get_address()
    last_address = xl_sheet.columns[last_column].rows[last_row].get_address()
    return xl_sheet.cells['{0}:{1}'.format(first_address, last_address)]


def get_value_from_range(xl_range):
    return xl_range.value.get()


def get_value_from_index(xl_sheet, row_index, column_index):
    return xl_sheet.columns[column_index].rows[row_index].value.get()


def clean_xl_data(data):
    return [[None if c == '' else c for c in row] for row in data]


def prepare_xl_data(data):
    return data


def set_value(xl_range, data):
    xl_range.value.set(data)


def get_selection_address(xl_app):
    return str(xl_app.selection.get_address())


def clear_contents_worksheet(xl_workbook, sheet):
    xl_workbook.sheets[sheet].used_range.clear_contents()


def clear_worksheet(xl_workbook, sheet):
    xl_workbook.sheets[sheet].used_range.clear_range()


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


def get_chart_object(xl_workbook, sheet, name_or_index):
    return xl_workbook.sheets[sheet].chart_objects[name_or_index]


def get_chart_index(xl_chart):
    return xl_chart.entry_index.get()


def get_chart_name(xl_chart):
    return xl_chart.name.get()


def add_chart(xl_workbook, sheet, left, top, width, height):
    return xl_workbook.make(at=xl_workbook.sheets[sheet],
                            new=kw.chart_object,
                            with_properties={kw.width: width, kw.top: top, kw.left_position: left, kw.height: height})


def set_chart_name(xl_chart, name):
    xl_chart.name.set(name)


def set_source_data_chart(xl_chart, xl_range):
    xl_chart.chart.set_source_data(source=xl_range)


def get_chart_type(xl_chart):
    return xl_chart.chart.chart_type.get()


def set_chart_type(xl_chart, chart_type):
    xl_chart.chart.chart_type.set(chart_type)