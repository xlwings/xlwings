import os
import datetime as dt
import subprocess
import unicodedata
from appscript import app, mactypes
from appscript import k as kw
from appscript.reference import CommandError
import psutil
import atexit
from .constants import ColorIndex, Calculation
from .utils import int_to_rgb, np_datetime_to_datetime
from . import mac_dict, PY3
try:
    import pandas as pd
except ImportError:
    pd = None
try:
    import numpy as np
except ImportError:
    np = None

# Time types
time_types = (dt.date, dt.datetime)
if hasattr(np, 'datetime64'):
    time_types = time_types + (np.datetime64,)

# We're only dealing with one instance of Excel on Mac
_xl_app = None

def set_xl_app(app_target=None):
    if app_target is None:
        app_target = 'Microsoft Excel'
    global _xl_app
    _xl_app = app(app_target, terms=mac_dict)


@atexit.register
def clean_up():
    """
    Since AppleScript cannot access Excel while a Macro is running, we have to run the Python call in a
    background process which makes the call return immediately: we rely on the StatusBar to give the user
    feedback.
    This function is triggered when the interpreter exits and runs the CleanUp Macro in VBA to show any
    errors and to reset the StatusBar.
    """
    if is_excel_running():
        # Prevents Excel from reopening if it has been closed manually or never been opened
        try:
            _xl_app.run_VB_macro('CleanUp')
        except (CommandError, AttributeError):
            # Excel files initiated from Python don't have the xlwings VBA module
            pass


def is_file_open(fullname):
    """
    Checks if the file is already open
    """
    for proc in psutil.process_iter():
        if proc.name() == 'Microsoft Excel':
            for i in proc.open_files():
                path = i.path
                if PY3:
                    if path.lower() == fullname.lower():
                        return True
                else:
                    if isinstance(path, str):
                        path = unicode(path, 'utf-8')
                        # Mac saves unicode data in decomposed form, e.g. an e with accent is stored as 2 code points
                        path = unicodedata.normalize('NFKC', path)
                    if isinstance(fullname, str):
                        fullname = unicode(fullname, 'utf-8')
                    if path.lower() == fullname.lower():
                        return True
    return False


def is_excel_running():
    for proc in psutil.process_iter():
        if proc.name() == 'Microsoft Excel':
            return True
    return False


def get_open_workbook(fullname, app_target=None):
    """
    Get the appscript Workbook object.
    On Mac, there's only ever one instance of Excel.
    """
    filename = os.path.basename(fullname)
    set_xl_app(app_target)
    xl_workbook = _xl_app.workbooks[filename]
    return _xl_app, xl_workbook


def get_active_workbook(app_target=None):
    set_xl_app(app_target)
    return _xl_app.active_workbook


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


def get_app(xl_workbook, app_target=None):
    set_xl_app(app_target)
    return _xl_app


def open_workbook(fullname, app_target=None):
    filename = os.path.basename(fullname)
    set_xl_app(app_target)
    _xl_app.open(fullname)
    xl_workbook = _xl_app.workbooks[filename]
    return _xl_app, xl_workbook


def close_workbook(xl_workbook):
    xl_workbook.close(saving=kw.no)


def new_workbook(app_target=None):
    is_running = is_excel_running()

    set_xl_app(app_target)

    if is_running:
        # If Excel is being fired up, a "Workbook1" is automatically added
        # If its already running, we create an new one that Excel unfortunately calls "Sheet1".
        # It's a feature though: See p.14 on Excel 2004 AppleScript Reference
        xl_workbook = _xl_app.make(new=kw.workbook)
    else:
        xl_workbook = _xl_app.workbooks[1]

    return _xl_app, xl_workbook


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
    Expects a 2d list.
    """
    # appscript returns empty cells as ''. So we replace those with None to be in line with pywin32
    return [[None if c == '' else c for c in row] for row in data]


def prepare_xl_data(data):
    """
    Expects a 2d list.
    """
    if hasattr(np, 'datetime64'):
        # handle numpy.datetime64
        data = [[np_datetime_to_datetime(c) if isinstance(c, np.datetime64) else c for c in row] for row in data]
    if hasattr(pd, 'tslib'):
        # This transformation seems to be only needed on Python 2.6 (?)
        data = [[c.to_datetime() if isinstance(c, pd.tslib.Timestamp) else c for c in row] for row in data]
    # Make datetime timezone naive
    data = [[c.replace(tzinfo=None) if isinstance(c, dt.datetime) else c for c in row] for row in data]
    # appscript packs integers larger than SInt32 but smaller than SInt64 as typeSInt64, and integers
    # larger than SInt64 as typeIEEE64BitFloatingPoint. Excel silently ignores typeSInt64. (GH 227)
    data = [[float(c) if isinstance(c, int) else c for c in row] for row in data]

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
    _xl_app.screen_updating.set(False)
    xl_range.clear_contents()
    _xl_app.screen_updating.set(True)

def clear_range(xl_range):
    _xl_app.screen_updating.set(False)
    xl_range.clear_range()
    _xl_app.screen_updating.set(True)

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


def get_column_width(xl_range):
    return xl_range.column_width.get()


def set_column_width(xl_range, value):
    xl_range.column_width.set(value)


def get_row_height(xl_range):
    return xl_range.row_height.get()


def set_row_height(xl_range, value):
    xl_range.row_height.set(value)


def get_width(xl_range):
    return xl_range.width.get()


def get_height(xl_range):
    return xl_range.height.get()


def autofit(range_, axis):
    address = range_.xl_range.get_address()
    _xl_app.screen_updating.set(False)
    if axis == 'rows' or axis == 'r':
        range_.xl_sheet.rows[address].autofit()
    elif axis == 'columns' or axis == 'c':
        range_.xl_sheet.columns[address].autofit()
    elif axis is None:
        range_.xl_sheet.rows[address].autofit()
        range_.xl_sheet.columns[address].autofit()
    _xl_app.screen_updating.set(True)


def autofit_sheet(sheet, axis):
    #TODO: combine with autofit that works on Range objects
    num_columns = sheet.xl_sheet.count(each=kw.column)
    num_rows = sheet.xl_sheet.count(each=kw.row)
    xl_range = get_range_from_indices(sheet.xl_sheet, 1, 1, num_rows, num_columns)
    address = xl_range.get_address()
    _xl_app.screen_updating.set(False)
    if axis == 'rows' or axis == 'r':
        sheet.xl_sheet.rows[address].autofit()
    elif axis == 'columns' or axis == 'c':
        sheet.xl_sheet.columns[address].autofit()
    elif axis is None:
        sheet.xl_sheet.rows[address].autofit()
        sheet.xl_sheet.columns[address].autofit()
    _xl_app.screen_updating.set(True)


def set_xl_workbook_current(xl_workbook):
    global xl_workbook_current
    xl_workbook_current = xl_workbook


def get_xl_workbook_current():
    try:
        return xl_workbook_current
    except NameError:
        return None


def get_number_format(range_):
    return range_.xl_range.number_format.get()


def set_number_format(range_, value):
    _xl_app.screen_updating.set(False)
    range_.xl_range.number_format.set(value)
    _xl_app.screen_updating.set(True)


def get_address(xl_range, row_absolute, col_absolute, external):
    return xl_range.get_address(row_absolute=row_absolute, column_absolute=col_absolute, external=external)


def add_sheet(xl_workbook, before, after):
    if before:
        position = before.xl_sheet.before
    else:
        position = after.xl_sheet.after
    return xl_workbook.make(new=kw.worksheet, at=position)


def count_worksheets(xl_workbook):
    return xl_workbook.count(each=kw.worksheet)


def get_hyperlink_address(xl_range):
    try:
        return xl_range.hyperlinks[1].address.get()
    except CommandError:
        raise Exception("The cell doesn't seem to contain a hyperlink!")


def set_hyperlink(xl_range, address, text_to_display=None, screen_tip=None):
    xl_range.make(at=xl_range, new=kw.hyperlink, with_properties={kw.address: address,
                                                                  kw.text_to_display: text_to_display,
                                                                  kw.screen_tip: screen_tip})


def set_color(xl_range, color_or_rgb):
    if color_or_rgb is None:
        xl_range.interior_object.color_index.set(ColorIndex.xlColorIndexNone)
    elif isinstance(color_or_rgb, int):
        xl_range.interior_object.color.set(int_to_rgb(color_or_rgb))
    else:
        xl_range.interior_object.color.set(color_or_rgb)


def get_color(xl_range):
    if xl_range.interior_object.color_index.get() == kw.color_index_none:
        return None
    else:
        return tuple(xl_range.interior_object.color.get())


def save_workbook(xl_workbook, path):
    saved_path = xl_workbook.properties().get(kw.path)
    if (saved_path != '') and (path is None):
        # Previously saved: Save under existing name
        xl_workbook.save()
    elif (saved_path == '') and (path is None):
        # Previously unsaved: Save under current name in current working directory
        path = os.path.join(os.getcwd(), xl_workbook.name.get() + '.xlsx')
        dir_name, file_name = os.path.split(path)
        dir_name_hfs = mactypes.Alias(dir_name).hfspath  # turn into HFS path format
        hfs_path = dir_name_hfs + ':' + file_name
        xl_workbook.save_workbook_as(filename=hfs_path, overwrite=True)
    elif path:
        # Save under new name/location
        dir_name, file_name = os.path.split(path)
        dir_name_hfs = mactypes.Alias(dir_name).hfspath  # turn into HFS path format
        hfs_path = dir_name_hfs + ':' + file_name
        xl_workbook.save_workbook_as(filename=hfs_path, overwrite=True)


def open_template(fullpath):
    subprocess.call(['open', fullpath])


def set_visible(xl_app, visible):
    if visible:
        xl_app.activate()
    else:
        app('System Events').processes['Microsoft Excel'].visible.set(visible)


def get_visible(xl_app):
    return app('System Events').processes['Microsoft Excel'].visible.get()


def get_fullname(xl_workbook):
    hfs_path = xl_workbook.properties().get(kw.full_name)
    if hfs_path == xl_workbook.properties().get(kw.name):
        return hfs_path
    url = mactypes.convertpathtourl(hfs_path, 1)  # kCFURLHFSPathStyle = 1
    return mactypes.converturltopath(url, 0)  # kCFURLPOSIXPathStyle = 0


def quit_app(xl_app):
    xl_app.quit(saving=kw.no)


def get_screen_updating(xl_app):
    return xl_app.screen_updating.get()


def set_screen_updating(xl_app, value):
    xl_app.screen_updating.set(value)


# TODO: Hack for Excel 2016, to be refactored
calculation = {kw.calculation_automatic: Calculation.xlCalculationAutomatic,
               kw.calculation_manual: Calculation.xlCalculationManual,
               kw.calculation_semiautomatic: Calculation.xlCalculationSemiautomatic}


def get_calculation(xl_app):
    return calculation[xl_app.calculation.get()]


def set_calculation(xl_app, value):
    calculation_reverse = dict(zip(calculation.values(), calculation.keys()))
    xl_app.calculation.set(calculation_reverse[value])


def calculate(xl_app):
    xl_app.calculate()


def get_named_range(range_):
    return range_.xl_range.name.get()


def set_named_range(range_, value):
    range_.xl_range.name.set(value)


def set_names(xl_workbook, names):
    try:
        for i in xl_workbook.named_items.get():
            names[i.name.get()] = i
    except TypeError:
        pass


def delete_name(xl_workbook, name):
    xl_workbook.named_items[name].delete()
