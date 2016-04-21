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
if np:
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


def posix_to_hfs_path(posix_path):
    """
    Turns a posix path (/Path/file.ext) into an HFS path (Macintosh HD:Path:file.ext)
    """
    dir_name, file_name = os.path.split(posix_path)
    dir_name_hfs = mactypes.Alias(dir_name).hfspath
    return dir_name_hfs + ':' + file_name


def hfs_to_posix_path(hfs_path):
    """
    Turns an HFS path (Macintosh HD:Path:file.ext) into a posix path (/Path/file.ext)
    """
    url = mactypes.convertpathtourl(hfs_path, 1)  # kCFURLHFSPathStyle = 1
    return mactypes.converturltopath(url, 0)  # kCFURLPOSIXPathStyle = 0


def is_file_open(fullname):
    """
    Checks if the file is already open
    """
    for proc in psutil.process_iter():
        try:
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
        except psutil.NoSuchProcess:
            pass
    return False


def is_excel_running():
    for proc in psutil.process_iter():
        try:
            if proc.name() == 'Microsoft Excel':
                return True
        except psutil.NoSuchProcess:
            pass
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


def get_sheet_workbook(xl_sheet):
    return xl_sheet.parent.get()


def get_range_sheet(xl_range):
    return xl_range.worksheet.get()


def get_range_coordinates(xl_range):
    row1 = xl_range.first_row_index.get()
    col1 = xl_range.first_column_index.get()
    row2 = row1 + xl_range.count(each=kw.row) - 1
    col2 = col1 + xl_range.count(each=kw.column) - 1
    return (row1, col1, row2, col2)



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

    if is_running or 0 == _xl_app.count(None, each=kw.workbook):
        # If Excel is being fired up, a "Workbook1" is automatically added
        # If its already running, we create an new one that Excel unfortunately calls "Sheet1".
        # It's a feature though: See p.14 on Excel 2004 AppleScript Reference
        xl_workbook = _xl_app.make(new=kw.workbook)
    else:
        xl_workbook = _xl_app.workbooks[1]

    return _xl_app, xl_workbook


def is_range_instance(xl_range):
    return False


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


def _clean_value_data_element(value, datetime_builder, empty_as, number_builder):
    if value == '':
        return empty_as
    if isinstance(value, dt.datetime) and datetime_builder is not dt.datetime:
        value = datetime_builder(
            month=value.month,
            day=value.day,
            year=value.year,
            hour=value.hour,
            minute=value.minute,
            second=value.second,
            microsecond=value.microsecond,
            tzinfo=None
        )
    elif number_builder is not None and type(value) == float:
        value = number_builder(value)
    return value


def clean_value_data(data, datetime_builder, empty_as, number_builder):
    return [[_clean_value_data_element(c, datetime_builder, empty_as, number_builder) for c in row] for row in data]


def get_value_from_index(xl_sheet, row_index, column_index):
    return xl_sheet.columns[column_index].rows[row_index].value.get()

def prepare_xl_data_element(x):
    if x is None:
        return ""
    elif np and isinstance(x, float) and np.isnan(x):
        return ""
    elif np and isinstance(x, np.datetime64):
        # handle numpy.datetime64
        return np_datetime_to_datetime(x).replace(tzinfo=None)
    elif pd and isinstance(x, pd.tslib.Timestamp):
        # This transformation seems to be only needed on Python 2.6 (?)
        return x.to_datetime().replace(tzinfo=None)
    elif isinstance(x, dt.datetime):
        # Make datetime timezone naive
        return x.replace(tzinfo=None)
    elif isinstance(x, int):
        # appscript packs integers larger than SInt32 but smaller than SInt64 as typeSInt64, and integers
        # larger than SInt64 as typeIEEE64BitFloatingPoint. Excel silently ignores typeSInt64. (GH 227)
        return float(x)

    return x

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


def get_formula_array(xl_range):
    return xl_range.formula_array.get()


def set_formula_array(xl_range, value):
    xl_range.formula_array.set(value)


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


def activate_shape(xl_shape):
    # xl_shape.activate_object() doesn't work
    xl_shape.select()


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


def get_left(xl_range):
    return xl_range.properties().get(kw.left_position)


def get_top(xl_range):
    return xl_range.properties().get(kw.top)


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
        hfs_path = posix_to_hfs_path(path)
        xl_workbook.save_workbook_as(filename=hfs_path, overwrite=True)
    elif path:
        # Save under new name/location
        hfs_path = posix_to_hfs_path(path)
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
    return hfs_to_posix_path(hfs_path)


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


def get_picture(picture):
    return picture.xl_workbook.sheets[picture.sheet_name_or_index].pictures[picture.name_or_index]


def get_picture_index(picture):
    # Workaround since picture.xl_picture.entry_index.get() is broken in AppleScript, returns k.missing_value
    # Also, count(each=kw.picture) returns count of shape nevertheless
    num_shapes = picture.xl_workbook.sheets[picture.sheet_name_or_index].count(each=kw.shape)
    picture_index = 0
    for i in range(1, num_shapes + 1):
        if picture.xl_workbook.sheets[picture.sheet_name_or_index].shapes[i].shape_type.get() == kw.shape_type_picture:
            picture_index += 1
        if picture.xl_workbook.sheets[picture.sheet_name_or_index].shapes[i].name.get() == picture.name:
            return picture_index


def get_picture_name(xl_picture):
    return xl_picture.name.get()


def get_shape(shape):
    return shape.xl_workbook.sheets[shape.sheet_name_or_index].shapes[shape.name_or_index]


def get_shape_name(shape):
    return shape.xl_shape.name.get()


def set_shape_name(xl_workbook, sheet_name_or_index, xl_shape, value):
    xl_workbook.sheets[sheet_name_or_index].shapes[xl_shape.name.get()].name.set(value)
    return xl_workbook.sheets[sheet_name_or_index].shapes[value]


def get_shapes_names(xl_workbook, sheet):
    shapes = xl_workbook.sheets[sheet].shapes.get()
    if shapes != kw.missing_value:
        return [i.name.get() for i in shapes]
    else:
        return []


def get_shape_left(shape):
    return shape.xl_shape.left_position.get()


def set_shape_left(shape, value):
    shape.xl_shape.left_position.set(value)


def get_shape_top(shape):
    return shape.xl_shape.top.get()


def set_shape_top(shape, value):
    shape.xl_shape.top.set(value)


def get_shape_width(shape):
    return shape.xl_shape.width.get()


def set_shape_width(shape, value):
    shape.xl_shape.width.set(value)


def get_shape_height(shape):
    return shape.xl_shape.height.get()


def set_shape_height(shape, value):
    shape.xl_shape.height.set(value)


def delete_shape(shape):
    shape.xl_shape.delete()


def add_picture(xl_workbook, sheet_name_or_index, filename, link_to_file, save_with_document, left, top, width, height):
    sheet_index = xl_workbook.sheets[sheet_name_or_index].entry_index.get()
    return xl_workbook.make(at=xl_workbook.sheets[sheet_index],
                            new=kw.picture,
                            with_properties={kw.file_name: posix_to_hfs_path(filename),
                                             kw.link_to_file: link_to_file,
                                             kw.save_with_document: save_with_document,
                                             kw.top: top,
                                             kw.left_position: left,
                                             kw.width: width,
                                             kw.height: height})


def get_app_version_string(xl_workbook):
    return _xl_app.version.get()


def get_major_app_version_number(xl_workbook):
    return int(get_app_version_string(xl_workbook).split('.')[0])


def delete_sheet(sheet):
    _xl_app.display_alerts.set(False)
    sheet.xl_sheet.delete()
    _xl_app.display_alerts.set(True)


def run(wb, command, app_, args):
    # kwargs = {'arg{0}'.format(i): n for i, n in enumerate(args, 1)}  # only for > PY 2.6
    kwargs = dict(('arg{0}'.format(i), n) for i, n in enumerate(args, 1))
    return app_.xl_app.run_VB_macro("'{0}'!{1}".format(wb.name, command), **kwargs)
