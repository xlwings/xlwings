import os
import datetime as dt
import subprocess
import unicodedata
import appscript
import struct
import aem
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


class Applications(object):

    def _iter_excel_instances(self):
        for proc in psutil.process_iter():
            try:
                if proc.name() == 'Microsoft Excel':
                    yield proc.pid
            except psutil.NoSuchProcess:
                pass

    def __iter__(self):
        for pid in self._iter_excel_instances():
            yield Application(xl=pid)

    def __len__(self):
        return len(list(self._iter_excel_instances()))

    def __getitem__(self, index):
        pids = list(self._iter_excel_instances())
        return Application(xl=pids[index])


class Application(object):

    def __init__(self, spec=None, xl=None):
        if xl is None:
            self.xl = app(name=spec or 'Microsoft Excel', newinstance=True, terms=mac_dict)
            # need to do *something* with the app otherwise it doesn't start up
            b = self.xl.visible
        elif isinstance(xl, int):
            self.xl = app(pid=xl, terms=mac_dict)
        else:
            self.xl = xl

    @property
    def api(self):
        return self.xl

    @property
    def pid(self):
        data = self.xl.AS_appdata.target().addressdesc.coerce(aem.kae.typeKernelProcessID).data
        pid, = struct.unpack('i', data)
        return pid

    @property
    def version(self):
        return self.xl.version.get()

    @property
    def active_workbook(self):
        return Workbook(self, self.xl.active_workbook.name.get())

    @property
    def active_sheet(self):
        book = self.active_workbook
        return Sheet(book, self.xl.active_sheet.name.get())

    @property
    def selection(self):
        sheet = self.active_sheet
        return Range(sheet, self.xl.selection.address.get())

    @property
    def visible(self):
        return app('System Events').processes['Microsoft Excel'].visible.get()

    @visible.setter
    def visible(self, visible):
        if visible:
            self.xl.activate()
        else:
            app('System Events').processes['Microsoft Excel'].visible.set(visible)

    def quit(self):
        self.xl.quit(saving=kw.no)

    def kill(self):
        psutil.Process(self.pid).kill()

    @property
    def screen_updating(self):
        return self.xl.screen_updating.get()

    @screen_updating.setter
    def screen_updating(self, value):
        self.xl.screen_updating.set(value)

    # TODO: Hack for Excel 2016, to be refactored
    _CALCULATION = {
        kw.calculation_automatic: Calculation.xlCalculationAutomatic,
        kw.calculation_manual: Calculation.xlCalculationManual,
        kw.calculation_semiautomatic: Calculation.xlCalculationSemiautomatic
    }

    _CALCULATION_REVERSE = {
        v: k for k, v in _CALCULATION.items()
    }

    @property
    def calculation(self):
        return Application._CALCULATION[self.calculation.get()]

    @calculation.setter
    def calculation(self, value):
        self.xl.calculation.set(Application._CALCULATION_REVERSE[value])

    def calculate(self):
        self.xl.calculate()

    @property
    def workbooks(self):
        return Workbooks(self)

    def range(self, arg1, arg2):
        return self.active_sheet.range(arg1, arg2)

    @property
    def hwnd(self):
        return None


class Workbooks(object):

    def __init__(self, app):
        self.app = app

    @property
    def api(self):
        return None

    def __call__(self, name_or_index):
        return Workbook(self.app, name_or_index)

    def __len__(self):
        return self.app.xl.count(each=kw.workbook)

    def add(self):
        xl = self.app.xl.make(new=kw.workbook)
        return Workbook(self.app, xl.name.get())

    def open(self, fullname):
        filename = os.path.basename(fullname)
        self.app.xl.open(fullname)
        return Workbook(self.app, filename)

    def __iter__(self):
        n = len(self)
        for i in range(n):
            yield Workbook(self.app, i+1)

    def open_workbook(self, fullname):
        filename = os.path.basename(fullname)
        self.xl.open(fullname)
        return Workbook(self, filename)

    def get_workbook(self, name):
        return Workbook(self, name)

    def add_sheet(xl_workbook, before, after):
        if before:
            position = before.xl_sheet.before
        else:
            position = after.xl_sheet.after
        return xl_workbook.make(new=kw.worksheet, at=position)

    def count_worksheets(xl_workbook):
        return xl_workbook.count(each=kw.worksheet)


class Workbook(object):
    def __init__(self, app, name_or_index):
        self.app = app
        self.xl = app.xl.workbooks[name_or_index]

    @property
    def api(self):
        return self.xl

    #@property
    #def xl(self):
    #    return self.application.xl.workbooks(self.name)

    def sheet(self, name_or_index):
        return Sheet(self, self.xl.name_or_index)

    @property
    def sheets(self):
        return Sheets(self)

    @property
    def name(self):
        return self.xl.name.get()

    @property
    def application(self):
        return Application(xl=self.app)

    @property
    def active_sheet(self):
        return Sheet(self, self.xl.active_sheet.name.get())

    def save(self, path):
        saved_path = self.xl.properties().get(kw.path)
        if (saved_path != '') and (path is None):
            # Previously saved: Save under existing name
            self.xl.save()
        elif (saved_path == '') and (path is None):
            # Previously unsaved: Save under current name in current working directory
            path = os.path.join(os.getcwd(), self.xl.name.get() + '.xlsx')
            hfs_path = posix_to_hfs_path(path)
            self.xl.save_workbook_as(filename=hfs_path, overwrite=True)
        elif path:
            # Save under new name/location
            hfs_path = posix_to_hfs_path(path)
            self.xl.save_workbook_as(filename=hfs_path, overwrite=True)

    @property
    def fullname(self):
        hfs_path = self.xl.properties().get(kw.full_name)
        # Excel 2011 returns HFS path, Excel 2016 returns POSIX path
        if hfs_path == self.xl.properties().get(kw.name) or int(self.app.version.split('.')[0]) >= 15:
            return hfs_path
        return hfs_to_posix_path(hfs_path)

    def set_names(xl_workbook, names):
        try:
            for i in xl_workbook.named_items.get():
                names[i.name.get()] = i
        except TypeError:
            pass

    def delete_name(xl_workbook, name):
        xl_workbook.named_items[name].delete()

    def add_picture(xl_workbook, sheet_name_or_index, filename, link_to_file, save_with_document, left, top, width, height):
        sheet_index = xl_workbook.sheets[sheet_name_or_index].entry_index.get()
        return xl_workbook.make(
            at=xl_workbook.sheets[sheet_index],
            new=kw.picture,
            with_properties={
                kw.file_name: posix_to_hfs_path(filename),
                kw.link_to_file: link_to_file,
                kw.save_with_document: save_with_document,
                kw.top: top,
                kw.left_position: left,
                kw.width: width,
                kw.height: height
            }
        )

    def close(self):
        self.xl.close(saving=kw.no)

def delete_sheet(sheet):
    _xl_app.display_alerts.set(False)
    sheet.xl_sheet.delete()
    _xl_app.display_alerts.set(True)


class Sheets(object):
    def __init__(self, workbook):
        self.workbook = workbook

    @property
    def api(self):
        return None

    def __call__(self, name_or_index):
        return Sheet(self.workbook, name_or_index)

    def __len__(self):
        return self.workbook.xl.count(each=kw.worksheet)

    def add(self, before=None, after=None):
        if before:
            position = before.xl.before
        else:
            position = after.xl.after
        xl = self.workbook.xl.make(new=kw.worksheet, at=position)
        return Sheet(self.workbook, xl.name.get())


class Sheet(object):
    def __init__(self, workbook, name_or_index):
        self.workbook = workbook
        self.xl = workbook.xl.sheets[name_or_index]

    def range(self, arg1, arg2):
        if isinstance(arg1, Range):
            if isinstance(arg2, Range):
                row1 = min(arg1.top, arg2.top)
                col1 = min(arg1.left, arg2.left)
                row2 = max(arg1.bottom, arg2.bottom)
                col2 = max(arg1.right, arg2.right)
                return Range(
                    self,
                    "{0}:{1}".format(
                        self.xl.rows[row1].columns[col1].address.get(),
                        self.xl.rows[row2].columns[col2].address.get(),
                    )
                )
            else:
                return Range(self, arg1.address)
        else:
            if arg2 is None:
                return Range(self, arg1)
            else:
                raise ValueError("Invalid parameters")

    @property
    def api(self):
        return self.xl

    @property
    def parent(self):
        return self.workbook

    @property
    def name(self):
        return self.xl.name.get()

    def set_name(self, value):
        self.xl.name.set(value)
        self.name = self.xl.name.get()

    def get_index(self):
        return self.xl.entry_index.get()

    def get_workbook(self):
        return self.workbook

    def activate(self):
        self.xl.activate_object()

    def get_value_from_index(self, row_index, column_index):
        return self.xl.columns[column_index].rows[row_index].value.get()

    def clear_contents(self):
        self.xl.used_range.clear_contents()

    def clear(self):
        self.xl.used_range.clear_range()

    def get_row_index_end_down(self, row_index, column_index):
        ix = self.xl.columns[column_index].rows[row_index].get_end(direction=kw.toward_the_bottom).first_row_index.get()
        return ix

    def get_column_index_end_right(self, row_index, column_index):
        ix = self.xl.columns[column_index].rows[row_index].get_end(direction=kw.toward_the_right).first_column_index.get()
        return ix

    def get_range_from_indices(self, first_row, first_column, last_row, last_column):
        first_address = self.xl.columns[first_column].rows[first_row].get_address()
        last_address = self.xl.columns[last_column].rows[last_row].get_address()
        return Range(self, '{0}:{1}'.format(first_address, last_address))

    def get_current_region_address(self, row_index, column_index):
        return str(self.xl.columns[column_index].rows[row_index].current_region.get_address())

    def get_chart_object(self, chart_name_or_index):
        return Chart(self, chart_name_or_index)

    def add_chart(self, left, top, width, height):
        # With the sheet name it won't find the chart later, so we go with the index (no idea why)
        sheet_index = self.xl.entry_index.get()
        return self.workbook.xl.make(
            at=self.xl,
            new=kw.chart_object,
            with_properties={
                kw.width: width,
                kw.top: top,
                kw.left_position: left,
                kw.height: height
            }
        )

    def autofit_sheet(self, axis):
        #TODO: combine with autofit that works on Range objects
        num_columns = self.xl.count(each=kw.column)
        num_rows = self.xl.count(each=kw.row)
        xl_range = self.get_range_from_indices(1, 1, num_rows, num_columns)
        address = xl_range.get_address()
        _xl_app.screen_updating.set(False)
        if axis == 'rows' or axis == 'r':
            self.xl.rows[address].autofit()
        elif axis == 'columns' or axis == 'c':
            self.xl.columns[address].autofit()
        elif axis is None:
            self.xl.rows[address].autofit()
            self.xl.columns[address].autofit()
        _xl_app.screen_updating.set(True)

    def get_shapes_names(self):
        shapes = self.xl.shapes.get()
        if shapes != kw.missing_value:
            return [i.name.get() for i in shapes]
        else:
            return []


class Range(object):
    def __init__(self, sheet, address):
        self.sheet = sheet
        self.xl = sheet.xl.cells[address]

    @property
    def api(self):
        return self.xl

    @property
    def raw_value(self):
        return self.xl.value.get()

    @raw_value.setter
    def raw_value(self, value):
        self.xl.value.set(value)

    def get_worksheet(self):
        return self.sheet

    def get_coordinates(self):
        row1 = self.xl.first_row_index.get()
        col1 = self.xl.first_column_index.get()
        row2 = row1 + self.xl.count(each=kw.row) - 1
        col2 = col1 + self.xl.count(each=kw.column) - 1
        return (row1, col1, row2, col2)

    def get_first_row(self):
        return self.xl.first_row_index.get()

    def get_first_column(self):
        return self.xl.first_column_index.get()

    def count_rows(self):
        return self.xl.count(each=kw.row)

    def count_columns(self):
        return self.xl.count(each=kw.column)

    def clear_contents(self):
        self.sheet.workbook.application.screen_updating.set(False)
        self.xl.clear_range()
        self.sheet.workbook.application.screen_updating.set(True)

    def get_formula(self):
        return self.xl.formula.get()

    def set_formula(self, value):
        self.xl.formula.set(value)

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

    def get_number_format(range_):
        return range_.xl_range.number_format.get()

    def set_number_format(range_, value):
        _xl_app.screen_updating.set(False)
        range_.xl_range.number_format.set(value)
        _xl_app.screen_updating.set(True)

    def get_address(xl_range, row_absolute, col_absolute, external):
        return xl_range.get_address(row_absolute=row_absolute, column_absolute=col_absolute, external=external)

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

    def get_named_range(range_):
        return range_.xl_range.name.get()

    def set_named_range(range_, value):
        range_.xl_range.name.set(value)

    def __len__(self):
        return self.xl.count(each=kw.cell)


class Shape(object):
    def __init__(self, sheet, name_or_index):
        self.sheet = sheet
        self.name_or_index = name_or_index

    @property
    def xl(self):
        return self.sheet.xl.shapes[self.name_or_index]
        #return self.sheet.__appscript__.chart_objects[self.name_or_index]

    def set_name(self, name):
        self.xl.set(name)
        self.name_or_index = name

    def get_index(self):
        return self.xl.entry_index.get()

    def get_name(self):
        return self.xl.name.get()

    def activate(self):
        # xl_shape.activate_object() doesn't work
        self.xl.select()


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


class Chart(Shape):

    def set_source_data_chart(xl_chart, xl_range):
        self.xl.chart.set_source_data(source=xl_range)

    def get_type(self):
        return self.xl.chart.chart_type.get()

    def set_type(self, chart_type):
        self.xl.chart.chart_type.set(chart_type)


class Names(object):
    def __init__(self, xl):
        self.xl = xl

    def __call__(self, name_or_index):
        return Name(xl=self.xl(name_or_index))

    def contains(self, name_or_index):
        try:
            self.xl(name_or_index)
        except pywintypes.com_error as e:
            if e.hresult == -2147352567:
                return False
            else:
                raise
        return True

    def __len__(self):
        return self.xl.Count

    def add(self, name, refers_to):
        return Name(xl=self.xl.Add(name, refers_to))


class Name(object):
    def __init__(self, xl):
        self.xl = xl

    def delete(self):
        self.xl.Delete()

    @property
    def name(self):
        return self.xl.Name

    @name.setter
    def name(self, value):
        self.xl.Name = value

    @property
    def refers_to(self):
        return self.xl.RefersTo

    @refers_to.setter
    def refers_to(self, value):
        self.xl.RefersTo = value

    @property
    def refers_to_range(self):
        return Range(sheet=self, xl=self.xl.RefersToRange)


def is_app_instance(xl_app):
    return type(xl_app) is appscript.reference.Application and '/Microsoft Excel.app' in str(xl_app)


def set_xl_app(app_target=None):
    if app_target is None:
        app_target = 'Microsoft Excel'
    global _xl_app
    _xl_app = app(app_target, terms=mac_dict)


def new_app(app_target='Microsoft Excel'):
    return app(app_target, terms=mac_dict)


def get_running_app():
    return app('Microsoft Excel', terms=mac_dict)


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
    app = Application()
    return Workbook(app, filename)


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
    return isinstance(xl_range, appscript.genericreference.GenericReference)



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



def open_template(fullpath):
    subprocess.call(['open', fullpath])


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




def run(wb, command, app_, args):
    # kwargs = {'arg{0}'.format(i): n for i, n in enumerate(args, 1)}  # only for > PY 2.6
    kwargs = dict(('arg{0}'.format(i), n) for i, n in enumerate(args, 1))
    return app_.xl_app.run_VB_macro("'{0}'!{1}".format(wb.name, command), **kwargs)
