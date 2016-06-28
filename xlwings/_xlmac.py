import os
import datetime as dt
import subprocess
import unicodedata
import struct
import aem
import appscript
from appscript import k as kw, mactypes, its
from appscript.reference import CommandError
import psutil
import atexit
from .constants import ColorIndex, Calculation
from .utils import int_to_rgb, np_datetime_to_datetime
from . import mac_dict, PY3, string_types
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

DIRECTIONS = {
    'd': kw.toward_the_bottom,
    'down': kw.toward_the_bottom,
    'l': kw.toward_the_left,
    'left': kw.toward_the_left,
    'r': kw.toward_the_right,
    'right': kw.toward_the_right,
    'u': kw.toward_the_top,
    'up': kw.toward_the_top
}


class Apps(object):

    def _iter_excel_instances(self):
        asn = subprocess.check_output(['lsappinfo', 'visibleprocesslist', '-includehidden']).decode('utf-8')
        for asn in asn.split(' '):
            if "Microsoft_Excel" in asn:
                pid_info = subprocess.check_output(['lsappinfo', 'info', '-only', 'pid', asn]).decode('utf-8')
                yield int(pid_info.split('=')[1])

    def __iter__(self):
        for pid in self._iter_excel_instances():
            yield App(xl=pid)

    def __len__(self):
        return len(list(self._iter_excel_instances()))

    def __getitem__(self, index):
        pids = list(self._iter_excel_instances())
        return App(xl=pids[index])


class App(object):

    def __init__(self, spec=None, xl=None):
        if xl is None:
            self.xl = appscript.app(name=spec or 'Microsoft Excel', newinstance=True, terms=mac_dict)
            self.activate()  # Makes it behave like on Windows
        elif isinstance(xl, int):
            self.xl = appscript.app(pid=xl, terms=mac_dict)
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
    def selection(self):
        sheet = self.books.active.sheets.active
        return Range(sheet, self.xl.selection.get_address())

    def activate(self, steal_focus=False):
        asn = subprocess.check_output(['lsappinfo', 'visibleprocesslist', '-includehidden']).decode('utf-8')
        frontmost_asn = asn.split(' ')[0]
        pid_info_frontmost = subprocess.check_output(['lsappinfo', 'info', '-only', 'pid', frontmost_asn]).decode('utf-8')
        pid_frontmost = int(pid_info_frontmost.split('=')[1])

        appscript.app('System Events').processes[its.unix_id == self.pid].processes[1].frontmost.set(True)
        if not steal_focus:
            appscript.app('System Events').processes[its.unix_id == pid_frontmost].processes[1].frontmost.set(True)

    @property
    def visible(self):
        return appscript.app('System Events').processes[its.unix_id == self.pid].visible.get()[0]

    @visible.setter
    def visible(self, visible):
        appscript.app('System Events').processes[its.unix_id == self.pid].visible.set(visible)

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
        return App._CALCULATION[self.calculation.get()]

    @calculation.setter
    def calculation(self, value):
        self.xl.calculation.set(App._CALCULATION_REVERSE[value])

    def calculate(self):
        self.xl.calculate()

    @property
    def books(self):
        return Books(self)

    def range(self, arg1, arg2):
        return self.active_sheet.range(arg1, arg2)

    @property
    def hwnd(self):
        return None

    def run(self, macro, args):
        # kwargs = {'arg{0}'.format(i): n for i, n in enumerate(args, 1)}  # only for > PY 2.6
        kwargs = dict(('arg{0}'.format(i), n) for i, n in enumerate(args, 1))
        return self.xl.run_VB_macro(macro, **kwargs)


class Books(object):

    def __init__(self, app):
        self.app = app

    @property
    def api(self):
        return None

    @property
    def active(self):
        return Book(self.app, self.app.xl.active_workbook.name.get())

    def __call__(self, name_or_index):
        return Book(self.app, name_or_index)

    def __len__(self):
        return self.app.xl.count(each=kw.workbook)

    def add(self):
        xl = self.app.xl.make(new=kw.workbook)
        return Book(self.app, xl.name.get())

    def open(self, fullname):
        filename = os.path.basename(fullname)
        self.app.xl.open(fullname)
        return Book(self.app, filename)

    def __iter__(self):
        n = len(self)
        for i in range(n):
            yield Book(self.app, i + 1)


class Book(object):
    def __init__(self, app, name_or_index):
        self.app = app
        self.xl = app.xl.workbooks[name_or_index]

    @property
    def api(self):
        return self.xl

    @property
    def name(self):
        return self.xl.name.get()

    @name.setter
    def name(self, value):
        self.xl.name.set(value)
        self.xl = self.app.xl.workbooks[value]

    @property
    def sheets(self):
        return Sheets(self)

    def close(self):
        self.xl.close(saving=kw.no)

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

    @property
    def names(self):
        return Names(book=self, xl=self.xl.named_items)

    def activate(self):
        self.xl.activate_object()


class Sheets(object):
    def __init__(self, workbook):
        self.workbook = workbook

    @property
    def api(self):
        return None

    @property
    def active(self):
        return Sheet(self.workbook, self.workbook.xl.active_sheet.name.get())

    def __call__(self, name_or_index):
        return Sheet(self.workbook, name_or_index)

    def __len__(self):
        return self.workbook.xl.count(each=kw.worksheet)

    def add(self, before=None, after=None):
        if before is None and after is None:
            before = self.workbook.app.books.active.sheets.active
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

    @property
    def api(self):
        return self.xl

    @property
    def name(self):
        return self.xl.name.get()

    @name.setter
    def name(self, value):
        self.xl.name.set(value)
        self.xl = self.workbook.xl.sheets[value]

    @property
    def names(self):
        return Names(self)

    @property
    def book(self):
        return self.workbook

    @property
    def index(self):
        return self.xl.entry_index.get()

    def range(self, arg1, arg2=None):
        if isinstance(arg1, tuple):
            if 0 in arg1:
                raise IndexError("Attempted to access 0-based Range. xlwings/Excel Ranges are 1-based.")
            row1 = arg1[0]
            col1 = arg1[1]
            address1 = self.xl.rows[row1].columns[col1].get_address()
        elif isinstance(arg1, Range):
            row1 = min(arg1.row, arg2.row)
            col1 = min(arg1.column, arg2.column)
            address1 = self.xl.rows[row1].columns[col1].get_address()
        elif isinstance(arg1, string_types):
            address1 = arg1.split(':')[0]
        else:
            raise ValueError("Invalid parameters")

        if isinstance(arg2, tuple):
            if 0 in arg2:
                raise IndexError("Attempted to access 0-based Range. xlwings/Excel Ranges are 1-based.")
            row2 = arg2[0]
            col2 = arg2[1]
            address2 = self.xl.rows[row2].columns[col2].get_address()
        elif isinstance(arg2, Range):
            row2 = max(arg1.row + arg1.row_count - 1, arg2.row + arg2.row_count - 1)
            col2 = max(arg1.column + arg1.column_count - 1, arg2.column + arg2.column_count - 1)
            address2 = self.xl.rows[row2].columns[col2].get_address()
        elif isinstance(arg2, string_types):
            address2 = arg2
        elif arg2 is None:
            if isinstance(arg1, string_types) and len(arg1.split(':')) == 2:
                address2 = arg1.split(':')[1]
            else:
                address2 = address1
        else:
            raise ValueError("Invalid parameters")

        return Range(self, "{0}:{1}".format(address1, address2))

    @property
    def cells(self):
        return self.range((1, 1), (self.xl.count(each=kw.row), self.xl.count(each=kw.column)))

    def activate(self):
        self.xl.activate_object()

    def clear_contents(self):
        self.xl.used_range.clear_contents()

    def clear(self):
        self.xl.used_range.clear_range()

    def autofit(self, axis=None):
        num_columns = self.xl.count(each=kw.column)
        num_rows = self.xl.count(each=kw.row)
        address = self.range((1, 1), (num_rows, num_columns)).address
        alerts_state = self.book.app.screen_updating
        self.book.app.screen_updating = False
        if axis == 'rows' or axis == 'r':
            self.xl.rows[address].autofit()
        elif axis == 'columns' or axis == 'c':
            self.xl.columns[address].autofit()
        elif axis is None:
            self.xl.rows[address].autofit()
            self.xl.columns[address].autofit()
        self.book.app.screen_updating = alerts_state

    def delete(self):
        alerts_state = self.book.app.screen_updating
        self.book.app.screen_updating = False
        self.xl.delete()
        self.book.app.screen_updating = alerts_state

    @property
    def charts(self):
        pass  # TODO

    @property
    def shapes(self):
        pass  # TODO

    @property
    def pictures(self):
        pass  # TODO


class Range(object):

    def __init__(self, sheet, address):
        self.sheet = sheet
        self.xl = sheet.xl.cells[address]

    @property
    def api(self):
        return self.xl

    def __len__(self):
        return self.xl.count(each=kw.cell)

    @property
    def row(self):
        return self.xl.first_row_index.get()

    @property
    def column(self):
        return self.xl.first_column_index.get()

    @property
    def row_count(self):
        return self.xl.count(each=kw.row)

    @property
    def column_count(self):
        return self.xl.count(each=kw.column)

    @property
    def raw_value(self):
        return self.xl.value.get()

    @raw_value.setter
    def raw_value(self, value):
        self.xl.value.set(value)

    def clear_contents(self):
        alerts_state = self.sheet.book.app.screen_updating
        self.sheet.book.app.screen_updating = False
        self.xl.clear_range()
        self.sheet.book.app.screen_updating = alerts_state

    def get_cell(self, row, col):
        return Range(self.sheet, self.xl.rows[row].columns[col].get_address())

    def clear(self):
        alerts_state = self.sheet.book.app.screen_updating
        self.sheet.book.app.screen_updating = False
        self.xl.clear_range()
        self.sheet.book.app.screen_updating = alerts_state

    @property
    def formula(self):
        return self.xl.formula.get()

    @formula.setter
    def formula(self, value):
        self.xl.formula.set(value)

    def end(self, direction):
        direction = DIRECTIONS.get(direction, direction)
        return Range(self.sheet, self.xl.get_end(direction=direction).get_address())

    @property
    def formula_array(self):
        return self.xl.formula_array.get()

    @formula_array.setter
    def formula_array(self, value):
        self.xl.formula_array.set(value)

    @property
    def column_width(self):
        return self.xl.column_width.get()

    @column_width.setter
    def column_width(self, value):
        self.xl.column_width.set(value)

    @property
    def row_height(self):
        return self.xl.row_height.get()

    @row_height.setter
    def row_height(self, value):
        self.xl.row_height.set(value)

    @property
    def width(self):
        return self.xl.width.get()

    @property
    def height(self):
        return self.xl.height.get()

    @property
    def left(self):
        return self.xl.properties().get(kw.left_position)

    @property
    def top(self):
        return self.xl.properties().get(kw.top)

    @property
    def number_format(self):
        return self.xl.number_format.get()

    @number_format.setter
    def number_format(self, value):
        alerts_state = self.sheet.book.app.screen_updating
        self.sheet.book.app.screen_updating = False
        self.xl.number_format.set(value)
        self.sheet.book.app.screen_updating = alerts_state

    def get_address(self, row_absolute, col_absolute, external):
        return self.xl.get_address(row_absolute=row_absolute, column_absolute=col_absolute, external=external)

    @property
    def address(self):
        return self.xl.get_address()

    @property
    def current_region(self):
        return Range(self.sheet, self.xl.current_region.get_address())

    def autofit(self, axis):
        address = self.address
        alerts_state = self.sheet.book.app.screen_updating
        self.sheet.book.app.screen_updating = False
        if axis == 'rows' or axis == 'r':
            self.sheet.xl.rows[address].autofit()
        elif axis == 'columns' or axis == 'c':
            self.sheet.xl.columns[address].autofit()
        elif axis is None:
            self.sheet.xl.rows[address].autofit()
            self.sheet.xl.columns[address].autofit()
        self.sheet.book.app.screen_updating = alerts_state

    def get_hyperlink_address(self):
        try:
            return self.xl.hyperlinks[1].address.get()
        except CommandError:
            raise Exception("The cell doesn't seem to contain a hyperlink!")

    def set_hyperlink(self, address, text_to_display=None, screen_tip=None):
        self.xl.make(at=self.xl, new=kw.hyperlink, with_properties={kw.address: address,
                                                                    kw.text_to_display: text_to_display,
                                                                    kw.screen_tip: screen_tip})

    @property
    def color(self):
        if self.xl.interior_object.color_index.get() == kw.color_index_none:
            return None
        else:
            return tuple(self.xl.interior_object.color.get  ())

    @color.setter
    def color(self, color_or_rgb):
        if color_or_rgb is None:
            self.xl.interior_object.color_index.set(ColorIndex.xlColorIndexNone)
        elif isinstance(color_or_rgb, int):
            self.xl.interior_object.color.set(int_to_rgb(color_or_rgb))
        else:
            self.xl.interior_object.color.set(color_or_rgb)

    @property
    def name(self):
        xl = self.xl.named_item
        if xl.get() == kw.missing_value:
            return None
        else:
            return Name(self.sheet.book, xl=xl)

    @name.setter
    def name(self, value):
        self.xl.name.set(value)

    def __call__(self, arg1, arg2=None):
        if arg2 is None:
            col = (arg1 - 1) % self.column_count
            row = int((arg1 - 1 - col) / self.column_count)
            return self(1 + row, 1 + col)
        else:
            return Range(self.sheet,
                         self.sheet.xl.rows[self.row + arg1 - 1].columns[self.column + arg2 - 1].get_address())

    @property
    def rows(self):
        row = self.row
        col1 = self.column
        col2 = col1 + self.column_count - 1
        sht = self.sheet
        return [
            self.sheet.range((row+i, col1), (row+i, col2))
            for i in range(self.row_count)
        ]

    @property
    def columns(self):
        col = self.column
        row1 = self.row
        row2 = row1 + self.row_count - 1
        sht = self.sheet
        return [
            sht.range((row1, col + i), (row2, col + i))
            for i in range(self.row_count)
        ]

    def select(self):
        # seems to only work reliably in this combination
        self.xl.activate()
        self.xl.select()

class RangeRows(object):

    def __init__(self, rng, step=1):
        self.rng = rng
        self.step = step

    def __call__(self, index):
        row = self.rng.row + index * self.step - 1
        col1 = self.rng.column
        col2 = self.rng.column_count
        return self.rng.sheet.range((row, col1), (row, col2))

    def slice(self, start, stop, step):
        row1 = self.rng.row + start * self.step
        row2 = row1 + (stop - start - 1) * self.step
        col1 = self.rng.column
        col2 = self.rng.column_count
        rng = self.rng.sheet.range((row1, col1), (row2, col2))
        return RangeRows(rng, self.step * step)

    def __len__(self):
        return len(range(0, self.rng.row_count, self.step))

    def __iter__(self):
        row = self.rng.row
        col1 = self.rng.column
        col2 = self.rng.column_count
        for i in range(0, self.rng.row_count, self.step):
            yield self.rng.sheet.range((row+i, col1), (row+i, col2))


class RangeRows(object):

    def __init__(self, rng, step=1):
        self.rng = rng
        self.step = step

    def __call__(self, index):
        row = self.rng.row + index * self.step - 1
        col1 = self.rng.column
        col2 = self.rng.column_count
        return self.rng.sheet.range((row, col1), (row, col2))

    def slice(self, start, stop, step):
        row1 = self.rng.row + start * self.step
        row2 = row1 + (stop - start - 1) * self.step
        col1 = self.rng.column
        col2 = self.rng.column_count
        rng = self.rng.sheet.range((row1, col1), (row2, col2))
        return RangeRows(rng, self.step * step)

    def __len__(self):
        return len(range(0, self.rng.row_count, self.step))

    def __iter__(self):
        row = self.rng.row
        col1 = self.rng.column
        col2 = self.rng.column_count
        for i in range(0, self.rng.row_count, self.step):
            yield self.rng.sheet.range((row+i, col1), (row+i, col2))


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
    def __init__(self, book, xl):
        self.book = book
        self.xl = xl

    def __call__(self, name_or_index):
        return Name(self.book, xl=self.xl[name_or_index])

    def contains(self, name_or_index):
        try:
            self.xl[name_or_index].get()
        except appscript.reference.CommandError as e:
            # TODO: make more specific
            return False
        return True

    def __len__(self):
        named_items = self.xl.get()
        if named_items == kw.missing_value:
            return 0
        else:
            return len(named_items)

    def add(self, name, refers_to):
        return Name(self.book, self.book.xl.make(at=self.book.xl,
                                                 new=kw.named_item,
                                                 with_properties={
                                                     kw.references: refers_to,
                                                     kw.name: name
                                                 }))


class Name(object):
    def __init__(self, book, xl):
        self.book = book
        self.xl = xl

    def delete(self):
        self.xl.delete()

    @property
    def name(self):
        return self.xl.name.get()

    @name.setter
    def name(self, value):
        self.xl.name.set(value)

    @property
    def refers_to(self):
        return self.xl.properties().get(kw.references)

    @refers_to.setter
    def refers_to(self, value):
        self.xl.properties(kw.references).set(value)

    @property
    def refers_to_range(self):
        ref = self.refers_to[1:].split('!')
        return Range(Sheet(self.book, ref[0]), ref[1])


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
        for app in Apps():
            try:
                app.xl.run_VB_macro('CleanUp')
            except (CommandError, AttributeError, aem.aemsend.EventError):
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
    app = App()
    return Book(app, filename)


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
