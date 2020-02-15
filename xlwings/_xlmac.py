import os
import re
import datetime as dt
import subprocess
import struct
import shutil
import atexit

import psutil
import aem
import appscript
from appscript import k as kw, mactypes, its
from appscript.reference import CommandError

from .constants import ColorIndex
from .utils import int_to_rgb, np_datetime_to_datetime, col_name, VersionNumber
from . import mac_dict

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


class Apps:

    def _iter_excel_instances(self):
        asn = subprocess.check_output(['lsappinfo', 'visibleprocesslist', '-includehidden']).decode('utf-8')
        for asn in asn.split(' '):
            if "Microsoft_Excel" in asn:
                pid_info = subprocess.check_output(['lsappinfo', 'info', '-only', 'pid', asn]).decode('utf-8')
                if pid_info != '"pid"=[ NULL ] \n':
                    yield int(pid_info.split('=')[1])

    def keys(self):
        return list(self._iter_excel_instances())

    def __iter__(self):
        for pid in self._iter_excel_instances():
            yield App(xl=pid)

    def __len__(self):
        return len(list(self._iter_excel_instances()))

    def __getitem__(self, pid):
        if pid not in self.keys():
            raise KeyError('Could not find an Excel instance with this PID.')
        return App(xl=pid)


class App:

    def __init__(self, spec=None, add_book=None, xl=None):
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
        try:
            # fails if e.g. chart is selected
            return Range(sheet, self.xl.selection.get_address())
        except CommandError:
            return None

    def activate(self, steal_focus=False):
        asn = subprocess.check_output(['lsappinfo', 'visibleprocesslist', '-includehidden']).decode('utf-8')
        frontmost_asn = asn.split(' ')[0]
        pid_info_frontmost = subprocess.check_output(['lsappinfo', 'info', '-only', 'pid', frontmost_asn]).decode('utf-8')
        pid_frontmost = int(pid_info_frontmost.split('=')[1])

        appscript.app('System Events').processes[its.unix_id == self.pid].frontmost.set(True)
        if not steal_focus:
            appscript.app('System Events').processes[its.unix_id == pid_frontmost].frontmost.set(True)

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

    @property
    def display_alerts(self):
        return self.xl.display_alerts.get()

    @display_alerts.setter
    def display_alerts(self, value):
        self.xl.display_alerts.set(value)

    @property
    def calculation(self):
        return calculation_k2s[self.xl.calculation.get()]

    @calculation.setter
    def calculation(self, value):
        self.xl.calculation.set(calculation_s2k[value])

    def calculate(self):
        self.xl.calculate()

    @property
    def books(self):
        return Books(self)

    def range(self, arg1, arg2):
        return self.books.active.sheets.active.range(arg1, arg2)

    @property
    def hwnd(self):
        return None

    def run(self, macro, args):
        kwargs = {'arg{0}'.format(i): n for i, n in enumerate(args, 1)}
        return self.xl.run_VB_macro(macro, **kwargs)


class Books:

    def __init__(self, app):
        self.app = app

    @property
    def api(self):
        return None

    @property
    def active(self):
        return Book(self.app, self.app.xl.active_workbook.name.get())

    def __call__(self, name_or_index):
        b = Book(self.app, name_or_index)
        if not b.xl.exists():
            raise KeyError(name_or_index)
        return b

    def __contains__(self, key):
        return Book(self.app, key).xl.exists()

    def __len__(self):
        return self.app.xl.count(each=kw.workbook)

    def add(self):
        self.app.activate()
        xl = self.app.xl.make(new=kw.workbook)
        wb = Book(self.app, xl.name.get())
        return wb

    def open(self, fullname, update_links=None, read_only=None, format=None, password=None, write_res_password=None,
             ignore_read_only_recommended=None, origin=None, delimiter=None, editable=None, notify=None, converter=None,
             add_to_mru=None, local=None, corrupt_load=None):
        # TODO: format and origin currently require a native appscript keyword, read_only doesn't seem to work
        # Unsupported params
        if local is not None:
            raise Exception('local is not supported on macOS')
        if corrupt_load is not None:
            raise Exception('corrupt_load is not supported on macOS')
        # update_links: on Windows only constants 0 and 3 seem to be supported in this context
        if update_links:
            update_links = kw.update_remote_and_external_links
        else:
            update_links = kw.do_not_update_links

        self.app.activate()
        filename = os.path.basename(fullname)
        self.app.xl.open_workbook(workbook_file_name=fullname, update_links=update_links, read_only=read_only,
                                  format=format, password=password, write_reserved_password=write_res_password,
                                  ignore_read_only_recommended=ignore_read_only_recommended,
                                  origin=origin, delimiter=delimiter, editable=editable, notify=notify,
                                  converter=converter, add_to_mru=add_to_mru)
        wb = Book(self.app, filename)
        return wb

    def __iter__(self):
        n = len(self)
        for i in range(n):
            yield Book(self.app, i + 1)


class Book:
    def __init__(self, app, name_or_index):
        self.app = app
        self.xl = app.xl.workbooks[name_or_index]

    @property
    def api(self):
        return self.xl

    @property
    def name(self):
        return self.xl.name.get()

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
        elif (saved_path != '') and (path is not None) and (os.path.split(path)[0] == ''):
            # Save existing book under new name in cwd if no path has been provided
            path = os.path.join(os.getcwd(), path)
            hfs_path = posix_to_hfs_path(os.path.realpath(path))
            self.xl.save_workbook_as(filename=hfs_path, overwrite=True)
        elif (saved_path == '') and (path is None):
            # Previously unsaved: Save under current name in current working directory
            path = os.path.join(os.getcwd(), self.xl.name.get() + '.xlsx')
            hfs_path = posix_to_hfs_path(os.path.realpath(path))
            self.xl.save_workbook_as(filename=hfs_path, overwrite=True)
        elif path:
            # Save under new name/location
            hfs_path = posix_to_hfs_path(os.path.realpath(path))
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
        return Names(parent=self, xl=self.xl.named_items)

    def activate(self):
        self.xl.activate_object()


class Sheets:

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

    def __iter__(self):
        for i in range(len(self)):
            yield self(i + 1)

    def add(self, before=None, after=None):
        if before is None and after is None:
            before = self.workbook.app.books.active.sheets.active
        if before:
            position = before.xl.before
        else:
            position = after.xl.after
        xl = self.workbook.xl.make(new=kw.worksheet, at=position)
        return Sheet(self.workbook, xl.name.get())


class Sheet:

    def __init__(self, workbook, name_or_index):
        self.workbook = workbook
        self.xl = workbook.xl.worksheets[name_or_index]

    @property
    def api(self):
        return self.xl

    @property
    def name(self):
        return self.xl.name.get()

    @name.setter
    def name(self, value):
        self.xl.name.set(value)
        self.xl = self.workbook.xl.worksheets[value]

    @property
    def names(self):
        return Names(parent=self, xl=self.xl.named_items)

    @property
    def book(self):
        return self.workbook

    @property
    def index(self):
        return self.xl.entry_index.get()

    def range(self, arg1, arg2=None):
        if isinstance(arg1, tuple):
            if len(arg1) == 2:
                if 0 in arg1:
                    raise IndexError("Attempted to access 0-based Range. xlwings/Excel Ranges are 1-based.")
                row1 = arg1[0]
                col1 = arg1[1]
                address1 = self.xl.rows[row1].columns[col1].get_address()
            elif len(arg1) == 4:
                return Range(self, arg1)
            else:
                raise ValueError("Invalid parameters")
        elif isinstance(arg1, Range):
            row1 = min(arg1.row, arg2.row)
            col1 = min(arg1.column, arg2.column)
            address1 = self.xl.rows[row1].columns[col1].get_address()
        elif isinstance(arg1, str):
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
            row2 = max(arg1.row + arg1.shape[0] - 1, arg2.row + arg2.shape[0] - 1)
            col2 = max(arg1.column + arg1.shape[1] - 1, arg2.column + arg2.shape[1] - 1)
            address2 = self.xl.rows[row2].columns[col2].get_address()
        elif isinstance(arg2, str):
            address2 = arg2
        elif arg2 is None:
            if isinstance(arg1, str) and len(arg1.split(':')) == 2:
                address2 = arg1.split(':')[1]
            else:
                return Range(self, "{0}".format(address1))
        else:
            raise ValueError("Invalid parameters")

        return Range(self, "{0}:{1}".format(address1, address2))

    @property
    def cells(self):
        return self.range((1, 1), (self.xl.count(each=kw.row), self.xl.count(each=kw.column)))

    def activate(self):
        self.xl.activate_object()

    def select(self):
        self.xl.select()

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
        alerts_state = self.book.app.xl.display_alerts.get()
        self.book.app.xl.display_alerts.set(False)
        self.xl.delete()
        self.book.app.xl.display_alerts.set(alerts_state)

    @property
    def charts(self):
        return Charts(self)

    @property
    def shapes(self):
        return Shapes(self)

    @property
    def pictures(self):
        return Pictures(self)

    @property
    def used_range(self):
        return Range(self, self.xl.used_range.get_address())


class Range:

    def __init__(self, sheet, address):
        self.sheet = sheet
        if isinstance(address, tuple):
            self._coords = address
            row, col, nrows, ncols = address
            if nrows and ncols:
                self.xl = sheet.xl.cells["%s:%s" % (
                    sheet.xl.rows[row].columns[col].get_address(),
                    sheet.xl.rows[row + nrows - 1].columns[col + ncols - 1].get_address(),
                )]
            else:
                self.xl = None
        else:
            self.xl = sheet.xl.cells[address]
            self._coords = None

    @property
    def coords(self):
        if self._coords is None:
            self._coords = (
                self.xl.first_row_index.get(),
                self.xl.first_column_index.get(),
                self.xl.count(each=kw.row),
                self.xl.count(each=kw.column)
            )
        return self._coords

    @property
    def api(self):
        return self.xl

    def __len__(self):
        return self.coords[2] * self.coords[3]

    @property
    def row(self):
        return self.coords[0]

    @property
    def column(self):
        return self.coords[1]

    @property
    def shape(self):
        return self.coords[2], self.coords[3]

    @property
    def raw_value(self):
        if self.xl is not None:
            return self.xl.value.get()

    @raw_value.setter
    def raw_value(self, value):
        if self.xl is not None:
            self.xl.value.set(value)

    def clear_contents(self):
        if self.xl is not None:
            alerts_state = self.sheet.book.app.screen_updating
            self.sheet.book.app.screen_updating = False
            self.xl.clear_contents()
            self.sheet.book.app.screen_updating = alerts_state

    def clear(self):
        if self.xl is not None:
            alerts_state = self.sheet.book.app.screen_updating
            self.sheet.book.app.screen_updating = False
            self.xl.clear_range()
            self.sheet.book.app.screen_updating = alerts_state

    def end(self, direction):
        direction = directions_s2k.get(direction, direction)
        return Range(self.sheet, self.xl.get_end(direction=direction).get_address())

    @property
    def formula(self):
        if self.xl is not None:
            return self.xl.formula.get()

    @formula.setter
    def formula(self, value):
        if self.xl is not None:
            self.xl.formula.set(value)

    @property
    def formula_array(self):
        if self.xl is not None:
            rv = self.xl.formula_array.get()
            return None if rv == kw.missing_value else rv

    @formula_array.setter
    def formula_array(self, value):
        if self.xl is not None:
            self.xl.formula_array.set(value)

    @property
    def column_width(self):
        if self.xl is not None:
            rv = self.xl.column_width.get()
            return None if rv == kw.missing_value else rv
        else:
            return 0

    @column_width.setter
    def column_width(self, value):
        if self.xl is not None:
            self.xl.column_width.set(value)

    @property
    def row_height(self):
        if self.xl is not None:
            rv = self.xl.row_height.get()
            return None if rv == kw.missing_value else rv
        else:
            return 0

    @row_height.setter
    def row_height(self, value):
        if self.xl is not None:
            self.xl.row_height.set(value)

    @property
    def width(self):
        if self.xl is not None:
            return self.xl.width.get()
        else:
            return 0

    @property
    def height(self):
        if self.xl is not None:
            return self.xl.height.get()
        else:
            return 0

    @property
    def left(self):
        return self.xl.properties().get(kw.left_position)

    @property
    def top(self):
        return self.xl.properties().get(kw.top)

    @property
    def number_format(self):
        if self.xl is not None:
            rv = self.xl.number_format.get()
            return None if rv == kw.missing_value else rv

    @number_format.setter
    def number_format(self, value):
        if self.xl is not None:
            alerts_state = self.sheet.book.app.screen_updating
            self.sheet.book.app.screen_updating = False
            self.xl.number_format.set(value)
            self.sheet.book.app.screen_updating = alerts_state

    def get_address(self, row_absolute, col_absolute, external):
        if self.xl is not None:
            return self.xl.get_address(row_absolute=row_absolute, column_absolute=col_absolute, external=external)

    @property
    def address(self):
        if self.xl is not None:
            return self.xl.get_address()
        else:
            row, col, nrows, ncols = self.coords
            return "$%s$%s{%sx%s}" % (col_name(col), row, nrows, ncols)

    @property
    def current_region(self):
        return Range(self.sheet, self.xl.current_region.get_address())

    def autofit(self, axis=None):
        if self.xl is not None:
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

    def insert(self, shift=None, copy_origin=None):
        # copy_origin is not supported on mac
        shifts = {'down': kw.shift_down, 'right': kw.shift_to_right, None: None}
        self.xl.insert_into_range(shift=shifts[shift])

    def delete(self, shift=None):
        shifts = {'up': kw.shift_up, 'left': kw.shift_to_left, None: None}
        self.xl.delete_range(shift=shifts[shift])

    def copy(self, destination=None):
        self.xl.copy_range(destination=destination.api if destination else None)

    def paste(self, paste=None, operation=None, skip_blanks=False, transpose=False):
        pastes = {
            # all_merging_conditional_formats unsupported on mac
            "all": kw.paste_all,
            "all_except_borders": kw.paste_all_except_borders,
            "all_using_source_theme": kw.paste_all_using_source_theme,
            "column_widths": kw.paste_column_widths,
            "comments": kw.paste_comments,
            "formats": kw.paste_formats,
            "formulas": kw.paste_formulas,
            "formulas_and_number_formats": kw.paste_formulas_and_number_formats,
            "validation": kw.paste_validation,
            "values": kw.paste_values,
            "values_and_number_formats": kw.paste_values_and_number_formats,
            None: None
        }

        operations = {
            "add": kw.paste_special_operation_add,
            "divide": kw.paste_special_operation_divide,
            "multiply": kw.paste_special_operation_multiply,
            "subtract": kw.paste_special_operation_subtract,
            None: None
        }

        self.xl.paste_special(what=pastes[paste], operation=operations[operation], skip_blanks=skip_blanks, transpose=transpose)

    @property
    def hyperlink(self):
        try:
            return self.xl.hyperlinks[1].address.get()
        except CommandError:
            raise Exception("The cell doesn't seem to contain a hyperlink!")

    def add_hyperlink(self, address, text_to_display=None, screen_tip=None):
        if self.xl is not None:
            self.xl.make(at=self.xl, new=kw.hyperlink, with_properties={kw.address: address,
                                                                        kw.text_to_display: text_to_display,
                                                                        kw.screen_tip: screen_tip})

    @property
    def color(self):
        if not self.xl or self.xl.interior_object.color_index.get() == kw.color_index_none:
            return None
        else:
            return tuple(self.xl.interior_object.color.get())

    @color.setter
    def color(self, color_or_rgb):
        if self.xl is not None:
            if color_or_rgb is None:
                self.xl.interior_object.color_index.set(ColorIndex.xlColorIndexNone)
            elif isinstance(color_or_rgb, int):
                self.xl.interior_object.color.set(int_to_rgb(color_or_rgb))
            else:
                self.xl.interior_object.color.set(color_or_rgb)

    @property
    def name(self):
        if not self.xl:
            return None
        xl = self.xl.named_item
        if xl.get() == kw.missing_value:
            return None
        else:
            return Name(self.sheet.book, xl=xl)

    @name.setter
    def name(self, value):
        if self.xl is not None:
            self.xl.name.set(value)

    def __call__(self, arg1, arg2=None):
        if arg2 is None:
            col = (arg1 - 1) % self.shape[1]
            row = int((arg1 - 1 - col) / self.shape[1])
            return self(1 + row, 1 + col)
        else:
            return Range(self.sheet,
                         self.sheet.xl.rows[self.row + arg1 - 1].columns[self.column + arg2 - 1].get_address())

    @property
    def rows(self):
        row = self.row
        col1 = self.column
        col2 = col1 + self.shape[1] - 1
        return [
            self.sheet.range((row + i, col1), (row + i, col2))
            for i in range(self.shape[0])
        ]

    @property
    def columns(self):
        col = self.column
        row1 = self.row
        row2 = row1 + self.shape[0] - 1
        sht = self.sheet
        return [
            sht.range((row1, col + i), (row2, col + i))
            for i in range(self.shape[1])
        ]

    def select(self):
        if self.xl is not None:
            return self.xl.select()

    @property
    def merge_area(self):
        return Range(self.sheet, self.xl.merge_area.get_address())

    @property
    def merge_cells(self):
        return self.xl.merge_cells.get()

    def merge(self, across):
        self.xl.merge(across=across)

    def unmerge(self):
        self.xl.unmerge()

class Shape:
    def __init__(self, parent, key):
        self.parent = parent
        self.xl = parent.xl.shapes[key]

    @property
    def api(self):
        return self.xl

    @property
    def name(self):
        return self.xl.name.get()

    @name.setter
    def name(self, value):
        self.xl.name.set(value)

    @property
    def type(self):
        return shape_types_k2s[self.xl.shape_type.get()]

    @property
    def left(self):
        return self.xl.left_position.get()

    @left.setter
    def left(self, value):
        self.xl.left_position.set(value)

    @property
    def top(self):
        return self.xl.top.get()

    @top.setter
    def top(self, value):
        self.xl.top.set(value)

    @property
    def width(self):
        return self.xl.width.get()

    @width.setter
    def width(self, value):
        self.xl.width.set(value)

    @property
    def height(self):
        return self.xl.height.get()

    @height.setter
    def height(self, value):
        self.xl.height.set(value)

    def delete(self):
        self.xl.delete()

    @property
    def index(self):
        return self.xl.entry_index.get()

    def activate(self):
        # self.xl.activate_object()  # doesn't work?
        self.xl.select()


class Collection:

    def __init__(self, parent):
        self.parent = parent
        self.xl = getattr(self.parent.xl, self._attr)

    @property
    def api(self):
        return self.xl

    def __call__(self, key):
        if not self.xl[key].exists():
            raise KeyError(key)
        return self._wrap(self.parent, key)

    def __len__(self):
        return self.parent.xl.count(each=self._kw)

    def __iter__(self):
        for i in range(len(self)):
            yield self(i + 1)

    def __contains__(self, key):
        return self.xl[key].exists()


class Chart:

    def __init__(self, parent, key):
        self.parent = parent
        if isinstance(parent, Sheet):
            self.xl_obj = parent.xl.chart_objects[key]
            self.xl = self.xl_obj.chart
        else:
            self.xl_obj = None
            self.xl = self.charts[key]

    @property
    def api(self):
        return self.xl_obj, self.xl

    def set_source_data(self, rng):
        self.xl.set_source_data(source=rng.xl)

    @property
    def name(self):
        if self.xl_obj is not None:
            return self.xl_obj.name.get()
        else:
            return self.xl.name.get()

    @name.setter
    def name(self, value):
        if self.xl_obj is not None:
            self.xl_obj.name.set(value)
        else:
            self.xl.name.get(value)

    @property
    def chart_type(self):
        return chart_types_k2s[self.xl.chart_type.get()]

    @chart_type.setter
    def chart_type(self, value):
        self.xl.chart_type.set(chart_types_s2k[value])

    @property
    def left(self):
        if self.xl_obj is None:
            raise Exception("This chart is not embedded.")
        return self.xl_obj.left_position.get()

    @left.setter
    def left(self, value):
        if self.xl_obj is None:
            raise Exception("This chart is not embedded.")
        self.xl_obj.left_position.set(value)

    @property
    def top(self):
        if self.xl_obj is None:
            raise Exception("This chart is not embedded.")
        return self.xl_obj.top.get()

    @top.setter
    def top(self, value):
        if self.xl_obj is None:
            raise Exception("This chart is not embedded.")
        self.xl_obj.top.set(value)

    @property
    def width(self):
        if self.xl_obj is None:
            raise Exception("This chart is not embedded.")
        return self.xl_obj.width.get()

    @width.setter
    def width(self, value):
        if self.xl_obj is None:
            raise Exception("This chart is not embedded.")
        self.xl_obj.width.set(value)

    @property
    def height(self):
        if self.xl_obj is None:
            raise Exception("This chart is not embedded.")
        return self.xl_obj.height.get()

    @height.setter
    def height(self, value):
        if self.xl_obj is None:
            raise Exception("This chart is not embedded.")
        self.xl_obj.height.set(value)

    def delete(self):
        # todo: what about chart sheets?
        self.xl_obj.delete()


class Charts(Collection):

    _attr = 'chart_objects'
    _kw = kw.chart_object
    _wrap = Chart

    def add(self, left, top, width, height):
        sheet_index = self.parent.xl.entry_index.get()
        return Chart(
            self.parent, self.parent.xl.make(
                at=self.parent.book.xl.sheets[sheet_index],
                new=kw.chart_object,
                with_properties={
                    kw.width: width,
                    kw.top: top,
                    kw.left_position: left,
                    kw.height: height
                }
            ).name.get()
        )


class Picture:

    def __init__(self, parent, key):
        self.parent = parent
        self.xl = parent.xl.pictures[key]

    @property
    def api(self):
        return self.xl

    @property
    def name(self):
        return self.xl.name.get()

    @name.setter
    def name(self, value):
        self.xl.name.set(value)

    @property
    def left(self):
        return self.xl.left_position.get()

    @left.setter
    def left(self, value):
        self.xl.left_position.set(value)

    @property
    def top(self):
        return self.xl.top.get()

    @top.setter
    def top(self, value):
        self.xl.top.set(value)

    @property
    def width(self):
        return self.xl.width.get()

    @width.setter
    def width(self, value):
        self.xl.width.set(value)

    @property
    def height(self):
        return self.xl.height.get()

    @height.setter
    def height(self, value):
        self.xl.height.set(value)

    def delete(self):
        self.xl.delete()


class Pictures(Collection):

    _attr = 'pictures'
    _kw = kw.picture
    _wrap = Picture

    def add(self, filename, link_to_file, save_with_document, left, top, width, height):

        version = VersionNumber(self.parent.book.app.version)

        if not link_to_file and version >= 15:
            # Office 2016 for Mac is sandboxed. This path seems to work without the need of granting access explicitly
            xlwings_picture = os.path.expanduser("~") + '/Library/Containers/com.microsoft.Excel/Data/xlwings_picture.png'
            shutil.copy2(filename, xlwings_picture)
            filename = xlwings_picture

        sheet_index = self.parent.xl.entry_index.get()
        picture = Picture(
            self.parent,
            self.parent.xl.make(
                at=self.parent.book.xl.sheets[sheet_index],
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
            ).name.get()
        )

        if not link_to_file and version >= 15:
            os.remove(filename)

        return picture


class Names:
    def __init__(self, parent, xl):
        self.parent = parent
        self.xl = xl

    def __call__(self, name_or_index):
        return Name(self.parent, xl=self.xl[name_or_index])

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
        return Name(self.parent, self.parent.xl.make(at=self.parent.xl,
                                                     new=kw.named_item,
                                                     with_properties={
                                                         kw.references: refers_to,
                                                         kw.name: name
                                                     }))


class Name:
    def __init__(self, parent, xl):
        self.parent = parent
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
        book = self.parent if isinstance(self.parent, Book) else self.parent.book
        external_address = self.xl.reference_range.get_address(external=True)
        match = re.search(r"\](.*?)'?!(.*)", external_address)
        return Range(Sheet(book, match.group(1)), match.group(2))


class Shapes(Collection):

    _attr = 'shapes'
    _kw = kw.shape
    _wrap = Shape


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


def is_excel_running():
    for proc in psutil.process_iter():
        try:
            if proc.name() == 'Microsoft Excel':
                return True
        except psutil.NoSuchProcess:
            pass
    return False


def _clean_value_data_element(value, datetime_builder, empty_as, number_builder):
    if value == '' or value == kw.missing_value:
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
    elif np and isinstance(x, (np.floating, float)) and np.isnan(x):
        return ""
    elif np and isinstance(x, np.datetime64):
        # handle numpy.datetime64
        return np_datetime_to_datetime(x).replace(tzinfo=None)
    elif np and isinstance(x, np.number):
        return float(x)
    elif pd and isinstance(x, pd.Timestamp):
        # This transformation seems to be only needed on Python 2.6 (?)
        return x.to_pydatetime().replace(tzinfo=None)
    elif pd and isinstance(x, type(pd.NaT)):
        return None
    elif isinstance(x, dt.datetime):
        # Make datetime timezone naive
        return x.replace(tzinfo=None)
    elif isinstance(x, int):
        # appscript packs integers larger than SInt32 but smaller than SInt64 as typeSInt64, and integers
        # larger than SInt64 as typeIEEE64BitFloatingPoint. Excel silently ignores typeSInt64. (GH 227)
        return float(x)

    return x


# --- constants ---

chart_types_k2s = {
    kw.ThreeD_area: '3d_area',
    kw.ThreeD_area_stacked: '3d_area_stacked',
    kw.ThreeD_area_stacked_100: '3d_area_stacked_100',
    kw.ThreeD_bar_clustered: '3d_bar_clustered',
    kw.ThreeD_bar_stacked: '3d_bar_stacked',
    kw.ThreeD_bar_stacked_100: '3d_bar_stacked_100',
    kw.ThreeD_column: '3d_column',
    kw.ThreeD_column_clustered: '3d_column_clustered',
    kw.ThreeD_column_stacked: '3d_column_stacked',
    kw.ThreeD_column_stacked_100: '3d_column_stacked_100',
    kw.ThreeD_line: '3d_line',
    kw.ThreeD_pie: '3d_pie',
    kw.ThreeD_pie_exploded: '3d_pie_exploded',
    kw.area_chart: 'area',
    kw.area_stacked: 'area_stacked',
    kw.area_stacked_100: 'area_stacked_100',
    kw.bar_clustered: 'bar_clustered',
    kw.bar_of_pie: 'bar_of_pie',
    kw.bar_stacked: 'bar_stacked',
    kw.bar_stacked_100: 'bar_stacked_100',
    kw.bubble: 'bubble',
    kw.bubble_ThreeD_effect: 'bubble_3d_effect',
    kw.column_clustered: 'column_clustered',
    kw.column_stacked: 'column_stacked',
    kw.column_stacked_100: 'column_stacked_100',
    kw.combination_chart: 'combination',
    kw.cone_bar_clustered: 'cone_bar_clustered',
    kw.cone_bar_stacked: 'cone_bar_stacked',
    kw.cone_bar_stacked_100: 'cone_bar_stacked_100',
    kw.cone_col: 'cone_col',
    kw.cone_column_clustered: 'cone_col_clustered',
    kw.cone_column_stacked: 'cone_col_stacked',
    kw.cone_column_stacked_100: 'cone_col_stacked_100',
    kw.cylinder_bar_clustered: 'cylinder_bar_clustered',
    kw.cylinder_bar_stacked: 'cylinder_bar_stacked',
    kw.cylinder_bar_stacked_100: 'cylinder_bar_stacked_100',
    kw.cylinder_column: 'cylinder_col',
    kw.cylinder_column_clustered: 'cylinder_col_clustered',
    kw.cylinder_column_stacked: 'cylinder_col_stacked',
    kw.cylinder_column_stacked_100: 'cylinder_col_stacked_100',
    kw.doughnut: 'doughnut',
    kw.doughnut_exploded: 'doughnut_exploded',
    kw.line_chart: 'line',
    kw.line_markers: 'line_markers',
    kw.line_markers_stacked: 'line_markers_stacked',
    kw.line_markers_stacked_100: 'line_markers_stacked_100',
    kw.line_stacked: 'line_stacked',
    kw.line_stacked_100: 'line_stacked_100',
    kw.pie_chart: 'pie',
    kw.pie_exploded: 'pie_exploded',
    kw.pie_of_pie: 'pie_of_pie',
    kw.pyramid_bar_clustered: 'pyramid_bar_clustered',
    kw.pyramid_bar_stacked: 'pyramid_bar_stacked',
    kw.pyramid_bar_stacked_100: 'pyramid_bar_stacked_100',
    kw.pyramid_column: 'pyramid_col',
    kw.pyramid_column_clustered: 'pyramid_col_clustered',
    kw.pyramid_column_stacked: 'pyramid_col_stacked',
    kw.pyramid_column_stacked_100: 'pyramid_col_stacked_100',
    kw.radar: 'radar',
    kw.radar_filled: 'radar_filled',
    kw.radar_markers: 'radar_markers',
    kw.stock_HLC: 'stock_hlc',
    kw.stock_OHLC: 'stock_ohlc',
    kw.stock_VHLC: 'stock_vhlc',
    kw.stock_VOHLC: 'stock_vohlc',
    kw.surface: 'surface',
    kw.surface_top_view: 'surface_top_view',
    kw.surface_top_view_wireframe: 'surface_top_view_wireframe',
    kw.surface_wireframe: 'surface_wireframe',
    kw.xy_scatter_lines: 'xy_scatter_lines',
    kw.xy_scatter_lines_no_markers: 'xy_scatter_lines_no_markers',
    kw.xy_scatter_smooth: 'xy_scatter_smooth',
    kw.xy_scatter_smooth_no_markers: 'xy_scatter_smooth_no_markers',
    kw.xyscatter: 'xy_scatter',
}

chart_types_s2k = {v: k for k, v in chart_types_k2s.items()}

directions_s2k = {
    'd': kw.toward_the_bottom,
    'down': kw.toward_the_bottom,
    'l': kw.toward_the_left,
    'left': kw.toward_the_left,
    'r': kw.toward_the_right,
    'right': kw.toward_the_right,
    'u': kw.toward_the_top,
    'up': kw.toward_the_top
}

directions_k2s = {
    kw.toward_the_bottom: 'down',
    kw.toward_the_left: 'left',
    kw.toward_the_right: 'right',
    kw.toward_the_top: 'up',
}

calculation_k2s = {
    kw.calculation_automatic: 'automatic',
    kw.calculation_manual: 'manual',
    kw.calculation_semiautomatic: 'semiautomatic'
}

calculation_s2k = {v: k for k, v in calculation_k2s.items()}

shape_types_k2s = {
    kw.shape_type_auto: 'auto_shape',
    kw.shape_type_callout: 'callout',
    kw.shape_type_canvas: 'canvas',
    kw.shape_type_chart: 'chart',
    kw.shape_type_comment: 'comment',
    kw.shape_type_content_application: 'content_app',
    kw.shape_type_diagram: 'diagram',
    kw.shape_type_free_form: 'free_form',
    kw.shape_type_group: 'group',
    kw.shape_type_embedded_OLE_control: 'embedded_ole_object',
    kw.shape_type_form_control: 'form_control',
    kw.shape_type_line: 'line',
    kw.shape_type_linked_OLE_object: 'linked_ole_object',
    kw.shape_type_linked_picture: 'linked_picture',
    kw.shape_type_OLE_control: 'ole_control_object',
    kw.shape_type_picture: 'picture',
    kw.shape_type_place_holder: 'placeholder',
    kw.shape_type_web_video: 'web_video',
    kw.shape_type_media: 'media',
    kw.shape_type_text_box: 'text_box',
    kw.shape_type_table: 'table',
    kw.shape_type_ink: 'ink',
    kw.shape_type_ink_comment: 'ink_comment',
}

shape_types_s2k = {v: k for k, v in shape_types_k2s.items()}
