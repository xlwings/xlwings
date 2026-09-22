import atexit
import datetime as dt
import numbers
import os
import re
import shutil
import struct
import subprocess
from collections import Counter
from contextlib import contextmanager
from functools import cache
from pathlib import Path
from uuid import uuid4
from weakref import WeakValueDictionary

import aem
import appscript
import osax
import psutil
from appscript import its, k as kw, mactypes
from appscript.reference import CommandError, Reference

import xlwings

from . import base_classes, mac_dict, utils
from ._names import NameIndex
from .constants import ColorIndex
from .utils import (
    VersionNumber,
    col_name,
    fullname_url_to_local_path,
    int_to_rgb,
    np_datetime_to_datetime,
    read_config_sheet,
)

try:
    import pandas as pd
except ImportError:
    pd = None
try:
    import numpy as np
except ImportError:
    np = None
try:
    from PIL import ImageGrab
except ImportError:
    PIL = None


# Time types
time_types = (dt.date, dt.datetime)
if np:
    time_types = time_types + (np.datetime64,)

cell_errors = (
    "#DIV/0!",
    "#N/A",
    "#NAME?",
    "#NULL!",
    "#NUM!",
    "#REF!",
    "#VALUE!",
)


def _parse_pid(pid_info):
    match = re.search(r'^\s*"?pid"?\s*=\s*(\d+)(?!\w)', pid_info, re.M)
    return int(match.group(1)) if match else None


def _clean_value_data_element(
    value, datetime_builder, empty_as, number_builder, err_to_str
):
    if value == "" or value == kw.missing_value:
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
            tzinfo=None,
        )
    elif number_builder is not None and isinstance(value, float):
        value = number_builder(value)
    return value


class Engine:
    @property
    def apps(self):
        return Apps()

    @property
    def name(self):
        return "excel"

    @property
    def type(self):
        return "desktop"

    @staticmethod
    def prepare_xl_data_element(x, options):
        if x is None:
            return ""
        elif pd and pd.isna(x):
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
        elif isinstance(x, bool):
            # Must be tested before int!
            return x
        elif isinstance(x, int):
            # appscript packs integers larger than SInt32 but smaller than SInt64 as
            # typeSInt64, and integers larger than SInt64 as typeIEEE64BitFloatingPoint.
            # Excel silently ignores typeSInt64. (GH 227)
            return float(x)
        return x

    @staticmethod
    def clean_value_data(data, datetime_builder, empty_as, number_builder, err_to_str):
        return [
            [
                _clean_value_data_element(
                    c, datetime_builder, empty_as, number_builder, err_to_str
                )
                for c in row
            ]
            for row in data
        ]


engine = Engine()


class Apps(base_classes.Apps):
    def _iter_excel_instances(self):
        asn = subprocess.check_output(
            ["lsappinfo", "visibleprocesslist", "-includehidden"]
        ).decode("utf-8")
        for asn in asn.split(" "):
            if "Microsoft_Excel" in asn:
                pid_info = subprocess.check_output(
                    ["lsappinfo", "info", "-only", "pid", asn]
                ).decode("utf-8")
                pid = _parse_pid(pid_info)
                if pid is not None:
                    yield pid

    def keys(self):
        return list(self._iter_excel_instances())

    def add(self, spec=None, add_book=None, visible=None):
        return App(spec=spec, add_book=add_book, visible=visible)

    def __iter__(self):
        for pid in self._iter_excel_instances():
            yield App(xl=pid)

    def __len__(self):
        return len(list(self._iter_excel_instances()))

    def __getitem__(self, pid):
        if pid not in self.keys():
            raise KeyError("Could not find an Excel instance with this PID.")
        return App(xl=pid)


class App(base_classes.App):
    def __init__(self, spec=None, add_book=None, xl=None, visible=True):
        if xl is None:
            self._xl = appscript.app(
                name=spec or "Microsoft Excel",
                newinstance=True,
                terms=mac_dict,
                hide=not visible,
            )
            if visible:
                self.activate()  # Makes it behave like on Windows
        elif isinstance(xl, int):
            self._xl = appscript.app(pid=xl, terms=mac_dict)
        else:
            self._xl = xl

    @property
    def xl(self):
        return self._xl

    @xl.setter
    def xl(self, value):
        self._xl = value

    @property
    def api(self):
        return self.xl

    @property
    def engine(self):
        return engine

    @property
    def path(self):
        return hfs_to_posix_path(self.xl.path.get())

    @property
    def pid(self):
        data = (
            self.xl.AS_appdata.target()
            .addressdesc.coerce(aem.kae.typeKernelProcessID)
            .data
        )
        (pid,) = struct.unpack("i", data)
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
        asn = subprocess.check_output(
            ["lsappinfo", "visibleprocesslist", "-includehidden"]
        ).decode("utf-8")
        frontmost_asn = asn.split(" ")[0]
        pid_info_frontmost = subprocess.check_output(
            ["lsappinfo", "info", "-only", "pid", frontmost_asn]
        ).decode("utf-8")
        pid_frontmost = _parse_pid(pid_info_frontmost)
        try:
            appscript.app("System Events").processes[
                its.unix_id == self.pid
            ].frontmost.set(True)
            if not steal_focus and pid_frontmost is not None:
                appscript.app("System Events").processes[
                    its.unix_id == pid_frontmost
                ].frontmost.set(True)
        except CommandError:
            pass  # may require root privileges (GH 1966)

    @property
    def visible(self):
        try:
            return (
                appscript.app("System Events")
                .processes[its.unix_id == self.pid]
                .visible.get()[0]
            )
        except CommandError:
            return None  # may require root privileges (GH 1966)

    @visible.setter
    def visible(self, visible):
        try:
            appscript.app("System Events").processes[
                its.unix_id == self.pid
            ].visible.set(visible)
        except CommandError:
            pass  # may require root privileges (GH 1966)

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
    def enable_events(self):
        return self.xl.enable_events.get()

    @enable_events.setter
    def enable_events(self, value):
        self.xl.enable_events.set(value)

    @property
    def interactive(self):
        # TODO: replace with specific error when Exceptions are refactored
        raise xlwings.XlwingsError(
            "Getting or setting 'app.interactive' isn't supported on macOS."
        )

    @interactive.setter
    def interactive(self, value):
        # TODO: replace with specific error when Exceptions are refactored
        raise xlwings.XlwingsError(
            "Getting or setting 'app.interactive' isn't supported on macOS."
        )

    @property
    def startup_path(self):
        return hfs_to_posix_path(self.xl.startup_path.get())

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

    @property
    def hwnd(self):
        return None

    def run(self, macro, args):
        kwargs = {"arg{0}".format(i): n for i, n in enumerate(args, 1)}
        return self.xl.run_VB_macro(macro, **kwargs)

    @property
    def status_bar(self):
        return self.xl.status_bar.get()

    @status_bar.setter
    def status_bar(self, value):
        self.xl.status_bar.set(value)

    @property
    def cut_copy_mode(self):
        modes = {kw.cut_mode: "cut", kw.copy_mode: "copy"}
        return modes.get(self.xl.cut_copy_mode.get())

    @cut_copy_mode.setter
    def cut_copy_mode(self, value):
        self.xl.cut_copy_mode.set(value)

    def alert(self, prompt, title, buttons, mode, callback):
        # OSAX (Open Scripting Architecture Extension) instance for StandardAdditions
        # See /System/Library/ScriptingAdditions/StandardAdditions.osax
        sa = osax.OSAX(pid=self.pid)
        sa.activate()  # Activate app so dialog box will be visible
        modes = {
            None: kw.informational,
            "info": kw.informational,
            "critical": kw.critical,
        }
        buttons_dict = {
            None: "OK",
            "ok": "OK",
            "ok_cancel": ["Cancel", "OK"],
            "yes_no": ["No", "Yes"],
            "yes_no_cancel": ["Cancel", "No", "Yes"],
        }
        rv = sa.display_alert(
            "" if title is None else title,
            message="" if prompt is None else prompt,
            buttons=buttons_dict[buttons],
            as_=modes[mode],
        )
        return rv[kw.button_returned].lower()


class Books(base_classes.Books):
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
        if self.app.visible:
            self.app.activate()
        xl = self.app.xl.make(new=kw.workbook)
        wb = Book(self.app, xl.name.get())
        return wb

    def open(
        self,
        fullname,
        update_links=None,
        read_only=None,
        format=None,
        password=None,
        write_res_password=None,
        ignore_read_only_recommended=None,
        origin=None,
        delimiter=None,
        editable=None,
        notify=None,
        converter=None,
        add_to_mru=None,
        local=None,
        corrupt_load=None,
    ):
        # TODO: format and origin currently require a native appscript keyword,
        #  read_only doesn't seem to work
        # Unsupported params
        if local is not None:
            # TODO: replace with specific error when Exceptions are refactored
            raise xlwings.XlwingsError("local is not supported on macOS")
        if corrupt_load is not None:
            # TODO: replace with specific error when Exceptions are refactored
            raise xlwings.XlwingsError("corrupt_load is not supported on macOS")
        # update_links: on Windows only constants 0 and 3 seem to be supported in
        # this context
        if update_links:
            update_links = kw.update_remote_and_external_links
        else:
            update_links = kw.do_not_update_links
        if self.app.visible:
            self.app.activate()
        filename = os.path.basename(fullname)
        self.app.xl.open_workbook(
            workbook_file_name=fullname,
            update_links=update_links,
            read_only=read_only,
            format=format,
            password=password,
            write_reserved_password=write_res_password,
            ignore_read_only_recommended=ignore_read_only_recommended,
            origin=origin,
            delimiter=delimiter,
            editable=editable,
            notify=notify,
            converter=converter,
            add_to_mru=add_to_mru,
            timeout=-1,
        )
        wb = Book(self.app, filename)
        return wb

    def __iter__(self):
        n = len(self)
        for i in range(n):
            yield Book(self.app, i + 1)


class Book(base_classes.Book):
    def __init__(self, app, name_or_index):
        self._app = app
        self.xl = app.xl.workbooks[name_or_index]

    @property
    def app(self):
        return self._app

    @property
    def api(self):
        return self.xl

    def json(self):
        raise NotImplementedError()

    @property
    def name(self):
        return self.xl.name.get()

    @property
    def sheets(self):
        return Sheets(self)

    def close(self):
        pivot_key = (self.app.pid, self.name)
        self.xl.close(saving=kw.no)
        _invalidate_pivot_value_states(pivot_key)

    def save(self, path, password):
        saved_path = self.xl.properties().get(kw.path)
        source_ext = os.path.splitext(self.name)[1] if saved_path else None
        target_ext = os.path.splitext(path)[1] if path else ".xlsx"
        if saved_path and source_ext == target_ext:
            file_format = self.xl.properties().get(kw.file_format)
        else:
            ext_to_file_format = {
                ".xlsx": kw.Excel_XML_file_format,
                ".xlsm": kw.macro_enabled_XML_file_format,
                ".xlsb": kw.Excel_binary_file_format,
                ".xltm": kw.macro_enabled_template_file_format,
                ".xltx": kw.template_file_format,
                ".xlam": kw.add_in_file_format,
                ".xls": kw.Excel98to2004_file_format,
                ".xlt": kw.Excel98to2004_template_file_format,
                ".xla": kw.Excel98to2004_add_in_file_format,
            }

            file_format = ext_to_file_format[target_ext]
        if (saved_path != "") and (path is None):
            # Previously saved: Save under existing name
            self.xl.save(timeout=-1)
        elif (
            (saved_path != "") and (path is not None) and (os.path.split(path)[0] == "")
        ):
            # Save existing book under new name in cwd if no path has been provided
            save_as_name = path
            path = os.path.join(os.getcwd(), path)
            hfs_path = posix_to_hfs_path(os.path.realpath(path))
            self.xl.save_workbook_as(
                filename=hfs_path,
                overwrite=True,
                file_format=file_format,
                timeout=-1,
                password=password,
            )
            self.xl = self.app.xl.workbooks[save_as_name]
        elif (saved_path == "") and (path is None):
            # Previously unsaved: Save under current name in current working directory
            save_as_name = self.xl.name.get() + ".xlsx"
            path = os.path.join(os.getcwd(), save_as_name)
            hfs_path = posix_to_hfs_path(os.path.realpath(path))
            self.xl.save_workbook_as(
                filename=hfs_path,
                overwrite=True,
                file_format=file_format,
                timeout=-1,
                password=password,
            )
            self.xl = self.app.xl.workbooks[save_as_name]
        elif path:
            # Save under new name/location
            hfs_path = posix_to_hfs_path(os.path.realpath(path))
            self.xl.save_workbook_as(
                filename=hfs_path,
                overwrite=True,
                file_format=file_format,
                timeout=-1,
                password=password,
            )
            self.xl = self.app.xl.workbooks[os.path.basename(path)]

    @property
    def fullname(self):
        display_alerts = self.app.display_alerts
        self.app.display_alerts = False
        # This causes a pop-up if there's a pw protected sheet, see #1377
        path = self.xl.properties().get(kw.full_name)
        if "://" in path:
            config = read_config_sheet(xlwings.Book(impl=self))
            self.app.display_alerts = display_alerts
            return fullname_url_to_local_path(
                url=path,
                sheet_onedrive_consumer_config=config.get("ONEDRIVE_CONSUMER_MAC"),
                sheet_onedrive_commercial_config=config.get("ONEDRIVE_COMMERCIAL_MAC"),
                sheet_sharepoint_config=config.get("SHAREPOINT_MAC"),
            )
        else:
            self.app.display_alerts = display_alerts
            return path

    @property
    def names(self):
        return Names(parent=self, xl=self.xl.named_items)

    def activate(self):
        self.xl.activate_object()

    def to_pdf(self, path, quality=None):
        # quality parameter for compatibility
        hfs_path = posix_to_hfs_path(path)
        display_alerts = self.app.display_alerts
        self.app.display_alerts = False
        if Path(path).exists():
            # Errors out with Parameter error (OSERROR: -50) otherwise
            os.unlink(path)
        self.xl.save(in_=hfs_path, as_=kw.PDF_file_format)
        self.app.display_alerts = display_alerts


class Sheets(base_classes.Sheets):
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

    def add(self, before=None, after=None, name=None):
        if before is None and after is None:
            before = self.workbook.app.books.active.sheets.active
        if before:
            position = before.xl.before
        else:
            position = after.xl.after
        xl = self.workbook.xl.make(new=kw.worksheet, at=position)
        if name is not None:
            xl.name.set(name)
            xl = self.workbook.xl.worksheets[name]
        return Sheet(self.workbook, xl.name.get())


class Sheet(base_classes.Sheet):
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
                    raise IndexError(
                        "Attempted to access 0-based Range. "
                        "xlwings/Excel Ranges are 1-based."
                    )
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
            address1 = arg1.split(":")[0]
        else:
            raise ValueError("Invalid parameters")

        if isinstance(arg2, tuple):
            if 0 in arg2:
                raise IndexError(
                    "Attempted to access 0-based Range. "
                    "xlwings/Excel Ranges are 1-based."
                )
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
            if isinstance(arg1, str) and len(arg1.split(":")) == 2:
                address2 = arg1.split(":")[1]
            else:
                return Range(self, "{0}".format(address1))
        else:
            raise ValueError("Invalid parameters")

        return Range(self, "{0}:{1}".format(address1, address2))

    @property
    def cells(self):
        return self.range(
            (1, 1), (self.xl.count(each=kw.row), self.xl.count(each=kw.column))
        )

    def activate(self):
        self.xl.activate_object()

    def select(self):
        self.xl.select()

    def clear_contents(self):
        self.xl.used_range.clear_contents()

    def clear_formats(self):
        self.xl.used_range.clear_formats()

    def clear(self):
        self.xl.used_range.clear_range()

    def autofit(self, axis=None):
        num_columns = self.xl.count(each=kw.column)
        num_rows = self.xl.count(each=kw.row)
        address = self.range((1, 1), (num_rows, num_columns)).address
        alerts_state = self.book.app.screen_updating
        self.book.app.screen_updating = False
        if axis == "rows" or axis == "r":
            self.xl.rows[address].autofit()
        elif axis == "columns" or axis == "c":
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

    def copy(self, before, after):
        if before:
            before = before.xl
        if after:
            after = after.xl
        self.xl.copy_worksheet(before_=before, after_=after)

    def move(self, before, after):
        destination = before.xl if before else after.xl.after
        self.xl.move(to=destination)

    @property
    def charts(self):
        return Charts(self)

    @property
    def shapes(self):
        return Shapes(self)

    @property
    def tables(self):
        return Tables(self)

    @property
    def pivot_tables(self):
        return PivotTables(self)

    @property
    def pictures(self):
        return Pictures(self)

    @property
    def used_range(self):
        return Range(self, self.xl.used_range.get_address())

    @property
    def visible(self):
        return True if self.xl.visible.get() == kw.sheet_visible else False

    @visible.setter
    def visible(self, value):
        self.xl.visible.set(value)

    def _window_property(self, name, *value):
        """Get (no `value`) or set (one `value`) a property of the book's window.

        Gridlines and the like are window properties, applying to the sheet
        that the window currently shows. A sheet that isn't active is
        therefore activated for the duration of the call and the previously
        active sheet restored afterwards, with screen updating off.
        """
        book = self.workbook.xl
        prop = getattr(book.windows[1], name)
        previous_book_sheet = book.active_sheet.name.get()
        if previous_book_sheet != self.xl.name.get():
            if self.xl.visible.get() != kw.sheet_visible:
                raise ValueError(
                    f"Sheet.{name}: hidden sheets can't be activated. Set "
                    "sheet.visible = True first."
                )
            app = self.workbook.app
            # Resolve now: appscript references are lazy, and `active_sheet`
            # would otherwise point at whatever is active by the time it's used.
            previous_book = app.xl.active_workbook.name.get()
            previous_sheet = app.xl.active_workbook.active_sheet.name.get()
            previous_screen_updating = app.screen_updating
            app.screen_updating = False
            try:
                self.xl.activate_object()
                if value:
                    prop.set(value[0])
                    return None
                return prop.get()
            finally:
                try:
                    book.sheets[previous_book_sheet].activate_object()
                finally:
                    try:
                        app.xl.workbooks[previous_book].sheets[
                            previous_sheet
                        ].activate_object()
                    finally:
                        app.screen_updating = previous_screen_updating
        if value:
            prop.set(value[0])
            return None
        return prop.get()

    @property
    def show_gridlines(self):
        return bool(self._window_property("display_gridlines"))

    @show_gridlines.setter
    def show_gridlines(self, value):
        self._window_property("display_gridlines", bool(value))

    @property
    def page_setup(self):
        return PageSetup(self, self.xl.page_setup_object)


_CONDITIONAL_FORMAT_TYPE_FROM_KW = {
    kw.cell_value: "cell_value",
    kw.expression: "custom",
    kw.color_scale: "color_scale",
    kw.databar: "data_bar",
    kw.icon_sets: "icon_set",
}
_CONDITIONAL_FORMAT_SPECIALIZED_COLLECTION_FROM_KW = {
    kw.color_scale: "color_scale_format_condition",
    kw.databar: "databar_format_condition",
    kw.icon_sets: "icon_set_format_condition",
}
_CONDITIONAL_FORMAT_OPERATOR_TO_KW = {
    "between": kw.operator_between,
    "not_between": kw.operator_not_between,
    "equal_to": kw.operator_equal,
    "not_equal_to": kw.operator_not_equal,
    "greater_than": kw.operator_greater,
    "less_than": kw.operator_less,
    "greater_than_or_equal": kw.operator_greater_equal,
    "less_than_or_equal": kw.operator_less_equal,
}
_CONDITIONAL_FORMAT_OPERATOR_FROM_KW = {
    value: key for key, value in _CONDITIONAL_FORMAT_OPERATOR_TO_KW.items()
}
_CONDITIONAL_FORMAT_THRESHOLD_TO_KW = {
    "lowest_value": kw.condition_value_lowest_value,
    "highest_value": kw.condition_value_highest_value,
    "number": kw.condition_value_number,
    "percent": kw.condition_value_percent,
    "percentile": kw.condition_value_percentile,
    "formula": kw.condition_value_formula,
}
_CONDITIONAL_FORMAT_THRESHOLD_FROM_KW = {
    **{value: key for key, value in _CONDITIONAL_FORMAT_THRESHOLD_TO_KW.items()},
    kw.condition_value_automatic_minimum: "automatic",
    kw.condition_value_automatic_maximum: "automatic",
}
_CONDITIONAL_FORMAT_ICON_SET_TO_KW = {
    "3_arrows": kw.icon_set_3_arrows,
    "3_arrows_gray": kw.icon_set_3_arrows_gray,
    "3_flags": kw.icon_set_3_flags,
    "3_traffic_lights_1": kw.icon_set_3_traffic_lights_1,
    "3_traffic_lights_2": kw.icon_set_3_traffic_lights_2,
    "3_signs": kw.icon_set_3_signs,
    "3_symbols": kw.icon_set_3_symbols,
    "3_symbols_2": kw.icon_set_3_symbols_2,
    "4_arrows": kw.icon_set_4_arrows,
    "4_arrows_gray": kw.icon_set_4_arrows_gray,
    "4_red_to_black": kw.icon_set_4_red_to_black,
    "4_rating": kw.icon_set_4_CRV,
    "4_traffic_lights": kw.icon_set_4_traffic_lights,
    "5_arrows": kw.icon_set_5_arrows,
    "5_arrows_gray": kw.icon_set_5_arrows_gray,
    "5_rating": kw.icon_set_5_CRV,
    "5_quarters": kw.icon_set_5_quarters,
    "3_stars": kw.icon_set_3_stars,
    "3_triangles": kw.icon_set_3_triangles,
    "5_boxes": kw.icon_set_5_boxes,
}
_CONDITIONAL_FORMAT_ICON_SET_FROM_KW = {
    value: key for key, value in _CONDITIONAL_FORMAT_ICON_SET_TO_KW.items()
}


@cache
def _conditional_format_icon_set_indexes():
    """Derive workbook IconSets indexes from Excel's generated enum values."""
    terminology = dict(mac_dict.enums)
    codes = {
        name: terminology[keyword.AS_name]
        for name, keyword in _CONDITIONAL_FORMAT_ICON_SET_TO_KW.items()
    }
    prefixes = {code[:2] for code in codes.values() if len(code) == 4}
    indexes = {name: int.from_bytes(code[2:], "big") for name, code in codes.items()}
    expected = set(range(1, len(codes) + 1))
    if len(prefixes) != 1 or set(indexes.values()) != expected:
        raise RuntimeError(
            "Excel's generated icon-set enumeration no longer matches its "
            "workbook IconSets collection."
        )
    return indexes


_CONDITIONAL_FORMAT_ICON_SET_INDEX = {
    name: index
    for index, name in enumerate(_CONDITIONAL_FORMAT_ICON_SET_TO_KW, start=1)
}


class Range(base_classes.Range):
    def __init__(self, sheet, address):
        self.sheet = sheet
        self.options = None  # Assigned by main.Range to keep API of sheet.range clean
        if isinstance(address, tuple):
            self._coords = address
            row, col, nrows, ncols = address
            if nrows and ncols:
                self.xl = sheet.xl.cells[
                    "%s:%s"
                    % (
                        sheet.xl.rows[row].columns[col].get_address(),
                        sheet.xl.rows[row + nrows - 1]
                        .columns[col + ncols - 1]
                        .get_address(),
                    )
                ]
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
                self.xl.count(each=kw.column),
            )
        return self._coords

    @property
    def api(self):
        return self.xl

    @property
    def autofilter(self):
        return AutoFilter(self)

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
    def max_cells_per_read(self):
        # Below the shared 4M default: a single Apple Event reply has a size
        # cap, and exceeding it fails with -1741 (buffer for AEFlattenDesc too
        # small). Measured with float cells: 2,000,000 fail, 1,500,000 work
        # (10M cells read in 35s vs. 44s with a 1M budget). The cap is in bytes,
        # not cells, so a text-heavy range can still hit it and needs a smaller
        # explicit chunksize.
        return 1_500_000

    @property
    def raw_value(self):
        def ensure_2d(values):
            # Usually done in converter, but macOS doesn't deliver any info about
            # errors with values
            if not isinstance(values, list):
                return [[values]]
            elif not isinstance(values[0], list):
                return [values]

        if self.xl is not None:
            values = self.xl.value.get()
            if self.options.get("err_to_str", False):
                string_values = self.xl.string_value.get()
                values = ensure_2d(values)
                string_values = ensure_2d(string_values)
                for row_ix, row in enumerate(string_values):
                    for col_ix, c in enumerate(row):
                        if c in cell_errors:
                            values[row_ix][col_ix] = c
            return values

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

    def clear_formats(self):
        if self.xl is not None:
            alerts_state = self.sheet.book.app.screen_updating
            self.sheet.book.app.screen_updating = False
            self.xl.clear_formats()
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
    def formula2(self):
        if self.xl is not None:
            return self.xl.formula2.get()

    @formula2.setter
    def formula2(self, value):
        if self.xl is not None:
            self.xl.formula2.set(value)

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
    def font(self):
        return Font(self, self.xl.font_object)

    @property
    def borders(self):
        return Borders(self, self.xl)

    @property
    def data_validation(self):
        return DataValidation(self)

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
    def has_array(self):
        if self.xl is not None:
            return self.xl.has_array.get()

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
            return self.xl.get_address(
                row_absolute=row_absolute,
                column_absolute=col_absolute,
                external=external,
            )

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
            if axis == "rows" or axis == "r":
                self.sheet.xl.rows[address].autofit()
            elif axis == "columns" or axis == "c":
                self.sheet.xl.columns[address].autofit()
            elif axis is None:
                self.sheet.xl.rows[address].autofit()
                self.sheet.xl.columns[address].autofit()
            self.sheet.book.app.screen_updating = alerts_state

    def insert(self, shift=None, copy_origin=None):
        # copy_origin is not supported on mac
        shifts = {"down": kw.shift_down, "right": kw.shift_to_right, None: None}
        self.xl.insert_into_range(shift=shifts[shift])

    def delete(self, shift=None):
        shifts = {"up": kw.shift_up, "left": kw.shift_to_left, None: None}
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
            None: None,
        }

        operations = {
            "add": kw.paste_special_operation_add,
            "divide": kw.paste_special_operation_divide,
            "multiply": kw.paste_special_operation_multiply,
            "subtract": kw.paste_special_operation_subtract,
            None: None,
        }

        self.xl.paste_special(
            what=pastes[paste],
            operation=operations[operation],
            skip_blanks=skip_blanks,
            transpose=transpose,
        )

    @property
    def hyperlink(self):
        try:
            return self.xl.hyperlinks[1].address.get()
        except CommandError:
            raise Exception("The cell doesn't seem to contain a hyperlink!")

    def add_hyperlink(self, address, text_to_display=None, screen_tip=None):
        if self.xl is not None:
            self.xl.make(
                at=self.xl,
                new=kw.hyperlink,
                with_properties={
                    kw.address: address,
                    kw.text_to_display: text_to_display,
                    kw.screen_tip: screen_tip,
                },
            )

    @property
    def color(self):
        if (
            not self.xl
            or self.xl.interior_object.color_index.get() == kw.color_index_none
        ):
            return None
        else:
            return tuple(self.xl.interior_object.color.get())

    @color.setter
    def color(self, color_or_rgb):
        if isinstance(color_or_rgb, str):
            color_or_rgb = utils.hex_to_rgb(color_or_rgb)
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
            return Range(
                self.sheet,
                self.sheet.xl.rows[self.row + arg1 - 1]
                .columns[self.column + arg2 - 1]
                .get_address(),
            )

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
            sht.range((row1, col + i), (row2, col + i)) for i in range(self.shape[1])
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

    @property
    def table(self):
        if self.xl.list_object.name.get() == kw.missing_value:
            return None
        else:
            return Table(self.sheet, self.xl.list_object.name.get())

    @property
    def characters(self):
        # This is broken with AppleScript/Excel 2016
        return Characters(parent=self, xl=self.xl.characters)

    @property
    def wrap_text(self):
        return self.xl.wrap_text.get()

    @wrap_text.setter
    def wrap_text(self, value):
        self.xl.wrap_text.set(value)

    @property
    def horizontal_alignment(self):
        # A range whose cells disagree returns an AEEnum that isn't one of the
        # documented keywords, so .get() maps it to None.
        return horizontal_alignments_k2s.get(self.xl.horizontal_alignment.get())

    @horizontal_alignment.setter
    def horizontal_alignment(self, value):
        self.xl.horizontal_alignment.set(horizontal_alignments_s2k[value])

    @property
    def vertical_alignment(self):
        return vertical_alignments_k2s.get(self.xl.vertical_alignment.get())

    @vertical_alignment.setter
    def vertical_alignment(self, value):
        self.xl.vertical_alignment.set(vertical_alignments_s2k[value])

    @property
    def note(self):
        try:
            # No easy way to check whether there's a comment like on Windows
            return (
                Note(parent=self, xl=self.xl.Excel_comment)
                if self.xl.Excel_comment.Excel_comment_text()
                else None
            )
        except appscript.reference.CommandError:
            return None

    @property
    def conditional_formats(self):
        return ConditionalFormats(self)

    def copy_picture(self, appearance, format):
        _appearance = {"screen": kw.screen, "printer": kw.printer}
        _format = {"picture": kw.picture, "bitmap": kw.bitmap}
        self.xl.copy_picture(appearance=_appearance[appearance], format=_format[format])

    def to_png(self, path):
        self.copy_picture(appearance="screen", format="bitmap")
        im = ImageGrab.grabclipboard()
        im.save(path)

    def to_pdf(self, path, quality=None):
        raise xlwings.XlwingsError("Range.to_pdf() isn't supported on macOS.")

    def autofill(self, destination, type_):
        types = {
            "fill_copy": kw.fill_copy,
            "fill_days": kw.fill_days,
            "fill_default": kw.fill_default,
            "fill_formats": kw.fill_formats,
            "fill_months": kw.fill_months,
            "fill_series": kw.fill_series,
            "fill_values": kw.fill_values,
            "fill_weekdays": kw.fill_weekdays,
            "fill_years": kw.fill_years,
            "growth_trend": kw.growth_trend,
            "linear_trend": kw.linear_trend,
            "flash_fill": kw.flashfill,
        }
        self.xl.autofill(destination=destination.api, type=types[type_])


class Shape(base_classes.Shape):
    def __init__(self, parent, key):
        self._parent = parent
        self.xl = parent.xl.shapes[key]

    @property
    def parent(self):
        return self._parent

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

    def scale_height(self, factor, relative_to_original_size, scale):
        self.xl.scale_height(
            scale=scaling[scale],
            relative_to_original_size=relative_to_original_size,
            factor=factor,
        )

    def scale_width(self, factor, relative_to_original_size, scale):
        self.xl.scale_width(
            scale=scaling[scale],
            relative_to_original_size=relative_to_original_size,
            factor=factor,
        )

    @property
    def text(self):
        if self.xl.shape_text_frame.has_text.get():
            return self.xl.shape_text_frame.text_range.content.get()

    @text.setter
    def text(self, value):
        self.xl.shape_text_frame.text_range.content.set(value)

    @property
    def font(self):
        return Font(self, self.xl.shape_text_frame.text_range.font)

    @property
    def characters(self):
        raise AttributeError("Characters isn't supported on macOS with shapes.")


# None is the normalized form of "none", see main.Borders
_BORDER_LINE_STYLE_TO_KW = {
    "continuous": kw.continuous,
    "dash": kw.dash,
    "dash_dot": kw.dash_dot,
    "dash_dot_dot": kw.dash_dot_dot,
    "dot": kw.dot,
    "double": kw.double,
    "slant_dash_dot": kw.slant_dash_dot,
    None: kw.line_style_none,
}
_BORDER_LINE_STYLE_FROM_KW = {
    keyword: name or "none" for name, keyword in _BORDER_LINE_STYLE_TO_KW.items()
}
_BORDER_WEIGHT_TO_KW = {
    "hairline": kw.border_weight_hairline,
    "thin": kw.border_weight_thin,
    "medium": kw.border_weight_medium,
    "thick": kw.border_weight_thick,
}
_BORDER_WEIGHT_FROM_KW = {
    keyword: name for name, keyword in _BORDER_WEIGHT_TO_KW.items()
}


class Border(base_classes.Border):
    def __init__(self, parent, side, xl):
        # xl is the reference returned by range.get_border(which_border=...)
        self.parent = parent
        self.side = side
        self.xl = xl

    @property
    def api(self):
        return self.xl

    def _value(self, attribute):
        """The value Excel reports for the range as a whole.

        Excel doesn't flag a range whose cells disagree, so a mixed range
        reports one of its values rather than None, and a multi-cell range
        reports no color for its diagonals even right after one was set.
        Read a single cell to get an unambiguous answer.
        """
        if self.xl is None:
            return None
        if attribute == "color" and self.xl.color_index.get() == kw.color_index_none:
            return None
        value = getattr(self.xl, attribute).get()
        if value == kw.missing_value:
            return None
        return tuple(value) if attribute == "color" else value

    @property
    def line_style(self):
        return _BORDER_LINE_STYLE_FROM_KW.get(self._value("line_style"))

    @line_style.setter
    def line_style(self, value):
        if self.xl is not None:
            self.xl.line_style.set(_BORDER_LINE_STYLE_TO_KW[value])

    @property
    def weight(self):
        # The dictionary's "weight" is the enum; "line_weight" is a plain int
        return _BORDER_WEIGHT_FROM_KW.get(self._value("weight"))

    @weight.setter
    def weight(self, value):
        if self.xl is not None:
            self.xl.weight.set(_BORDER_WEIGHT_TO_KW[value])

    @property
    def color(self):
        return self._value("color")

    @color.setter
    def color(self, color_or_rgb):
        if isinstance(color_or_rgb, str):
            color_or_rgb = utils.hex_to_rgb(color_or_rgb)
        if self.xl is not None:
            if isinstance(color_or_rgb, int):
                self.xl.color.set(int_to_rgb(color_or_rgb))
            else:
                self.xl.color.set(color_or_rgb)


class DataValidation(base_classes.DataValidation):
    _TYPE_FROM_KW = {
        kw.validate_whole_number: "whole_number",
        kw.validate_decimal: "decimal",
        kw.validate_list: "list",
        kw.validated_date: "date",
        kw.validate_time: "time",
        kw.validate_text_length: "text_length",
        kw.validate_custom: "custom",
    }
    _TYPE_TO_KW = {value: key for key, value in _TYPE_FROM_KW.items()}
    _OPERATOR_FROM_KW = {
        kw.operator_between: "between",
        kw.operator_not_between: "not_between",
        kw.operator_equal: "equal_to",
        kw.operator_not_equal: "not_equal_to",
        kw.operator_greater: "greater_than",
        kw.operator_less: "less_than",
        kw.operator_greater_equal: "greater_than_or_equal",
        kw.operator_less_equal: "less_than_or_equal",
    }
    _OPERATOR_TO_KW = {value: key for key, value in _OPERATOR_FROM_KW.items()}
    _ALERT_STYLE_FROM_KW = {
        kw.valid_alert_stop: "stop",
        kw.valid_alert_warning: "warning",
        kw.valid_alert_information: "information",
    }

    def __init__(self, parent):
        self.parent = parent

    @property
    def api(self):
        return self.parent.xl.validation

    def _nonuniform_type(self):
        try:
            validation_cells = self.parent.xl.special_cells(
                type=kw.cell_type_all_validation
            )
            intersection = self.parent.sheet.book.app.xl.intersect(
                range1=self.parent.xl,
                range2=validation_cells,
            )
        except CommandError:
            return "none"
        try:
            validated_count = intersection.count(each=kw.cell)
        except CommandError:
            return "inconsistent"
        target_count = self.parent.shape[0] * self.parent.shape[1]
        return "mixed_criteria" if validated_count < target_count else "inconsistent"

    @property
    def type(self):
        try:
            native_type = self.parent.xl.validation.validation_type.get()
        except CommandError:
            return self._nonuniform_type()
        if native_type == kw.missing_value:
            return self._nonuniform_type()
        return self._TYPE_FROM_KW.get(native_type, "unknown")

    def _property(self, name):
        if self.type in {"none", "mixed_criteria", "inconsistent"}:
            return None
        try:
            value = getattr(self.parent.xl.validation, name).get()
        except CommandError:
            return None
        return None if value == kw.missing_value else value

    @property
    def operator(self):
        if self.type not in {
            "whole_number",
            "decimal",
            "date",
            "time",
            "text_length",
        }:
            return None
        return self._OPERATOR_FROM_KW.get(self._property("validation_operator"))

    @property
    def formula1(self):
        if self.type not in {
            "whole_number",
            "decimal",
            "date",
            "time",
            "text_length",
        }:
            return None
        value = self._property("formula1")
        return None if value is None else str(value)

    @property
    def formula2(self):
        if self.operator not in {"between", "not_between"}:
            return None
        value = self._property("formula2")
        return None if value is None else str(value)

    @property
    def formula(self):
        if self.type != "custom":
            return None
        value = self._property("formula1")
        return None if value is None else str(value)

    @property
    def source(self):
        if self.type != "list":
            return None
        value = self._property("formula1")
        return None if value is None else str(value)

    @property
    def in_cell_dropdown(self):
        value = self._property("in_cell_dropdown") if self.type == "list" else None
        return None if value is None else bool(value)

    @property
    def ignore_blank(self):
        value = self._property("ignore_blank")
        return None if value is None else bool(value)

    @property
    def input_title(self):
        return self._property("input_title")

    @property
    def input_message(self):
        return self._property("input_message")

    @property
    def show_input(self):
        value = self._property("show_input")
        return None if value is None else bool(value)

    @property
    def error_title(self):
        return self._property("error_title")

    @property
    def error_message(self):
        return self._property("error_message")

    @property
    def show_error(self):
        value = self._property("show_error")
        return None if value is None else bool(value)

    @property
    def alert_style(self):
        return self._ALERT_STYLE_FROM_KW.get(self._property("alert_style"))

    def _formula(self, source):
        if isinstance(source, base_classes.Range):
            formula = f"={source.get_address(True, True, True)}"
        elif isinstance(source, base_classes.Name):
            formula = f"={source.name}"
        else:
            separator = self.parent.sheet.book.app.xl.get_international(
                data_type=kw.list_separator
            )
            formula = separator.join(source)
        if len(formula) > 255:
            raise ValueError(
                "the Excel data validation source cannot exceed 255 characters"
            )
        return formula

    def _set(self, rule_type, operator=None, formula1=None, formula2=None):
        current_type = self.type
        if current_type in {"mixed_criteria", "inconsistent"}:
            raise xlwings.XlwingsError(
                "Cannot update data validation because the target cells have "
                "different validation rules."
            )
        kwargs = {"type": self._TYPE_TO_KW[rule_type], "formula1": formula1}
        if operator is not None:
            kwargs["operator"] = self._OPERATOR_TO_KW[operator]
        if formula2 is not None:
            kwargs["formula2"] = formula2
        validation = self.parent.xl.validation
        if current_type == "none":
            validation.add_data_validation(**kwargs)
        else:
            validation.modify(**kwargs)

    def set_list(self, source, in_cell_dropdown):
        formula = self._formula(source)
        self._set("list", formula1=formula)
        self.parent.xl.validation.in_cell_dropdown.set(in_cell_dropdown)

    def set_rule(self, rule_type, operator, formula1, formula2):
        self._set(rule_type, operator, formula1, formula2)

    def delete(self):
        self.parent.xl.validation.delete()


class Borders(base_classes.Borders):
    def __init__(self, parent, xl):
        # xl is the range reference: the sides are looked up via get_border
        self.parent = parent
        self.xl = xl

    @property
    def api(self):
        return self.xl

    def __getitem__(self, side):
        if self.xl is not None:
            return Border(
                self.parent, side, self.xl.get_border(which_border=getattr(kw, side))
            )
        return Border(self.parent, side, None)

    def _common_value(self, attribute):
        """The value the existing grid sides share, or None if they differ."""
        if self.xl is None:
            return None
        values = {getattr(self[side], attribute) for side in self._grid_sides()}
        return values.pop() if len(values) == 1 else None

    @property
    def line_style(self):
        return self._common_value("line_style")

    @line_style.setter
    def line_style(self, value):
        self.set(base_classes.BORDER_GRID_SIDES, line_style=value)

    @property
    def weight(self):
        return self._common_value("weight")

    @weight.setter
    def weight(self, value):
        self.set(base_classes.BORDER_GRID_SIDES, weight=value)

    @property
    def color(self):
        return self._common_value("color")

    @color.setter
    def color(self, color_or_rgb):
        self.set(base_classes.BORDER_GRID_SIDES, color=color_or_rgb)

    def set(
        self,
        which,
        *,
        line_style=base_classes._UNSET,
        weight=base_classes._UNSET,
        color=base_classes._UNSET,
    ):
        # `which` arrives validated and expanded by main.Borders. The fixed
        # order color, weight, line style is documented: Excel's border
        # attributes interfere, and this makes the line style win.
        if self.xl is None:
            return
        # Writing borders with screen updating on is about 3x slower
        app = self.parent.sheet.book.app
        screen_updating_state = app.screen_updating
        app.screen_updating = False
        try:
            for side in which:
                border = self[side]
                if color is not base_classes._UNSET:
                    border.color = color
                if weight is not base_classes._UNSET:
                    border.weight = weight
                if line_style is not base_classes._UNSET:
                    border.line_style = line_style
        finally:
            app.screen_updating = screen_updating_state

    def clear(self, which):
        self.set(which, line_style=None)


class Font(base_classes.Font):
    def __init__(self, parent, xl):
        # xl can be font or font_object
        self.parent = parent
        self.xl = xl

    @property
    def api(self):
        return self.xl

    @property
    def bold(self):
        return self.xl.bold.get()

    @bold.setter
    def bold(self, value):
        self.xl.bold.set(value)

    @property
    def italic(self):
        return self.xl.italic.get()

    @italic.setter
    def italic(self, value):
        self.xl.italic.set(value)

    @property
    def size(self):
        return self.xl.font_size.get()

    @size.setter
    def size(self, value):
        self.xl.font_size.set(value)

    @property
    def color(self):
        if isinstance(self.parent, Range):
            return tuple(self.xl.color.get())
        elif isinstance(self.parent, Shape):
            return tuple(self.xl.font_color.get())

    @color.setter
    def color(self, color_or_rgb):
        if isinstance(color_or_rgb, str):
            color_or_rgb = utils.hex_to_rgb(color_or_rgb)
        if self.xl is not None:
            if isinstance(self.parent, (Range, Characters)):
                obj = self.xl.color
            elif isinstance(self.parent, Shape):
                obj = self.xl.font_color

            if isinstance(color_or_rgb, int):
                obj.set(int_to_rgb(color_or_rgb))
            else:
                obj.set(color_or_rgb)

    @property
    def name(self):
        if isinstance(self.parent, Range):
            return self.xl.name.get()
        elif isinstance(self.parent, Shape):
            return self.xl.font_name.get()

    @name.setter
    def name(self, value):
        if isinstance(self.parent, Range):
            self.xl.name.set(value)
        elif isinstance(self.parent, Shape):
            self.xl.font_name.set(value)


class Characters(base_classes.Characters):
    def __init__(self, parent, xl):
        self.parent = parent
        self.xl = xl

    @property
    def api(self):
        return self.xl

    @property
    def text(self):
        return self.xl.content.get()

    @property
    def font(self):
        return Font(self, self.xl.font_object)

    def __getitem__(self, item):
        # TODO: This is broken with AppleScript and Excel 2016:
        # set bold of font object of (characters 5 thru 9 of range "A1") to true
        # https://answers.microsoft.com/en-us/msoffice/forum/
        #  msoffice_excel-mso_mac-msoversion_other/applescript-and-excel-problem/
        #  6e5a50b1-6209-4fbf-91f4-6d6674f1e488
        if isinstance(item, slice):
            return Characters(
                parent=self.parent,
                xl=self.xl[
                    item.start + 1
                    if item.start
                    else None : item.stop
                    if item.stop
                    else len(self.text)
                ],
            )
        else:
            return Characters(parent=self.parent, xl=self.xl[item + 1 : item + 1])


class PageSetup(base_classes.PageSetup):
    def __init__(self, parent, xl):
        self.parent = parent
        self.xl = xl

    @property
    def api(self):
        return self.xl

    @property
    def print_area(self):
        value = self.xl.print_area.get()
        if value == kw.missing_value:
            return None
        else:
            return self.xl.print_area.get()

    @print_area.setter
    def print_area(self, value):
        self.xl.print_area.set("" if value is None else value)


class Note(base_classes.Note):
    def __init__(self, parent, xl):
        self.parent = parent
        self.xl = xl

    def api(self):
        return self.xl

    @property
    def text(self):
        return self.xl.Excel_comment_text()

    @text.setter
    def text(self, value):
        self.xl.Excel_comment_text(text=value)

    def delete(self):
        self.parent.xl.clear_Excel_comments()


class Collection(base_classes.Collection):
    def __init__(self, parent):
        self._parent = parent
        self.xl = getattr(self.parent.xl, self._attr)

    @property
    def parent(self):
        return self._parent

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


class ConditionalFormat(base_classes.ConditionalFormat):
    def __init__(self, parent, key):
        self.parent = parent
        generic = parent.xl.format_conditions[key]
        native_type = generic.format_condition_type.get()
        specialized = _CONDITIONAL_FORMAT_SPECIALIZED_COLLECTION_FROM_KW.get(
            native_type
        )
        if specialized is None:
            self.xl = generic
            return
        # Excel exposes family-specific properties only through a `range` element
        # reference, not through the `cells` reference Range normally uses. The
        # specialized collections retain the global format-condition index.
        sheet = parent.sheet.xl
        range_ref = Reference(
            sheet.AS_appdata,
            sheet.AS_aemreference.elements(b"X117").byname(
                parent.address.replace("$", "")
            ),
        )
        self.xl = getattr(range_ref, specialized)[key]

    @property
    def api(self):
        return self.xl

    @property
    def type(self):
        return _CONDITIONAL_FORMAT_TYPE_FROM_KW.get(
            self.xl.format_condition_type.get(), "unknown"
        )

    @property
    def stop_if_true(self):
        if self.type in {"color_scale", "data_bar", "icon_set"}:
            return None
        return self.xl.stop_if_true.get()

    @property
    def operator(self):
        if self.type != "cell_value":
            return None
        return _CONDITIONAL_FORMAT_OPERATOR_FROM_KW.get(
            self.xl.condition_operator.get()
        )

    @property
    def formula1(self):
        return self.xl.formula_1.get() if self.type == "cell_value" else None

    @property
    def formula2(self):
        if self.type != "cell_value" or self.operator not in {
            "between",
            "not_between",
        }:
            return None
        value = self.xl.formula_2.get()
        return None if value == kw.missing_value else value

    @property
    def formula(self):
        return self.xl.formula_1.get() if self.type == "custom" else None

    @staticmethod
    def _color(obj, color_index=None):
        color_index = (obj.color_index if color_index is None else color_index).get()
        if color_index in {
            kw.color_index_none,
            kw.color_index_automatic,
            kw.missing_value,
        }:
            return None
        value = obj.color.get()
        return None if value is None or value == kw.missing_value else tuple(value)

    @property
    def fill_color(self):
        if self.type not in {"cell_value", "custom"}:
            return None
        return self._color(self.xl.interior_object)

    @property
    def font_color(self):
        if self.type not in {"cell_value", "custom"}:
            return None
        return self._color(self.xl.font_object, self.xl.font_object.font_color_index)

    @property
    def font_bold(self):
        if self.type not in {"cell_value", "custom"}:
            return None
        value = self.xl.font_object.bold.get()
        return None if value is None or value == kw.missing_value else value

    @property
    def font_italic(self):
        if self.type not in {"cell_value", "custom"}:
            return None
        value = self.xl.font_object.italic.get()
        return None if value is None or value == kw.missing_value else value

    @staticmethod
    def _threshold(criterion, type_property, value_property):
        criterion_type = _CONDITIONAL_FORMAT_THRESHOLD_FROM_KW.get(
            type_property.get(), "unknown"
        )
        if criterion_type in {
            "automatic",
            "lowest_value",
            "highest_value",
            "unknown",
        }:
            value = None
        else:
            value = value_property.get()
            if value == kw.missing_value:
                value = None
        return criterion_type, value

    @property
    def colors(self):
        if self.type != "color_scale":
            return None
        count = self.xl.count(each=kw.color_scale_criterion)
        # Criteria are direct children of the rule. Addressing them through the
        # color_scale_criteria property can terminate Excel's Apple-event process.
        criteria = self.xl.color_scale_criterion
        return tuple(
            self._color(criteria[index].format_color) for index in range(1, count + 1)
        )

    @property
    def bar_color(self):
        if self.type != "data_bar":
            return None
        return self._color(self.xl.databar_bar_color)

    @property
    def gradient(self):
        if self.type != "data_bar":
            return None
        return self.xl.databar_fill_type.get() == kw.databar_fill_gradient

    @property
    def show_value(self):
        if self.type == "data_bar":
            return self.xl.format_condition_show_value.get()
        if self.type == "icon_set":
            return not self.xl.show_icon_only.get()
        return None

    @property
    def icon_set(self):
        if self.type != "icon_set":
            return None
        return _CONDITIONAL_FORMAT_ICON_SET_FROM_KW.get(
            self.xl.format_condition_icon_set.icon_set_id.get()
        )

    @property
    def reverse_order(self):
        return self.xl.reverse_icon_set_order.get() if self.type == "icon_set" else None

    def _threshold_pairs(self):
        if self.type == "color_scale":
            count = self.xl.count(each=kw.color_scale_criterion)
            criteria = self.xl.color_scale_criterion
            return tuple(
                self._threshold(
                    criteria[index],
                    criteria[index].color_scale_criterion_type,
                    criteria[index].color_scale_criterion_value,
                )
                for index in range(1, count + 1)
            )
        if self.type == "data_bar":
            return tuple(
                self._threshold(
                    criterion,
                    criterion.condition_value_type,
                    criterion.condition_value_value,
                )
                for criterion in (
                    self.xl.min_point_condition_value,
                    self.xl.max_point_condition_value,
                )
            )
        if self.type == "icon_set":
            count = self.xl.count(each=kw.icon_criterion)
            criteria = self.xl.icon_criterion
            return tuple(
                self._threshold(
                    criteria[index],
                    criteria[index].icon_criterion_type,
                    criteria[index].icon_criterion_value,
                )
                for index in range(2, count + 1)
            )
        return None

    @property
    def threshold_types(self):
        pairs = self._threshold_pairs()
        return None if pairs is None else tuple(pair[0] for pair in pairs)

    @property
    def thresholds(self):
        pairs = self._threshold_pairs()
        return None if pairs is None else tuple(pair[1] for pair in pairs)

    def set(self, changes):
        criteria = {"operator", "formula1", "formula2", "formula"} & changes.keys()
        if criteria:
            if self.type == "cell_value":
                kwargs = {
                    "type": kw.cell_value,
                    "operator": _CONDITIONAL_FORMAT_OPERATOR_TO_KW[
                        changes.get("operator", self.operator)
                    ],
                    "formula1": changes.get("formula1", self.formula1),
                }
                formula2 = changes.get("formula2", self.formula2)
                if formula2 is not None:
                    kwargs["formula2"] = formula2
                self.xl.modify_condition(**kwargs)
            else:
                self.xl.modify_condition(
                    type=kw.expression,
                    formula1=changes.get("formula", self.formula),
                )
        if "fill_color" in changes:
            self.xl.interior_object.color.set(changes["fill_color"])
        if "font_color" in changes:
            self.xl.font_object.color.set(changes["font_color"])
        if "font_bold" in changes:
            self.xl.font_object.bold.set(changes["font_bold"])
        if "font_italic" in changes:
            self.xl.font_object.italic.set(changes["font_italic"])
        if "stop_if_true" in changes:
            self.xl.stop_if_true.set(changes["stop_if_true"])

    def delete(self):
        self.xl.delete()


class ConditionalFormats(Collection, base_classes.ConditionalFormats):
    _attr = "format_conditions"
    _kw = kw.format_condition
    _wrap = ConditionalFormat

    def _finish_add(self, rule, spec):
        rule.set_first_priority()
        wrapped = ConditionalFormat(self.parent, 1)
        wrapped.set(
            {
                key: value
                for key, value in spec.items()
                if key
                in {
                    "fill_color",
                    "font_color",
                    "font_bold",
                    "font_italic",
                    "stop_if_true",
                }
            }
        )
        return wrapped

    def add_cell_value(self, spec):
        properties = {
            kw.format_condition_type: kw.cell_value,
            kw.condition_operator: _CONDITIONAL_FORMAT_OPERATOR_TO_KW[spec["operator"]],
            kw.formula_1: spec["formula1"],
        }
        if spec["formula2"] is not None:
            properties[kw.formula_2] = spec["formula2"]
        rule = self.parent.xl.make(
            at=self.parent.xl,
            new=kw.format_condition,
            with_properties=properties,
        )
        return self._finish_add(rule, spec)

    def add_custom(self, spec):
        rule = self.parent.xl.make(
            at=self.parent.xl,
            new=kw.format_condition,
            with_properties={
                kw.format_condition_type: kw.expression,
                kw.formula_1: spec["formula"],
            },
        )
        return self._finish_add(rule, spec)

    @staticmethod
    def _set_threshold(
        criterion,
        type_property,
        value_property,
        criterion_type,
        value,
        *,
        automatic_type=None,
        modify=False,
    ):
        native_type = (
            automatic_type
            if criterion_type == "automatic"
            else _CONDITIONAL_FORMAT_THRESHOLD_TO_KW[criterion_type]
        )
        if modify:
            kwargs = {"type": native_type}
            if value is not None:
                kwargs["condition_value"] = value
            criterion.modify_condition_value(**kwargs)
        else:
            type_property.set(native_type)
            if value is not None:
                value_property.set(value)

    def _finish_visual_add(self, rule):
        rule.set_first_priority()
        return ConditionalFormat(self.parent, 1)

    def add_color_scale(self, spec):
        rule = self.parent.xl.make(
            at=self.parent.xl,
            new=kw.color_scale_format_condition,
            with_properties={kw.color_scale_type: len(spec["colors"])},
        )
        criteria = rule.color_scale_criterion
        for index, (color, criterion_type, value) in enumerate(
            zip(spec["colors"], spec["threshold_types"], spec["thresholds"]),
            start=1,
        ):
            criterion = criteria[index]
            self._set_threshold(
                criterion,
                criterion.color_scale_criterion_type,
                criterion.color_scale_criterion_value,
                criterion_type,
                value,
            )
            criterion.format_color.color.set(color)
        return self._finish_visual_add(rule)

    def add_data_bar(self, spec):
        rule = self.parent.xl.make(at=self.parent.xl, new=kw.databar_format_condition)
        rule.databar_bar_color.color.set(spec["bar_color"])
        rule.databar_fill_type.set(
            kw.databar_fill_gradient if spec["gradient"] else kw.databar_fill_solid
        )
        rule.format_condition_show_value.set(spec["show_value"])
        for criterion, criterion_type, value, automatic_type in zip(
            (rule.min_point_condition_value, rule.max_point_condition_value),
            spec["threshold_types"],
            spec["thresholds"],
            (
                kw.condition_value_automatic_minimum,
                kw.condition_value_automatic_maximum,
            ),
        ):
            self._set_threshold(
                criterion,
                criterion.condition_value_type,
                criterion.condition_value_value,
                criterion_type,
                value,
                automatic_type=automatic_type,
                modify=True,
            )
        return self._finish_visual_add(rule)

    def add_icon_set(self, spec):
        rule = self.parent.xl.make(at=self.parent.xl, new=kw.icon_set_format_condition)
        # `format condition icon sets` is declared as a property in Excel's
        # AppleScript dictionary, so appscript can't index it as a collection.
        # The singular elements do exist directly below the workbook, however.
        workbook = self.parent.sheet.book.xl
        icon_set = Reference(
            workbook.AS_appdata,
            workbook.AS_aemreference.elements(b"X319").byindex(
                _conditional_format_icon_set_indexes()[spec["icon_set"]]
            ),
        )
        rule.format_condition_icon_set.set(icon_set)
        rule.show_icon_only.set(not spec["show_value"])
        rule.reverse_icon_set_order.set(spec["reverse_order"])
        criteria = rule.icon_criterion
        for index, (criterion_type, value) in enumerate(
            zip(spec["threshold_types"], spec["thresholds"]), start=2
        ):
            criterion = criteria[index]
            self._set_threshold(
                criterion,
                criterion.icon_criterion_type,
                criterion.icon_criterion_value,
                criterion_type,
                value,
            )
            criterion.condition_operator.set(kw.operator_greater_equal)
        return self._finish_visual_add(rule)

    def clear(self):
        # Deleting the collection through the `cells` reference can terminate
        # Excel, and generic format-condition references don't delete visual
        # rules. Resolve each current first-priority rule to its native family.
        for _ in range(len(self)):
            ConditionalFormat(self.parent, 1).delete()


class AutoFilter(base_classes.AutoFilter):
    _COMPARISON_PREFIXES = {
        "equal_to": "=",
        "not_equal_to": "<>",
        "greater_than": ">",
        "less_than": "<",
        "greater_than_or_equal": ">=",
        "less_than_or_equal": "<=",
    }

    def __init__(self, parent, is_table=False):
        self.parent = parent
        self.is_table = is_table

    @staticmethod
    def _escape(value):
        return value.replace("~", "~~").replace("*", "~*").replace("?", "~?")

    @staticmethod
    def _criterion_value(value):
        if not isinstance(value, dict):
            return value
        if value.get("type") == "date":
            parsed = dt.date.fromisoformat(value["value"])
            return f"{parsed.month}/{parsed.day}/{parsed.year}"
        if value.get("type") == "datetime":
            parsed = dt.datetime.fromisoformat(value["value"])
            seconds = f"{parsed.second + parsed.microsecond / 1_000_000:g}"
            return (
                f"{parsed.month}/{parsed.day}/{parsed.year} "
                f"{parsed.hour}:{parsed.minute}:{seconds}"
            )
        raise ValueError("Unknown AutoFilter comparison value type")

    def _criteria(self, operator, value1, value2):
        if value1 is None:
            return ("=" if operator == "equal_to" else "<>"), None, None
        value1 = self._escape(self._criterion_value(value1))
        if operator == "between":
            return (
                f">={value1}",
                kw.autofilter_and,
                f"<={self._escape(self._criterion_value(value2))}",
            )
        if operator == "not_between":
            return (
                f"<{value1}",
                kw.autofilter_or,
                f">{self._escape(self._criterion_value(value2))}",
            )
        return f"{self._COMPARISON_PREFIXES[operator]}{value1}", None, None

    @property
    def _range(self):
        return self.parent.xl.range_object if self.is_table else self.parent.xl

    @property
    def _column_count(self):
        return self.parent.range.shape[1] if self.is_table else self.parent.shape[1]

    def _worksheet_filter_range(self):
        sheet = self.parent.sheet.xl
        if not sheet.autofilter_mode.get():
            return None
        return sheet.autofilter_object.range_object.get_address()

    def _ensure_target(self):
        if self.is_table:
            return
        existing = self._worksheet_filter_range()
        if existing is not None and existing != self.parent.xl.get_address():
            raise ValueError(
                "This worksheet already has an AutoFilter on a different range"
            )

    def _native_autofilter(self):
        if self.is_table:
            return self.parent.xl.autofilter_object
        existing = self._worksheet_filter_range()
        if existing is None or existing != self.parent.xl.get_address():
            return None
        return self.parent.sheet.xl.autofilter_object

    @property
    def criteria(self):
        autofilter = self._native_autofilter()
        if autofilter is None:
            return [
                base_classes.empty_autofilter_criteria(field)
                for field in range(1, self._column_count + 1)
            ]
        snapshots = []
        operator_types = {
            kw.filter_by_value: "values",
            kw.top_10_items: "top_items",
            kw.bottom_10_items: "bottom_items",
            kw.top_10_percent: "top_percent",
            kw.bottom_10_percent: "bottom_percent",
        }
        comparison_operators = {
            kw.autofilter_and: "and",
            kw.autofilter_or: "or",
        }
        for field in range(1, self._column_count + 1):
            native_filter = autofilter.filters[field]
            if not native_filter.filter_on.get():
                snapshots.append(base_classes.empty_autofilter_criteria(field))
                continue
            try:
                operator = native_filter.operator.get()
            except (CommandError, AttributeError):
                operator = None
            type_ = operator_types.get(operator, "comparison")
            if operator not in (
                None,
                kw.missing_value,
                *comparison_operators,
                *operator_types,
            ):
                type_ = "unknown"
            try:
                criteria2 = native_filter.criteria2.get()
                if criteria2 == kw.missing_value:
                    criteria2 = None
            except (CommandError, AttributeError):
                criteria2 = None
            try:
                criteria1 = native_filter.criteria1.get()
                if criteria1 == kw.missing_value:
                    criteria1 = None
            except (CommandError, AttributeError):
                criteria1 = None
            snapshots.append(
                base_classes.autofilter_criteria_snapshot(
                    field,
                    type_,
                    criteria1,
                    criteria2,
                    comparison_operators.get(operator),
                )
            )
        return snapshots

    def apply_values(self, field, values):
        self._ensure_target()
        self._range.autofilter_range(
            field=field, criteria1=values, operator=kw.filter_by_value
        )

    def apply_comparison(self, field, operator, value1, value2):
        self._ensure_target()
        criteria1, native_operator, criteria2 = self._criteria(operator, value1, value2)
        kwargs = {"field": field, "criteria1": criteria1}
        if native_operator is not None:
            kwargs["operator"] = native_operator
            kwargs["criteria2"] = criteria2
        self._range.autofilter_range(**kwargs)

    def _apply_top_bottom(self, field, value, operator):
        self._ensure_target()
        self._range.autofilter_range(
            field=field, criteria1=str(value), operator=operator
        )

    def apply_top_items(self, field, count):
        self._apply_top_bottom(field, count, kw.top_10_items)

    def apply_bottom_items(self, field, count):
        self._apply_top_bottom(field, count, kw.bottom_10_items)

    def apply_top_percent(self, field, percent):
        self._apply_top_bottom(field, percent, kw.top_10_percent)

    def apply_bottom_percent(self, field, percent):
        self._apply_top_bottom(field, percent, kw.bottom_10_percent)

    def clear(self, field):
        if not self.is_table:
            existing = self._worksheet_filter_range()
            if existing is None or existing != self.parent.xl.get_address():
                return
        fields = [field] if field is not None else range(1, self._column_count + 1)
        for field_index in fields:
            self._range.autofilter_range(field=field_index)


class Table(base_classes.Table):
    def __init__(self, parent, key):
        self._parent = parent
        self.xl = parent.xl.list_objects[key]

    @property
    def parent(self):
        return self._parent

    @property
    def api(self):
        return self.xl

    @property
    def name(self):
        return self.xl.name.get()

    @name.setter
    def name(self, value):
        self.xl.name.set(value)
        self.xl = self.parent.xl.list_objects[value]

    @property
    def data_body_range(self):
        if self.xl.cell_table.get() == kw.missing_value:
            return
        else:
            return Range(self.parent, self.xl.cell_table.get_address())

    @property
    def display_name(self):
        return self.xl.display_name.get()

    @display_name.setter
    def display_name(self, value):
        # Changing the display_name also changes the name
        self.xl.display_name.set(value)
        self.xl = self.parent.xl.list_objects[value]

    @property
    def header_row_range(self):
        if self.xl.header_row.get() == kw.missing_value:
            return
        else:
            return Range(self.parent, self.xl.header_row.get_address())

    @property
    def insert_row_range(self):
        if self.xl.insert_row.get() == kw.missing_value:
            return
        else:
            return Range(self.parent, self.xl.insert_row.get_address())

    @property
    def range(self):
        return Range(self.parent, self.xl.range_object.get_address())

    @property
    def autofilter(self):
        return AutoFilter(self, is_table=True)

    @property
    def show_autofilter(self):
        return self.xl.show_autofilter.get()

    @show_autofilter.setter
    def show_autofilter(self, value):
        self.xl.show_autofilter.set(value)

    @property
    def show_headers(self):
        return self.xl.show_headers.get()

    @show_headers.setter
    def show_headers(self, value):
        self.xl.show_headers.set(value)

    @property
    def show_table_style_column_stripes(self):
        return self.xl.show_table_style_column_stripes.get()

    @show_table_style_column_stripes.setter
    def show_table_style_column_stripes(self, value):
        self.xl.show_table_style_column_stripes.set(value)

    @property
    def show_table_style_first_column(self):
        return self.xl.show_table_style_first_column.get()

    @show_table_style_first_column.setter
    def show_table_style_first_column(self, value):
        self.xl.show_table_style_first_column.set(value)

    @property
    def show_table_style_last_column(self):
        return self.xl.show_table_style_last_column.get()

    @show_table_style_last_column.setter
    def show_table_style_last_column(self, value):
        self.xl.show_table_style_last_column.set(value)

    @property
    def show_table_style_row_stripes(self):
        return self.xl.show_table_style_row_stripes.get()

    @show_table_style_row_stripes.setter
    def show_table_style_row_stripes(self, value):
        self.xl.show_table_style_row_stripes.set(value)

    @property
    def show_totals(self):
        return self.xl.total.get()

    @show_totals.setter
    def show_totals(self, value):
        self.xl.total.set(value)

    @property
    def table_style(self):
        return self.xl.table_style.properties().get(kw.name)

    @table_style.setter
    def table_style(self, value):
        self.xl.table_style.set(value)

    @property
    def totals_row_range(self):
        if self.xl.total_row.get() == kw.missing_value:
            return
        else:
            return Range(self.parent, self.xl.total_row.get_address())

    def resize(self, range):
        self.xl.resize(range=range.api)


class Tables(Collection, base_classes.Tables):
    _attr = "list_objects"
    _kw = kw.list_object
    _wrap = Table

    def add(
        self,
        source_type=None,
        source=None,
        link_source=None,
        has_headers=None,
        destination=None,
        table_style_name=None,
        name=None,
    ):
        header_row = {
            True: kw.header_yes,
            False: kw.header_no,
            "guess": kw.header_guess,
        }
        sheet_index = self.parent.xl.entry_index.get()
        table = Table(
            self.parent,
            self.parent.xl.make(
                at=self.parent.book.xl.sheets[sheet_index],
                new=kw.list_object,
                with_properties={
                    kw.source_type: kw.src_range,
                    kw.range_object: source.api,
                    kw.header_row: header_row[has_headers],
                    kw.table_style: table_style_name,
                },
            ).name.get(),
        )
        if name is not None:
            table.name = name
        return table


class Chart(base_classes.Chart):
    def __init__(self, parent, key):
        self._parent = parent
        if isinstance(parent, Sheet):
            self.xl_obj = parent.xl.chart_objects[key]
            self.xl = self.xl_obj.chart
        else:
            # chart sheet
            self.xl_obj = None
            self.xl = parent.xl.chart_sheets[key]

    @property
    def parent(self):
        return self._parent

    @property
    def api(self):
        return self.xl_obj, self.xl

    def set_source_data(self, rng, plot_by=None):
        if plot_by is None:
            self.xl.set_source_data(source=rng.xl)
        else:
            self.xl.set_source_data(source=rng.xl, plot_by=plot_by_s2k[plot_by])

    def set_x_axis_values(self, rng):
        for series in _mac_list(self.xl.series_collection):
            series.xvalues.set(rng.xl)

    @property
    def name(self):
        if self.xl_obj is not None:
            return self.xl_obj.name.get()
        else:
            return self.xl.name.get()

    @name.setter
    def name(self, value):
        # Charts are addressed by name, so the references have to be
        # re-resolved after renaming (same as Table.name)
        if self.xl_obj is not None:
            self.xl_obj.name.set(value)
            self.xl_obj = self._parent.xl.chart_objects[value]
            self.xl = self.xl_obj.chart
        else:
            self.xl.name.set(value)
            self.xl = self._parent.xl.chart_sheets[value]

    @property
    def chart_type(self):
        return chart_types_k2s[self.xl.chart_type.get()]

    @chart_type.setter
    def chart_type(self, value):
        self.xl.chart_type.set(chart_types_s2k[value])

    @property
    def title(self):
        if not self.xl.has_title.get():
            return None
        text = self.xl.chart_title.chart_title_text.get()
        return None if text == kw.missing_value else text

    @title.setter
    def title(self, value):
        if value is None:
            self.xl.has_title.set(False)
        else:
            self.xl.has_title.set(True)
            self.xl.chart_title.chart_title_text.set(value)

    @property
    def legend(self):
        return ChartLegend(self)

    @property
    def plot_by(self):
        return plot_by_k2s[self.xl.plot_by.get()]

    @plot_by.setter
    def plot_by(self, value):
        self.xl.plot_by.set(plot_by_s2k[value])

    @property
    def style(self):
        return self.xl.chart_style.get()

    @style.setter
    def style(self, value):
        self.xl.chart_style.set(value)

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
        if self.xl_obj is None:
            # chart sheet: Excel asks for confirmation like for any sheet
            app_xl = self._parent.app.xl
            alerts_state = app_xl.display_alerts.get()
            app_xl.display_alerts.set(False)
            try:
                self.xl.delete()
            finally:
                app_xl.display_alerts.set(alerts_state)
        else:
            self.xl_obj.delete()

    def to_png(self, path):
        raise xlwings.XlwingsError("Chart.to_png() isn't supported on macOS.")
        # Both versions should work, but seem to be broken with Excel 2016
        #
        # Version 1
        # import uuid
        # temp_path = posix_to_hfs_path(os.path.expanduser("~")
        #                               + f"/Library/Containers/com.microsoft.Excel/"
        #                                 f"Data/{uuid.uuid4()}.png")
        # self.xl.save_as(filename=temp_path)
        # shutil.copy2(temp_path, path)
        # try:
        #     os.unlink(temp_path)
        # except:
        #     pass
        #
        # Version 2
        # self.xl_obj.save_as_picture(file_name=posix_to_hfs_path('...'),
        #                             picture_type=kw.save_as_PNG_file)

    def to_pdf(self, path, quality=None):
        raise xlwings.XlwingsError("Chart.to_pdf() isn't supported on macOS.")


class ChartLegend(base_classes.ChartLegend):
    def __init__(self, parent):
        # Only the parent is kept: the chart's native reference is name-based
        # and gets replaced when the chart is renamed
        self.parent = parent

    @property
    def xl(self):
        return self.parent.xl

    @property
    def api(self):
        return self.xl.legend_object

    @property
    def visible(self):
        return self.xl.has_legend.get()

    @visible.setter
    def visible(self, value):
        self.xl.has_legend.set(value)

    @property
    def position(self):
        if not self.xl.has_legend.get():
            return None
        return legend_positions_k2s[self.xl.legend_object.position.get()]

    @position.setter
    def position(self, value):
        self.xl.has_legend.set(True)
        self.xl.legend_object.position.set(legend_positions_s2k[value])


class Charts(Collection, base_classes.Charts):
    _attr = "chart_objects"
    _kw = kw.chart_object
    _wrap = Chart

    def add(
        self,
        left,
        top,
        width,
        height,
        chart_type=None,
        source=None,
        plot_by=None,
        name=None,
        anchor=None,
        style=227,
    ):
        if anchor:
            top, left = anchor.top, anchor.left
        sheet_index = self.parent.xl.entry_index.get()
        chart = Chart(
            self.parent,
            self.parent.xl.make(
                at=self.parent.book.xl.sheets[sheet_index],
                new=kw.chart_object,
                with_properties={
                    kw.width: width,
                    kw.top: top,
                    kw.left_position: left,
                    kw.height: height,
                },
            ).name.get(),
        )
        # data before type: stock/xy types need series to exist; name last as
        # the chart is addressed by name
        if source is not None:
            chart.set_source_data(source, plot_by)
        if chart_type is not None:
            chart.chart_type = chart_type
        if style is not None:
            chart.style = style
        if name is not None:
            chart.name = name
        return chart


def _mac_list(ref):
    """The items of an appscript element list, or [] when Excel reports
    `missing value` for an empty collection."""
    items = ref.get()
    return [] if items == kw.missing_value else items


_pivot_value_states = WeakValueDictionary()

# Stored on the native pivot so defaults survive collection lookups and saves.
_pivot_defaults_pending = "xlwings:defaults:pending-values"
_pivot_defaults_applied = "xlwings:defaults"


def _pivot_key(pivot):
    sheet = pivot.parent
    book = sheet.book
    # Names normalize references obtained by numeric and string lookup.
    return (book.app.pid, book.name, sheet.name, pivot.name)


def _invalidate_pivot_value_states(prefix):
    for key, state in list(_pivot_value_states.items()):
        if key[: len(prefix)] == prefix:
            state.deleted = True
            del _pivot_value_states[key]


class _PivotValueState:
    def __init__(self, key):
        self.key = key
        self.name = key[-1]
        self.deleted = False

    def rename(self, name):
        _pivot_value_states.pop(self.key, None)
        self.name = name
        self.key = (*self.key[:-1], name)
        _pivot_value_states[self.key] = self


class PivotTable(base_classes.PivotTable):
    def __init__(self, parent, key):
        self._parent = parent
        self.xl = parent.xl.pivot_tables[key]

    @property
    def api(self):
        return self.xl

    @property
    def parent(self):
        return self._parent

    @property
    def name(self):
        return self.xl.name.get()

    @name.setter
    def name(self, value):
        old_key = _pivot_key(self)
        self.xl.name.set(value)
        # the native reference is name-based
        self.xl = self.parent.xl.pivot_tables[value]
        for key, state in list(_pivot_value_states.items()):
            if key[:-1] == old_key:
                del _pivot_value_states[key]
                state.key = (*old_key[:-1], value, state.name)
                _pivot_value_states[state.key] = state

    @property
    def field_names(self):
        # pivot_fields lists the source fields plus, with two or more value
        # fields, the "Values" pseudo field (data_pivot_field)
        try:
            values_name = self.xl.data_pivot_field.name.get()
        except CommandError:
            values_name = None
        return [
            field.name.get()
            for field in _mac_list(self.xl.pivot_fields)
            if field.name.get() != values_name
            and field.pivot_field_orientation.get() != kw.orient_as_data_field
        ]

    @property
    def rows(self):
        return PivotFields(pivot=self, area="rows")

    @property
    def columns(self):
        return PivotFields(pivot=self, area="columns")

    @property
    def filters(self):
        return PivotFields(pivot=self, area="filters")

    @property
    def values(self):
        return PivotValueFields(pivot=self)

    @property
    def layout(self):
        # layout_row_default only applies to fields added later, so read the
        # actual layout off the row fields; None when they disagree.
        row_fields = _mac_list(self.xl.row_fields)
        if not row_fields:
            return pivot_layouts_k2s.get(self.xl.layout_row_default.get())
        layouts = set()
        for field in row_fields:
            # address the source field: some properties don't resolve via
            # the row_fields element reference
            field = self.xl.pivot_fields[field.name.get()]
            if field.layout_compact_row.get():
                layouts.add("compact")
            elif field.layout_form.get() == kw.tabular:
                layouts.add("tabular")
            elif field.layout_form.get() == kw.outline:
                layouts.add("outline")
            else:
                layouts.add(pivot_layouts_k2s.get(self.xl.layout_row_default.get()))
        return layouts.pop() if len(layouts) == 1 else None

    @layout.setter
    def layout(self, value):
        self.xl.row_axis_layout(layout=pivot_layouts_s2k[value])
        self.xl.layout_row_default.set(pivot_layouts_s2k[value])

    @property
    def show_row_grand_totals(self):
        return self.xl.row_grand.get()

    @show_row_grand_totals.setter
    def show_row_grand_totals(self, value):
        self.xl.row_grand.set(value)

    @property
    def show_column_grand_totals(self):
        return self.xl.column_grand.get()

    @show_column_grand_totals.setter
    def show_column_grand_totals(self, value):
        self.xl.column_grand.set(value)

    @property
    def range(self):
        return Range(self.parent, self.xl.table_range1.get_address())

    @property
    def data_body_range(self):
        # Excel for Mac answers with the row labels area when there are no
        # value fields
        if not _mac_list(self.xl.data_fields):
            return None
        return Range(self.parent, self.xl.data_body_range.get_address())

    def refresh(self):
        self.xl.refresh_table()

    def delete(self):
        # There is no delete command; clearing the full report range (incl.
        # the filters area) removes it.
        key = _pivot_key(self)
        self.xl.table_range2.clear_range()
        _invalidate_pivot_value_states(key)


class PivotField(base_classes.PivotField):
    def __init__(self, pivot, name):
        # addressed as the source field, so the wrapper follows the field
        # when it is moved to another area
        self._pivot = pivot
        self._name = name

    @property
    def xl(self):
        return self._pivot.xl.pivot_fields[self._name]

    @property
    def api(self):
        return self.xl

    @property
    def parent(self):
        return self._pivot

    @property
    def name(self):
        return self._name

    def remove(self):
        self.xl.pivot_field_orientation.set(kw.orient_as_hidden)


class PivotFields(base_classes.PivotFields):
    def __init__(self, pivot, area):
        self._pivot = pivot
        self._area = area

    @property
    def xl(self):
        return getattr(self._pivot.xl, pivot_area_elements[self._area])

    @property
    def api(self):
        return self.xl

    @property
    def parent(self):
        return self._pivot

    @property
    def area(self):
        return self._area

    def _names(self):
        # in position order; Excel's "Values" pseudo field isn't a source
        # field, so hide it, like field_names does and like Office.js
        try:
            values_name = self._pivot.xl.data_pivot_field.name.get()
        except CommandError:
            values_name = None
        return [
            name
            for name in (field.name.get() for field in _mac_list(self.xl))
            if name != values_name
        ]

    def __call__(self, key):
        names = self._names()
        if isinstance(key, numbers.Number):
            if key < 1 or key > len(names):
                raise KeyError(key)
            return PivotField(self._pivot, names[key - 1])
        if key not in names:
            raise KeyError(key)
        return PivotField(self._pivot, key)

    def __len__(self):
        return len(self._names())

    def __iter__(self):
        for name in self._names():
            yield PivotField(self._pivot, name)

    def __contains__(self, key):
        return key in self._names()

    def add(self, name):
        field = self._pivot.xl.pivot_fields[name]
        if not field.exists():
            raise KeyError(name)
        orientation = pivot_area_orientations[self._area]
        # setting the orientation appends the field to the area; leave a
        # field that is already here where it is
        if field.pivot_field_orientation.get() != orientation:
            field.pivot_field_orientation.set(orientation)
        if self._pivot.xl.tag.get() in (
            _pivot_defaults_pending,
            _pivot_defaults_applied,
        ):
            # see PivotTables.add: fields added this way ignore the pivot
            # table's default layout, so re-apply it to all of them
            self._pivot.xl.row_axis_layout(
                layout=self._pivot.xl.layout_row_default.get()
            )
        return PivotField(self._pivot, name)


class PivotValueField(base_classes.PivotValueField):
    def __init__(self, pivot, name):
        self._pivot = pivot
        key = (*_pivot_key(pivot), name)
        state = _pivot_value_states.get(key)
        if state is None:
            state = _PivotValueState(key)
            _pivot_value_states[key] = state
        self._state = state

    @property
    def xl(self):
        if self._state.deleted:
            raise KeyError("The value field has been removed.")
        # Resolve from the shared state so aliases also survive a pivot rename.
        _, book, sheet, pivot, name = self._state.key
        return (
            self._pivot.parent.book.app.xl.workbooks[book]
            .worksheets[sheet]
            .pivot_tables[pivot]
            .data_fields[name]
        )

    @property
    def api(self):
        return self.xl

    @property
    def parent(self):
        return self._pivot

    @property
    def name(self):
        return self._state.name

    @name.setter
    def name(self, value):
        self.xl.name.set(value)
        self._state.rename(value)

    @property
    def source_field(self):
        return self.xl.source_name.get()

    @property
    def function(self):
        return pivot_functions_k2s.get(self.xl.function.get())

    @function.setter
    def function(self, value):
        # Excel renames an automatic caption ("Sum of X" -> "Count of X")
        # along with the function, so re-resolve the field by its position
        field = self.xl
        position = field.position.get()
        field.function.set(pivot_functions_s2k[value])
        _, book, sheet, pivot, _ = self._state.key
        fields = (
            self._pivot.parent.book.app.xl.workbooks[book]
            .worksheets[sheet]
            .pivot_tables[pivot]
            .data_fields
        )
        self._state.rename(fields[position].name.get())

    @property
    def number_format(self):
        return self.xl.number_format.get()

    @number_format.setter
    def number_format(self, value):
        self.xl.number_format.set(value)

    def remove(self):
        self.xl.pivot_field_orientation.set(kw.orient_as_hidden)
        self._state.deleted = True
        _pivot_value_states.pop(self._state.key, None)


class PivotValueFields(base_classes.PivotValueFields):
    def __init__(self, pivot):
        self._pivot = pivot

    @property
    def xl(self):
        return self._pivot.xl.data_fields

    @property
    def api(self):
        return self.xl

    @property
    def parent(self):
        return self._pivot

    def _names(self):
        return [field.name.get() for field in _mac_list(self.xl)]

    def __call__(self, key):
        names = self._names()
        if isinstance(key, numbers.Number):
            if key < 1 or key > len(names):
                raise KeyError(key)
            return PivotValueField(self._pivot, names[key - 1])
        if key not in names:
            raise KeyError(key)
        return PivotValueField(self._pivot, key)

    def __len__(self):
        return len(self._names())

    def __iter__(self):
        for name in self._names():
            yield PivotValueField(self._pivot, name)

    def __contains__(self, key):
        return key in self._names()

    def add(self, field, function=None, name=None, number_format=None):
        source = self._pivot.xl.pivot_fields[field]
        if not source.exists():
            raise KeyError(field)
        # add_data_field is broken in Excel's AppleScript interface (it either
        # does nothing or crashes Excel); setting the orientation of the source
        # field appends a value field with Excel's default function
        source.pivot_field_orientation.set(kw.orient_as_data_field)
        data_fields = _mac_list(self.xl)
        defaults = self._pivot.xl.tag.get()
        if defaults == _pivot_defaults_pending and len(data_fields) >= 2:
            # see PivotTables.add: the pseudo field only accepts changes once
            # it's shown, i.e. with two or more value fields
            pseudo = self._pivot.xl.data_pivot_field
            # Captions share a namespace with the value fields. Also reserve
            # the requested caption, which is applied below.
            captions = {field.name.get().casefold() for field in data_fields}
            if name is not None:
                captions.add(name.casefold())
            caption = "Values"
            suffix = 2
            while caption.casefold() in captions:
                caption = f"Values{suffix}"
                suffix += 1
            pseudo.name.set(caption)
            pseudo.pivot_field_orientation.set(kw.orient_as_column_field)
            self._pivot.xl.tag.set(_pivot_defaults_applied)
        if defaults in (_pivot_defaults_pending, _pivot_defaults_applied):
            # adding a value field reverts the captions to the classic form,
            # see PivotFields.add
            self._pivot.xl.row_axis_layout(
                layout=self._pivot.xl.layout_row_default.get()
            )
        new = PivotValueField(self._pivot, data_fields[-1].name.get())
        # function first: it resets an automatic caption
        if function is not None:
            new.function = function
        if name is not None:
            new.name = name
        if number_format is not None:
            new.number_format = number_format
        return new


class PivotTables(Collection, base_classes.PivotTables):
    _attr = "pivot_tables"
    _kw = kw.pivot_table
    _wrap = PivotTable

    def add(self, source, destination, name=None):
        # `make new pivot table` creates the pivot cache itself. It takes the
        # source as an A1-style reference with the sheet name, a defined name
        # or a structured reference; an R1C1 string is rejected across sheets
        # with a bare "parameter error", as is `make new pivot cache`.
        # Qualify the workbook: Excel resolves unqualified sources against
        # the active workbook even when `make` targets a different book.
        if isinstance(source, Table):
            sheet = source.parent
            prefix = f"[{sheet.book.name}]{sheet.name}".replace("'", "''")
            source_data = f"'{prefix}'!{source.name}[#All]"
        else:
            source_data = source.get_address(True, True, True)
        # On a sheet that already has a pivot table, `make` silently answers
        # with the existing one instead of creating another (whatever the
        # `at` target), and the pivot table's location property can't move
        # one in from elsewhere, so a second pivot table per sheet is out.
        before = [pt.name.get() for pt in _mac_list(self.parent.xl.pivot_tables)]
        if before:
            raise NotImplementedError(
                "On macOS, only the first pivot table on a sheet can be created; "
                f"sheet {self.parent.name!r} already has {before!r}. Create it on "
                "another sheet."
            )
        top_left = Range(self.parent, (destination.row, destination.column, 1, 1))
        self.parent.book.xl.make(
            at=self.parent.xl,
            new=kw.pivot_table,
            with_properties={
                kw.source_data: source_data,
                kw.table_range1: top_left.xl,
            },
        )
        # `make` may answer with an index-based reference, so address the
        # new pivot table by its name
        after = [pt.name.get() for pt in _mac_list(self.parent.xl.pivot_tables)]
        if len(after) != 1:
            raise xlwings.XlwingsError(
                f"Excel didn't create the pivot table on sheet {self.parent.name!r}."
            )
        pivot = PivotTable(self.parent, after[0])
        # `make` answers with a classic (Excel 2003 style) pivot table: no
        # table style, tabular layout with in-grid drop zones, and the values
        # pseudo field captioned "Data" and laid out down the rows. Apply
        # Excel's defaults for a new pivot table instead, as Windows and
        # Office.js do, so the report looks the same on all platforms. The
        # values pseudo field only accepts changes once it's shown, and the
        # default layout isn't applied to fields added by script, so
        # PivotValueFields.add and PivotFields.add finish the job.
        pivot.xl.table_style2.set("PivotStyleLight16")
        pivot.xl.in_grid_drop_zones.set(False)
        pivot.xl.layout_row_default.set(kw.compact_row)
        pivot.xl.tag.set(_pivot_defaults_pending)
        if name:
            pivot.name = name
        return pivot


class Picture(base_classes.Picture):
    def __init__(self, parent, key):
        self._parent = parent
        self.xl = parent.xl.pictures[key]

    @property
    def parent(self):
        return self._parent

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

    @property
    def lock_aspect_ratio(self):
        return self.xl.lock_aspect_ratio.get()

    @lock_aspect_ratio.setter
    def lock_aspect_ratio(self, value):
        self.xl.lock_aspect_ratio.set(value)

    def update(self, filename):
        return utils.excel_update_picture(self, filename)


class Pictures(Collection, base_classes.Pictures):
    _attr = "pictures"
    _kw = kw.picture
    _wrap = Picture

    def add(
        self,
        filename,
        link_to_file,
        save_with_document,
        left,
        top,
        width,
        height,
        anchor,
    ):
        if anchor:
            top, left = anchor.top, anchor.left

        version = VersionNumber(self.parent.book.app.version)

        if not link_to_file and version >= 15:
            # Office 2016 for Mac is sandboxed. This path seems to work without the
            # need of granting access explicitly.
            xlwings_picture = (
                os.path.expanduser("~")
                + "/Library/Containers/com.microsoft.Excel/Data/xlwings_picture.png"
            )
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
                    kw.width: width,
                    kw.height: height,
                    # Top and left: see below
                    kw.top: 0,
                    kw.left_position: 0,
                },
            ).name.get(),
        )

        # Top and left cause an issue in the make command above
        # if they are not set to 0 when width & height are -1
        picture.top = top if top else 0
        picture.left = left if left else 0

        if not link_to_file and version >= 15:
            os.remove(filename)

        return picture


class Names(base_classes.Names):
    def __init__(self, parent, xl):
        self.parent = parent
        self.xl = xl

    def __call__(self, name_or_index):
        if isinstance(name_or_index, numbers.Number):
            name = self.xl[name_or_index].name.get()
            return Name(
                self.parent,
                collection=self,
                index=NameIndex(
                    name_or_index, name, self._sheet_names() if "!" in name else ()
                ),
            )
        return Name(self.parent, xl=self.xl[name_or_index])

    def _name_strings(self):
        names = self.xl.name.get()
        if names == kw.missing_value:
            return []
        # Excel's bulk read can repeat a sheet-local name in place of a shadowed
        # workbook name. Indexed reads distinguish them; verify only collisions.
        counts = Counter(names)
        return [
            self.xl[i].name.get() if counts[name] > 1 else name
            for i, name in enumerate(names, 1)
        ]

    def _name_at_index(self, index):
        try:
            return self.xl[index].name.get()
        except CommandError:
            # An earlier deletion can leave the cached index past the end.
            return None

    def _sheet_names(self):
        book = self.parent if isinstance(self.parent, Book) else self.parent.book
        names = book.xl.worksheets.name.get()
        return () if names == kw.missing_value else tuple(names)

    def snapshot(self):
        names = self._name_strings()
        sheets = self._sheet_names() if any("!" in name for name in names) else ()
        return [
            (name, Name(self.parent, collection=self, index=NameIndex(i, name, sheets)))
            for i, name in enumerate(names, 1)
        ]

    def contains(self, name_or_index):
        try:
            self.xl[name_or_index].get()
        except appscript.reference.CommandError:
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
        return Name(
            self.parent,
            self.parent.xl.make(
                at=self.parent.xl,
                new=kw.named_item,
                with_properties={kw.references: refers_to, kw.name: name},
            ),
        )


class Name(base_classes.Name):
    def __init__(self, parent, xl=None, collection=None, index=None):
        self.parent = parent
        self._xl = xl
        self._collection = collection
        self._index = index

    @property
    def xl(self):
        if self._index is None:
            return self._xl
        index = self._index.resolve(
            self._collection._name_at_index,
            self._collection._name_strings,
            self._collection._sheet_names,
        )
        return self._collection.xl[index]

    @contextmanager
    def _mutation_context(self, new_name=None, refers_to=None):
        # Excel can route even indexed mutations of a workbook name to a local
        # name on the active sheet. Use a sheet without that shadow while writing.
        name = self.name
        if "!" in name:
            yield refers_to
            return
        names = {name.lower()}
        if new_name is not None and "!" not in new_name:
            names.add(new_name.lower())
        book = self.parent if isinstance(self.parent, Book) else self.parent.book
        try:
            active_sheet = book.sheets.active
            local_names = (
                active_sheet.names._name_strings() if active_sheet.xl.exists() else None
            )
        except CommandError:
            # A chart sheet cannot be addressed through the worksheets collection.
            local_names = None
        if local_names is not None and not any(
            item.rsplit("!", 1)[-1].lower() in names for item in local_names
        ):
            yield refers_to
            return
        previous_book = book.app.books.active
        # Keep a native sheets reference so chart sheets can also be restored.
        previous_sheet = book.xl.sheets[book.xl.active_sheet.name.get()]
        temporary_sheet = None
        if refers_to is not None and local_names is not None:
            # Excel interprets input relative references from A1, while reads
            # depend on the selected cell. Normalize on the original sheet at A1.
            selection = book.app.selection
            try:
                book.sheets.active.range("A1").select()
                temporary_name = book.names.add(f"xw_tmp_{uuid4().hex[:16]}", refers_to)
                try:
                    refers_to = temporary_name.refers_to
                finally:
                    temporary_name.delete()
            finally:
                if selection is not None:
                    selection.select()
        try:
            if refers_to is not None or local_names is None:
                temporary_sheet = self._add_mutation_sheet(book)
                temporary_sheet.activate()
            else:
                for sheet in book.sheets:
                    if sheet.visible and not any(
                        item.rsplit("!", 1)[-1].lower() in names
                        for item in sheet.names._name_strings()
                    ):
                        sheet.activate()
                        break
                else:
                    temporary_sheet = self._add_mutation_sheet(book)
                    temporary_sheet.activate()
            yield refers_to
        finally:
            try:
                previous_sheet.activate_object()
                if temporary_sheet is not None:
                    temporary_sheet.delete()
            finally:
                previous_book.activate()

    def _add_mutation_sheet(self, book):
        try:
            return book.sheets.add(after=book.sheets(len(book.sheets)))
        except CommandError as exc:
            raise xlwings.XlwingsError(
                f"Cannot modify defined name {self.name!r}: Excel could not create "
                "the temporary worksheet needed to preserve its scope. "
                "Check whether the workbook structure is protected."
            ) from exc

    def delete(self):
        with self._mutation_context():
            self.xl.delete()

    @property
    def name(self):
        if self._index is None:
            return self.xl.name.get()
        self._index.resolve(
            self._collection._name_at_index,
            self._collection._name_strings,
            self._collection._sheet_names,
        )
        return self._index.name

    @name.setter
    def name(self, value):
        with self._mutation_context(new_name=value):
            native = self.xl
            old_name = native.name.get()
            expected_name = value
            if "!" not in value and "!" in old_name:
                scope = old_name.rsplit("!", 1)[0]
                expected_name = f"{scope}!{value}"
            native.name.set(value)
            collection = (
                self._collection if self._collection is not None else self.parent.names
            )
            names = collection._name_strings()
            if expected_name not in names or (
                old_name != expected_name and old_name in names
            ):
                # Excel can display an alert and return normally without renaming.
                # Keep the old identity until the native collection confirms it.
                raise xlwings.XlwingsError(
                    f"Excel did not rename defined name {old_name!r} to {value!r}."
                )
            self._collection = collection
            self._index = NameIndex(
                names.index(expected_name) + 1,
                expected_name,
                collection._sheet_names() if "!" in expected_name else (),
            )

    @property
    def refers_to(self):
        return self.xl.properties().get(kw.references)

    @refers_to.setter
    def refers_to(self, value):
        with self._mutation_context(refers_to=value) as refers_to:
            self.xl.references.set(refers_to)

    @property
    def refers_to_range(self):
        book = self.parent if isinstance(self.parent, Book) else self.parent.book
        external_address = self.xl.reference_range.get_address(external=True)
        match = re.search(r"\](.*?)'?!(.*)", external_address)
        return Range(Sheet(book, match.group(1)), match.group(2))


class Shapes(Collection):
    _attr = "shapes"
    _kw = kw.shape
    _wrap = Shape


@atexit.register
def cleanup():
    """
    Since AppleScript cannot access Excel while a Macro is running, we have to run the
    Python call in a background process which makes the call return immediately: we
    rely on the StatusBar to give the user feedback.
    This function is triggered when the interpreter exits and runs the CleanUp Macro in
    VBA to show any errors and to reset the StatusBar.
    """
    if is_excel_running():
        # Prevents Excel from reopening
        # if it has been closed manually or never been opened
        for app in Apps():
            try:
                app.xl.run_VB_macro("CleanUp")
            except (CommandError, AttributeError, aem.aemsend.EventError):
                # Excel files initiated from Python don't have the xlwings VBA module
                pass


def posix_to_hfs_path(posix_path):
    """
    Turns a posix path (/Path/file.ext) into an HFS path (Macintosh HD:Path:file.ext)
    """
    dir_name, file_name = os.path.split(posix_path)
    dir_name_hfs = mactypes.Alias(dir_name).hfspath
    return dir_name_hfs + ":" + file_name


def hfs_to_posix_path(hfs_path):
    """
    Turns an HFS path (Macintosh HD:Path:file.ext) into a posix path (/Path/file.ext)
    """
    url = mactypes.convertpathtourl(hfs_path, 1)  # kCFURLHFSPathStyle = 1
    return mactypes.converturltopath(url, 0)  # kCFURLPOSIXPathStyle = 0


def is_excel_running():
    for proc in psutil.process_iter():
        try:
            if proc.name() == "Microsoft Excel":
                return True
        except psutil.NoSuchProcess:
            pass
    return False


# --- constants ---

chart_types_k2s = {
    kw.ThreeD_area: "3d_area",
    kw.ThreeD_area_stacked: "3d_area_stacked",
    kw.ThreeD_area_stacked_100: "3d_area_stacked_100",
    kw.ThreeD_bar_clustered: "3d_bar_clustered",
    kw.ThreeD_bar_stacked: "3d_bar_stacked",
    kw.ThreeD_bar_stacked_100: "3d_bar_stacked_100",
    kw.ThreeD_column: "3d_column",
    kw.ThreeD_column_clustered: "3d_column_clustered",
    kw.ThreeD_column_stacked: "3d_column_stacked",
    kw.ThreeD_column_stacked_100: "3d_column_stacked_100",
    kw.ThreeD_line: "3d_line",
    kw.ThreeD_pie: "3d_pie",
    kw.ThreeD_pie_exploded: "3d_pie_exploded",
    kw.area_chart: "area",
    kw.area_stacked: "area_stacked",
    kw.area_stacked_100: "area_stacked_100",
    kw.bar_clustered: "bar_clustered",
    kw.bar_of_pie: "bar_of_pie",
    kw.bar_stacked: "bar_stacked",
    kw.bar_stacked_100: "bar_stacked_100",
    kw.bubble: "bubble",
    kw.bubble_ThreeD_effect: "bubble_3d_effect",
    kw.column_clustered: "column_clustered",
    kw.column_stacked: "column_stacked",
    kw.column_stacked_100: "column_stacked_100",
    kw.combination_chart: "combination",
    kw.cone_bar_clustered: "cone_bar_clustered",
    kw.cone_bar_stacked: "cone_bar_stacked",
    kw.cone_bar_stacked_100: "cone_bar_stacked_100",
    kw.cone_col: "cone_col",
    kw.cone_column_clustered: "cone_col_clustered",
    kw.cone_column_stacked: "cone_col_stacked",
    kw.cone_column_stacked_100: "cone_col_stacked_100",
    kw.cylinder_bar_clustered: "cylinder_bar_clustered",
    kw.cylinder_bar_stacked: "cylinder_bar_stacked",
    kw.cylinder_bar_stacked_100: "cylinder_bar_stacked_100",
    kw.cylinder_column: "cylinder_col",
    kw.cylinder_column_clustered: "cylinder_col_clustered",
    kw.cylinder_column_stacked: "cylinder_col_stacked",
    kw.cylinder_column_stacked_100: "cylinder_col_stacked_100",
    kw.doughnut: "doughnut",
    kw.doughnut_exploded: "doughnut_exploded",
    kw.line_chart: "line",
    kw.line_markers: "line_markers",
    kw.line_markers_stacked: "line_markers_stacked",
    kw.line_markers_stacked_100: "line_markers_stacked_100",
    kw.line_stacked: "line_stacked",
    kw.line_stacked_100: "line_stacked_100",
    kw.pie_chart: "pie",
    kw.pie_exploded: "pie_exploded",
    kw.pie_of_pie: "pie_of_pie",
    kw.pyramid_bar_clustered: "pyramid_bar_clustered",
    kw.pyramid_bar_stacked: "pyramid_bar_stacked",
    kw.pyramid_bar_stacked_100: "pyramid_bar_stacked_100",
    kw.pyramid_column: "pyramid_col",
    kw.pyramid_column_clustered: "pyramid_col_clustered",
    kw.pyramid_column_stacked: "pyramid_col_stacked",
    kw.pyramid_column_stacked_100: "pyramid_col_stacked_100",
    kw.radar: "radar",
    kw.radar_filled: "radar_filled",
    kw.radar_markers: "radar_markers",
    kw.stock_HLC: "stock_hlc",
    kw.stock_OHLC: "stock_ohlc",
    kw.stock_VHLC: "stock_vhlc",
    kw.stock_VOHLC: "stock_vohlc",
    kw.surface: "surface",
    kw.surface_top_view: "surface_top_view",
    kw.surface_top_view_wireframe: "surface_top_view_wireframe",
    kw.surface_wireframe: "surface_wireframe",
    kw.xy_scatter_lines: "xy_scatter_lines",
    kw.xy_scatter_lines_no_markers: "xy_scatter_lines_no_markers",
    kw.xy_scatter_smooth: "xy_scatter_smooth",
    kw.xy_scatter_smooth_no_markers: "xy_scatter_smooth_no_markers",
    kw.xyscatter: "xy_scatter",
}

chart_types_s2k = {v: k for k, v in chart_types_k2s.items()}

legend_positions_k2s = {
    kw.legend_position_top: "top",
    kw.legend_position_bottom: "bottom",
    kw.legend_position_left: "left",
    kw.legend_position_right: "right",
    kw.legend_position_corner: "corner",
}
legend_positions_s2k = {v: k for k, v in legend_positions_k2s.items()}

# Note the differing keyword prefixes: horizontal_align_* vs vertical_alignment_*
horizontal_alignments_s2k = {
    "general": kw.horizontal_align_general,
    "left": kw.horizontal_align_left,
    "center": kw.horizontal_align_center,
    "right": kw.horizontal_align_right,
    "fill": kw.horizontal_align_fill,
    "justify": kw.horizontal_align_justify,
    "center_across_selection": kw.horizontal_align_center_across_selection,
    "distributed": kw.horizontal_align_distributed,
}
horizontal_alignments_k2s = {v: k for k, v in horizontal_alignments_s2k.items()}

vertical_alignments_s2k = {
    "top": kw.vertical_alignment_top,
    "center": kw.vertical_alignment_center,
    "bottom": kw.vertical_alignment_bottom,
    "justify": kw.vertical_alignment_justify,
    "distributed": kw.vertical_alignment_distributed,
}
vertical_alignments_k2s = {v: k for k, v in vertical_alignments_s2k.items()}

# by_rows is defined twice in mac_dict (XlRowCol and XlSearchOrder); appscript
# packs the first definition, which is the XlRowCol one that plot_by expects
plot_by_k2s = {kw.by_rows: "rows", kw.by_columns: "columns"}
plot_by_s2k = {v: k for k, v in plot_by_k2s.items()}


directions_s2k = {
    "d": kw.toward_the_bottom,
    "down": kw.toward_the_bottom,
    "l": kw.toward_the_left,
    "left": kw.toward_the_left,
    "r": kw.toward_the_right,
    "right": kw.toward_the_right,
    "u": kw.toward_the_top,
    "up": kw.toward_the_top,
}

directions_k2s = {
    kw.toward_the_bottom: "down",
    kw.toward_the_left: "left",
    kw.toward_the_right: "right",
    kw.toward_the_top: "up",
}

calculation_k2s = {
    kw.calculation_automatic: "automatic",
    kw.calculation_manual: "manual",
    kw.calculation_semiautomatic: "semiautomatic",
}

calculation_s2k = {v: k for k, v in calculation_k2s.items()}

shape_types_k2s = {
    kw.shape_type_3d_model: "3d_model",
    kw.shape_type_auto: "auto_shape",
    kw.shape_type_callout: "callout",
    kw.shape_type_canvas: "canvas",
    kw.shape_type_chart: "chart",
    kw.shape_type_comment: "comment",
    kw.shape_type_content_application: "content_app",
    kw.shape_type_diagram: "diagram",
    kw.shape_type_free_form: "free_form",
    kw.shape_type_graphic: "graphic",
    kw.shape_type_group: "group",
    kw.shape_type_embedded_OLE_control: "embedded_ole_object",
    kw.shape_type_form_control: "form_control",
    kw.shape_type_line: "line",
    kw.shape_type_linked_3d_model: "linked_3d_model",
    kw.shape_type_linked_graphic: "linked_graphic",
    kw.shape_type_linked_OLE_object: "linked_ole_object",
    kw.shape_type_linked_picture: "linked_picture",
    kw.shape_type_OLE_control: "ole_control_object",
    kw.shape_type_picture: "picture",
    kw.shape_type_place_holder: "placeholder",
    kw.shape_type_web_video: "web_video",
    kw.shape_type_media: "media",
    kw.shape_type_text_box: "text_box",
    kw.shape_type_table: "table",
    kw.shape_type_ink: "ink",
    kw.shape_type_ink_comment: "ink_comment",
    kw.shape_type_unset: "unset",
    kw.shape_type_slicer: "slicer",
}

scaling = {
    "scale_from_top_left": kw.scale_from_top_left,
    "scale_from_bottom_right": kw.scale_from_bottom_right,
    "scale_from_middle": kw.scale_from_middle,
}

shape_types_s2k = {v: k for k, v in shape_types_k2s.items()}

pivot_functions_s2k = {
    "sum": kw.do_sum,
    "count": kw.do_count,
    "average": kw.do_average,
    "max": kw.do_maximum,
    "min": kw.do_minimum,
    "product": kw.do_product,
    "count_numbers": kw.do_count_numbers,
    "stdev": kw.do_standard_deviation,
    "stdevp": kw.do_standard_deviation_p,
    "var": kw.do_var,
    "varp": kw.do_var_p,
}
pivot_functions_k2s = {v: k for k, v in pivot_functions_s2k.items()}

pivot_layouts_s2k = {
    "compact": kw.compact_row,
    "outline": kw.outline_row,
    "tabular": kw.tabular_row,
}
pivot_layouts_k2s = {v: k for k, v in pivot_layouts_s2k.items()}

# xlwings' field areas -> the pivot table element / the orientation
pivot_area_elements = {
    "rows": "row_fields",
    "columns": "column_fields",
    "filters": "page_fields",
}
pivot_area_orientations = {
    "rows": kw.orient_as_row_field,
    "columns": kw.orient_as_column_field,
    "filters": kw.orient_as_page_field,
}
