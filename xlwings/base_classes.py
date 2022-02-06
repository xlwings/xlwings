# TODO: add all classes and use in _mac.py and _windows.py


class Apps:
    def keys(self):
        raise NotImplementedError()

    def add(self, spec=None, add_book=None, xl=None, visible=None):
        raise NotImplementedError()

    def __iter__(self):
        raise NotImplementedError()

    def __len__(self):
        raise NotImplementedError()

    def __getitem__(self, pid):
        raise NotImplementedError()


class App:
    @property
    def xl(self):
        raise NotImplementedError()

    @xl.setter
    def xl(self, value):
        raise NotImplementedError()

    @property
    def api(self):
        raise NotImplementedError()

    @property
    def selection(self):
        raise NotImplementedError()

    def activate(self, steal_focus=False):
        raise NotImplementedError()

    @property
    def visible(self):
        raise NotImplementedError()

    @visible.setter
    def visible(self, visible):
        raise NotImplementedError()

    def quit(self):
        raise NotImplementedError()

    def kill(self):
        raise NotImplementedError()

    @property
    def screen_updating(self):
        raise NotImplementedError()

    @screen_updating.setter
    def screen_updating(self, value):
        raise NotImplementedError()

    @property
    def display_alerts(self):
        raise NotImplementedError()

    @display_alerts.setter
    def display_alerts(self, value):
        raise NotImplementedError()

    @property
    def enable_events(self):
        raise NotImplementedError()

    @enable_events.setter
    def enable_events(self, value):
        raise NotImplementedError()

    @property
    def interactive(self):
        raise NotImplementedError()

    @interactive.setter
    def interactive(self, value):
        raise NotImplementedError()

    @property
    def startup_path(self):
        raise NotImplementedError()

    @property
    def calculation(self):
        raise NotImplementedError()

    @calculation.setter
    def calculation(self, value):
        raise NotImplementedError()

    def calculate(self):
        raise NotImplementedError()

    @property
    def version(self):
        raise NotImplementedError()

    @property
    def books(self):
        raise NotImplementedError()

    @property
    def hwnd(self):
        raise NotImplementedError()

    @property
    def pid(self):
        raise NotImplementedError()

    def range(self, arg1, arg2=None):
        raise NotImplementedError()

    def run(self, macro, args):
        raise NotImplementedError()

    @property
    def status_bar(self):
        raise NotImplementedError()

    @status_bar.setter
    def status_bar(self, value):
        raise NotImplementedError()

    @property
    def cut_copy_mode(self):
        raise NotImplementedError()

    @cut_copy_mode.setter
    def cut_copy_mode(self, value):
        raise NotImplementedError()


class Books:
    @property
    def api(self):
        raise NotImplementedError()

    @property
    def active(self):
        raise NotImplementedError()

    def __call__(self, name_or_index):
        raise NotImplementedError()

    def __len__(self):
        raise NotImplementedError()

    def add(self):
        raise NotImplementedError()

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
        raise NotImplementedError()

    def __iter__(self):
        raise NotImplementedError()


class Book:
    @property
    def api(self):
        raise NotImplementedError()

    def json(self):
        raise NotImplementedError()

    @property
    def name(self):
        raise NotImplementedError()

    @property
    def sheets(self):
        raise NotImplementedError()

    @property
    def app(self):
        raise NotImplementedError()

    def close(self):
        raise NotImplementedError()

    def save(self, path=None, password=None):
        raise NotImplementedError()

    @property
    def fullname(self):
        raise NotImplementedError()

    @property
    def names(self):
        raise NotImplementedError()

    def activate(self):
        raise NotImplementedError()

    def to_pdf(self, path, quality):
        raise NotImplementedError()


class Sheets:
    @property
    def api(self):
        raise NotImplementedError()

    @property
    def active(self):
        raise NotImplementedError()

    def __call__(self, name_or_index):
        raise NotImplementedError()

    def __len__(self):
        raise NotImplementedError()

    def __iter__(self):
        raise NotImplementedError()

    def add(self, before=None, after=None):
        raise NotImplementedError()


class Sheet:
    @property
    def api(self):
        raise NotImplementedError()

    @property
    def name(self):
        raise NotImplementedError()

    @name.setter
    def name(self, value):
        raise NotImplementedError()

    @property
    def names(self):
        raise NotImplementedError()

    @property
    def book(self):
        raise NotImplementedError()

    @property
    def index(self):
        raise NotImplementedError()

    def range(self, arg1, arg2=None):
        raise NotImplementedError()

    @property
    def cells(self):
        raise NotImplementedError()

    def activate(self):
        raise NotImplementedError()

    def select(self):
        raise NotImplementedError()

    def clear_contents(self):
        raise NotImplementedError()

    def clear_formats(self):
        raise NotImplementedError()

    def clear(self):
        raise NotImplementedError()

    def autofit(self, axis=None):
        raise NotImplementedError()

    def delete(self):
        raise NotImplementedError()

    def copy(self, before, after):
        raise NotImplementedError()

    @property
    def charts(self):
        raise NotImplementedError()

    @property
    def shapes(self):
        raise NotImplementedError()

    @property
    def tables(self):
        raise NotImplementedError()

    @property
    def pictures(self):
        raise NotImplementedError()

    @property
    def used_range(self):
        raise NotImplementedError()

    @property
    def visible(self):
        raise NotImplementedError()

    @visible.setter
    def visible(self, value):
        raise NotImplementedError()

    @property
    def page_setup(self):
        raise NotImplementedError()

    def to_pdf(self, path, quality):
        raise NotImplementedError()


class Range:
    @property
    def coords(self):
        raise NotImplementedError()

    @property
    def api(self):
        raise NotImplementedError()

    def __len__(self):
        raise NotImplementedError()

    @property
    def row(self):
        raise NotImplementedError()

    @property
    def column(self):
        raise NotImplementedError()

    @property
    def shape(self):
        raise NotImplementedError()

    @property
    def raw_value(self):
        raise NotImplementedError()

    @raw_value.setter
    def raw_value(self, value):
        raise NotImplementedError()

    def clear_contents(self):
        raise NotImplementedError()

    def clear_formats(self):
        raise NotImplementedError()

    def clear(self):
        raise NotImplementedError()

    def end(self, direction):
        raise NotImplementedError()

    @property
    def formula(self):
        raise NotImplementedError()

    @formula.setter
    def formula(self, value):
        raise NotImplementedError()

    @property
    def formula2(self):
        raise NotImplementedError()

    @formula2.setter
    def formula2(self, value):
        raise NotImplementedError()

    @property
    def formula_array(self):
        raise NotImplementedError()

    @formula_array.setter
    def formula_array(self, value):
        raise NotImplementedError()

    @property
    def font(self):
        raise NotImplementedError()

    @property
    def column_width(self):
        raise NotImplementedError()

    @column_width.setter
    def column_width(self, value):
        raise NotImplementedError()

    @property
    def row_height(self):
        raise NotImplementedError()

    @row_height.setter
    def row_height(self, value):
        raise NotImplementedError()

    @property
    def width(self):
        raise NotImplementedError()

    @property
    def height(self):
        raise NotImplementedError()

    @property
    def left(self):
        raise NotImplementedError()

    @property
    def top(self):
        raise NotImplementedError()

    @property
    def has_array(self):
        raise NotImplementedError()

    @property
    def number_format(self):
        raise NotImplementedError()

    @number_format.setter
    def number_format(self, value):
        raise NotImplementedError()

    def get_address(self, row_absolute, col_absolute, external):
        raise NotImplementedError()

    @property
    def address(self):
        raise NotImplementedError()

    @property
    def current_region(self):
        raise NotImplementedError()

    def autofit(self, axis=None):
        raise NotImplementedError()

    def insert(self, shift=None, copy_origin=None):
        raise NotImplementedError()

    def delete(self, shift=None):
        raise NotImplementedError()

    def copy(self, destination=None):
        raise NotImplementedError()

    def paste(self, paste=None, operation=None, skip_blanks=False, transpose=False):
        raise NotImplementedError()

    @property
    def hyperlink(self):
        raise NotImplementedError()

    def add_hyperlink(self, address, text_to_display=None, screen_tip=None):
        raise NotImplementedError()

    @property
    def color(self):
        raise NotImplementedError()

    @color.setter
    def color(self, color_or_rgb):
        raise NotImplementedError()

    @property
    def name(self):
        raise NotImplementedError()

    @name.setter
    def name(self, value):
        raise NotImplementedError()

    def __call__(self, arg1, arg2=None):
        raise NotImplementedError()

    @property
    def rows(self):
        raise NotImplementedError()

    @property
    def columns(self):
        raise NotImplementedError()

    def select(self):
        raise NotImplementedError()

    @property
    def merge_area(self):
        raise NotImplementedError()

    @property
    def merge_cells(self):
        raise NotImplementedError()

    def merge(self, across):
        raise NotImplementedError()

    def unmerge(self):
        raise NotImplementedError()

    @property
    def table(self):
        raise NotImplementedError()

    @property
    def characters(self):
        raise NotImplementedError()

    @property
    def wrap_text(self):
        raise NotImplementedError()

    @wrap_text.setter
    def wrap_text(self, value):
        raise NotImplementedError()

    @property
    def note(self):
        raise NotImplementedError()

    def copy_picture(self, appearance, format):
        raise NotImplementedError()

    def to_png(self, path):
        raise NotImplementedError()

    def to_pdf(self, path, quality):
        raise NotImplementedError()
