# TODO: add all classes and use in _mac.py and _windows.py


class Apps:
    def keys(self):
        raise NotImplementedError()

    def add(self, spec=None, add_book=None, xl=None, visible=None):
        raise NotImplementedError()

    @staticmethod
    def cleanup():
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
    def path(self):
        raise NotImplementedError()

    @property
    def pid(self):
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

    def alert(self, prompt, title, buttons, mode, callback):
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
    def freeze_panes(self):
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

    def to_html(self, path):
        raise NotImplementedError()


class Range:
    def adjust_indent(self, amount):
        raise NotImplementedError()

    def group(self, by):
        raise NotImplementedError()

    def ungroup(self, by):
        raise NotImplementedError()

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

    def copy_from(
        self, source_range, copy_type="all", skip_blanks=False, transpose=False
    ):
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

    def autofill(self, destination, type_):
        raise NotImplementedError()


class Picture:
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
    def parent(self):
        raise NotImplementedError()

    @property
    def left(self):
        raise NotImplementedError()

    @left.setter
    def left(self, value):
        raise NotImplementedError()

    @property
    def top(self):
        raise NotImplementedError()

    @top.setter
    def top(self, value):
        raise NotImplementedError()

    @property
    def width(self):
        raise NotImplementedError()

    @width.setter
    def width(self, value):
        raise NotImplementedError()

    @property
    def height(self):
        raise NotImplementedError()

    @height.setter
    def height(self, value):
        raise NotImplementedError()

    def delete(self):
        raise NotImplementedError()

    @property
    def lock_aspect_ratio(self):
        raise NotImplementedError()

    @lock_aspect_ratio.setter
    def lock_aspect_ratio(self, value):
        raise NotImplementedError()

    def index(self):
        raise NotImplementedError()


class Collection:
    @property
    def api(self):
        raise NotImplementedError()

    @property
    def parent(self):
        raise NotImplementedError()

    def __call__(self, key):
        raise NotImplementedError()

    def __len__(self):
        raise NotImplementedError()

    def __iter__(self):
        raise NotImplementedError()

    def __contains__(self, key):
        raise NotImplementedError()


class Pictures:
    def add(self, filename, link_to_file, save_with_document, left, top, width, height):
        raise NotImplementedError()


class Names:
    # @property
    # def api(self):
    #     raise NotImplementedError()

    def __call__(self, name_or_index):
        raise NotImplementedError()

    def contains(self, name_or_index):
        raise NotImplementedError()

    def __len__(self):
        raise NotImplementedError()

    def add(self, name, refers_to):
        raise NotImplementedError()


class Name:
    # @property
    # def api(self):
    #     raise NotImplementedError()

    def delete(self):
        raise NotImplementedError()

    @property
    def name(self):
        raise NotImplementedError()

    @name.setter
    def name(self, value):
        raise NotImplementedError()

    @property
    def refers_to(self):
        raise NotImplementedError()

    @refers_to.setter
    def refers_to(self, value):
        raise NotImplementedError()

    @property
    def refers_to_range(self):
        raise NotImplementedError()


class Shape:
    @property
    def api(self):
        raise NotImplementedError()

    @property
    def name(self):
        raise NotImplementedError()

    @property
    def parent(self):
        raise NotImplementedError()

    @property
    def type(self):
        raise NotImplementedError()

    @property
    def left(self):
        raise NotImplementedError()

    @left.setter
    def left(self, value):
        raise NotImplementedError()

    @property
    def top(self):
        raise NotImplementedError()

    @top.setter
    def top(self, value):
        raise NotImplementedError()

    @property
    def width(self):
        raise NotImplementedError()

    @width.setter
    def width(self, value):
        raise NotImplementedError()

    @property
    def height(self):
        raise NotImplementedError()

    @height.setter
    def height(self, value):
        raise NotImplementedError()

    def delete(self):
        raise NotImplementedError()

    @name.setter
    def name(self, value):
        raise NotImplementedError()

    @property
    def index(self):
        raise NotImplementedError()

    def activate(self):
        raise NotImplementedError()

    def scale_height(self, factor, relative_to_original_size, scale):
        raise NotImplementedError()

    def scale_width(self, factor, relative_to_original_size, scale):
        raise NotImplementedError()

    @property
    def text(self):
        raise NotImplementedError()

    @text.setter
    def text(self, value):
        raise NotImplementedError()

    @property
    def font(self):
        raise NotImplementedError()

    @property
    def characters(self):
        raise NotImplementedError()


class Font:
    @property
    def api(self):
        raise NotImplementedError()

    @property
    def bold(self):
        raise NotImplementedError()

    @bold.setter
    def bold(self, value):
        raise NotImplementedError()

    @property
    def italic(self):
        raise NotImplementedError()

    @italic.setter
    def italic(self, value):
        raise NotImplementedError()

    @property
    def size(self):
        raise NotImplementedError()

    @size.setter
    def size(self, value):
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


class Characters:
    @property
    def api(self):
        raise NotImplementedError()

    @property
    def text(self):
        raise NotImplementedError()

    @property
    def font(self):
        raise NotImplementedError()

    def __getitem__(self, item):
        raise NotImplementedError()


class PageSetup:
    @property
    def api(self):
        raise NotImplementedError()

    @property
    def print_area(self):
        raise NotImplementedError()

    @print_area.setter
    def print_area(self, value):
        raise NotImplementedError()


class Note:
    @property
    def api(self):
        raise NotImplementedError()

    @property
    def text(self):
        raise NotImplementedError()

    @text.setter
    def text(self, value):
        raise NotImplementedError()

    def delete(self):
        raise NotImplementedError()


class Table:
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
    def data_body_range(self):
        raise NotImplementedError()

    @property
    def display_name(self):
        # This seems to be equivalent to name and Office Scripts has dropped it
        raise NotImplementedError()

    @display_name.setter
    def display_name(self, value):
        raise NotImplementedError()

    @property
    def header_row_range(self):
        raise NotImplementedError()

    @property
    def insert_row_range(self):
        raise NotImplementedError()

    @property
    def parent(self):
        raise NotImplementedError()

    @property
    def range(self):
        raise NotImplementedError()

    @property
    def show_autofilter(self):
        raise NotImplementedError()

    @show_autofilter.setter
    def show_autofilter(self, value):
        raise NotImplementedError()

    @property
    def show_headers(self):
        raise NotImplementedError()

    @show_headers.setter
    def show_headers(self, value):
        raise NotImplementedError()

    @property
    def show_table_style_column_stripes(self):
        raise NotImplementedError()

    @show_table_style_column_stripes.setter
    def show_table_style_column_stripes(self, value):
        raise NotImplementedError()

    @property
    def show_table_style_first_column(self):
        raise NotImplementedError()

    @show_table_style_first_column.setter
    def show_table_style_first_column(self, value):
        raise NotImplementedError()

    @property
    def show_table_style_last_column(self):
        raise NotImplementedError()

    @show_table_style_last_column.setter
    def show_table_style_last_column(self, value):
        raise NotImplementedError()

    @property
    def show_table_style_row_stripes(self):
        raise NotImplementedError()

    @show_table_style_row_stripes.setter
    def show_table_style_row_stripes(self, value):
        raise NotImplementedError()

    @property
    def show_totals(self):
        raise NotImplementedError()

    @show_totals.setter
    def show_totals(self, value):
        raise NotImplementedError()

    @property
    def table_style(self):
        raise NotImplementedError()

    @table_style.setter
    def table_style(self, value):
        raise NotImplementedError()

    @property
    def totals_row_range(self):
        raise NotImplementedError()

    def resize(self, range):
        raise NotImplementedError()


class Tables:
    def add(
        self,
        source_type=None,
        source=None,
        link_source=None,
        has_headers=None,
        destination=None,
        table_style_name=None,
    ):
        raise NotImplementedError()


class Chart:
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
    def parent(self):
        raise NotImplementedError()

    def set_source_data(self, rng):
        raise NotImplementedError()

    @property
    def chart_type(self):
        raise NotImplementedError()

    @chart_type.setter
    def chart_type(self, chart_type):
        raise NotImplementedError()

    @property
    def left(self):
        raise NotImplementedError()

    @left.setter
    def left(self, value):
        raise NotImplementedError()

    @property
    def top(self):
        raise NotImplementedError()

    @top.setter
    def top(self, value):
        raise NotImplementedError()

    @property
    def width(self):
        raise NotImplementedError()

    @width.setter
    def width(self, value):
        raise NotImplementedError()

    @property
    def height(self):
        raise NotImplementedError()

    @height.setter
    def height(self, value):
        raise NotImplementedError()

    def delete(self):
        raise NotImplementedError()

    def to_png(self, path):
        raise NotImplementedError()

    def to_pdf(self, path, quality):
        raise NotImplementedError()


class Charts:
    def _wrap(self, xl):
        raise NotImplementedError()

    def add(self, left, top, width, height):
        raise NotImplementedError()


class FreezePanes:
    def freeze_at(self, frozen_range):
        raise NotImplementedError()

    def unfreeze():
        raise NotImplementedError()
