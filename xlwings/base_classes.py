from typing import Any, Literal, get_args


class _Unset:
    """Type of the `_UNSET` sentinel; the repr keeps autodoc signatures readable."""

    def __repr__(self) -> str:
        return "..."


# Sentinel for "attribute not supplied" in Borders.set(). Shared by main.Borders and
# the engine implementations; always compare by identity (``value is _UNSET``).
# Typed as Any so that it can be the default of a typed keyword argument.
_UNSET: Any = _Unset()

# The border vocabulary. Plain lowercase strings are the API; the Literal
# aliases give editors autocomplete and type checkers typo detection.
BorderSide = Literal[
    "edge_top",
    "edge_bottom",
    "edge_left",
    "edge_right",
    "inside_vertical",
    "inside_horizontal",
    "diagonal_down",
    "diagonal_up",
]
BorderGroup = Literal["outside", "inside", "all", "everything"]

# Office.js limits range get operations to 5,000,000 cells on all platforms.
# A lower cross-engine default also reduces timeout and memory pressure on
# desktop engines. Individual engines may override or disable it.
DEFAULT_MAX_CELLS_PER_READ = 4_000_000

# Initial write heuristic; cell count does not guarantee a payload byte size.
# Engines can tune this independently of the read budget.
DEFAULT_MAX_CELLS_PER_WRITE = 100_000
BorderLineStyle = Literal[
    "continuous",
    "dash",
    "dash_dot",
    "dash_dot_dot",
    "dot",
    "double",
    "slant_dash_dot",
    "none",
]
BorderWeight = Literal["hairline", "thin", "medium", "thick"]

# Border side names in canonical order. main.Borders validates and expands the
# user-facing selectors into these before calling an engine, so engines only
# ever see the canonical names. The first six are the grid sides that the
# collection-level properties read and write; the diagonals are reachable
# individually only.
BORDER_SIDES: tuple[str, ...] = get_args(BorderSide)
BORDER_GRID_SIDES = BORDER_SIDES[:6]

# The chart vocabulary, same idea as the borders: lowercase strings are the API,
# main.Chart validates them so engines only ever see the canonical names.
ChartLegendPosition = Literal["top", "bottom", "left", "right", "corner"]
ChartPlotBy = Literal["rows", "columns"]
CHART_LEGEND_POSITIONS: tuple[str, ...] = get_args(ChartLegendPosition)
CHART_PLOT_BY: tuple[str, ...] = get_args(ChartPlotBy)
# The documented chart type names (see main.Chart.chart_type); the desktop
# engines map all of them, the remote engine all but "combination".
CHART_TYPES: tuple[str, ...] = (
    "3d_area",
    "3d_area_stacked",
    "3d_area_stacked_100",
    "3d_bar_clustered",
    "3d_bar_stacked",
    "3d_bar_stacked_100",
    "3d_column",
    "3d_column_clustered",
    "3d_column_stacked",
    "3d_column_stacked_100",
    "3d_line",
    "3d_pie",
    "3d_pie_exploded",
    "area",
    "area_stacked",
    "area_stacked_100",
    "bar_clustered",
    "bar_of_pie",
    "bar_stacked",
    "bar_stacked_100",
    "bubble",
    "bubble_3d_effect",
    "column_clustered",
    "column_stacked",
    "column_stacked_100",
    "combination",
    "cone_bar_clustered",
    "cone_bar_stacked",
    "cone_bar_stacked_100",
    "cone_col",
    "cone_col_clustered",
    "cone_col_stacked",
    "cone_col_stacked_100",
    "cylinder_bar_clustered",
    "cylinder_bar_stacked",
    "cylinder_bar_stacked_100",
    "cylinder_col",
    "cylinder_col_clustered",
    "cylinder_col_stacked",
    "cylinder_col_stacked_100",
    "doughnut",
    "doughnut_exploded",
    "line",
    "line_markers",
    "line_markers_stacked",
    "line_markers_stacked_100",
    "line_stacked",
    "line_stacked_100",
    "pie",
    "pie_exploded",
    "pie_of_pie",
    "pyramid_bar_clustered",
    "pyramid_bar_stacked",
    "pyramid_bar_stacked_100",
    "pyramid_col",
    "pyramid_col_clustered",
    "pyramid_col_stacked",
    "pyramid_col_stacked_100",
    "radar",
    "radar_filled",
    "radar_markers",
    "stock_hlc",
    "stock_ohlc",
    "stock_vhlc",
    "stock_vohlc",
    "surface",
    "surface_top_view",
    "surface_top_view_wireframe",
    "surface_wireframe",
    "xy_scatter",
    "xy_scatter_lines",
    "xy_scatter_lines_no_markers",
    "xy_scatter_smooth",
    "xy_scatter_smooth_no_markers",
)


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

    async def get_selection(self):
        raise NotImplementedError(
            "App.get_selection() is only supported in xlwings Lite"
        )


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

    async def get_active(self):
        raise NotImplementedError(
            "Books.get_active() is only supported in xlwings Lite"
        )

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

    async def load(self, values=None):
        raise NotImplementedError("Book.load() is only supported in xlwings Lite")

    async def flush(self):
        raise NotImplementedError("Book.flush() is only supported in xlwings Lite")


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

    async def get_active(self):
        raise NotImplementedError(
            "Sheets.get_active() is only supported in xlwings Lite"
        )


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
    def show_gridlines(self):
        raise NotImplementedError()

    @show_gridlines.setter
    def show_gridlines(self, value):
        raise NotImplementedError()

    @property
    def page_setup(self):
        raise NotImplementedError()

    def to_html(self, path):
        raise NotImplementedError()

    async def load(self, values=None):
        raise NotImplementedError("Sheet.load() is only supported in xlwings Lite")


class Range:
    def get_async_pipeline_overrides(self, options):
        raise NotImplementedError("get_value() is only supported in xlwings Lite")

    async def get_formula(self):
        raise NotImplementedError("get_formula() is only supported in xlwings Lite")

    async def get_formula_array(self):
        raise NotImplementedError(
            "get_formula_array() is only supported in xlwings Lite"
        )

    async def get_number_format(self):
        raise NotImplementedError(
            "get_number_format() is only supported in xlwings Lite"
        )

    async def get_wrap_text(self):
        raise NotImplementedError("get_wrap_text() is only supported in xlwings Lite")

    async def get_column_width(self):
        raise NotImplementedError(
            "get_column_width() is only supported in xlwings Lite"
        )

    async def get_row_height(self):
        raise NotImplementedError("get_row_height() is only supported in xlwings Lite")

    async def get_left(self):
        raise NotImplementedError("get_left() is only supported in xlwings Lite")

    async def get_top(self):
        raise NotImplementedError("get_top() is only supported in xlwings Lite")

    async def get_width(self):
        raise NotImplementedError("get_width() is only supported in xlwings Lite")

    async def get_height(self):
        raise NotImplementedError("get_height() is only supported in xlwings Lite")

    async def get_hyperlink(self):
        raise NotImplementedError("get_hyperlink() is only supported in xlwings Lite")

    async def get_current_region(self):
        raise NotImplementedError(
            "get_current_region() is only supported in xlwings Lite"
        )

    async def get_merge_area(self):
        raise NotImplementedError("get_merge_area() is only supported in xlwings Lite")

    async def get_merge_cells(self):
        raise NotImplementedError("get_merge_cells() is only supported in xlwings Lite")

    async def get_table(self):
        raise NotImplementedError("get_table() is only supported in xlwings Lite")

    async def get_color(self):
        raise NotImplementedError("Range.get_color() is only supported in xlwings Lite")

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

    @property
    def max_cells_per_read(self):
        """Cell budget above which value reads are chunked automatically when the
        user hasn't passed an explicit ``chunksize``. Engines may override this;
        ``None`` disables implicit read chunking for the engine."""
        return DEFAULT_MAX_CELLS_PER_READ

    @property
    def max_cells_per_write(self):
        """Cell budget above which value writes are chunked automatically when the
        user hasn't passed an explicit ``chunksize``. Engines may override this;
        ``None`` disables implicit write chunking for the engine."""
        return DEFAULT_MAX_CELLS_PER_WRITE

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
    def borders(self):
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
        """Return a handle that retains its identity when native indices shift.

        Reading a broken reference can make Excel insert internal names. Existing
        handles must still address the original entry, including its scope, after
        that insertion. Do not rebind indices through ambiguous string lookups.
        """
        raise NotImplementedError()

    def snapshot(self):
        """Return (name text, stable handle) pairs without resolving references.

        Engines can override this to read all name strings in one native call.
        """
        names = [self(i + 1) for i in range(len(self))]
        return [(name.name, name) for name in names]

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

    async def get_text(self):
        raise NotImplementedError("Shape.get_text() is only supported in xlwings Lite")


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

    async def get_bold(self):
        raise NotImplementedError("get_bold() is only supported in xlwings Lite")

    async def get_italic(self):
        raise NotImplementedError("get_italic() is only supported in xlwings Lite")

    async def get_size(self):
        raise NotImplementedError("get_size() is only supported in xlwings Lite")

    async def get_name(self):
        raise NotImplementedError("Font.get_name() is only supported in xlwings Lite")

    async def get_color(self):
        raise NotImplementedError("Font.get_color() is only supported in xlwings Lite")


class Border:
    @property
    def api(self):
        raise NotImplementedError()

    @property
    def line_style(self):
        raise NotImplementedError()

    @line_style.setter
    def line_style(self, value):
        raise NotImplementedError()

    @property
    def weight(self):
        raise NotImplementedError()

    @weight.setter
    def weight(self, value):
        raise NotImplementedError()

    @property
    def color(self):
        raise NotImplementedError()

    @color.setter
    def color(self, color_or_rgb):
        raise NotImplementedError()

    async def get_line_style(self):
        raise NotImplementedError(
            "Border.get_line_style() is only supported in xlwings Lite"
        )

    async def get_weight(self):
        raise NotImplementedError(
            "Border.get_weight() is only supported in xlwings Lite"
        )

    async def get_color(self):
        raise NotImplementedError(
            "Border.get_color() is only supported in xlwings Lite"
        )


class Borders:
    def _grid_sides(self):
        """Grid sides that exist for this range's dimensions."""
        nrows, ncols = self.parent.shape
        return tuple(
            side
            for side in BORDER_GRID_SIDES
            if (side != "inside_vertical" or ncols > 1)
            and (side != "inside_horizontal" or nrows > 1)
        )

    @property
    def api(self):
        raise NotImplementedError()

    @property
    def line_style(self):
        raise NotImplementedError()

    @line_style.setter
    def line_style(self, value):
        raise NotImplementedError()

    @property
    def weight(self):
        raise NotImplementedError()

    @weight.setter
    def weight(self, value):
        raise NotImplementedError()

    @property
    def color(self):
        raise NotImplementedError()

    @color.setter
    def color(self, color_or_rgb):
        raise NotImplementedError()

    async def get_line_style(self):
        raise NotImplementedError(
            "Borders.get_line_style() is only supported in xlwings Lite"
        )

    async def get_weight(self):
        raise NotImplementedError(
            "Borders.get_weight() is only supported in xlwings Lite"
        )

    async def get_color(self):
        raise NotImplementedError(
            "Borders.get_color() is only supported in xlwings Lite"
        )

    def __getitem__(self, key):
        raise NotImplementedError()

    def set(self, which="all", *, line_style=_UNSET, weight=_UNSET, color=_UNSET):
        raise NotImplementedError()

    def clear(self, which="everything"):
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

    async def get_text(self):
        raise NotImplementedError(
            "Characters.get_text() is only supported in xlwings Lite"
        )


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

    async def get_text(self):
        raise NotImplementedError("Note.get_text() is only supported in xlwings Lite")

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

    def set_source_data(self, rng, plot_by=None):
        raise NotImplementedError()

    @property
    def chart_type(self):
        raise NotImplementedError()

    @chart_type.setter
    def chart_type(self, chart_type):
        raise NotImplementedError()

    @property
    def title(self):
        raise NotImplementedError()

    @title.setter
    def title(self, value):
        raise NotImplementedError()

    @property
    def legend(self):
        raise NotImplementedError()

    @property
    def plot_by(self):
        raise NotImplementedError()

    @plot_by.setter
    def plot_by(self, value):
        raise NotImplementedError()

    @property
    def style(self):
        raise NotImplementedError()

    @style.setter
    def style(self, value):
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

    async def get_png(self):
        raise NotImplementedError("get_png() is only supported in xlwings Lite")


class ChartLegend:
    @property
    def api(self):
        raise NotImplementedError()

    @property
    def visible(self):
        raise NotImplementedError()

    @visible.setter
    def visible(self, value):
        raise NotImplementedError()

    @property
    def position(self):
        raise NotImplementedError()

    @position.setter
    def position(self, value):
        raise NotImplementedError()


class Charts:
    def _wrap(self, xl):
        raise NotImplementedError()

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
    ):
        raise NotImplementedError()


class FreezePanes:
    def freeze_at(self, frozen_range):
        raise NotImplementedError()

    def unfreeze():
        raise NotImplementedError()
