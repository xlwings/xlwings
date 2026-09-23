import math
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

DataValidationType = Literal[
    "none",
    "whole_number",
    "decimal",
    "list",
    "date",
    "time",
    "text_length",
    "custom",
    "inconsistent",
    "mixed_criteria",
    "unknown",
]
DATA_VALIDATION_TYPES: tuple[str, ...] = get_args(DataValidationType)
DataValidationOperator = Literal[
    "between",
    "not_between",
    "equal_to",
    "not_equal_to",
    "greater_than",
    "less_than",
    "greater_than_or_equal",
    "less_than_or_equal",
]
DATA_VALIDATION_OPERATORS: tuple[str, ...] = get_args(DataValidationOperator)
DataValidationAlertStyle = Literal["stop", "warning", "information"]

AutoFilterComparisonOperator = Literal[
    "between",
    "not_between",
    "equal_to",
    "not_equal_to",
    "greater_than",
    "less_than",
    "greater_than_or_equal",
    "less_than_or_equal",
]
AUTOFILTER_COMPARISON_OPERATORS: tuple[str, ...] = get_args(
    AutoFilterComparisonOperator
)
AutoFilterCriteriaType = Literal[
    "none",
    "values",
    "comparison",
    "top_items",
    "bottom_items",
    "top_percent",
    "bottom_percent",
    "unknown",
]
AUTOFILTER_CRITERIA_TYPES: tuple[str, ...] = get_args(AutoFilterCriteriaType)


def empty_autofilter_criteria(field: int, type_: str = "none") -> dict[str, Any]:
    return {
        "field": field,
        "type": type_,
        "values": None,
        "operator": None,
        "value1": None,
        "value2": None,
        "count": None,
        "percent": None,
    }


def _unescape_autofilter_value(value: str) -> str:
    result = []
    index = 0
    while index < len(value):
        if value[index] == "~" and index + 1 < len(value) and value[index + 1] in "~*?":
            index += 1
        result.append(value[index])
        index += 1
    return "".join(result)


def _split_autofilter_comparison(value: Any) -> tuple[str | None, str | None]:
    if not isinstance(value, str):
        return None, None
    for prefix in (">=", "<=", "<>", ">", "<", "="):
        if value.startswith(prefix):
            operand = value[len(prefix) :]
            return prefix, _unescape_autofilter_value(operand) if operand else None
    return "=", _unescape_autofilter_value(value)


def autofilter_criteria_snapshot(
    field: int,
    type_: str,
    criteria1: Any = None,
    criteria2: Any = None,
    join_operator: str | None = None,
) -> dict[str, Any]:
    snapshot = empty_autofilter_criteria(field, type_)
    if type_ == "values":
        values = criteria1 if isinstance(criteria1, (list, tuple)) else [criteria1]
        if any(not isinstance(value, str) for value in values):
            return empty_autofilter_criteria(field, "unknown")
        snapshot["values"] = [
            value[1:] if value.startswith("=") else value for value in values
        ]
        return snapshot
    if type_ in ("top_items", "bottom_items"):
        value = str(criteria1).removeprefix("=")
        if value.startswith((">", "<")):
            return snapshot
        try:
            snapshot["count"] = int(value)
        except (TypeError, ValueError):
            return empty_autofilter_criteria(field, "unknown")
        return snapshot
    if type_ in ("top_percent", "bottom_percent"):
        value = str(criteria1).removeprefix("=")
        if value.startswith((">", "<")):
            return snapshot
        try:
            snapshot["percent"] = float(value)
        except (TypeError, ValueError):
            return empty_autofilter_criteria(field, "unknown")
        if not math.isfinite(snapshot["percent"]):
            return empty_autofilter_criteria(field, "unknown")
        return snapshot
    if type_ != "comparison":
        return snapshot

    prefix1, value1 = _split_autofilter_comparison(criteria1)
    prefix2, value2 = _split_autofilter_comparison(criteria2)
    if prefix1 is None:
        return empty_autofilter_criteria(field, "unknown")
    if criteria2 is not None:
        if (
            join_operator == "or"
            and (prefix1, prefix2) == ("=", "=")
            and value1 is not None
            and value2 is not None
        ):
            snapshot["type"] = "values"
            snapshot["values"] = [value1, value2]
            return snapshot
        if join_operator == "and" and (prefix1, prefix2) == (">=", "<="):
            operator = "between"
        elif join_operator == "or" and (prefix1, prefix2) == ("<", ">"):
            operator = "not_between"
        else:
            return empty_autofilter_criteria(field, "unknown")
    else:
        operator = {
            "=": "equal_to",
            "<>": "not_equal_to",
            ">": "greater_than",
            "<": "less_than",
            ">=": "greater_than_or_equal",
            "<=": "less_than_or_equal",
        }.get(prefix1)
        if operator is None:
            return empty_autofilter_criteria(field, "unknown")
    snapshot["operator"] = operator
    snapshot["value1"] = value1
    snapshot["value2"] = value2
    return snapshot


# Conditional-format types supported by the first public rule model. Other
# native rule types remain visible as `unknown` so callers can inspect and
# delete them without the engines silently dropping them from the collection.
ConditionalFormatType = Literal[
    "cell_value",
    "custom",
    "color_scale",
    "data_bar",
    "icon_set",
    "unknown",
]
CONDITIONAL_FORMAT_TYPES: tuple[str, ...] = get_args(ConditionalFormatType)
ConditionalFormatOperator = Literal[
    "between",
    "not_between",
    "equal_to",
    "not_equal_to",
    "greater_than",
    "less_than",
    "greater_than_or_equal",
    "less_than_or_equal",
]
CONDITIONAL_FORMAT_OPERATORS: tuple[str, ...] = get_args(ConditionalFormatOperator)
ConditionalFormatThresholdType = Literal["number", "percent", "percentile"]
CONDITIONAL_FORMAT_THRESHOLD_TYPES: tuple[str, ...] = get_args(
    ConditionalFormatThresholdType
)
ConditionalFormatCriterionType = Literal[
    "automatic",
    "lowest_value",
    "highest_value",
    "number",
    "percent",
    "percentile",
    "formula",
    "unknown",
]
ConditionalFormatIconSet = Literal[
    "3_arrows",
    "3_arrows_gray",
    "3_flags",
    "3_traffic_lights_1",
    "3_traffic_lights_2",
    "3_signs",
    "3_symbols",
    "3_symbols_2",
    "4_arrows",
    "4_arrows_gray",
    "4_red_to_black",
    "4_rating",
    "4_traffic_lights",
    "5_arrows",
    "5_arrows_gray",
    "5_rating",
    "5_quarters",
    "3_stars",
    "3_triangles",
    "5_boxes",
]
CONDITIONAL_FORMAT_ICON_SETS: tuple[str, ...] = get_args(ConditionalFormatIconSet)

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
ChartMarkerStyle = Literal[
    "automatic",
    "none",
    "square",
    "diamond",
    "triangle",
    "x",
    "star",
    "dot",
    "dash",
    "circle",
    "plus",
]
CHART_LEGEND_POSITIONS: tuple[str, ...] = get_args(ChartLegendPosition)
CHART_PLOT_BY: tuple[str, ...] = get_args(ChartPlotBy)
CHART_MARKER_STYLES: tuple[str, ...] = get_args(ChartMarkerStyle)

# The pivot table vocabulary. PivotFunction lists Excel's "Summarize Values By"
# options, named after the worksheet functions: note that "count" counts
# non-empty cells (COUNTA), while "count_numbers" is the worksheet COUNT.
# main.PivotTables/PivotValueFields validate them so engines only ever see the
# canonical names.
PivotFunction = Literal[
    "sum",
    "count",
    "average",
    "max",
    "min",
    "product",
    "count_numbers",
    "stdev",
    "stdevp",
    "var",
    "varp",
]
PivotLayout = Literal["compact", "outline", "tabular"]
PIVOT_FUNCTIONS: tuple[str, ...] = get_args(PivotFunction)
PIVOT_LAYOUTS: tuple[str, ...] = get_args(PivotLayout)
# The three field areas that PivotFields can stand for (the values area is a
# separate class); engines receive these names.
PIVOT_AREAS: tuple[str, ...] = ("rows", "columns", "filters")

# The alignment vocabulary. Excel, macOS and Office.js all support the same
# eight horizontal and five vertical values; main.Range validates them so
# engines only ever see the canonical names.
HorizontalAlignment = Literal[
    "general",
    "left",
    "center",
    "right",
    "fill",
    "justify",
    "center_across_selection",
    "distributed",
]
VerticalAlignment = Literal[
    "top",
    "center",
    "bottom",
    "justify",
    "distributed",
]
HORIZONTAL_ALIGNMENTS: tuple[str, ...] = get_args(HorizontalAlignment)
VERTICAL_ALIGNMENTS: tuple[str, ...] = get_args(VerticalAlignment)
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

    def move(self, before, after):
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
    @property
    def autofilter(self):
        raise NotImplementedError()

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

    async def get_horizontal_alignment(self):
        raise NotImplementedError(
            "get_horizontal_alignment() is only supported in xlwings Lite"
        )

    async def get_vertical_alignment(self):
        raise NotImplementedError(
            "get_vertical_alignment() is only supported in xlwings Lite"
        )

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

    async def get_data_validation(self):
        raise NotImplementedError(
            "get_data_validation() is only supported in xlwings Lite"
        )

    async def get_color(self):
        raise NotImplementedError("Range.get_color() is only supported in xlwings Lite")

    async def get_conditional_formats(self):
        raise NotImplementedError(
            "get_conditional_formats() is only supported in xlwings Lite"
        )

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
    def data_validation(self):
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
    def conditional_formats(self):
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
    def horizontal_alignment(self):
        raise NotImplementedError()

    @horizontal_alignment.setter
    def horizontal_alignment(self, value):
        raise NotImplementedError()

    @property
    def vertical_alignment(self):
        raise NotImplementedError()

    @vertical_alignment.setter
    def vertical_alignment(self, value):
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


class DataValidation:
    @property
    def api(self):
        raise NotImplementedError()

    def set_list(self, source, in_cell_dropdown):
        raise NotImplementedError()

    def set_rule(self, rule_type, operator, formula1, formula2):
        raise NotImplementedError()

    def delete(self):
        raise NotImplementedError()


class AutoFilter:
    @property
    def criteria(self):
        raise NotImplementedError()

    async def get_criteria(self):
        raise NotImplementedError("get_criteria() is only supported in xlwings Lite")

    def apply_values(self, field, values):
        raise NotImplementedError()

    def apply_comparison(self, field, operator, value1, value2):
        raise NotImplementedError()

    def apply_top_items(self, field, count):
        raise NotImplementedError()

    def apply_bottom_items(self, field, count):
        raise NotImplementedError()

    def apply_top_percent(self, field, percent):
        raise NotImplementedError()

    def apply_bottom_percent(self, field, percent):
        raise NotImplementedError()

    def clear(self, field):
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


class ConditionalFormat:
    @property
    def api(self):
        raise NotImplementedError()

    @property
    def type(self):
        raise NotImplementedError()

    @property
    def stop_if_true(self):
        raise NotImplementedError()

    @property
    def operator(self):
        raise NotImplementedError()

    @property
    def formula1(self):
        raise NotImplementedError()

    @property
    def formula2(self):
        raise NotImplementedError()

    @property
    def formula(self):
        raise NotImplementedError()

    @property
    def fill_color(self):
        raise NotImplementedError()

    @property
    def font_color(self):
        raise NotImplementedError()

    @property
    def font_bold(self):
        raise NotImplementedError()

    @property
    def font_italic(self):
        raise NotImplementedError()

    @property
    def colors(self):
        raise NotImplementedError()

    @property
    def bar_color(self):
        raise NotImplementedError()

    @property
    def gradient(self):
        raise NotImplementedError()

    @property
    def show_value(self):
        raise NotImplementedError()

    @property
    def icon_set(self):
        raise NotImplementedError()

    @property
    def reverse_order(self):
        raise NotImplementedError()

    @property
    def threshold_types(self):
        raise NotImplementedError()

    @property
    def thresholds(self):
        raise NotImplementedError()

    def set(self, changes):
        raise NotImplementedError()

    def delete(self):
        raise NotImplementedError()


class ConditionalFormats(Collection):
    def add_cell_value(self, spec):
        raise NotImplementedError()

    def add_custom(self, spec):
        raise NotImplementedError()

    def add_color_scale(self, spec):
        raise NotImplementedError()

    def add_data_bar(self, spec):
        raise NotImplementedError()

    def add_icon_set(self, spec):
        raise NotImplementedError()

    def clear(self):
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
    def autofilter(self):
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

    def set_x_axis_values(self, rng):
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
    def category_axis(self):
        raise NotImplementedError()

    @property
    def value_axis(self):
        raise NotImplementedError()

    @property
    def series(self):
        raise NotImplementedError()

    async def get_series(self):
        raise NotImplementedError("get_series() is only supported in xlwings Lite")

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


class ChartAxis:
    @property
    def api(self):
        raise NotImplementedError()

    @property
    def title(self):
        raise NotImplementedError()

    @title.setter
    def title(self, value):
        raise NotImplementedError()

    @property
    def minimum_scale(self):
        raise NotImplementedError()

    @minimum_scale.setter
    def minimum_scale(self, value):
        raise NotImplementedError()

    @property
    def maximum_scale(self):
        raise NotImplementedError()

    @maximum_scale.setter
    def maximum_scale(self, value):
        raise NotImplementedError()

    @property
    def major_unit(self):
        raise NotImplementedError()

    @major_unit.setter
    def major_unit(self, value):
        raise NotImplementedError()

    @property
    def number_format(self):
        raise NotImplementedError()

    @number_format.setter
    def number_format(self, value):
        raise NotImplementedError()

    @property
    def visible(self):
        raise NotImplementedError()

    @visible.setter
    def visible(self, value):
        raise NotImplementedError()

    def set(
        self,
        *,
        title=_UNSET,
        minimum_scale=_UNSET,
        maximum_scale=_UNSET,
        major_unit=_UNSET,
        number_format=_UNSET,
        visible=_UNSET,
    ):
        raise NotImplementedError()

    async def get_title(self):
        raise NotImplementedError("get_title() is only supported in xlwings Lite")

    async def get_minimum_scale(self):
        raise NotImplementedError(
            "get_minimum_scale() is only supported in xlwings Lite"
        )

    async def get_maximum_scale(self):
        raise NotImplementedError(
            "get_maximum_scale() is only supported in xlwings Lite"
        )

    async def get_major_unit(self):
        raise NotImplementedError("get_major_unit() is only supported in xlwings Lite")

    async def get_number_format(self):
        raise NotImplementedError(
            "get_number_format() is only supported in xlwings Lite"
        )

    async def get_visible(self):
        raise NotImplementedError("get_visible() is only supported in xlwings Lite")


class ChartSeries:
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
    def marker_style(self):
        raise NotImplementedError()

    @marker_style.setter
    def marker_style(self, value):
        raise NotImplementedError()

    @property
    def marker_size(self):
        raise NotImplementedError()

    @marker_size.setter
    def marker_size(self, value):
        raise NotImplementedError()

    @property
    def marker_foreground_color(self):
        raise NotImplementedError()

    @marker_foreground_color.setter
    def marker_foreground_color(self, value):
        raise NotImplementedError()

    @property
    def marker_background_color(self):
        raise NotImplementedError()

    @marker_background_color.setter
    def marker_background_color(self, value):
        raise NotImplementedError()

    @property
    def line_color(self):
        raise NotImplementedError()

    @line_color.setter
    def line_color(self, value):
        raise NotImplementedError()

    @property
    def fill_color(self):
        raise NotImplementedError()

    @fill_color.setter
    def fill_color(self, value):
        raise NotImplementedError()

    def set(
        self,
        *,
        name=_UNSET,
        marker_style=_UNSET,
        marker_size=_UNSET,
        marker_foreground_color=_UNSET,
        marker_background_color=_UNSET,
        line_color=_UNSET,
        fill_color=_UNSET,
    ):
        raise NotImplementedError()

    async def get_name(self):
        raise NotImplementedError("get_name() is only supported in xlwings Lite")

    async def get_marker_style(self):
        raise NotImplementedError(
            "get_marker_style() is only supported in xlwings Lite"
        )

    async def get_marker_size(self):
        raise NotImplementedError("get_marker_size() is only supported in xlwings Lite")

    async def get_marker_foreground_color(self):
        raise NotImplementedError(
            "get_marker_foreground_color() is only supported in xlwings Lite"
        )

    async def get_marker_background_color(self):
        raise NotImplementedError(
            "get_marker_background_color() is only supported in xlwings Lite"
        )

    async def get_line_color(self):
        raise NotImplementedError("get_line_color() is only supported in xlwings Lite")

    async def get_fill_color(self):
        raise NotImplementedError("get_fill_color() is only supported in xlwings Lite")


class ChartSeriesCollection(Collection):
    pass


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
        style=227,
    ):
        raise NotImplementedError()


class PivotTable:
    @property
    def api(self):
        raise NotImplementedError()

    @property
    def parent(self):
        raise NotImplementedError()

    @property
    def name(self):
        raise NotImplementedError()

    @name.setter
    def name(self, value):
        raise NotImplementedError()

    @property
    def field_names(self):
        raise NotImplementedError()

    @property
    def rows(self):
        raise NotImplementedError()

    @property
    def columns(self):
        raise NotImplementedError()

    @property
    def filters(self):
        raise NotImplementedError()

    @property
    def values(self):
        raise NotImplementedError()

    @property
    def layout(self):
        raise NotImplementedError()

    @layout.setter
    def layout(self, value):
        raise NotImplementedError()

    @property
    def show_row_grand_totals(self):
        raise NotImplementedError()

    @show_row_grand_totals.setter
    def show_row_grand_totals(self, value):
        raise NotImplementedError()

    @property
    def show_column_grand_totals(self):
        raise NotImplementedError()

    @show_column_grand_totals.setter
    def show_column_grand_totals(self, value):
        raise NotImplementedError()

    @property
    def range(self):
        raise NotImplementedError()

    @property
    def data_body_range(self):
        raise NotImplementedError()

    def refresh(self):
        raise NotImplementedError()

    def delete(self):
        raise NotImplementedError()


class PivotTables(Collection):
    def add(self, source, destination, name=None):
        raise NotImplementedError()


class PivotFields(Collection):
    """One of the rows/columns/filters areas of a pivot table."""

    @property
    def area(self):
        raise NotImplementedError()

    def add(self, name):
        raise NotImplementedError()


class PivotField:
    @property
    def api(self):
        raise NotImplementedError()

    @property
    def parent(self):
        raise NotImplementedError()

    @property
    def name(self):
        raise NotImplementedError()

    def remove(self):
        raise NotImplementedError()


class PivotValueFields(Collection):
    def add(self, field, function=None, name=None, number_format=None):
        raise NotImplementedError()


class PivotValueField:
    @property
    def api(self):
        raise NotImplementedError()

    @property
    def parent(self):
        raise NotImplementedError()

    @property
    def name(self):
        raise NotImplementedError()

    @name.setter
    def name(self, value):
        raise NotImplementedError()

    @property
    def source_field(self):
        raise NotImplementedError()

    @property
    def function(self):
        raise NotImplementedError()

    @function.setter
    def function(self, value):
        raise NotImplementedError()

    @property
    def number_format(self):
        raise NotImplementedError()

    @number_format.setter
    def number_format(self, value):
        raise NotImplementedError()

    def remove(self):
        raise NotImplementedError()


class FreezePanes:
    def freeze_at(self, frozen_range):
        raise NotImplementedError()

    def unfreeze():
        raise NotImplementedError()
