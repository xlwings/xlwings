"""
Required Notice: Copyright (C) Zoomer Analytics GmbH.

xlwings PRO is dual-licensed under one of the following licenses:

* PolyForm Noncommercial License 1.0.0 (for noncommercial use):
  https://polyformproject.org/licenses/noncommercial/1.0.0
* xlwings PRO License (for commercial use):
  https://github.com/xlwings/xlwings/blob/main/LICENSE_PRO.txt

Commercial licenses can be purchased at https://www.xlwings.org
"""

import asyncio
import base64
import copy
import datetime as dt
import numbers
import re
import sys
from functools import lru_cache

try:
    import numpy as np
except ImportError:
    np = None
try:
    import pandas as pd
except ImportError:
    pd = None

from .. import (
    NoSuchObjectError,
    ShapeAlreadyExists,
    XlwingsError,
    __version__,
    base_classes,
    utils,
)
from ..constants import MAX_COLUMNS, MAX_ROWS

# Private marker set on a sheet's api dict once its cell values have been loaded
# on an async (lazy) book. Stored on the dict (not keyed by sheet name) so it
# survives renames; the sheet's api dict is preserved across metadata reloads by
# _update_api_in_place. Only read locally (never serialized back to JS).
_SHEET_VALUES_LOADED_KEY = "_xlwings_values_loaded"

# xlwings' public calculation modes mapped to Office.js' Excel.CalculationMode.
# xlwings' "semiautomatic" is Office.js' "AutomaticExceptTables".
_CALCULATION_PY2JS = {
    "automatic": "Automatic",
    "manual": "Manual",
    "semiautomatic": "AutomaticExceptTables",
}
_CALCULATION_JS2PY = {v: k for k, v in _CALCULATION_PY2JS.items()}


def _color_to_hex(color_or_rgb):
    """Normalize xlwings' accepted color forms to the `#RRGGBB` Office.js wants.

    The public API takes an RGB tuple, a hex string or an integer, matching the
    desktop engines. `None` passes through, since it means "no fill" rather
    than a color.
    """
    if color_or_rgb is None:
        return None
    if isinstance(color_or_rgb, str):
        # Already hex; normalize a missing "#" so the client always gets the
        # same shape.
        return color_or_rgb if color_or_rgb.startswith("#") else f"#{color_or_rgb}"
    if isinstance(color_or_rgb, int) and not isinstance(color_or_rgb, bool):
        return utils.rgb_to_hex(*utils.int_to_rgb(color_or_rgb))
    try:
        red, green, blue = color_or_rgb
    except (TypeError, ValueError):
        raise ValueError(
            "Color must be an RGB tuple like (255, 0, 0), a hex string like "
            f'"#FFA500", or an integer --- got {color_or_rgb!r}.'
        ) from None
    return utils.rgb_to_hex(red, green, blue)


def _address_key(address):
    """Return a comparison key for equivalent absolute/relative A1 addresses."""
    return address.replace("$", "").upper() if address is not None else None


# xlwings' chart types mapped to Office.js' Excel.ChartType. These are the
# enum's *values* ("Line"), which is what Office.js sends and accepts --
# not its member names ("line"). All 73 xlwings types have an equivalent;
# the names differ only in casing and a few spellings.
_CHART_TYPE_PY2JS = {
    "3d_area": "3DArea",
    "3d_area_stacked": "3DAreaStacked",
    "3d_area_stacked_100": "3DAreaStacked100",
    "3d_bar_clustered": "3DBarClustered",
    "3d_bar_stacked": "3DBarStacked",
    "3d_bar_stacked_100": "3DBarStacked100",
    "3d_column": "3DColumn",
    "3d_column_clustered": "3DColumnClustered",
    "3d_column_stacked": "3DColumnStacked",
    "3d_column_stacked_100": "3DColumnStacked100",
    "3d_line": "3DLine",
    "3d_pie": "3DPie",
    "3d_pie_exploded": "3DPieExploded",
    "area": "Area",
    "area_stacked": "AreaStacked",
    "area_stacked_100": "AreaStacked100",
    "bar_clustered": "BarClustered",
    "bar_of_pie": "BarOfPie",
    "bar_stacked": "BarStacked",
    "bar_stacked_100": "BarStacked100",
    "bubble": "Bubble",
    "bubble_3d_effect": "Bubble3DEffect",
    "column_clustered": "ColumnClustered",
    "column_stacked": "ColumnStacked",
    "column_stacked_100": "ColumnStacked100",
    "cone_bar_clustered": "ConeBarClustered",
    "cone_bar_stacked": "ConeBarStacked",
    "cone_bar_stacked_100": "ConeBarStacked100",
    "cone_col": "ConeCol",
    "cone_col_clustered": "ConeColClustered",
    "cone_col_stacked": "ConeColStacked",
    "cone_col_stacked_100": "ConeColStacked100",
    "cylinder_bar_clustered": "CylinderBarClustered",
    "cylinder_bar_stacked": "CylinderBarStacked",
    "cylinder_bar_stacked_100": "CylinderBarStacked100",
    "cylinder_col": "CylinderCol",
    "cylinder_col_clustered": "CylinderColClustered",
    "cylinder_col_stacked": "CylinderColStacked",
    "cylinder_col_stacked_100": "CylinderColStacked100",
    "doughnut": "Doughnut",
    "doughnut_exploded": "DoughnutExploded",
    "line": "Line",
    "line_markers": "LineMarkers",
    "line_markers_stacked": "LineMarkersStacked",
    "line_markers_stacked_100": "LineMarkersStacked100",
    "line_stacked": "LineStacked",
    "line_stacked_100": "LineStacked100",
    "pie": "Pie",
    "pie_exploded": "PieExploded",
    "pie_of_pie": "PieOfPie",
    "pyramid_bar_clustered": "PyramidBarClustered",
    "pyramid_bar_stacked": "PyramidBarStacked",
    "pyramid_bar_stacked_100": "PyramidBarStacked100",
    "pyramid_col": "PyramidCol",
    "pyramid_col_clustered": "PyramidColClustered",
    "pyramid_col_stacked": "PyramidColStacked",
    "pyramid_col_stacked_100": "PyramidColStacked100",
    "radar": "Radar",
    "radar_filled": "RadarFilled",
    "radar_markers": "RadarMarkers",
    "stock_hlc": "StockHLC",
    "stock_ohlc": "StockOHLC",
    "stock_vhlc": "StockVHLC",
    "stock_vohlc": "StockVOHLC",
    "surface": "Surface",
    "surface_top_view": "SurfaceTopView",
    "surface_top_view_wireframe": "SurfaceTopViewWireframe",
    "surface_wireframe": "SurfaceWireframe",
    "xy_scatter": "XYScatter",
    "xy_scatter_lines": "XYScatterLines",
    "xy_scatter_lines_no_markers": "XYScatterLinesNoMarkers",
    "xy_scatter_smooth": "XYScatterSmooth",
    "xy_scatter_smooth_no_markers": "XYScatterSmoothNoMarkers",
}
_CHART_TYPE_JS2PY = {v: k for k, v in _CHART_TYPE_PY2JS.items()}

# Office.js' Excel.ShapeType mapped onto xlwings' shape type names, so
# Shape.type speaks the same vocabulary as the desktop engines. Office.js'
# set is much coarser (5 vs 32), but each of its members has an xlwings
# equivalent except "Unsupported", which passes through as-is.
_SHAPE_TYPE_JS2PY = {
    "Image": "picture",
    "GeometricShape": "auto_shape",
    "Group": "group",
    "Line": "line",
}

# xlwings' scale anchors mapped to Office.js' Excel.ShapeScaleFrom.
_SHAPE_SCALE_FROM = {
    "scale_from_top_left": "ScaleFromTopLeft",
    "scale_from_middle": "ScaleFromMiddle",
    "scale_from_bottom_right": "ScaleFromBottomRight",
}

# Chart vocabulary mapped to Office.js' Excel.ChartLegendPosition and
# Excel.ChartPlotBy / Excel.ChartSeriesBy.
_LEGEND_POSITION_PY2JS = {
    "top": "Top",
    "bottom": "Bottom",
    "left": "Left",
    "right": "Right",
    "corner": "Corner",
}
_PLOT_BY_PY2JS = {"rows": "Rows", "columns": "Columns"}
_MARKER_STYLE_PY2JS = {
    "automatic": "Automatic",
    "none": "None",
    "square": "Square",
    "diamond": "Diamond",
    "triangle": "Triangle",
    "x": "X",
    "star": "Star",
    "dot": "Dot",
    "dash": "Dash",
    "circle": "Circle",
    "plus": "Plus",
}
_MARKER_STYLE_JS2PY = {v: k for k, v in _MARKER_STYLE_PY2JS.items()}

# xlwings' pivot vocabulary mapped to Office.js' Excel.AggregationFunction and
# Excel.PivotLayoutType values. main.py validates, so plain lookups suffice.
_PIVOT_FUNCTION_PY2JS = {
    "sum": "Sum",
    "count": "Count",
    "average": "Average",
    "max": "Max",
    "min": "Min",
    "product": "Product",
    "count_numbers": "CountNumbers",
    "stdev": "StandardDeviation",
    "stdevp": "StandardDeviationP",
    "var": "Variance",
    "varp": "VarianceP",
}
_PIVOT_FUNCTION_JS2PY = {v: k for k, v in _PIVOT_FUNCTION_PY2JS.items()}
_PIVOT_LAYOUT_PY2JS = {"compact": "Compact", "outline": "Outline", "tabular": "Tabular"}
_PIVOT_LAYOUT_JS2PY = {v: k for k, v in _PIVOT_LAYOUT_PY2JS.items()}
_PIVOT_AREAS = ("rows", "columns", "filters")

# Office.js Excel.HorizontalAlignment / Excel.VerticalAlignment member names.
_HORIZONTAL_ALIGNMENT_PY2JS = {
    "general": "General",
    "left": "Left",
    "center": "Center",
    "right": "Right",
    "fill": "Fill",
    "justify": "Justify",
    "center_across_selection": "CenterAcrossSelection",
    "distributed": "Distributed",
}
_HORIZONTAL_ALIGNMENT_JS2PY = {v: k for k, v in _HORIZONTAL_ALIGNMENT_PY2JS.items()}
_VERTICAL_ALIGNMENT_PY2JS = {
    "top": "Top",
    "center": "Center",
    "bottom": "Bottom",
    "justify": "Justify",
    "distributed": "Distributed",
}
_VERTICAL_ALIGNMENT_JS2PY = {v: k for k, v in _VERTICAL_ALIGNMENT_PY2JS.items()}

_CONDITIONAL_FORMAT_TYPE_JS2PY = {
    "CellValue": "cell_value",
    "Custom": "custom",
    "ColorScale": "color_scale",
    "DataBar": "data_bar",
    "IconSet": "icon_set",
}
_CONDITIONAL_FORMAT_OPERATOR_PY2JS = {
    "between": "Between",
    "not_between": "NotBetween",
    "equal_to": "EqualTo",
    "not_equal_to": "NotEqualTo",
    "greater_than": "GreaterThan",
    "less_than": "LessThan",
    "greater_than_or_equal": "GreaterThanOrEqual",
    "less_than_or_equal": "LessThanOrEqual",
}
_CONDITIONAL_FORMAT_OPERATOR_JS2PY = {
    value: key for key, value in _CONDITIONAL_FORMAT_OPERATOR_PY2JS.items()
}
_CONDITIONAL_FORMAT_THRESHOLD_PY2JS = {
    "automatic": "Automatic",
    "lowest_value": "LowestValue",
    "highest_value": "HighestValue",
    "number": "Number",
    "percent": "Percent",
    "percentile": "Percentile",
    "formula": "Formula",
}
_CONDITIONAL_FORMAT_THRESHOLD_JS2PY = {
    value: key for key, value in _CONDITIONAL_FORMAT_THRESHOLD_PY2JS.items()
}
_CONDITIONAL_FORMAT_ICON_SET_PY2JS = {
    "3_arrows": "ThreeArrows",
    "3_arrows_gray": "ThreeArrowsGray",
    "3_flags": "ThreeFlags",
    "3_traffic_lights_1": "ThreeTrafficLights1",
    "3_traffic_lights_2": "ThreeTrafficLights2",
    "3_signs": "ThreeSigns",
    "3_symbols": "ThreeSymbols",
    "3_symbols_2": "ThreeSymbols2",
    "4_arrows": "FourArrows",
    "4_arrows_gray": "FourArrowsGray",
    "4_red_to_black": "FourRedToBlack",
    "4_rating": "FourRating",
    "4_traffic_lights": "FourTrafficLights",
    "5_arrows": "FiveArrows",
    "5_arrows_gray": "FiveArrowsGray",
    "5_rating": "FiveRating",
    "5_quarters": "FiveQuarters",
    "3_stars": "ThreeStars",
    "3_triangles": "ThreeTriangles",
    "5_boxes": "FiveBoxes",
}
_CONDITIONAL_FORMAT_ICON_SET_JS2PY = {
    value: key for key, value in _CONDITIONAL_FORMAT_ICON_SET_PY2JS.items()
}


def _mark_sheet_values_loaded(sheet_api):
    sheet_api[_SHEET_VALUES_LOADED_KEY] = True


def _sheet_values_loaded(sheet_api):
    return sheet_api.get(_SHEET_VALUES_LOADED_KEY, False)


def _normalize_jsnull(obj, *, normalize_values=False):
    """Recursively replace Pyodide's `JsNull` sentinel with Python `None`.

    Pyodide >= 0.28 converts JS `null` to `pyodide.ffi.jsnull` instead of
    `None`. Book data coming back from Office.js via `.to_py()` therefore
    contains `JsNull` for empty/absent fields (e.g. `scope_sheet_index` on
    book-scoped names, `address` for non-cell selections), which breaks the
    `is None` checks throughout this module. Normalize everything back to
    `None` at the JS boundary so downstream code stays Pyodide-version
    agnostic.

    The `values` arrays (cell data) are skipped by default for two reasons. First,
    *book* data (`getBookData()`) represents empty cells as `""`, not
    `null`, so a book's `values` cannot contain `JsNull`. Second, walking
    them would mean an extra full pass over every cell of an eagerly-loaded book
    (e.g. `xw.Book(json=...)` in xlwings Lite's notebook runner).

    Callers whose `values` fields are metadata rather than cell matrices can set
    `normalize_values=True`.

    Note this is *not* true for custom function *arguments*: Excel's custom
    functions runtime sends empty cells in a range argument as JS `null` ->
    `JsNull`. Those don't pass through here — UDFs use a separate engine, so
    they're normalized in `_xlofficejs._clean_value_data_element` instead.
    """
    try:
        from pyodide.ffi import JsNull
    except ImportError:
        return obj

    def walk(o):
        if isinstance(o, JsNull):
            return None
        if isinstance(o, dict):
            return {
                k: v if k == "values" and not normalize_values else walk(v)
                for k, v in o.items()
            }
        if isinstance(o, list):
            return [walk(v) for v in o]
        return o

    return walk(obj)


# Time types (doesn't contain dt.date)
time_types = (dt.datetime,)
if np:
    time_types = time_types + (np.datetime64,)
if pd:
    time_types = time_types + (pd.Timestamp,)

datetime_pattern = r"^(-?(?:[1-9][0-9]*)?[0-9]{4})-(1[0-2]|0[1-9])-(3[01]|0[1-9]|[12][0-9])T(2[0-3]|[01][0-9]):([0-5][0-9]):([0-5][0-9])(\.[0-9]+)?(Z|[+-](?:2[0-3]|[01][0-9]):[0-5][0-9])?$"  # noqa: E501
datetime_regex = re.compile(datetime_pattern)


def _update_pivots_in_place(target, source):
    """Keep pivot and value wrappers alive across a metadata reload.

    Loaded objects match by Office.js ID, including after an external rename.
    A pivot created in this script has no ID yet, but its name is known. New
    value fields have neither an ID nor necessarily a caption; their source
    and position identify them after the queued actions have been flushed.
    """
    previous = target.get("pivot_tables")
    incoming = source.get("pivot_tables")
    if not isinstance(previous, list) or not isinstance(incoming, list):
        return

    def merge(old, new, *, values=False):
        by_id = {item["id"]: item for item in old if "id" in item}
        by_name = {item["name"]: item for item in old if "name" in item}
        matched = set()
        result = []
        for index, item in enumerate(new):
            item = dict(item)
            entry = by_id.get(item.get("id"))
            if entry is None:
                candidate = by_name.get(item.get("name"))
                if candidate is not None and (
                    "id" not in candidate or "id" not in item
                ):
                    entry = candidate
            if entry is None and values and index < len(old):
                candidate = old[index]
                if (
                    "id" not in candidate
                    and "name" not in candidate
                    and candidate.get("source_field") == item.get("source_field")
                ):
                    entry = candidate
            if entry is None or id(entry) in matched:
                result.append(item)
                continue
            matched.add(id(entry))
            if not values:
                item = {
                    **item,
                    "values": merge(entry["values"], item["values"], values=True),
                }
            # Drop private assumptions (such as _name_explicit) when reading
            # authoritative metadata from Excel.
            entry.clear()
            entry.update(item)
            result.append(entry)
        old[:] = result
        return old

    source["pivot_tables"] = merge(previous, incoming)


def _update_api_in_place(target, source):
    """Update target dict in-place from source, preserving references to nested dicts
    inside lists (matched by 'name' key). This ensures that e.g. Sheet or Name objects
    holding references to dicts inside target['sheets'] see the updated values."""
    for key, value in source.items():
        if isinstance(value, list) and key in target and isinstance(target[key], list):
            old_by_name = {
                item["name"]: item
                for item in target[key]
                if isinstance(item, dict) and "name" in item
            }
            new_list = []
            for item in value:
                if isinstance(item, dict) and item.get("name") in old_by_name:
                    entry = old_by_name[item["name"]]
                    if key == "sheets":
                        item = dict(item)
                        _update_pivots_in_place(entry, item)
                    entry.update(item)
                    new_list.append(entry)
                else:
                    new_list.append(item)
            target[key] = new_list
        elif (
            isinstance(value, dict) and key in target and isinstance(target[key], dict)
        ):
            target[key].update(value)
        else:
            target[key] = value


def _clean_value_data_element(
    value, datetime_builder, empty_as, number_builder, err_to_str
):
    if value == "":
        return empty_as
    if isinstance(value, str):
        # TODO: Send arrays back and forth with indices of the location of dt values
        if datetime_regex.match(value):
            value = dt.datetime.fromisoformat(
                value[:-1]
            )  # cut off "Z" (Python doesn't accept it and Excel doesn't support tz)
        elif not err_to_str and value in [
            "#DIV/0!",
            "#N/A",
            "#NAME?",
            "#NULL!",
            "#NUM!",
            "#REF!",
            "#VALUE!",
            "#DATA!",
        ]:
            value = None
        else:
            value = value
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
    def __init__(self):
        self.apps = Apps()

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

    @staticmethod
    def prepare_xl_data_element(x, options):
        if x is None:
            return ""
        elif pd and pd.isna(x):
            return ""
        elif np and isinstance(x, (np.floating, float)) and np.isnan(x):
            return ""
        elif np and isinstance(x, np.number):
            return float(x)
        elif pd and isinstance(x, type(pd.NaT)):
            return None
        elif isinstance(x, time_types):
            if np and isinstance(x, np.datetime64):
                x = utils.np_datetime_to_datetime(x)
            elif pd and isinstance(x, pd.Timestamp):
                x = x.to_pydatetime()
            if x.time() == dt.time(0, 0):
                x = x.date().isoformat()
            else:
                x = x.replace(tzinfo=None).isoformat(sep=" ").split(".")[0]
        elif isinstance(x, dt.date):
            x = x.isoformat()
        return x

    @property
    def name(self):
        return "remote"

    @property
    def type(self):
        return "remote"


class Apps(base_classes.Apps):
    def __init__(self):
        self._apps = [App(self)]

    def __iter__(self):
        return iter(self._apps)

    def __len__(self):
        return len(self._apps)

    def __getitem__(self, index):
        return self._apps[index]

    def add(self, **kwargs):
        self._apps.insert(0, App(self, **kwargs))
        return self._apps[0]


class App(base_classes.App):
    _next_pid = -1

    def __init__(self, apps, add_book=True, **kwargs):
        self.apps = apps
        self._pid = App._next_pid
        App._next_pid -= 1
        self._display_alerts = True
        self._books = Books(self)
        if add_book:
            self._books.add()

    def kill(self):
        self.apps._apps.remove(self)
        self.apps = None

    @property
    def engine(self):
        return engine

    @property
    def books(self):
        return self._books

    @property
    def pid(self):
        return self._pid

    @property
    def display_alerts(self):
        # Office.js never shows the alerts that this suppresses on the desktop
        # engines, but Range.merge() toggles it, so it has to be readable and
        # writable rather than raising.
        return self._display_alerts

    @display_alerts.setter
    def display_alerts(self, value):
        self._display_alerts = value

    def _unsupported(name, detail, read_only=False):
        # Office.js' Excel.Application doesn't expose these, so they raise with
        # the reason rather than a bare NotImplementedError. read_only mirrors
        # the public API: path, startup_path and version have no setter there,
        # so defining one here would turn AttributeError into the wrong error.
        message = f"App.{name} is not supported on this engine: {detail}"

        def getter(self):
            raise NotImplementedError(message)

        if read_only:
            return property(getter)

        def setter(self, value):
            raise NotImplementedError(message)

        return property(getter, setter)

    cut_copy_mode = _unsupported("cut_copy_mode", "it has no clipboard access.")
    enable_events = _unsupported("enable_events", "it has no equivalent setting.")
    interactive = _unsupported("interactive", "it has no equivalent setting.")
    status_bar = _unsupported("status_bar", "it has no equivalent setting.")
    path = _unsupported(
        "path",
        "an add-in has no access to the Excel installation's paths.",
        read_only=True,
    )
    startup_path = _unsupported(
        "startup_path",
        "an add-in has no access to the Excel installation's paths.",
        read_only=True,
    )
    version = _unsupported(
        "version",
        "the Excel application version isn't available to an add-in.",
        read_only=True,
    )
    del _unsupported

    def quit(self):
        raise NotImplementedError(
            "App.quit() is not supported on this engine: an add-in can't close "
            "the Excel application."
        )

    @property
    def calculation(self):
        calculation = self.books.active.api["book"].get("calculation")
        if calculation is None:
            raise NotImplementedError(
                "This client doesn't send the calculation mode. It requires a "
                "newer version of the xlwings JavaScript module."
            )
        return _CALCULATION_JS2PY[calculation]

    @calculation.setter
    def calculation(self, value):
        try:
            js_value = _CALCULATION_PY2JS[value]
        except KeyError:
            raise ValueError(
                f"Invalid calculation mode: {value!r}. Must be one of "
                f"{sorted(_CALCULATION_PY2JS)}."
            ) from None
        self.books.active.api["book"]["calculation"] = js_value
        self.books.active.append_json_action(func="setCalculation", args=[js_value])

    @property
    def screen_updating(self):
        # Office.js has no screen updating flag to read back, only a
        # suspend-until-next-sync call. See the setter.
        raise NotImplementedError(
            "App.screen_updating can't be read on this engine, which can only "
            "suspend screen updating until its next sync rather than reporting a "
            "setting."
        )

    @screen_updating.setter
    def screen_updating(self, value):
        # Office.js can only suspend screen updating until its next sync, not
        # turn it off indefinitely. Setting it back to True is therefore a
        # no-op: the suspension ends on its own at the next sync.
        self.books.active.append_json_action(
            func="setScreenUpdating", args=[bool(value)]
        )

    def calculate(self):
        # args=[] rather than omitting it: append_json_action wraps a missing
        # args as [None], and this action takes no arguments.
        self.books.active.append_json_action(func="calculate", args=[])

    @property
    def selection(self):
        book = self.books.active
        return Range(sheet=book.sheets.active, arg1=book.api["book"]["selection"])

    async def get_selection(self):
        if sys.platform != "emscripten":
            raise NotImplementedError(
                "App.get_selection() is only supported in xlwings Lite"
            )
        import js

        result = _normalize_jsnull((await js.xlwings.getSelection()).to_py())
        sheet_index = int(result["sheetIndex"])
        address = result["address"]
        book = self.books.active
        sheet = Sheet(
            api=book.api["sheets"][sheet_index],
            sheets=book.sheets,
            index=sheet_index + 1,
        )
        if address is None:
            return None  # Non-cell selection (e.g., shape)
        return Range(sheet=sheet, arg1=address)

    @property
    def visible(self):
        return True

    @visible.setter
    def visible(self, value):
        pass

    def activate(self, steal_focus=None):
        pass

    def alert(self, prompt, title, buttons, mode, callback):
        self.books.active.append_json_action(
            func="alert",
            args=[
                "" if prompt is None else prompt,
                "" if title is None else title,
                "" if buttons is None else buttons,
                "" if mode is None else mode,
                "" if callback is None else callback,
            ],
        )

    def run(self, macro, args):
        self.books.active.append_json_action(
            func="runMacro",
            args=[macro] + [args] if not isinstance(args, list) else [macro] + args,
        )


class Books(base_classes.Books):
    def __init__(self, app):
        self.app = app
        self.books = []
        self._active = None

    @property
    def active(self):
        return self._active

    async def get_active(self):
        if sys.platform != "emscripten":
            raise NotImplementedError(
                "Books.get_active() is only supported in xlwings Lite"
            )
        import js
        from pyodide.ffi import to_js

        book_data_js = await js.xlwings.getBookData(
            to_js({"lazy": True}, dict_converter=js.Object.fromEntries)
        )
        # open() normalizes JsNull -> None
        book_data = book_data_js.to_py()
        # The book was fetched lazily (structure only, no cell values), so mark
        # it: sync `.value` reads raise until values are loaded, and `load()`
        # defaults to metadata-only. See Book.load / Range.api.
        return self.open(book_data, lazy=True)

    def open(self, json, lazy=False):
        # Normalize here (rather than only at the getBookData boundary) so that
        # callers passing raw `.to_py()` data straight to `xw.Book(json=...)`
        # (e.g. xlwings Lite's notebook runner) also get JsNull -> None.
        book = Book(api=_normalize_jsnull(json), books=self, lazy=lazy)
        self.books.append(book)
        self._active = book
        return book

    def add(self):
        book = Book(
            api={
                "version": __version__,
                "book": {"name": f"Book{len(self) + 1}", "active_sheet_index": 0},
                "sheets": [
                    {
                        "name": "Sheet1",
                        "values": [[]],
                    },
                ],
            },
            books=self,
        )
        self.books.append(book)
        self._active = book
        return book

    def _try_find_book_by_name(self, name):
        for book in self.books:
            if book.name == name or book.fullname == name:
                return book
        return None

    def __len__(self):
        return len(self.books)

    def __iter__(self):
        for book in self.books:
            yield book

    def __call__(self, name_or_index):
        if isinstance(name_or_index, numbers.Number):
            return self.books[name_or_index - 1]
        else:
            book = self._try_find_book_by_name(name_or_index)
            if book is None:
                raise KeyError(name_or_index)
            return book


class Book(base_classes.Book):
    def __init__(self, api, books, lazy=False):
        self.books = books
        self._api = api
        self._json = {"actions": []}
        # Async/lazy book: fetched with structure only (no cell values). While
        # lazy, sync `.value` reads on a sheet whose values haven't been loaded
        # raise (use `await get_value()` or `await load(values=True)`), and
        # `load()` defaults to metadata-only. Whether a sheet's values are loaded
        # is marked per sheet on its api dict (see `_SHEET_VALUES_LOADED_KEY`), so
        # it survives sheet renames.
        self._lazy = lazy
        if api["version"] != __version__ and api["client"] != "Office.js":
            raise XlwingsError(
                f"Your xlwings version is different on the client ({api['version']}) "
                f"and server ({__version__})."
            )

    def append_json_action(self, **kwargs):
        args = kwargs.get("args")
        self._json["actions"].append(
            {
                "func": kwargs.get("func"),
                "args": [args] if not isinstance(args, list) else args,
                "values": kwargs.get("values"),
                "sheet_position": kwargs.get("sheet_position"),
                "start_row": kwargs.get("start_row"),
                "start_column": kwargs.get("start_column"),
                "row_count": kwargs.get("row_count"),
                "column_count": kwargs.get("column_count"),
            }
        )

    @property
    def api(self):
        return self._api

    def json(self):
        return self._json

    async def flush(self):
        if sys.platform != "emscripten":
            raise NotImplementedError("Book.flush() is only supported in xlwings Lite")
        import js
        from pyodide.ffi import to_js

        actions = self._json.get("actions", [])
        if actions:
            actions_js = to_js(
                {"actions": actions}, dict_converter=js.Object.fromEntries
            )
            # runActions may fail after applying only part of the batch. Never retain
            # that indeterminate batch: replaying it can duplicate non-idempotent
            # actions such as adding a defined name.
            self._json["actions"] = []
            await js.xlwings.runActions(actions_js)
        # Yield to the browser event loop so it can repaint (to print to output pane)
        await asyncio.sleep(0.01)

    async def load(self, values=None):
        """(Re)load the book's data from Excel on demand.

        On an async (lazy) book, this defaults to loading only *metadata* -
        sheet structure, tables, pictures, names - and not cell values, since
        bulk-loading values would defeat the point of the async API. Pass
        `values=True` to also snapshot all cell values (after which sync
        `.value` reads work again).

        On a regular (eager) book, everything including values is loaded, as
        before.
        """
        if sys.platform != "emscripten":
            raise NotImplementedError("Book.load() is only supported in xlwings Lite")
        import js
        from pyodide.ffi import to_js

        # Default: metadata-only for lazy books, full for eager books.
        load_values = (not self._lazy) if values is None else bool(values)
        # getBookData(lazy=True) returns structure only (empty values); a plain
        # call returns everything.
        opts = to_js({"lazy": not load_values}, dict_converter=js.Object.fromEntries)
        data = _normalize_jsnull((await js.xlwings.getBookData(opts)).to_py())
        if not load_values:
            # Metadata-only: don't let the empty `values` payload clobber any
            # values already loaded on this book.
            for sheet in data.get("sheets", []):
                sheet.pop("values", None)
        _update_api_in_place(self._api, data)
        if load_values:
            for sheet in self._api.get("sheets", []):
                _mark_sheet_values_loaded(sheet)
        get_range_api.cache_clear()

    @property
    def name(self):
        return self.api["book"]["name"]

    @property
    def fullname(self):
        return self.name

    @property
    def names(self):
        return Names(parent=self, api=self.api["names"])

    @property
    def sheets(self):
        return Sheets(api=self.api["sheets"], book=self)

    async def get_comments(self):
        if sys.platform != "emscripten":
            raise NotImplementedError(
                "get_comments() is only supported in xlwings Lite"
            )
        import js

        entries = _normalize_jsnull((await js.xlwings.getComments()).to_py())
        return [
            Comment(
                Range(sheet=self.sheets(entry["sheet"]), arg1=entry["address"]),
                entry["id"],
            )
            for entry in entries
        ]

    @property
    def app(self):
        return self.books.app

    def save(self, path=None, password=None):
        if path is not None:
            raise NotImplementedError(
                "Book.save() can't take a path on this engine, which can only "
                "save the book in place. Call save() without arguments instead."
            )
        if password is not None:
            raise NotImplementedError(
                "Book.save() can't take a password on this engine, which has no "
                "way to set one."
            )
        self.append_json_action(func="save", args=[])

    def to_pdf(self, path, quality):
        raise NotImplementedError(
            "Book.to_pdf() is not supported on this engine, which has no PDF export."
        )

    def close(self):
        assert self.api is not None, "Seems this book was already closed."
        self.books.books.remove(self)
        self.books = None
        self._api = None

    def activate(self):
        pass


class Sheets(base_classes.Sheets):
    def __init__(self, api, book):
        self._api = api
        self.book = book

    @property
    def active(self):
        ix = self.book.api["book"]["active_sheet_index"]
        return Sheet(api=self.api[ix], sheets=self, index=ix + 1)

    async def get_active(self):
        if sys.platform != "emscripten":
            raise NotImplementedError(
                "Sheets.get_active() is only supported in xlwings Lite"
            )
        import js

        ix = int(await js.xlwings.getActiveSheetIndex())
        return Sheet(api=self.api[ix], sheets=self, index=ix + 1)

    @property
    def api(self):
        return self._api

    def __call__(self, name_or_index):
        if isinstance(name_or_index, int):
            return Sheet(
                api=self.api[name_or_index - 1], sheets=self, index=name_or_index
            )
        else:
            for ix, sheet in enumerate(self.api):
                if (
                    isinstance(name_or_index, str)
                    and sheet["name"].casefold() == name_or_index.casefold()
                ):
                    return Sheet(api=sheet, sheets=self, index=ix + 1)
        raise ValueError(f"Sheet '{name_or_index}' doesn't exist!")

    def add(self, before=None, after=None, name=None):
        # TODO: this is hardcoded to English
        if name is None:
            sheet_number = 1
            while True:
                if f"Sheet{sheet_number}" in [sheet.name for sheet in self]:
                    sheet_number += 1
                else:
                    break
            name = f"Sheet{sheet_number}"

        api = {
            "name": name,
            "visibility": "Visible",
            "values": [[]],
            "pictures": [],
            "shapes": [],
            "charts": [],
            "notes": [],
            "tables": [],
            "print_area": None,
        }
        if self.book.api["client"] == "Office.js":
            api["show_gridlines"] = True
            api["pivot_tables"] = []

        if before:
            if before.index == 1:
                ix = 1
            else:
                ix = before.index - 1
        elif after:
            ix = after.index + 1
        else:
            # Default position is different from Desktop apps!
            ix = len(self) + 1

        self.api.insert(ix - 1, api)
        self.book.append_json_action(func="addSheet", args=[ix - 1, name])
        self.book.api["book"]["active_sheet_index"] = ix - 1

        return Sheet(api=api, sheets=self, index=ix)

    def __len__(self):
        return len(self.api)

    def __iter__(self):
        for ix, sheet in enumerate(self.api):
            yield Sheet(api=sheet, sheets=self, index=ix + 1)


class Sheet(base_classes.Sheet):
    def __init__(self, api, sheets, index):
        self._api = api
        self._index = index
        self.sheets = sheets

    def append_json_action(self, **kwargs):
        self.book.append_json_action(
            **{
                **kwargs,
                **{
                    "sheet_position": self.index - 1,
                },
            }
        )

    @property
    def api(self):
        return self._api

    @property
    def name(self):
        return self.api["name"]

    @name.setter
    def name(self, value):
        old_name = self.name
        self.append_json_action(
            func="setSheetName",
            args=value,
        )
        self.api["name"] = value
        # Excel rewrites these definitions when it applies the rename. Mirror
        # that in the snapshot so later reads/deletes in this script agree.
        for name in self.book.api["names"]:
            if name.get("refers_to") is not None:
                name["refers_to"] = _rename_sheet_in_formula(
                    name["refers_to"], old_name, value
                )
            if not name["book_scope"] and name["scope_sheet_index"] == self.index - 1:
                name["scope_sheet_name"] = value
                if "!" in name["name"]:
                    name[
                        "name"
                    ] = f"{_sheet_name_prefix(value)}!{name['name'].rsplit('!', 1)[-1]}"

    @property
    def visible(self):
        # Office.js also knows "VeryHidden", which maps to False like "Hidden"
        # does: xlwings' public API is a bool, so the distinction is dropped.
        return self.api["visibility"] == "Visible"

    @visible.setter
    def visible(self, value):
        visibility = "Visible" if value else "Hidden"
        self.append_json_action(
            func="setSheetVisibility",
            args=visibility,
        )
        self.api["visibility"] = visibility

    @property
    def show_gridlines(self):
        try:
            return self.api["show_gridlines"]
        except KeyError:
            # Only the Office.js client sends this key.
            raise NotImplementedError(
                "Sheet.show_gridlines is only supported with Office.js clients"
            ) from None

    @show_gridlines.setter
    def show_gridlines(self, value):
        if self.book.api["client"] != "Office.js":
            raise NotImplementedError(
                "Sheet.show_gridlines is only supported with Office.js clients"
            )
        value = bool(value)
        self.append_json_action(func="setShowGridlines", args=[value])
        self.api["show_gridlines"] = value

    @property
    def index(self):
        return self._index

    @property
    def book(self):
        return self.sheets.book

    @property
    def notes(self):
        if self.api.get("notes_supported") is False:
            raise NotImplementedError("Notes require ExcelApi 1.18")
        return [
            Note(Range(sheet=self, arg1=entry["address"]))
            for entry in self.api.get("notes", [])
        ]

    @property
    def comments(self):
        raise NotImplementedError(
            "Reading comments synchronously isn't supported on this engine. "
            "Use 'await sheet.get_comments()'."
        )

    async def get_comments(self):
        if sys.platform != "emscripten":
            raise NotImplementedError(
                "get_comments() is only supported in xlwings Lite"
            )
        import js

        entries = _normalize_jsnull((await js.xlwings.getComments(self.name)).to_py())
        return [
            Comment(Range(sheet=self, arg1=entry["address"]), entry["id"])
            for entry in entries
        ]

    def range(self, arg1, arg2=None):
        return Range(sheet=self, arg1=arg1, arg2=arg2)

    @property
    def cells(self):
        return Range(
            sheet=self,
            arg1=(1, 1),
            arg2=(MAX_ROWS, MAX_COLUMNS),
        )

    @property
    def names(self):
        api = [
            name
            for name in self.book.api["names"]
            if name["scope_sheet_index"] is not None
            and name["scope_sheet_index"] + 1 == self.index
            and not name["book_scope"]
        ]
        return Names(parent=self, api=api)

    def activate(self):
        ix = self.index - 1
        self.book.api["book"]["active_sheet_index"] = ix
        self.append_json_action(func="activateSheet", args=ix)

    @property
    def pictures(self):
        return Pictures(self)

    @property
    def shapes(self):
        return Shapes(self)

    @property
    def charts(self):
        return Charts(self)

    @property
    def page_setup(self):
        return PageSetup(self)

    @property
    def tables(self):
        return Tables(parent=self)

    @property
    def pivot_tables(self):
        # Only the Office.js client sends this key, and it sends null when
        # Excel lacks the ExcelApi 1.8 pivot table API.
        if self.api.get("pivot_tables") is None:
            raise NotImplementedError(
                "Pivot tables are only supported with Office.js clients "
                "(ExcelApi 1.8 or later)."
            )
        return PivotTables(parent=self)

    def autofit(self, axis=None):
        if axis in ("rows", "r"):
            self.append_json_action(func="setSheetAutofit", args="rows")
        elif axis in ("columns", "c"):
            self.append_json_action(func="setSheetAutofit", args="columns")
        elif axis is None:
            self.append_json_action(func="setSheetAutofit", args="rows")
            self.append_json_action(func="setSheetAutofit", args="columns")
        else:
            raise ValueError(
                f"Invalid axis: {axis!r}. Use 'rows'/'r', 'columns'/'c' or None."
            )

    def select(self):
        # Office.js only has activate(); selecting a sheet is activating it.
        self.activate()

    def copy(self, before=None, after=None):
        # The public Sheet.copy() finds the new sheet by diffing sheet names
        # before and after this call, so the local api list has to gain the
        # copy right away -- a queued action alone would leave it empty and
        # the diff would raise. Same approach as Sheets.add().
        target = before if before is not None else after
        if target.book is not self.book:
            raise NotImplementedError(
                "Sheet.copy() can't copy to a different book on this engine, "
                "which only positions the copy within the same workbook."
            )
        if before is not None:
            position, target_ix = "Before", before.index - 1
        else:
            position, target_ix = "After", after.index - 1
        sheets_api = self.book.api["sheets"]
        existing = {sheet["name"] for sheet in sheets_api}
        # Excel names the copy "<name> (2)", incrementing until it's free.
        suffix = 2
        while f"{self.name} ({suffix})" in existing:
            suffix += 1
        name = f"{self.name} ({suffix})"
        new_ix = target_ix if position == "Before" else target_ix + 1
        # A copied worksheet starts with the same metadata, but it must own its
        # nested collections. A shallow copy would make changes to charts,
        # shapes, notes, tables, pictures or cached values leak back to the
        # source sheet's local representation.
        api = copy.deepcopy(self.api)
        api["name"] = name
        sheets_api.insert(new_ix, api)
        self.append_json_action(func="copySheet", args=[position, target_ix, name])

    def move(self, before=None, after=None):
        target = before if before is not None else after
        if target.book is not self.book:
            raise ValueError("Sheets must belong to the same book.")

        source_ix = self.index - 1
        target_ix = target.index - 1
        if source_ix == target_ix:
            raise ValueError("A sheet can't be moved relative to itself.")
        if before is not None:
            new_ix = target_ix - 1 if source_ix < target_ix else target_ix
        else:
            new_ix = target_ix if source_ix < target_ix else target_ix + 1

        # Queue the action against the source's position before changing the local
        # snapshot. Office.js Worksheet.position is zero-based.
        self.book.append_json_action(
            func="setSheetPosition",
            args=[new_ix],
            sheet_position=source_ix,
        )

        sheets_api = self.book.api["sheets"]
        moved_sheet = sheets_api.pop(source_ix)
        sheets_api.insert(new_ix, moved_sheet)

        def moved_index(index):
            if index == source_ix:
                return new_ix
            if source_ix < new_ix and source_ix < index <= new_ix:
                return index - 1
            if new_ix < source_ix and new_ix <= index < source_ix:
                return index + 1
            return index

        book_api = self.book.api["book"]
        book_api["active_sheet_index"] = moved_index(book_api["active_sheet_index"])
        for name in self.book.api["names"]:
            for key in ("sheet_index", "scope_sheet_index"):
                if name.get(key) is not None:
                    name[key] = moved_index(name[key])
        self._index = new_ix + 1

    def to_html(self, path):
        raise NotImplementedError(
            "Sheet.to_html() is not supported on this engine, which has no HTML export."
        )

    def delete(self):
        ix = self.index - 1
        del self.book.api["sheets"][ix]
        # Keep the locally tracked active sheet pointing at an existing sheet:
        # deleting a sheet at or before it shifts the remaining ones down, and
        # deleting the last sheet would otherwise leave the index out of range.
        book_api = self.book.api["book"]
        active_ix = book_api["active_sheet_index"]
        if active_ix > ix:
            book_api["active_sheet_index"] = active_ix - 1
        elif active_ix == ix:
            book_api["active_sheet_index"] = min(ix, len(self.book.api["sheets"]) - 1)
        self.append_json_action(func="sheetDelete")

    def clear(self):
        self.append_json_action(func="sheetClear")

    def clear_contents(self):
        self.append_json_action(func="sheetClearContents")

    def clear_formats(self):
        self.append_json_action(func="sheetClearFormats")

    @property
    def used_range(self):
        address = self.api.get("used_range_address")
        if address:
            return Range(sheet=self, arg1=address)
        if address is None and "used_range_address" in self.api:
            # The client reported an empty sheet, which has no used range.
            # Excel's COM API reports A1 in that case.
            return Range(sheet=self, arg1=(1, 1))
        # Fallback for clients that don't send `used_range_address` yet: derive
        # the extent from the shape of the values payload. Since those values
        # are anchored at A1, the used range's real top-left corner is lost,
        # i.e. a sheet whose used range is C5:D10 reports A1:D10.
        if self.book._lazy and not _sheet_values_loaded(self.api):
            raise XlwingsError(
                f"Cell values of sheet '{self.name}' haven't been loaded "
                "(async book). Use 'await book.load(values=True)' or "
                "'await sheet.load(values=True)' first."
            )
        values = self.api["values"]
        nrows = len(values)
        ncols = max((len(row) for row in values), default=0)
        if nrows == 0 or ncols == 0:
            # Empty sheet: Excel reports A1 as the used range
            return Range(sheet=self, arg1=(1, 1))
        return Range(sheet=self, arg1=(1, 1), arg2=(nrows, ncols))

    async def get_used_range(self, values_only=False):
        if sys.platform != "emscripten":
            raise NotImplementedError(
                "Sheet.get_used_range() is only supported in xlwings Lite"
            )
        import js

        address = _normalize_jsnull(
            await js.xlwings.getUsedRangeAddress(self.name, values_only)
        )
        return Range(sheet=self, arg1=address) if address else None

    @property
    def freeze_panes(self):
        return FreezePanes(self)

    async def load(self, values=None):
        """(Re)load this sheet's data from Excel on demand.

        Like `Book.load`, this loads only *metadata* (tables, pictures, names,
        structure) by default on an async (lazy) book, and everything including
        values on a regular book. Pass `values=True` to also load this sheet's
        cell values.
        """
        if sys.platform != "emscripten":
            raise NotImplementedError("Sheet.load() is only supported in xlwings Lite")
        import js
        from pyodide.ffi import to_js

        book = self.book
        load_values = (not book._lazy) if values is None else bool(values)
        opts = to_js(
            {"include": self.name, "lazy": not load_values},
            dict_converter=js.Object.fromEntries,
        )
        book_data_js = await js.xlwings.getBookData(opts)
        book_data = _normalize_jsnull(book_data_js.to_py())
        for sheet_data in book_data["sheets"]:
            if sheet_data["name"] == self.name:
                if not load_values:
                    # Don't clobber any values already present with the empty
                    # metadata-only payload.
                    sheet_data.pop("values", None)
                _update_pivots_in_place(self._api, sheet_data)
                self._api.update(sheet_data)
                break
        if load_values:
            _mark_sheet_values_loaded(self._api)
        get_range_api.cache_clear()


@lru_cache(None)
def get_range_api(api_values, arg1, arg2=None):
    # Keeping this outside of the Range class allows us to cache it across multiple
    # instances of the same range
    if arg2:
        values = [
            row[arg1[1] - 1 : arg2[1]] for row in api_values[arg1[0] - 1 : arg2[0]]
        ]
        if not values:
            # Completely outside the used range
            return [(None,) * (arg2[1] + 1 - arg1[1])] * (arg2[0] + 1 - arg1[0])
        else:
            # Partly outside the used range
            nrows = arg2[0] + 1 - arg1[0]
            ncols = arg2[1] + 1 - arg1[1]
            nrows_actual = len(values)
            ncols_actual = len(values[0])
            delta_rows = nrows - nrows_actual
            delta_cols = ncols - ncols_actual
            if delta_rows != 0:
                values.extend([(None,) * ncols_actual] * delta_rows)
            if delta_cols != 0:
                v = []
                for row in values:
                    v.append(row + (None,) * delta_cols)
                values = v
            return values
    else:
        try:
            values = [(api_values[arg1[0] - 1][arg1[1] - 1],)]
            return values
        except IndexError:
            # Outside the used range
            return [(None,)]


class Range(base_classes.Range):
    def __init__(self, sheet, arg1, arg2=None):
        self.sheet = sheet
        self.arg1_input = arg1
        self.arg2_input = arg2

        # Handle None case (for app.selection if e.g., a shape is selected)
        if arg1 is None:
            self.arg1 = None
            self.arg2 = None
            return

        # Range
        if isinstance(arg1, Range) and isinstance(arg2, Range):
            cell1 = arg1.coords[1], arg1.coords[2]
            cell2 = arg2.coords[1], arg2.coords[2]
            arg1 = min(cell1[0], cell2[0]), min(cell1[1], cell2[1])
            arg2 = max(cell1[0], cell2[0]), max(cell1[1], cell2[1])
        # A1 notation
        if isinstance(arg1, str):
            # A1 notation
            tuple1, tuple2 = utils.a1_to_tuples(arg1)
            if not tuple1:
                # Local names shadow workbook names. Resolve the name before
                # its coordinates: constants/formulas exist but aren't ranges,
                # and a local name may point to a different worksheet.
                candidates = [
                    name
                    for name in sheet.book.api["names"]
                    if name["name"].rsplit("!", 1)[-1].casefold() == arg1.casefold()
                    and (
                        name["book_scope"]
                        or name["scope_sheet_index"] == sheet.index - 1
                    )
                ]
                if candidates:
                    api_name = min(candidates, key=lambda name: name["book_scope"])
                    resolved = Name(
                        parent=sheet.book if api_name["book_scope"] else sheet,
                        api=api_name,
                    ).refers_to_range
                    sheet = resolved.sheet
                    tuple1, tuple2 = resolved.arg1, resolved.arg2
            if not tuple1:
                # Tables
                for api_table in sheet.api["tables"]:
                    if api_table["name"] == arg1:
                        tuple1, tuple2 = utils.a1_to_tuples(
                            api_table["data_body_range_address"]
                        )
                        break
            if not tuple1:
                raise NoSuchObjectError(
                    f"The address/named range '{arg1}' doesn't exist."
                )
            arg1, arg2 = tuple1, tuple2

        # Coordinates
        if len(arg1) == 4:
            row, col, nrows, ncols = arg1
            arg1 = (row, col)
            if nrows > 1 or ncols > 1:
                arg2 = (row + nrows - 1, col + ncols - 1)

        self.arg1 = arg1  # 1-based tuple
        self.arg2 = arg2  # 1-based tuple
        self.sheet = sheet

    def append_json_action(self, **kwargs):
        # Do nothing if the range is None (e.g., if a Shape is selected)
        if self.arg1 is None:
            return
        nrows, ncols = self.shape
        self.sheet.book.append_json_action(
            **{
                **kwargs,
                **{
                    "sheet_position": self.sheet.index - 1,
                    "start_row": self.row - 1,
                    "start_column": self.column - 1,
                    "row_count": nrows,
                    "column_count": ncols,
                },
            }
        )

    @property
    def autofilter(self):
        return AutoFilter(self)

    @property
    def api(self):
        return get_range_api(
            tuple(tuple(row) for row in self.sheet.api["values"]), self.arg1, self.arg2
        )

    @property
    def coords(self):
        return self.sheet.name, self.row, self.column, len(self.api), len(self.api[0])

    @property
    def row(self):
        return self.arg1[0]

    @property
    def column(self):
        return self.arg1[1]

    @property
    def shape(self):
        if self.arg2:
            return self.arg2[0] - self.arg1[0] + 1, self.arg2[1] - self.arg1[1] + 1
        else:
            return 1, 1

    def get_async_pipeline_overrides(self, options):
        """Return async stage replacements for the converter pipeline."""
        if sys.platform != "emscripten":
            raise NotImplementedError("get_value() is only supported in xlwings Lite")
        from ..conversion.standard import (
            AsyncExpandRangeStage,
            AsyncReadValueFromRangeStage,
            ExpandRangeStage,
            ReadValueFromRangeStage,
        )

        overrides = {ReadValueFromRangeStage: AsyncReadValueFromRangeStage(options)}
        if options.get("expand", None):
            overrides[ExpandRangeStage] = AsyncExpandRangeStage(options)
        return overrides

    @property
    def raw_value(self):
        # On an async (lazy) book, cell values aren't pre-loaded. Reading `.value`
        # synchronously would silently return None; raise instead and point to
        # the async API. `await get_value()` doesn't go through here (it fetches
        # via getRangeValues), and `await book.load(values=True)` marks the sheet
        # as loaded, after which sync reads work again.
        if self.sheet.book._lazy and not _sheet_values_loaded(self.sheet.api):
            raise XlwingsError(
                f"Cell values of sheet '{self.sheet.name}' haven't been loaded "
                "(async book). Use 'await myrange.get_value()' to read on demand, "
                "or 'await book.load(values=True)' to load all values first."
            )
        return self.api

    @raw_value.setter
    def raw_value(self, value):
        if not isinstance(value, list):
            # Covers also this case: myrange['A1:B2'].value = 'xyz'
            nrows, ncols = self.shape
            values = [[value] * ncols] * nrows
        else:
            values = value
        self.append_json_action(
            func="setValues",
            values=values,
        )

    def clear_contents(self):
        self.append_json_action(
            func="rangeClearContents",
        )

    def clear(self):
        self.append_json_action(
            func="rangeClear",
        )

    def sort(self, keys, ascending, has_headers):
        self.append_json_action(
            func="rangeSort",
            args=[keys, ascending, has_headers],
        )

    def remove_duplicates(self, columns, has_headers):
        self.append_json_action(
            func="rangeRemoveDuplicates",
            args=[columns, has_headers],
        )

    async def get_special_cells(self, cell_type, value_type):
        if sys.platform != "emscripten":
            raise NotImplementedError(
                "Range.get_special_cells() requires xlwings Lite on this engine"
            )
        import js

        addresses = await js.xlwings.getSpecialCells(
            self.sheet.name, self.address, cell_type, value_type
        )
        return [Range(self.sheet, address) for address in addresses.to_py()]

    async def find(self, text, whole, direction, order, match_case):
        if sys.platform != "emscripten":
            raise NotImplementedError(
                "Range.find() requires xlwings Lite on this engine"
            )
        import js

        address = await js.xlwings.findRange(
            self.sheet.name,
            self.address,
            text,
            whole,
            direction,
            order,
            match_case,
        )
        return Range(self.sheet, address) if address else None

    def replace_all(self, old, new, whole, match_case):
        self.append_json_action(
            func="rangeReplaceAll", args=[old, new, whole, match_case]
        )

    def clear_formats(self):
        self.append_json_action(
            func="rangeClearFormats",
        )

    def get_address(self, row_absolute, col_absolute, external):
        # Purely local: the engine already knows the range's coordinates, so
        # this is string formatting rather than something to fetch.
        if self.arg1 is None:
            return
        row_prefix = "$" if row_absolute else ""
        col_prefix = "$" if col_absolute else ""
        nrows, ncols = self.shape

        def cell(row, column):
            return f"{col_prefix}{utils.col_name(column)}{row_prefix}{row}"

        address = cell(self.row, self.column)
        if nrows != 1 or ncols != 1:
            address += f":{cell(self.row + nrows - 1, self.column + ncols - 1)}"
        if not external:
            return address
        # Excel quotes the sheet name when it or the book name has spaces.
        book_name = self.sheet.book.name
        sheet_name = self.sheet.name
        prefix = f"[{book_name}]{sheet_name}"
        if " " in book_name or " " in sheet_name:
            prefix = f"'{prefix}'"
        return f"{prefix}!{address}"

    @property
    def address(self):
        # Handle non-cell selection
        if self.arg1 is None:
            return
        nrows, ncols = self.shape
        address = f"${utils.col_name(self.column)}${self.row}"
        if nrows == 1 and ncols == 1:
            return address
        else:
            return (
                f"{address}"
                f":${utils.col_name(self.column + ncols - 1)}${self.row + nrows - 1}"
            )

    @property
    def has_array(self):
        # Not supported, but since this is only used for legacy CSE arrays, probably
        # not much of an issue. Here as there's currently a dependency in expansion.py.
        return None

    def end(self, direction):
        if direction == "down":
            i = 1
            while True:
                try:
                    if self.sheet.api["values"][self.row - 1 + i][
                        self.column - 1
                    ] not in (None, ""):
                        i += 1
                    else:
                        break
                except IndexError:
                    break  # outside used range
            nrows = i - 1
            return self.sheet.range((self.row + nrows, self.column))
        if direction == "up":
            i = -1
            while True:
                row_ix = self.row - 1 + i
                if row_ix >= 0 and self.sheet.api["values"][row_ix][
                    self.column - 1
                ] not in (None, ""):
                    i -= 1
                else:
                    break
            nrows = i + 1
            return self.sheet.range((self.row + nrows, self.column))
        if direction == "right":
            i = 1
            while True:
                try:
                    if self.sheet.api["values"][self.row - 1][
                        self.column - 1 + i
                    ] not in (None, ""):
                        i += 1
                    else:
                        break
                except IndexError:
                    break  # outside used range
            ncols = i - 1
            return self.sheet.range((self.row, self.column + ncols))
        if direction == "left":
            i = -1
            while True:
                col_ix = self.column - 1 + i
                if col_ix >= 0 and self.sheet.api["values"][self.row - 1][
                    col_ix
                ] not in (None, ""):
                    i -= 1
                else:
                    break
            ncols = i + 1
            return self.sheet.range((self.row, self.column + ncols))

    def autofit(self, axis=None):
        if axis == "rows" or axis == "r":
            self.append_json_action(func="setAutofit", args="rows")
        elif axis == "columns" or axis == "c":
            self.append_json_action(func="setAutofit", args="columns")
        elif axis is None:
            self.append_json_action(func="setAutofit", args="rows")
            self.append_json_action(func="setAutofit", args="columns")

    @property
    def color(self):
        raise NotImplementedError(
            "Reading the fill color synchronously isn't supported on this "
            "engine. Use 'await myrange.get_color()' to fetch it on demand."
        )

    @color.setter
    def color(self, value):
        self.append_json_action(func="setRangeColor", args=_color_to_hex(value))

    def set_colors(self, colors):
        matrix = [
            ["keep" if color is ... else _color_to_hex(color) for color in row]
            for row in colors
        ]
        self.append_json_action(
            func="setRangeColors",
            args=[matrix],
        )

    @property
    def formula(self):
        # Formulas aren't part of the payload that the client sends with every
        # request, to keep it small. Use `await get_formula()` instead, which
        # fetches them on demand.
        raise NotImplementedError(
            "Reading formulas synchronously isn't supported on this engine. "
            "Use 'await myrange.get_formula()' to fetch them on demand."
        )

    @property
    def left(self):
        raise NotImplementedError(
            "Reading the left position synchronously isn't supported on this "
            "engine. Use 'await myrange.get_left()' to fetch it on demand."
        )

    @property
    def top(self):
        raise NotImplementedError(
            "Reading the top position synchronously isn't supported on this "
            "engine. Use 'await myrange.get_top()' to fetch it on demand."
        )

    @property
    def width(self):
        raise NotImplementedError(
            "Reading the width synchronously isn't supported on this engine. "
            "Use 'await myrange.get_width()' to fetch it on demand."
        )

    @property
    def height(self):
        raise NotImplementedError(
            "Reading the height synchronously isn't supported on this engine. "
            "Use 'await myrange.get_height()' to fetch it on demand."
        )

    @property
    def current_region(self):
        raise NotImplementedError(
            "Reading the current region synchronously isn't supported on this "
            "engine. Use 'await myrange.get_current_region()' to fetch it on "
            "demand."
        )

    @property
    def merge_area(self):
        raise NotImplementedError(
            "Reading the merge area synchronously isn't supported on this "
            "engine. Use 'await myrange.get_merge_area()' to fetch it on demand."
        )

    @property
    def merge_cells(self):
        raise NotImplementedError(
            "Reading whether cells are merged synchronously isn't supported on "
            "this engine. Use 'await myrange.get_merge_cells()' to fetch it on "
            "demand."
        )

    @property
    def table(self):
        raise NotImplementedError(
            "Reading the table synchronously isn't supported on this engine. "
            "Use 'await myrange.get_table()' to fetch it on demand."
        )

    @property
    def characters(self):
        # Office.js has no character-range object for cells: TextRange (and so
        # getSubstring) only exists on shapes. The cell-level `textRuns` in
        # getCellProperties is a runs model that would mean reimplementing
        # Excel's rich-text splitting, so this raises rather than half-doing it.
        raise NotImplementedError(
            "Range.characters is not supported on this engine, which can't "
            "address a range of characters within a cell. Shape.characters works."
        )

    @property
    def conditional_formats(self):
        # The collection itself is useful without a read because clear() is a
        # queued mutation. Dynamic inspection requires the async getter below.
        return ConditionalFormats(self)

    @property
    def note(self):
        # The payload carries the sheet's notes keyed by address, so this
        # knows whether one exists without a fetch -- as the sync property
        # requires. Returns None when there's no note, like the other engines.
        for note in self.sheet.api.get("notes") or []:
            if _address_key(note["address"]) == _address_key(self.address):
                return Note(self)
        return None

    def add_note(self, text):
        if self.sheet.book.api["client"] != "Office.js":
            raise NotImplementedError("Adding notes requires an Office.js client")
        notes = self.sheet.api.get("notes")
        if notes is None or self.sheet.api.get("notes_supported") is False:
            raise NotImplementedError("Notes require ExcelApi 1.18")
        self.append_json_action(func="addNote", args=[self.address, text])
        notes.append({"address": self.address})
        return Note(self)

    @property
    def comment(self):
        raise NotImplementedError(
            "Reading a comment synchronously isn't supported on this engine. "
            "Use 'await range.get_comment()'."
        )

    async def get_comment(self):
        if sys.platform != "emscripten":
            raise NotImplementedError("get_comment() is only supported in xlwings Lite")
        import js

        result = _normalize_jsnull(
            await js.xlwings.getCommentAt(self.sheet.name, self.address)
        )
        if result is None:
            return None
        entry = _normalize_jsnull(result.to_py())
        return Comment(self, entry["id"])

    def add_comment(self, text):
        if self.sheet.book.api["client"] != "Office.js":
            raise NotImplementedError(
                "Adding threaded comments requires an Office.js client"
            )
        self.append_json_action(func="addComment", args=[self.address, text])
        return Comment(self)

    @property
    def hyperlink(self):
        raise NotImplementedError(
            "Reading the hyperlink synchronously isn't supported on this engine. "
            "Use 'await myrange.get_hyperlink()' to fetch it on demand."
        )

    async def get_hyperlink(self):
        return await self._get_range_data("hyperlink")

    async def get_current_region(self):
        address = await self._get_range_data("current_region")
        return Range(sheet=self.sheet, arg1=address)

    async def get_merge_area(self):
        address = await self._get_range_data("merge_area")
        # Office.js reports no merged areas for an unmerged cell; xlwings
        # returns the cell itself in that case.
        return Range(sheet=self.sheet, arg1=address) if address else self

    async def get_merge_cells(self):
        return await self._get_range_data("merge_cells")

    async def get_table(self):
        name = await self._get_range_data("table")
        if not name:
            return None
        # Table's constructor indexes the sheet's tables list, so resolve the
        # name the client reported to its position there.
        for ix, table in enumerate(self.sheet.api["tables"]):
            if table["name"] == name:
                return Table(self.sheet, ix + 1)
        raise KeyError(name)

    async def get_data_validation(self):
        entry = await self._get_range_data("data_validation")
        return DataValidation(self, entry)

    async def get_conditional_formats(self):
        entries = await self._get_range_data("conditional_formats")
        return ConditionalFormats(self, entries)

    async def _get_range_data(self, key, method=None):
        """Fetch one on-demand property for this range from the client.

        `getRangeData` takes a list of read keys and returns them under the
        same names, so this is a thin wrapper around it. `method` names the
        calling method in the error message, for the keys whose name differs
        from it.
        """
        if sys.platform != "emscripten":
            raise NotImplementedError(
                f"{method or f'get_{key}'}() is only supported in xlwings Lite"
            )
        import js
        from pyodide.ffi import to_js

        data_js = await js.xlwings.getRangeData(
            self.sheet.name, self.address, to_js([key])
        )
        return _normalize_jsnull(data_js.to_py())[key]

    async def get_formula_array(self):
        return await self._get_range_data("formula_array")

    async def get_number_format(self):
        return await self._get_range_data("number_format")

    async def get_color(self):
        color = await self._get_range_data("color")
        return utils.hex_to_rgb(color) if color else None

    async def get_colors(self):
        colors = await self._get_range_data("colors")
        return [
            [utils.hex_to_rgb(color) if color is not None else None for color in row]
            for row in colors
        ]

    async def get_wrap_text(self):
        return await self._get_range_data("wrap_text")

    async def get_horizontal_alignment(self):
        # Office.js reports None for a range whose cells disagree.
        return _HORIZONTAL_ALIGNMENT_JS2PY.get(
            await self._get_range_data("horizontal_alignment")
        )

    async def get_vertical_alignment(self):
        return _VERTICAL_ALIGNMENT_JS2PY.get(
            await self._get_range_data("vertical_alignment")
        )

    async def get_column_width(self):
        return await self._get_range_data("column_width")

    async def get_row_height(self):
        return await self._get_range_data("row_height")

    async def get_left(self):
        return await self._get_range_data("left")

    async def get_top(self):
        return await self._get_range_data("top")

    async def get_width(self):
        return await self._get_range_data("width")

    async def get_height(self):
        return await self._get_range_data("height")

    async def get_formula(self):
        """Fetch this range's formulas as a raw 2D list.

        Always 2D and unsqueezed, which is what `AdjustDimensionsStage` expects.
        The public `Range.get_formula` applies the `ndim` option on top, so that
        what the *user* gets back matches the shape of reading `.value`.
        """
        return await self._get_range_data("formulas", method="get_formula")

    @formula.setter
    def formula(self, value):
        nrows, ncols = self.shape
        if not isinstance(value, list):
            # Scalars broadcast over the whole range, like on the other engines.
            self.append_json_action(func="setFormula", values=[[value] * ncols] * nrows)
            return
        if value and not isinstance(value[0], list):
            # A flat list is a row, unless the target is a single column. This
            # mirrors how `.value` treats flat lists.
            value = [[item] for item in value] if ncols == 1 and nrows != 1 else [value]
        if not value or not value[0]:
            return
        # Like `.value`, the data wins over the target's current shape: writing
        # more formulas than the range holds expands it instead of raising.
        target = self
        if len(value) != nrows or len(value[0]) != ncols:
            target = Range(
                self.sheet,
                self.arg1,
                (self.arg1[0] + len(value) - 1, self.arg1[1] + len(value[0]) - 1),
            )
        target.append_json_action(func="setFormula", values=value)

    @property
    def formula2(self):
        # See `formula`: Office.js has no separate formula2, and reading is
        # async on this engine.
        raise NotImplementedError(
            "Reading formulas synchronously isn't supported on this engine. "
            "Use 'await myrange.get_formula()' to fetch them on demand."
        )

    @formula2.setter
    def formula2(self, value):
        # Office.js has no separate formula2: range.formulas already writes
        # dynamic array formulas, which is what formula2 stands for.
        self.formula = value

    @property
    def column_width(self):
        raise NotImplementedError(
            "Reading the column width synchronously isn't supported on this engine. "
            "Use 'await myrange.get_column_width()' to fetch it on demand."
        )

    @column_width.setter
    def column_width(self, value):
        if isinstance(value, bool) or not isinstance(value, numbers.Real):
            raise ValueError("column_width must be a number.")
        if value < 0:
            raise ValueError("column_width can't be negative.")
        # Points, which is what this host measures column widths in -- the
        # desktop engines pass their own raw unit (characters) through the same
        # way. Converting between the two can't be done exactly without
        # measuring the workbook's font, so the raw value is passed on and
        # cross-engine code converts if it needs to.
        self.append_json_action(func="setColumnWidth", args=value)

    @property
    def row_height(self):
        raise NotImplementedError(
            "Reading the row height synchronously isn't supported on this engine. "
            "Use 'await myrange.get_row_height()' to fetch it on demand."
        )

    @row_height.setter
    def row_height(self, value):
        self.append_json_action(func="setRowHeight", args=value)

    @property
    def wrap_text(self):
        raise NotImplementedError(
            "Reading the wrap text flag synchronously isn't supported on this engine. "
            "Use 'await myrange.get_wrap_text()' to fetch it on demand."
        )

    @wrap_text.setter
    def wrap_text(self, value):
        self.append_json_action(func="setWrapText", args=bool(value))

    def _require_officejs(self, name):
        # Only the Office.js client implements the alignment callbacks; the
        # other clients would fail with an opaque error when dispatching them.
        if self.sheet.book.api["client"] != "Office.js":
            raise NotImplementedError(
                f"Range.{name} is only supported with Office.js clients"
            )

    @property
    def horizontal_alignment(self):
        raise NotImplementedError(
            "Reading the horizontal alignment synchronously isn't supported on this "
            "engine. Use 'await myrange.get_horizontal_alignment()' to fetch it on "
            "demand."
        )

    @horizontal_alignment.setter
    def horizontal_alignment(self, value):
        self._require_officejs("horizontal_alignment")
        self.append_json_action(
            func="setHorizontalAlignment", args=_HORIZONTAL_ALIGNMENT_PY2JS[value]
        )

    @property
    def vertical_alignment(self):
        raise NotImplementedError(
            "Reading the vertical alignment synchronously isn't supported on this "
            "engine. Use 'await myrange.get_vertical_alignment()' to fetch it on "
            "demand."
        )

    @vertical_alignment.setter
    def vertical_alignment(self, value):
        self._require_officejs("vertical_alignment")
        self.append_json_action(
            func="setVerticalAlignment", args=_VERTICAL_ALIGNMENT_PY2JS[value]
        )

    @property
    def formula_array(self):
        raise NotImplementedError(
            "Reading array formulas synchronously isn't supported on this engine. "
            "Use 'await myrange.get_formula_array()' to fetch it on demand."
        )

    @formula_array.setter
    def formula_array(self, value):
        self.append_json_action(func="setFormulaArray", args=value)

    def add_hyperlink(self, address, text_to_display=None, screen_tip=None):
        self.append_json_action(
            func="addHyperlink", args=[address, text_to_display, screen_tip]
        )

    @property
    def number_format(self):
        raise NotImplementedError(
            "Reading the number format synchronously isn't supported on this engine. "
            "Use 'await myrange.get_number_format()' to fetch it on demand."
        )

    @number_format.setter
    def number_format(self, value):
        self.append_json_action(func="setNumberFormat", args=value)

    @property
    def name(self):
        for name in self.sheet.book.api["names"]:
            if name["sheet_index"] == self.sheet.index - 1 and name[
                "address"
            ] == self.address.replace("$", ""):
                return Name(
                    parent=self.sheet.book if name["book_scope"] else self.sheet,
                    api=name,
                )

    @name.setter
    def name(self, value):
        self.append_json_action(
            func="setRangeName",
            args=value,
        )

    def autofill(self, destination, type_):
        types = {
            "fill_copy": "FillCopy",
            "fill_days": "FillDays",
            "fill_default": "FillDefault",
            "fill_formats": "FillFormats",
            "fill_months": "FillMonths",
            "fill_series": "FillSeries",
            "fill_values": "FillValues",
            "fill_weekdays": "FillWeekdays",
            "fill_years": "FillYears",
            "growth_trend": "GrowthTrend",
            "linear_trend": "LinearTrend",
            "flash_fill": "FlashFill",
        }
        if type_ not in types:
            raise XlwingsError(
                f"Invalid autofill type '{type_}'. "
                f"Must be one of: {', '.join(sorted(types))}."
            )
        destination_impl = destination.impl
        if (
            destination_impl.sheet.book is not self.sheet.book
            or destination_impl.sheet.api is not self.sheet.api
        ):
            raise XlwingsError(
                "range.autofill() requires the destination to be on the same sheet."
            )
        self.append_json_action(
            func="rangeAutofill", args=[destination.address, types[type_]]
        )

    def copy(self, destination=None):
        if destination is None:
            raise XlwingsError("range.copy() requires a destination argument.")
        self.append_json_action(
            func="copyRange",
            args=[destination.sheet.index - 1, destination.address],
        )

    def copy_from(self, source_range, copy_type=None, skip_blanks=None, transpose=None):
        self.append_json_action(
            func="copyFromRange",
            args=[
                source_range.sheet.index - 1,
                source_range.address,
                copy_type,
                skip_blanks,
                transpose,
            ],
        )

    def delete(self, shift=None):
        if shift not in ("up", "left"):
            # Non-remote version allows shift=None
            raise XlwingsError(
                "range.delete() requires either 'up' or 'left' as shift arguments."
            )
        self.append_json_action(func="rangeDelete", args=shift)

    def insert(self, shift=None, copy_origin=None):
        if shift not in ("down", "right"):
            raise XlwingsError(
                "range.insert() requires either 'down' or 'right' as shift arguments."
            )
        if copy_origin not in (
            "format_from_left_or_above",
            "format_from_right_or_below",
        ):
            raise XlwingsError(
                "range.insert() requires either 'format_from_left_or_above' or "
                "'format_from_right_or_below' as copy_origin arguments."
            )
        # copy_origin is only supported by VBA clients
        self.append_json_action(func="rangeInsert", args=[shift, copy_origin])

    def select(self):
        self.append_json_action(
            func="rangeSelect",
        )

    def merge(self, across):
        self.append_json_action(func="rangeMerge", args=bool(across))

    def unmerge(self):
        self.append_json_action(func="rangeUnmerge")

    def group(self, by):
        self.append_json_action(func="rangeGroup", args=[by])

    def ungroup(self, by):
        self.append_json_action(func="rangeUngroup", args=[by])

    def adjust_indent(self, amount):
        self.append_json_action(func="rangeAdjustIndent", args=amount)

    def to_png(self, path):
        self.append_json_action(func="rangeToPng", args=[path])

    def copy_picture(self, appearance, format):
        # Copies to the system clipboard, which Office.js has no API for --
        # the same reason paste() can't work. Range.to_png() covers the
        # "get this range as an image" case.
        raise NotImplementedError(
            "Range.copy_picture() is not supported on this engine, which has no "
            "clipboard access. Use 'to_png()' to export the range as an image."
        )

    def paste(self, paste=None, operation=None, skip_blanks=False, transpose=False):
        raise NotImplementedError(
            "Range.paste() is not supported on this engine, which has no "
            "clipboard access. Use 'copy()' with an explicit source range instead."
        )

    def to_pdf(self, path, quality):
        raise NotImplementedError(
            "Range.to_pdf() is not supported on this engine, which has no PDF export."
        )

    @property
    def font(self):
        return Font(self, self.sheet.book.api)

    @property
    def borders(self):
        return Borders(self, self.sheet.book.api)

    @property
    def data_validation(self):
        return DataValidation(self)

    def __len__(self):
        nrows, ncols = self.shape
        return nrows * ncols

    def __call__(self, arg1, arg2=None):
        if arg2 is None:
            col = (arg1 - 1) % self.shape[1]
            row = int((arg1 - 1 - col) / self.shape[1])
            return self(row + 1, col + 1)
        else:
            return Range(
                sheet=self.sheet,
                arg1=(self.row + arg1 - 1, self.column + arg2 - 1),
            )


class DataValidation(base_classes.DataValidation):
    def __init__(self, parent, entry=None):
        self.parent = parent
        self._entry = entry

    @property
    def api(self):
        return None

    def _read(self, name):
        if self._entry is None:
            raise NotImplementedError(
                "Reading data validation synchronously isn't supported on this "
                "engine. Use 'await myrange.get_data_validation()' to fetch it on "
                "demand."
            )
        return self._entry.get(name)

    @property
    def type(self):
        return self._read("type")

    @property
    def operator(self):
        return self._read("operator")

    @property
    def formula1(self):
        return self._read("formula1")

    @property
    def formula2(self):
        return self._read("formula2")

    @property
    def formula(self):
        return self._read("formula")

    @property
    def source(self):
        return self._read("source")

    @property
    def in_cell_dropdown(self):
        return self._read("in_cell_dropdown")

    @property
    def ignore_blank(self):
        return self._read("ignore_blank")

    @property
    def input_title(self):
        return self._read("input_title")

    @property
    def input_message(self):
        return self._read("input_message")

    @property
    def show_input(self):
        return self._read("show_input")

    @property
    def error_title(self):
        return self._read("error_title")

    @property
    def error_message(self):
        return self._read("error_message")

    @property
    def show_error(self):
        return self._read("show_error")

    @property
    def alert_style(self):
        return self._read("alert_style")

    def set_list(self, source, in_cell_dropdown):
        self.parent._require_officejs("data_validation")
        if isinstance(source, base_classes.Range):
            source_payload = {
                "type": "range",
                "sheet_position": source.sheet.index - 1,
                "start_row": source.row - 1,
                "start_column": source.column - 1,
                "row_count": source.shape[0],
                "column_count": source.shape[1],
            }
        elif isinstance(source, base_classes.Name):
            source_payload = {"type": "name", "name": source.name}
        else:
            source_payload = {"type": "literal", "values": source}
        self.parent.append_json_action(
            func="setDataValidationList",
            args=[source_payload, in_cell_dropdown],
        )

    def set_rule(self, rule_type, operator, formula1, formula2):
        self.parent._require_officejs("data_validation")
        self.parent.append_json_action(
            func="setDataValidationRule",
            args=[
                {
                    "type": rule_type,
                    "operator": operator,
                    "formula1": formula1,
                    "formula2": formula2,
                }
            ],
        )

    def delete(self):
        self.parent._require_officejs("data_validation")
        self.parent.append_json_action(func="deleteDataValidation")


class Collection(base_classes.Collection):
    def __init__(self, parent):
        self._parent = parent
        # Default to empty rather than raising KeyError: a client older than
        # the payload field reports "no shapes"/"no charts", which is the same
        # answer it would give for a sheet that genuinely has none.
        self._api = parent.api.get(self._attr, [])

    @property
    def api(self):
        return self._api

    @property
    def parent(self):
        return self._parent

    def __call__(self, key):
        if isinstance(key, numbers.Number):
            if key > len(self):
                raise KeyError(key)
            else:
                return self._wrap(self.parent, key)
        else:
            for ix, i in enumerate(self.api):
                if i["name"] == key:
                    return self._wrap(self.parent, ix + 1)
            raise KeyError(key)

    def __len__(self):
        return len(self.api)

    def __iter__(self):
        for ix, api in enumerate(self.api):
            yield self._wrap(self._parent, ix + 1)

    def __contains__(self, key):
        if isinstance(key, numbers.Number):
            return 1 <= key <= len(self)
        else:
            for i in self.api:
                if i["name"] == key:
                    return True
            return False


class Picture(base_classes.Picture):
    def __init__(self, parent, key):
        self._parent = parent
        self._api = self.parent.api["pictures"][key - 1]
        self.key = key

    def append_json_action(self, **kwargs):
        self.parent.book.append_json_action(
            **{
                **kwargs,
                **{
                    "sheet_position": self.parent.index - 1,
                },
            }
        )

    @property
    def api(self):
        return self._api

    @property
    def parent(self):
        return self._parent

    @property
    def name(self):
        return self.api["name"]

    @name.setter
    def name(self, value):
        self.api["name"] = value
        self.append_json_action(func="setPictureName", args=[self.index - 1, value])

    @property
    def left(self):
        return self.api["left"]

    @left.setter
    def left(self, value):
        self.api["left"] = value
        self.append_json_action(func="setPictureLeft", args=[self.index - 1, value])

    @property
    def top(self):
        return self.api["top"]

    @top.setter
    def top(self, value):
        self.api["top"] = value
        self.append_json_action(func="setPictureTop", args=[self.index - 1, value])

    @property
    def lock_aspect_ratio(self):
        return self.api["lock_aspect_ratio"]

    @lock_aspect_ratio.setter
    def lock_aspect_ratio(self, value):
        self.api["lock_aspect_ratio"] = value
        self.append_json_action(
            func="setPictureLockAspectRatio", args=[self.index - 1, value]
        )

    @property
    def width(self):
        return self.api["width"]

    @width.setter
    def width(self, value):
        self.api["width"] = value
        self.append_json_action(func="setPictureWidth", args=[self.index - 1, value])

    @property
    def height(self):
        return self.api["height"]

    @height.setter
    def height(self, value):
        self.api["height"] = value
        self.append_json_action(func="setPictureHeight", args=[self.index - 1, value])

    @property
    def index(self):
        # TODO: make available in public API
        if isinstance(self.key, numbers.Number):
            return self.key
        else:
            for ix, obj in self.api:
                if obj["name"] == self.key:
                    return ix + 1
            raise KeyError(self.key)

    def delete(self):
        self.parent._api["pictures"].pop(self.index - 1)
        self.append_json_action(func="deletePicture", args=self.index - 1)

    def update(self, filename):
        with open(filename, "rb") as image_file:
            encoded_image_string = base64.b64encode(image_file.read()).decode("utf-8")
        self.append_json_action(
            func="updatePicture",
            args=[
                encoded_image_string,
                self.index - 1,
                self.name,
                self.width,
                self.height,
            ],
        )
        return self


class Pictures(Collection, base_classes.Pictures):
    _attr = "pictures"
    _wrap = Picture

    def append_json_action(self, **kwargs):
        self.parent.book.append_json_action(
            **{
                **kwargs,
                **{
                    "sheet_position": self.parent.index - 1,
                },
            }
        )

    def add(
        self,
        filename,
        link_to_file=None,
        save_with_document=None,
        left=None,
        top=None,
        width=None,
        height=None,
        anchor=None,
    ):
        if self.parent.book.api["client"] == "Google Apps Script" and (left or top):
            raise ValueError(
                "'left' and 'top' are not supported with Google Sheets. "
                "Use 'anchor' instead."
            )
        if anchor is None:
            column_index = 0
            row_index = 0
        else:
            column_index = anchor.column - 1
            row_index = anchor.row - 1
        # Google Sheets allows a max size of 1 million pixels. For matplotlib, you
        # can control the pixels like so: fig = plt.figure(figsize=(6, 4), dpi=200)
        # This sample has (6 * 200) * (4 * 200) = 960,000 px
        # Note that savefig(bbox_inches="tight") crops the image and therefore will
        # reduce the number of pixels in a non-deterministic way. Existing figure
        # size can be checked via fig.get_size_inches(). pandas accepts figsize also:
        # ax = df.plot(figsize=(3,3))
        # fig = ax.get_figure()
        with open(filename, "rb") as image_file:
            encoded_image_string = base64.b64encode(image_file.read()).decode("utf-8")
        # TODO: width and height are currently ignored but can be set via obj properties
        self.append_json_action(
            func="addPicture",
            args=[
                encoded_image_string,
                column_index,
                row_index,
                left if left else 0,
                top if top else 0,
            ],
        )
        self.parent._api["pictures"].append(
            {
                "name": "Image",
                "width": None,
                "height": None,
                # Seeded so the getters work before the next round-trip
                # refreshes the payload, rather than raising KeyError.
                "left": left if left else 0,
                "top": top if top else 0,
                "lock_aspect_ratio": None,
            }
        )
        return Picture(self.parent, len(self.parent.api["pictures"]))


def _sheet_name_prefix(name):
    # Quote punctuation, spaces, numeric names, and cell-reference-like names.
    if re.fullmatch(r"[^\W\d][\w.]*", name) and not re.fullmatch(
        r"[A-Za-z]{1,3}[1-9][0-9]*|[Rr][0-9]+[Cc][0-9]+", name
    ):
        return name
    return "'" + name.replace("'", "''") + "'"


# Tokenize only what a sheet rename needs. Strings, external qualifiers and
# structured-reference brackets are opaque; this is not a formula evaluator.
_SHEET_REFERENCE_TOKEN = re.compile(
    r'"(?:[^"]|"")*"'
    r"|'(?P<quoted>(?:[^']|'')*)'!"
    r"|\[[^\]]*\][\w.]+(?::[\w.]+)?!"
    r"|(?P<unquoted>[\w.]+(?::[\w.]+)?)!"
    r"|(?P<bracket>\[)"
)


def _rename_sheet_in_formula(formula, old_name, new_name):
    parts = []
    position = 0
    while match := _SHEET_REFERENCE_TOKEN.search(formula, position):
        parts.append(formula[position : match.start()])
        end = match.end()
        replacement = match[0]
        if match["bracket"]:
            # Structured references can nest and use apostrophe escapes.
            depth = 1
            while end < len(formula) and depth:
                char = formula[end]
                if char == "'" and end + 1 < len(formula):
                    end += 2
                    continue
                if char == "[":
                    depth += 1
                elif char == "]":
                    depth -= 1
                end += 1
            replacement = formula[match.start() : end]
        else:
            qualifier = match["quoted"] or match["unquoted"]
            if qualifier is not None and "[" not in qualifier:
                # A colon separates the endpoints of a 3-D sheet reference.
                names = qualifier.replace("''", "'").split(":")
                if any(name.casefold() == old_name.casefold() for name in names):
                    names = [
                        new_name if name.casefold() == old_name.casefold() else name
                        for name in names
                    ]
                    replacement = _sheet_name_prefix(":".join(names)) + "!"
        parts.append(replacement)
        position = end
    parts.append(formula[position:])
    return "".join(parts)


def _name_reference_metadata(book, refers_to):
    """Keep definitions verbatim; only infer coordinates for simple A1 references.

    Excel resolves other expressions (including formulas returning ranges) when
    the workbook is next loaded. An exclamation mark alone is not proof that a
    definition is a range: LAMBDAs can contain sheet references too.
    """
    metadata = {"refers_to": refers_to, "sheet_index": None, "address": None}
    match = re.fullmatch(
        r"=(?:'((?:[^']|'')+)'|([^'!\[\]():,=+*/^&<>\"%{}-]+))!"
        r"(\$?[A-Za-z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]*)?"
        r"|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[1-9][0-9]*:\$?[1-9][0-9]*)",
        refers_to,
    )
    if match is None:
        return metadata
    sheet_name = (match[1] or match[2]).replace("''", "'")
    sheet = book.sheets(sheet_name)
    metadata["sheet_index"] = sheet.index - 1
    metadata["address"] = match[3].replace("$", "").upper()
    return metadata


class Name(base_classes.Name):
    def __init__(self, parent, api):
        self.parent = parent
        self.api = api

    @property
    def name(self):
        if self.api["book_scope"]:
            return self.api["name"]
        else:
            sheet_name = self.api["scope_sheet_name"]
            if "!" not in self.api["name"]:
                # VBA/Google Sheets already do this
                sheet_name = _sheet_name_prefix(sheet_name)
                return f"{sheet_name}!{self.api['name']}"
            else:
                return self.api["name"]

    @property
    def refers_to(self):
        if "refers_to" in self.api:
            return self.api["refers_to"]
        # Older clients only send range coordinates. Non-range/multi-area
        # names may have neither coordinates nor a definition in that payload.
        if self.api["sheet_index"] is None or self.api["address"] is None:
            return None
        book = self.parent if isinstance(self.parent, Book) else self.parent.book
        sheet = book.sheets(self.api["sheet_index"] + 1)
        sheet_name = _sheet_name_prefix(sheet.name)
        return f"={sheet_name}!{sheet.range(self.api['address']).address}"

    @name.setter
    def name(self, value):
        # Excel.NamedItem.name is readonly in Office.js: a named item can't be
        # renamed, only deleted and recreated -- which changes its identity and
        # drops its comment and visibility, so it isn't done implicitly here.
        raise NotImplementedError(
            "Name.name can't be set on this engine, where a name is read-only. "
            "Delete the name and add it again under the new name."
        )

    @property
    def refers_to_range(self):
        if self.api["sheet_index"] is None or self.api["address"] is None:
            raise ValueError(f"Name '{self.name}' does not refer to a single range.")
        book = self.parent if isinstance(self.parent, Book) else self.parent.book
        sheet = book.sheets(self.api["sheet_index"] + 1)
        return sheet.range(self.api["address"])

    @refers_to.setter
    def refers_to(self, value):
        book = self.parent if isinstance(self.parent, Book) else self.parent.book
        metadata = _name_reference_metadata(book, value)
        self.parent.append_json_action(
            func="setNameRefersTo",
            args=[
                self.api["name"],
                self.api["book_scope"],
                self.api["scope_sheet_index"],
                value,
            ],
        )
        self.api.update(metadata)

    def delete(self):
        # Drop the local entry too, so `name in book.names` is right straight
        # away rather than only after the next round-trip -- the other
        # delete() methods do the same. The parent is a Book or a Sheet, so
        # the list lives on its `names` collection rather than on `api`.
        book = self.parent if isinstance(self.parent, Book) else self.parent.book
        names = book.api["names"]
        if self.api in names:
            names.remove(self.api)
        self.parent.append_json_action(
            func="nameDelete",
            args=[
                self.name,  # this includes the sheet name for sheet scope
                self.refers_to,
                self.api["name"],  # no sheet name
                self.api["sheet_index"],
                self.api["book_scope"],
                self.api["scope_sheet_index"],
            ],
        )


class Names(base_classes.Names):
    def __init__(self, parent, api):
        self.parent = parent
        self.api = api

    def add(self, name, refers_to):
        # TODO: raise backend error in case of duplicates
        if isinstance(self.parent, Book):
            is_parent_book = True
        else:
            is_parent_book = False
        book = self.parent if is_parent_book else self.parent.book
        metadata = _name_reference_metadata(book, refers_to)
        self.parent.append_json_action(func="namesAdd", args=[name, refers_to])

        api = {
            "name": name,
            **metadata,
            "book_scope": True if is_parent_book else False,
            # A sheet-scoped name is scoped to the sheet it was added through;
            # a book-scoped one has no scope sheet. Both are part of the
            # payload, so delete() and the refers_to setter read them --
            # without them those raise KeyError on a name added mid-script.
            "scope_sheet_name": None if is_parent_book else self.parent.name,
            "scope_sheet_index": None if is_parent_book else self.parent.index - 1,
        }
        # Register it on the book's list, which is the one the payload owns.
        # Sheet.names builds a filtered copy of that list on each access, so
        # appending to self.api would be thrown away for a sheet-scoped name.
        book.api["names"].append(api)
        if not is_parent_book:
            self.api.append(api)
        return Name(self.parent, api)

    def __call__(self, name_or_index):
        if isinstance(name_or_index, numbers.Number):
            name_or_index -= 1
            if name_or_index > len(self):
                raise KeyError(name_or_index)
            else:
                return Name(self.parent, api=self.api[name_or_index])
        else:
            for ix, i in enumerate(self.api):
                name = Name(self.parent, api=self.api[ix])
                if name.name == name_or_index:
                    # Sheet scope names have the sheet name prepended
                    return name
            raise KeyError(name_or_index)

    def contains(self, name_or_index):
        if isinstance(name_or_index, numbers.Number):
            return 1 <= name_or_index <= len(self)
        else:
            for i in self.api:
                if i["name"] == name_or_index:
                    return True
            return False

    def __len__(self):
        return len(self.api)


engine = Engine()


class AutoFilter(base_classes.AutoFilter):
    def __init__(self, parent, table_index=None):
        self.parent = parent
        self.table_index = table_index

    def _append(self, range_func, table_func, args):
        self.parent._require_officejs("autofilter")
        if self.table_index is None:
            self.parent.append_json_action(func=range_func, args=args)
        else:
            self.parent.append_json_action(
                func=table_func, args=[self.table_index, *args]
            )

    @property
    def criteria(self):
        raise NotImplementedError(
            "AutoFilter.criteria isn't available on this engine; in xlwings Lite, "
            "use await autofilter.get_criteria()"
        )

    async def get_criteria(self):
        if sys.platform != "emscripten":
            raise NotImplementedError(
                "get_criteria() is only supported in xlwings Lite"
            )
        import js

        data_js = await js.xlwings.getAutoFilterCriteria(
            self.parent.sheet.name,
            self.parent.address,
            self.table_index,
        )
        return _normalize_jsnull(data_js.to_py(), normalize_values=True)

    def apply_values(self, field, values):
        self._append(
            "applyAutoFilterRange",
            "applyAutoFilterTable",
            [field, {"type": "values", "values": values}],
        )

    def apply_comparison(self, field, operator, value1, value2):
        spec = {
            "type": "comparison",
            "operator": operator,
            "value1": value1,
        }
        if value2 is not None:
            spec["value2"] = value2
        self._append("applyAutoFilterRange", "applyAutoFilterTable", [field, spec])

    def _apply_top_bottom(self, field, type_, value):
        self._append(
            "applyAutoFilterRange",
            "applyAutoFilterTable",
            [field, {"type": type_, "value": value}],
        )

    def apply_top_items(self, field, count):
        self._apply_top_bottom(field, "top_items", count)

    def apply_bottom_items(self, field, count):
        self._apply_top_bottom(field, "bottom_items", count)

    def apply_top_percent(self, field, percent):
        self._apply_top_bottom(field, "top_percent", percent)

    def apply_bottom_percent(self, field, percent):
        self._apply_top_bottom(field, "bottom_percent", percent)

    def clear(self, field):
        self._append("clearAutoFilterRange", "clearAutoFilterTable", [field])


class Table(base_classes.Table):
    @property
    def show_autofilter(self):
        return self.api["show_autofilter"]

    @show_autofilter.setter
    def show_autofilter(self, value):
        self.append_json_action(
            func="showAutofilterTable", args=[self.index - 1, value]
        )

    @property
    def autofilter(self):
        return AutoFilter(self.range, table_index=self.index - 1)

    def __init__(self, parent, key):
        self._parent = parent
        self._api = self.parent.api["tables"][key - 1]
        self.key = key

    def append_json_action(self, **kwargs):
        self.parent.book.append_json_action(
            **{
                **kwargs,
                **{
                    "sheet_position": self.parent.index - 1,
                },
            }
        )

    @property
    def api(self):
        return self._api

    @property
    def parent(self):
        return self._parent

    @property
    def name(self):
        return self.api["name"]

    @name.setter
    def name(self, value):
        self.api["name"] = value
        self.append_json_action(func="setTableName", args=[self.index - 1, value])

    @property
    def range(self):
        if self.api["range_address"]:
            return self.parent.range(self.api["range_address"])
        else:
            return None

    @property
    def header_row_range(self):
        if self.api["header_row_range_address"]:
            return self.parent.range(self.api["header_row_range_address"])
        else:
            return None

    @property
    def data_body_range(self):
        if self.api["data_body_range_address"]:
            return self.parent.range(self.api["data_body_range_address"])
        else:
            return None

    @property
    def totals_row_range(self):
        if self.api["total_row_range_address"]:
            return self.parent.range(self.api["total_row_range_address"])
        else:
            return None

    @property
    def show_headers(self):
        return self.api["show_headers"]

    @show_headers.setter
    def show_headers(self, value):
        self.api["row_count"] = (
            self.api.get("row_count", 0)
            + int(value)
            - int(self.api.get("show_headers", True))
        )
        self.api["show_headers"] = value
        self.append_json_action(func="showHeadersTable", args=[self.index - 1, value])

    @property
    def show_totals(self):
        return self.api["show_totals"]

    @show_totals.setter
    def show_totals(self, value):
        self.api["row_count"] = (
            self.api.get("row_count", 0)
            + int(value)
            - int(self.api.get("show_totals", False))
        )
        self.api["show_totals"] = value
        self.append_json_action(func="showTotalsTable", args=[self.index - 1, value])

    @property
    def insert_row_range(self):
        # Office.js' Excel.Table has no InsertRowRange equivalent: its only
        # range accessors are getRange(), getDataBodyRange(),
        # getHeaderRowRange() and getTotalRowRange(). Returning None would be
        # indistinguishable from the documented "table isn't empty" answer, so
        # raise instead of answering wrongly.
        raise NotImplementedError(
            "Table.insert_row_range is not supported on this engine, which has "
            "no equivalent."
        )

    @property
    def display_name(self):
        # Office.js' Excel.Table only has `name`, so display_name aliases it.
        # The two are equivalent in practice anyway: on macOS, setting
        # display_name changes the name too, and Office Scripts dropped the
        # distinction as well.
        return self.name

    @display_name.setter
    def display_name(self, value):
        self.name = value

    @property
    def show_table_style_first_column(self):
        return self.api["show_table_style_first_column"]

    @show_table_style_first_column.setter
    def show_table_style_first_column(self, value):
        self.api["show_table_style_first_column"] = value
        self.append_json_action(
            func="showTableStyleFirstColumn", args=[self.index - 1, value]
        )

    @property
    def show_table_style_last_column(self):
        return self.api["show_table_style_last_column"]

    @show_table_style_last_column.setter
    def show_table_style_last_column(self, value):
        self.api["show_table_style_last_column"] = value
        self.append_json_action(
            func="showTableStyleLastColumn", args=[self.index - 1, value]
        )

    @property
    def show_table_style_row_stripes(self):
        return self.api["show_table_style_row_stripes"]

    @show_table_style_row_stripes.setter
    def show_table_style_row_stripes(self, value):
        self.api["show_table_style_row_stripes"] = value
        self.append_json_action(
            func="showTableStyleRowStripes", args=[self.index - 1, value]
        )

    @property
    def show_table_style_column_stripes(self):
        return self.api["show_table_style_column_stripes"]

    @show_table_style_column_stripes.setter
    def show_table_style_column_stripes(self, value):
        self.api["show_table_style_column_stripes"] = value
        self.append_json_action(
            func="showTableStyleColumnStripes", args=[self.index - 1, value]
        )

    @property
    def table_style(self):
        return self.api["table_style"]

    @table_style.setter
    def table_style(self, value):
        self.append_json_action(func="setTableStyle", args=[self.index - 1, value])

    @property
    def index(self):
        # TODO: make available in public API
        if isinstance(self.key, numbers.Number):
            return self.key
        else:
            for ix, obj in self.api:
                if obj["name"] == self.key:
                    return ix + 1
            raise KeyError(self.key)

    def resize(self, range):
        self.append_json_action(
            func="resizeTable", args=[self.index - 1, range.address]
        )

    @property
    def rows(self):
        return TableRows(self)


class TableRow(base_classes.TableRow):
    def __init__(self, parent, index):
        self.parent = parent
        self._index = index

    @property
    def index(self):
        return self._index

    @property
    def range(self):
        raise NotImplementedError(
            "TableRow.range is not supported in xlwings Lite. Use 'await row.get_range()'."
        )

    async def get_range(self):
        if sys.platform != "emscripten":
            raise NotImplementedError("TableRow.get_range() requires xlwings Lite")
        import js

        address = await js.xlwings.getTableRowRangeAddress(
            self.parent.parent.parent.name, self.parent.parent.index - 1, self.index
        )
        return self.parent.parent.parent.range(str(address))

    def delete(self):
        self.parent.parent.append_json_action(
            func="deleteTableRow",
            args=[self.parent.parent.index - 1, self.index],
        )
        self.parent.parent.api["row_count"] -= 1


class TableRows(base_classes.TableRows):
    def __init__(self, parent):
        self._parent = parent

    @property
    def parent(self):
        return self._parent

    @property
    def column_count(self):
        return self.parent.api.get("column_count", 0)

    def __len__(self):
        metadata = self.parent.api
        return max(
            0,
            metadata.get("row_count", 0)
            - bool(metadata.get("show_headers", True))
            - bool(metadata.get("show_totals", False)),
        )

    def __call__(self, key):
        if (
            not isinstance(key, numbers.Integral)
            or isinstance(key, bool)
            or not 1 <= key <= len(self)
        ):
            raise KeyError(key)
        return TableRow(self, key - 1)

    def __iter__(self):
        for index in range(len(self)):
            yield TableRow(self, index)

    async def get_count(self):
        if sys.platform != "emscripten":
            raise NotImplementedError("TableRows.get_count() requires xlwings Lite")
        import js

        count = int(
            await js.xlwings.getTableRowCount(
                self.parent.parent.name, self.parent.index - 1
            )
        )
        self.parent.api["row_count"] = (
            count
            + bool(self.parent.api.get("show_headers", True))
            + bool(self.parent.api.get("show_totals", False))
        )
        return count

    def add(self, values, index):
        position = len(self) if index is None else index
        self.parent.append_json_action(
            func="addTableRow",
            args=[self.parent.index - 1, position, values],
        )
        self.parent.api["row_count"] = self.parent.api.get("row_count", 0) + 1
        return TableRow(self, position)


class Tables(Collection, base_classes.Tables):
    _attr = "tables"
    _wrap = Table

    def append_json_action(self, **kwargs):
        self.parent.book.append_json_action(
            **{
                **kwargs,
                **{
                    "sheet_position": self.parent.index - 1,
                },
            }
        )

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
        self.append_json_action(
            func="addTable",
            args=[source.address, has_headers, table_style_name, name],
        )
        self.parent._api["tables"].append(
            {
                # What the caller asked for, so reading these back before the
                # next round-trip gives the requested values rather than
                # placeholders. An unnamed table gets its name from Excel, so
                # there's nothing to seed until the payload refreshes.
                "name": name if name else "",
                "range_address": source.address if source else None,
                "row_count": source.shape[0] if source else 0,
                "column_count": source.shape[1] if source else 0,
                "header_row_range_address": None,
                "data_body_range_address": None,
                "total_row_range_address": None,
                "show_headers": has_headers if has_headers is not None else True,
                "show_totals": False,
                "table_style": table_style_name if table_style_name else "",
                # Excel's defaults for a new table, so the getters work before
                # the next round-trip refreshes the payload.
                "show_autofilter": True,
                "show_table_style_first_column": False,
                "show_table_style_last_column": False,
                "show_table_style_row_stripes": True,
                "show_table_style_column_stripes": False,
            }
        )
        return Table(self.parent, len(self.parent.api["tables"]))


class FreezePanes(base_classes.FreezePanes):
    def __init__(self, sheet):
        self.sheet = sheet

    def append_json_action(self, **kwargs):
        self.sheet.book.append_json_action(
            **{
                **kwargs,
                **{
                    "sheet_position": self.sheet.index - 1,
                },
            }
        )

    def freeze_at(self, frozen_range):
        self.append_json_action(func="freezePaneAtRange", args=[frozen_range])

    def unfreeze(self):
        self.append_json_action(func="freezePaneUnfreeze")


class Chart(base_classes.Chart):
    def __init__(self, parent, key, pending=None):
        self._parent = parent
        self.key = key
        # A chart added via Charts.add() isn't created in Excel until it gets
        # source data, since Office.js' charts.add() requires it. `pending`
        # holds the geometry until then.
        self._pending = pending
        self._uses_default_name = pending is not None
        self._api = None if pending is not None else parent.api["charts"][key - 1]
        # Formatting written while pending, replayed in call order once the
        # chart exists (order matters: a legend position implies visibility).
        self._pending_actions = []

    def append_json_action(self, **kwargs):
        self.parent.book.append_json_action(
            **{
                **kwargs,
                **{
                    "sheet_position": self.parent.index - 1,
                },
            }
        )

    @property
    def api(self):
        return self._api if self._api is not None else self._pending

    @property
    def parent(self):
        return self._parent

    @property
    def index(self):
        if isinstance(self.key, numbers.Number):
            return self.key
        for ix, obj in enumerate(self.parent.api["charts"]):
            if obj["name"] == self.key:
                return ix + 1
        raise KeyError(self.key)

    @property
    def name(self):
        return self.api["name"]

    @name.setter
    def name(self, value):
        self.api["name"] = value
        self._uses_default_name = False
        if self._pending is None:
            self.append_json_action(func="setChartName", args=[self.index - 1, value])

    @property
    def chart_type(self):
        chart_type = self.api["chart_type"]
        return _CHART_TYPE_JS2PY.get(chart_type, chart_type)

    @chart_type.setter
    def chart_type(self, value):
        try:
            js_value = _CHART_TYPE_PY2JS[value]
        except KeyError:
            raise ValueError(
                f"Invalid chart type: {value!r}. Must be one of "
                f"{sorted(_CHART_TYPE_PY2JS)}."
            ) from None
        self.api["chart_type"] = js_value
        if self._pending is None:
            self.append_json_action(
                func="setChartType", args=[self.index - 1, js_value]
            )

    def _local_or_raise(self, key, what):
        """Local state written earlier in this script, or an error.

        The payload only carries a chart's name, type and geometry; everything
        else is known only once it has been set here.
        """
        if key in self.api:
            return self.api[key]
        raise NotImplementedError(
            f"Reading a chart's {what} isn't supported on this engine."
        )

    def _set_format(self, key, value, func, args):
        """Write through to the local state and queue (or buffer) the action."""
        self.api[key] = value
        if self._pending is not None:
            self._pending_actions.append((func, list(args)))
        else:
            self.append_json_action(func=func, args=[self.index - 1, *args])

    @property
    def title(self):
        return self._local_or_raise("title", "title")

    @title.setter
    def title(self, value):
        self._set_format("title", value, "setChartTitle", [value])

    @property
    def legend(self):
        return ChartLegend(self)

    @property
    def category_axis(self):
        return ChartAxis(self, "category")

    @property
    def value_axis(self):
        return ChartAxis(self, "value")

    @property
    def series(self):
        if self._pending is not None:
            raise XlwingsError(
                "Chart series require source data. Call Chart.set_source_data() "
                "and await book.flush() first."
            )
        return ChartSeriesCollection(self)

    async def get_series(self):
        if self._pending is not None:
            raise XlwingsError(
                "Chart series reads require source data. Call Chart.set_source_data() "
                "and await book.flush() first."
            )
        if sys.platform != "emscripten":
            raise NotImplementedError("get_series() is only supported in xlwings Lite")
        import js

        count = await js.xlwings.getChartSeriesCount(self.parent.name, self.index - 1)
        return ChartSeriesCollection(self, count=int(count))

    @property
    def plot_by(self):
        return self._local_or_raise("plot_by", "plot_by")

    @plot_by.setter
    def plot_by(self, value):
        # While pending, the orientation rides on the addChart action instead
        self.api["plot_by"] = value
        if self._pending is None:
            self.append_json_action(
                func="setChartPlotBy", args=[self.index - 1, _PLOT_BY_PY2JS[value]]
            )

    @property
    def style(self):
        return self._local_or_raise("style", "style")

    @style.setter
    def style(self, value):
        if self._pending is not None:
            # The style rides on addChart so creation doesn't need a redundant
            # formatting action. A later assignment replaces the default.
            self.api["style"] = value
        else:
            self._set_format("style", value, "setChartStyle", [value])

    def set_source_data(self, rng, plot_by=None):
        if self._pending is not None:
            # First source data: this is where the chart can finally be
            # created, since Office.js needs the type and data together.
            pending = self._pending
            # A name can be taken while this chart waits for source data.
            # Check before changing pending state so callers can rename and retry.
            if not self._uses_default_name and any(
                chart["name"] == pending["name"] for chart in self.parent.api["charts"]
            ):
                raise ShapeAlreadyExists(
                    f"'{pending['name']}' is already present on {self.parent.name}."
                )
            if plot_by is None:
                # a plot_by set while pending applies now
                plot_by = pending.get("plot_by")
            if plot_by is None:
                pending.pop("plot_by", None)
            else:
                pending["plot_by"] = plot_by
            anchor = pending.pop("anchor", None)
            if self._uses_default_name:
                pending["name"] = Charts.unique_default_name(self.parent.api["charts"])
            self.parent.api["charts"].append(pending)
            self.key = len(self.parent.api["charts"])
            self._api = pending
            self._pending = None
            self.append_json_action(
                func="addChart",
                args=[
                    pending["name"],
                    pending["chart_type"],
                    rng.sheet.name,
                    rng.address,
                    pending["left"],
                    pending["top"],
                    pending["width"],
                    pending["height"],
                    None if plot_by is None else _PLOT_BY_PY2JS[plot_by],
                    anchor,
                    pending.get("style"),
                ],
            )
            for func, args in self._pending_actions:
                self.append_json_action(func=func, args=[self.index - 1, *args])
            self._pending_actions = []
            return
        if plot_by is None:
            # Office.js' setData() defaults to Auto, so the cached orientation
            # is no longer known
            self.api.pop("plot_by", None)
            self.append_json_action(
                func="setChartSourceData",
                args=[self.index - 1, rng.sheet.name, rng.address],
            )
        else:
            self.api["plot_by"] = plot_by
            self.append_json_action(
                func="setChartSourceData",
                args=[
                    self.index - 1,
                    rng.sheet.name,
                    rng.address,
                    _PLOT_BY_PY2JS[plot_by],
                ],
            )

    def set_x_axis_values(self, rng):
        args = [rng.sheet.name, rng.address]
        if self._pending is not None:
            self._pending_actions.append(("setChartXAxisValues", args))
        else:
            self.append_json_action(
                func="setChartXAxisValues", args=[self.index - 1, *args]
            )

    def _set_position(self, attribute, value):
        self.api[attribute] = value
        if self._pending is None:
            self.append_json_action(
                func="setChartPosition", args=[self.index - 1, attribute, value]
            )

    @property
    def left(self):
        return self.api["left"]

    @left.setter
    def left(self, value):
        self._set_position("left", value)

    @property
    def top(self):
        return self.api["top"]

    @top.setter
    def top(self, value):
        self._set_position("top", value)

    @property
    def width(self):
        return self.api["width"]

    @width.setter
    def width(self, value):
        self._set_position("width", value)

    @property
    def height(self):
        return self.api["height"]

    @height.setter
    def height(self, value):
        self._set_position("height", value)

    def delete(self):
        if self._pending is not None:
            # Never created in Excel, so there's nothing to delete there.
            self._pending = None
            self._pending_actions = []
            return
        del self.parent.api["charts"][self.index - 1]
        self.append_json_action(func="deleteChart", args=[self.index - 1])

    async def get_png(self):
        """Fetch the chart as a base64-encoded PNG."""
        if sys.platform != "emscripten":
            raise NotImplementedError("get_png() is only supported in xlwings Lite")
        import js

        return await js.xlwings.getChartImage(self.parent.name, self.index - 1)

    def to_png(self, path):
        if self._pending is not None:
            raise XlwingsError(
                "Chart.to_png() requires source data. Call Chart.set_source_data() "
                "first."
            )
        # Like Range.to_png(), this queues an action that writes the file; it
        # lands when the script returns or on the next `await book.flush()`.
        self.append_json_action(func="chartToPng", args=[self.index - 1, path])

    def to_pdf(self, path, quality):
        raise NotImplementedError(
            "Chart.to_pdf() is not supported on this engine, which has no PDF export."
        )


class ChartSeries(base_classes.ChartSeries):
    def __init__(self, parent, key):
        self.parent = parent
        self.key = key

    @property
    def index(self):
        return self.key

    @property
    def api(self):
        raise NotImplementedError(
            "ChartSeries.api isn't available on this engine: there is no native "
            "series object, only queued actions."
        )

    def _state_key(self, attribute):
        return f"series_{self.index}_{attribute}"

    def _local_or_raise(self, attribute):
        try:
            return self.parent._local_or_raise(
                self._state_key(attribute), f"series {self.index} {attribute}"
            )
        except NotImplementedError:
            raise NotImplementedError(
                "Reading chart series attributes synchronously isn't supported "
                f"on this engine. Use 'await series.get_{attribute}()' to fetch "
                "the value."
            ) from None

    @property
    def name(self):
        return self._local_or_raise("name")

    @name.setter
    def name(self, value):
        self.set(name=value)

    @property
    def marker_style(self):
        return self._local_or_raise("marker_style")

    @marker_style.setter
    def marker_style(self, value):
        self.set(marker_style=value)

    @property
    def marker_size(self):
        return self._local_or_raise("marker_size")

    @marker_size.setter
    def marker_size(self, value):
        self.set(marker_size=value)

    @property
    def marker_foreground_color(self):
        return self._local_or_raise("marker_foreground_color")

    @marker_foreground_color.setter
    def marker_foreground_color(self, value):
        self.set(marker_foreground_color=value)

    @property
    def marker_background_color(self):
        return self._local_or_raise("marker_background_color")

    @marker_background_color.setter
    def marker_background_color(self, value):
        self.set(marker_background_color=value)

    @property
    def line_color(self):
        return self._local_or_raise("line_color")

    @line_color.setter
    def line_color(self, value):
        self.set(line_color=value)

    @property
    def fill_color(self):
        return self._local_or_raise("fill_color")

    @fill_color.setter
    def fill_color(self, value):
        self.set(fill_color=value)

    def set(
        self,
        *,
        name=base_classes._UNSET,
        marker_style=base_classes._UNSET,
        marker_size=base_classes._UNSET,
        marker_foreground_color=base_classes._UNSET,
        marker_background_color=base_classes._UNSET,
        line_color=base_classes._UNSET,
        fill_color=base_classes._UNSET,
    ):
        values = {}
        local_values = {}
        for attribute, value in (
            ("name", name),
            ("marker_style", marker_style),
            ("marker_size", marker_size),
            ("marker_foreground_color", marker_foreground_color),
            ("marker_background_color", marker_background_color),
            ("line_color", line_color),
            ("fill_color", fill_color),
        ):
            if value is base_classes._UNSET:
                continue
            local_values[attribute] = value
            if attribute == "marker_style":
                values[attribute] = _MARKER_STYLE_PY2JS[value]
            elif attribute.endswith("_color"):
                values[attribute] = _color_to_hex(value)
            else:
                values[attribute] = value
        if not values:
            return
        for attribute, value in local_values.items():
            self.parent.api[self._state_key(attribute)] = value
        self.parent.append_json_action(
            func="setChartSeries",
            args=[self.parent.index - 1, self.index - 1, values],
        )

    async def _get_series_data(self, key):
        if sys.platform != "emscripten":
            raise NotImplementedError(f"get_{key}() is only supported in xlwings Lite")
        import js
        from pyodide.ffi import to_js

        data_js = await js.xlwings.getChartSeriesData(
            self.parent.parent.name,
            self.parent.index - 1,
            self.index - 1,
            to_js([key]),
        )
        value = _normalize_jsnull(data_js.to_py())[key]
        if key == "marker_style":
            return _MARKER_STYLE_JS2PY.get(value, value)
        if key == "marker_size" and value is not None:
            return int(value)
        if key.endswith("_color"):
            return utils.hex_to_rgb(value) if value else None
        return value

    async def get_name(self):
        return await self._get_series_data("name")

    async def get_marker_style(self):
        return await self._get_series_data("marker_style")

    async def get_marker_size(self):
        return await self._get_series_data("marker_size")

    async def get_marker_foreground_color(self):
        return await self._get_series_data("marker_foreground_color")

    async def get_marker_background_color(self):
        return await self._get_series_data("marker_background_color")

    async def get_line_color(self):
        return await self._get_series_data("line_color")

    async def get_fill_color(self):
        return await self._get_series_data("fill_color")


class ChartSeriesCollection(base_classes.ChartSeriesCollection):
    def __init__(self, parent, count=None):
        self._parent = parent
        self._count = count

    @property
    def parent(self):
        return self._parent

    @property
    def api(self):
        raise NotImplementedError(
            "ChartSeriesCollection.api isn't available on this engine."
        )

    def _loaded_count(self):
        if self._count is None:
            raise NotImplementedError(
                "Inspecting chart series synchronously isn't supported on this "
                "engine. Use 'await chart.get_series()' to fetch the collection."
            )
        return self._count

    def __call__(self, key):
        count = self._loaded_count()
        if (
            not isinstance(key, numbers.Integral)
            or isinstance(key, bool)
            or key < 1
            or key > count
        ):
            raise KeyError(key)
        return ChartSeries(self.parent, int(key))

    def __len__(self):
        return self._loaded_count()

    def __iter__(self):
        for key in range(1, self._loaded_count() + 1):
            yield ChartSeries(self.parent, key)

    def __contains__(self, key):
        return (
            isinstance(key, numbers.Integral)
            and not isinstance(key, bool)
            and 1 <= key <= self._loaded_count()
        )


class ChartAxis(base_classes.ChartAxis):
    def __init__(self, parent, axis_type):
        self.parent = parent
        self.axis_type = axis_type

    @property
    def api(self):
        raise NotImplementedError(
            "ChartAxis.api isn't available on this engine: there is no native "
            "axis object, only queued actions."
        )

    def _state_key(self, attribute):
        return f"{self.axis_type}_axis_{attribute}"

    def _local_or_raise(self, attribute):
        return self.parent._local_or_raise(
            self._state_key(attribute),
            f"primary {self.axis_type} axis {attribute}",
        )

    def _queue(self, values):
        args = [self.axis_type, values]
        if self.parent._pending is not None:
            self.parent._pending_actions.append(("setChartAxis", args))
        else:
            self.parent.append_json_action(
                func="setChartAxis", args=[self.parent.index - 1, *args]
            )

    def _sync_read_error(self, getter):
        return NotImplementedError(
            "Reading chart axis attributes synchronously isn't supported on this "
            f"engine. Use 'await chart.{self.axis_type}_axis.{getter}()' to fetch "
            "the value."
        )

    @property
    def title(self):
        if self.parent.api.get(self._state_key("visible")) is False:
            return None
        try:
            return self._local_or_raise("title")
        except NotImplementedError:
            raise self._sync_read_error("get_title") from None

    @title.setter
    def title(self, value):
        self.set(title=value)

    @property
    def minimum_scale(self):
        try:
            return self._local_or_raise("minimum_scale")
        except NotImplementedError:
            raise self._sync_read_error("get_minimum_scale") from None

    @minimum_scale.setter
    def minimum_scale(self, value):
        self.set(minimum_scale=value)

    @property
    def maximum_scale(self):
        try:
            return self._local_or_raise("maximum_scale")
        except NotImplementedError:
            raise self._sync_read_error("get_maximum_scale") from None

    @maximum_scale.setter
    def maximum_scale(self, value):
        self.set(maximum_scale=value)

    @property
    def major_unit(self):
        try:
            return self._local_or_raise("major_unit")
        except NotImplementedError:
            raise self._sync_read_error("get_major_unit") from None

    @major_unit.setter
    def major_unit(self, value):
        self.set(major_unit=value)

    @property
    def number_format(self):
        try:
            return self._local_or_raise("number_format")
        except NotImplementedError:
            raise self._sync_read_error("get_number_format") from None

    @number_format.setter
    def number_format(self, value):
        self.set(number_format=value)

    @property
    def visible(self):
        try:
            return self._local_or_raise("visible")
        except NotImplementedError:
            raise self._sync_read_error("get_visible") from None

    @visible.setter
    def visible(self, value):
        self.set(visible=value)

    def set(
        self,
        *,
        title=base_classes._UNSET,
        minimum_scale=base_classes._UNSET,
        maximum_scale=base_classes._UNSET,
        major_unit=base_classes._UNSET,
        number_format=base_classes._UNSET,
        visible=base_classes._UNSET,
    ):
        values = {}
        for attribute, value in (
            ("title", title),
            ("minimum_scale", minimum_scale),
            ("maximum_scale", maximum_scale),
            ("major_unit", major_unit),
            ("number_format", number_format),
            ("visible", visible),
        ):
            if value is base_classes._UNSET:
                continue
            values[attribute] = value
            key = self._state_key(attribute)
            if (
                attribute in {"minimum_scale", "maximum_scale", "major_unit"}
                and value is None
            ):
                # None restores Excel's automatic value, whose resolved numeric
                # result isn't known until a synchronized read.
                self.parent.api.pop(key, None)
            else:
                self.parent.api[key] = value
        if title is not base_classes._UNSET and title is not None:
            self.parent.api[self._state_key("visible")] = True
        if visible is False:
            self.parent.api[self._state_key("visible")] = False
        if values:
            self._queue(values)

    async def _get_axis_data(self, key):
        if self.parent._pending is not None:
            raise XlwingsError(
                "Chart axis reads require source data. Call Chart.set_source_data() "
                "and await book.flush() first."
            )
        if sys.platform != "emscripten":
            raise NotImplementedError(f"get_{key}() is only supported in xlwings Lite")
        import js
        from pyodide.ffi import to_js

        data_js = await js.xlwings.getChartAxisData(
            self.parent.parent.name,
            self.parent.index - 1,
            self.axis_type,
            to_js([key]),
        )
        return _normalize_jsnull(data_js.to_py())[key]

    async def get_title(self):
        return await self._get_axis_data("title")

    async def get_minimum_scale(self):
        return await self._get_axis_data("minimum_scale")

    async def get_maximum_scale(self):
        return await self._get_axis_data("maximum_scale")

    async def get_major_unit(self):
        return await self._get_axis_data("major_unit")

    async def get_number_format(self):
        return await self._get_axis_data("number_format")

    async def get_visible(self):
        return await self._get_axis_data("visible")


class ChartLegend(base_classes.ChartLegend):
    def __init__(self, parent):
        self.parent = parent

    @property
    def api(self):
        raise NotImplementedError(
            "ChartLegend.api isn't available on this engine: there is no native "
            "legend object, only queued actions."
        )

    @property
    def visible(self):
        return self.parent._local_or_raise("legend_visible", "legend visibility")

    @visible.setter
    def visible(self, value):
        value = bool(value)
        if not value:
            # hiding forgets the position; showing again doesn't restore it
            self.parent.api.pop("legend_position", None)
        self.parent._set_format(
            "legend_visible", value, "setChartLegend", ["visible", value]
        )

    @property
    def position(self):
        if self.parent.api.get("legend_visible") is False:
            return None
        return self.parent._local_or_raise("legend_position", "legend position")

    @position.setter
    def position(self, value):
        # setting a position shows the legend, on the JS side too
        self.parent.api["legend_visible"] = True
        self.parent._set_format(
            "legend_position",
            value,
            "setChartLegend",
            ["position", _LEGEND_POSITION_PY2JS[value]],
        )


class Charts(Collection, base_classes.Charts):
    _attr = "charts"
    _wrap = Chart

    @staticmethod
    def unique_default_name(charts):
        existing = {chart["name"].casefold() for chart in charts}
        name = "Chart"
        suffix = 2
        while name.casefold() in existing:
            name = f"Chart {suffix}"
            suffix += 1
        return name

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
        # Office.js' charts.add() needs a type and source data, which xlwings
        # doesn't necessarily have yet at this point -- so hold the geometry and
        # create the chart on the first set_source_data().
        chart_type = chart_type or "column_clustered"
        try:
            js_type = _CHART_TYPE_PY2JS[chart_type]
        except KeyError:
            raise ValueError(
                f"Invalid chart type: {chart_type!r}. Must be one of "
                f"{sorted(_CHART_TYPE_PY2JS)}."
            ) from None
        pending = {
            "name": name if name else self.unique_default_name(self.api),
            "chart_type": js_type,
            # Range.left/top aren't readable synchronously on this engine, so
            # the anchor's address goes to the client instead
            "left": None if anchor is not None else left,
            "top": None if anchor is not None else top,
            "width": width,
            "height": height,
        }
        if style is not None:
            pending["style"] = style
        if anchor is not None:
            pending["anchor"] = anchor.address
        chart = Chart(self.parent, len(self.api) + 1, pending=pending)
        chart._uses_default_name = name is None
        if source is not None:
            chart.set_source_data(source, plot_by)
        return chart


def _quote_sheet_name(name):
    return "'" + name.replace("'", "''") + "'"


class PivotTable(base_classes.PivotTable):
    def __init__(self, parent, key):
        self._parent = parent
        # Hold on to the payload dict itself: the index is resolved afresh
        # for every action so that deleting an earlier pivot table doesn't
        # invalidate this wrapper.
        self._api = parent.api["pivot_tables"][key - 1]

    def append_json_action(self, **kwargs):
        self.parent.book.append_json_action(
            **{
                **kwargs,
                **{
                    "sheet_position": self.parent.index - 1,
                },
            }
        )

    @property
    def api(self):
        return self._api

    @property
    def parent(self):
        return self._parent

    @property
    def index(self):
        for ix, entry in enumerate(self.parent.api["pivot_tables"]):
            if entry is self._api:
                return ix + 1
        raise KeyError("The pivot table has been deleted.")

    def _queue(self, func, *args):
        self.append_json_action(func=func, args=[self.index - 1, *args])

    @property
    def name(self):
        return self.api["name"]

    @name.setter
    def name(self, value):
        self._queue("setPivotTableName", value)
        self.api["name"] = value

    @property
    def field_names(self):
        return list(self.api["field_names"])

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
        layout = self.api.get("layout")
        # Office.js reports null when the row fields use mixed layouts
        return None if layout is None else _PIVOT_LAYOUT_JS2PY.get(layout, layout)

    @layout.setter
    def layout(self, value):
        js_value = _PIVOT_LAYOUT_PY2JS[value]
        self._queue("setPivotLayout", "layout", js_value)
        self.api["layout"] = js_value

    @property
    def show_row_grand_totals(self):
        return self.api["show_row_grand_totals"]

    @show_row_grand_totals.setter
    def show_row_grand_totals(self, value):
        self._queue("setPivotLayout", "show_row_grand_totals", value)
        self.api["show_row_grand_totals"] = value

    @property
    def show_column_grand_totals(self):
        return self.api["show_column_grand_totals"]

    @show_column_grand_totals.setter
    def show_column_grand_totals(self, value):
        self._queue("setPivotLayout", "show_column_grand_totals", value)
        self.api["show_column_grand_totals"] = value

    @property
    def range(self):
        raise NotImplementedError(
            "PivotTable.range isn't supported on this engine: the payload doesn't "
            "carry the pivot table's range."
        )

    @property
    def data_body_range(self):
        raise NotImplementedError(
            "PivotTable.data_body_range isn't supported on this engine: the "
            "payload doesn't carry the pivot table's range."
        )

    def refresh(self):
        self._queue("refreshPivotTable")

    def delete(self):
        ix = self.index - 1
        self._queue("deletePivotTable")
        del self.parent.api["pivot_tables"][ix]


class PivotField(base_classes.PivotField):
    def __init__(self, pivot, name):
        self._pivot = pivot
        self._name = name

    @property
    def api(self):
        raise NotImplementedError(
            "PivotField.api isn't available on this engine: there is no native "
            "field object, only queued actions."
        )

    @property
    def parent(self):
        return self._pivot

    @property
    def name(self):
        return self._name

    def _area(self):
        """The area the field is currently in, or None."""
        for area in _PIVOT_AREAS:
            if self._name in self._pivot.api[area]:
                return area
        return None

    def remove(self):
        area = self._area()
        if area is None:
            return
        self._pivot._queue("removePivotField", area, self._name)
        self._pivot.api[area].remove(self._name)


class PivotFields(base_classes.PivotFields):
    def __init__(self, pivot, area):
        self._pivot = pivot
        self._area = area

    @property
    def api(self):
        # the list of field names in this area, in order
        return self._pivot.api[self._area]

    @property
    def parent(self):
        return self._pivot

    @property
    def area(self):
        return self._area

    def __call__(self, key):
        if isinstance(key, numbers.Number):
            if key < 1 or key > len(self):
                raise KeyError(key)
            return PivotField(self._pivot, self.api[key - 1])
        if key not in self.api:
            raise KeyError(key)
        return PivotField(self._pivot, key)

    def __len__(self):
        return len(self.api)

    def __iter__(self):
        for name in list(self.api):
            yield PivotField(self._pivot, name)

    def __contains__(self, key):
        if isinstance(key, numbers.Number):
            return 1 <= key <= len(self)
        return key in self.api

    def add(self, name):
        if name in self.api:
            # already here: Excel would leave it in place, too
            return PivotField(self._pivot, name)
        # Office.js moves a hierarchy off another axis when adding it here
        for area in _PIVOT_AREAS:
            if name in self._pivot.api[area]:
                self._pivot.api[area].remove(name)
        self._pivot._queue("addPivotField", self._area, name)
        self.api.append(name)
        return PivotField(self._pivot, name)


class PivotValueField(base_classes.PivotValueField):
    def __init__(self, pivot, entry):
        self._pivot = pivot
        # the payload dict of this value field; the position is resolved
        # afresh for every action
        self._entry = entry

    @property
    def api(self):
        return self._entry

    @property
    def parent(self):
        return self._pivot

    @property
    def index(self):
        for ix, entry in enumerate(self._pivot.api["values"]):
            if entry is self._entry:
                return ix + 1
        raise KeyError("The value field has been removed.")

    def _queue(self, attribute, value):
        self._pivot._queue("setPivotValueField", self.index - 1, attribute, value)

    def _local_or_raise(self, key, what):
        if key in self._entry:
            return self._entry[key]
        raise NotImplementedError(
            f"Reading a value field's {what} isn't supported on this engine "
            "unless the value field already existed when the script started "
            "or it was set in this script."
        )

    @property
    def name(self):
        return self._local_or_raise("name", "name")

    @name.setter
    def name(self, value):
        self._queue("name", value)
        self._entry["name"] = value
        self._entry["_name_explicit"] = True

    @property
    def source_field(self):
        return self._entry["source_field"]

    @property
    def function(self):
        function = self._entry.get("function")
        if function in (None, "Automatic", "Unknown"):
            return None
        return _PIVOT_FUNCTION_JS2PY.get(function, function)

    @function.setter
    def function(self, value):
        js_value = _PIVOT_FUNCTION_PY2JS[value]
        self._queue("function", js_value)
        self._entry["function"] = js_value
        if not self._entry.get("_name_explicit"):
            # Excel renames an automatic caption ("Sum of X" -> "Count of X")
            # along with the function, so the local name is no longer known
            self._entry.pop("name", None)

    @property
    def number_format(self):
        return self._local_or_raise("number_format", "number format")

    @number_format.setter
    def number_format(self, value):
        self._queue("number_format", value)
        self._entry["number_format"] = value

    def remove(self):
        ix = self.index - 1
        self._pivot._queue("removePivotValueField", ix)
        del self._pivot.api["values"][ix]


class PivotValueFields(base_classes.PivotValueFields):
    def __init__(self, pivot):
        self._pivot = pivot

    @property
    def api(self):
        return self._pivot.api["values"]

    @property
    def parent(self):
        return self._pivot

    def __call__(self, key):
        if isinstance(key, numbers.Number):
            if key < 1 or key > len(self):
                raise KeyError(key)
            return PivotValueField(self._pivot, self.api[key - 1])
        for entry in self.api:
            if entry.get("name") == key:
                return PivotValueField(self._pivot, entry)
        raise KeyError(key)

    def __len__(self):
        return len(self.api)

    def __iter__(self):
        for entry in list(self.api):
            yield PivotValueField(self._pivot, entry)

    def __contains__(self, key):
        if isinstance(key, numbers.Number):
            return 1 <= key <= len(self)
        return any(entry.get("name") == key for entry in self.api)

    def add(self, field, function=None, name=None, number_format=None):
        js_function = None if function is None else _PIVOT_FUNCTION_PY2JS[function]
        self._pivot._queue(
            "addPivotValueField", field, js_function, name, number_format
        )
        # Only what the caller asked for: Excel picks the caption (and the
        # function) otherwise, and those aren't known until the next payload.
        entry = {"source_field": field, "function": js_function}
        if name is not None:
            entry["name"] = name
            entry["_name_explicit"] = True
        if number_format is not None:
            entry["number_format"] = number_format
        self.api.append(entry)
        return PivotValueField(self._pivot, entry)


class PivotTables(Collection, base_classes.PivotTables):
    _attr = "pivot_tables"
    _wrap = PivotTable

    def append_json_action(self, **kwargs):
        self.parent.book.append_json_action(
            **{
                **kwargs,
                **{
                    "sheet_position": self.parent.index - 1,
                },
            }
        )

    def unique_default_name(self):
        # Office.js requires a name, and Excel numbers pivot tables across
        # the whole workbook
        existing = {
            pt["name"].casefold()
            for sheet in self.parent.book.api["sheets"]
            for pt in (sheet.get("pivot_tables") or [])
        }
        number = 1
        while f"pivottable{number}" in existing:
            number += 1
        return f"PivotTable{number}"

    def add(self, source, destination, name=None):
        name = name if name else self.unique_default_name()
        if isinstance(source, Table):
            source_kind, source_ref = "table", source.name
        else:
            # Office.js accepts a sheet-qualified address string as source
            source_kind = "range"
            source_ref = f"{_quote_sheet_name(source.sheet.name)}!{source.address}"
        self.append_json_action(
            func="addPivotTable",
            args=[name, source_kind, source_ref, destination.address],
        )
        self.api.append(
            {
                "name": name,
                # The source fields aren't known until the next payload
                "field_names": [],
                "rows": [],
                "columns": [],
                "filters": [],
                "values": [],
                # Excel's defaults for a new pivot table
                "layout": "Compact",
                "show_row_grand_totals": True,
                "show_column_grand_totals": True,
            }
        )
        return PivotTable(self.parent, len(self.api))


class Characters(base_classes.Characters):
    """A character slice of a shape's text.

    Office.js exposes character ranges through `TextRange.getSubstring()`,
    which only shapes have --- `Range.characters` raises, see there.
    """

    def __init__(self, parent, start=None, length=None):
        self.parent = parent
        self.start = start
        self.length = length

    @property
    def api(self):
        raise NotImplementedError(
            "Characters.api isn't available on this engine, which addresses "
            "shapes by index rather than holding a native object."
        )

    @property
    def text(self):
        raise NotImplementedError(
            "Reading characters synchronously isn't supported on this engine. "
            "Use 'await mycharacters.get_text()' to fetch it on demand."
        )

    async def get_text(self):
        return await self.parent._get_shape_data(
            "characters_text", start=self.start, length=self.length
        )

    @property
    def font(self):
        return Font(self, None)

    def __getitem__(self, item):
        if isinstance(item, slice):
            start = item.start or 0
            length = None if item.stop is None else item.stop - start
            return Characters(self.parent, start=start, length=length)
        return Characters(self.parent, start=item, length=1)


class ConditionalFormat(base_classes.ConditionalFormat):
    def __init__(self, parent, entry):
        self.parent = parent
        self._entry = entry

    @property
    def api(self):
        return self._entry

    @property
    def type(self):
        return _CONDITIONAL_FORMAT_TYPE_JS2PY.get(self._entry.get("type"), "unknown")

    @property
    def stop_if_true(self):
        if self.type in {"color_scale", "data_bar", "icon_set"}:
            return None
        return self._entry.get("stop_if_true")

    @property
    def operator(self):
        if self.type != "cell_value":
            return None
        return _CONDITIONAL_FORMAT_OPERATOR_JS2PY.get(self._entry.get("operator"))

    @property
    def formula1(self):
        return self._entry.get("formula1") if self.type == "cell_value" else None

    @property
    def formula2(self):
        if self.type != "cell_value" or self.operator not in {
            "between",
            "not_between",
        }:
            return None
        return self._entry.get("formula2")

    @property
    def formula(self):
        return self._entry.get("formula") if self.type == "custom" else None

    @staticmethod
    def _rgb(value):
        return utils.hex_to_rgb(value) if value else None

    @property
    def fill_color(self):
        return self._rgb(self._entry.get("fill_color"))

    @property
    def font_color(self):
        return self._rgb(self._entry.get("font_color"))

    @property
    def font_bold(self):
        return self._entry.get("font_bold")

    @property
    def font_italic(self):
        return self._entry.get("font_italic")

    @property
    def colors(self):
        if self.type != "color_scale":
            return None
        return tuple(self._rgb(value) for value in self._entry.get("colors", ()))

    @property
    def bar_color(self):
        return (
            self._rgb(self._entry.get("bar_color")) if self.type == "data_bar" else None
        )

    @property
    def gradient(self):
        return self._entry.get("gradient") if self.type == "data_bar" else None

    @property
    def show_value(self):
        if self.type not in {"data_bar", "icon_set"}:
            return None
        return self._entry.get("show_value")

    @property
    def icon_set(self):
        if self.type != "icon_set":
            return None
        return _CONDITIONAL_FORMAT_ICON_SET_JS2PY.get(self._entry.get("icon_set"))

    @property
    def reverse_order(self):
        return self._entry.get("reverse_order") if self.type == "icon_set" else None

    @property
    def threshold_types(self):
        if self.type not in {"color_scale", "data_bar", "icon_set"}:
            return None
        return tuple(
            _CONDITIONAL_FORMAT_THRESHOLD_JS2PY.get(value, "unknown")
            for value in self._entry.get("threshold_types", ())
        )

    @property
    def thresholds(self):
        if self.type not in {"color_scale", "data_bar", "icon_set"}:
            return None
        return tuple(self._entry.get("thresholds", ()))

    def _snapshot(self):
        keys = (
            "type",
            "stop_if_true",
            "operator",
            "formula1",
            "formula2",
            "formula",
            "fill_color",
            "font_color",
            "font_bold",
            "font_italic",
            "colors",
            "bar_color",
            "gradient",
            "show_value",
            "icon_set",
            "reverse_order",
            "threshold_types",
            "thresholds",
        )
        snapshot = {}
        for key in keys:
            if key not in self._entry:
                continue
            value = self._entry[key]
            if isinstance(value, (list, tuple)):
                value = list(value)
            if key == "thresholds":
                value = [None if item is None else str(item) for item in value]
            snapshot[key] = value
        return snapshot

    def set(self, changes):
        position = self.parent.position(self._entry)
        expected = self._snapshot()
        serialized = dict(changes)
        if "operator" in serialized:
            serialized["operator"] = _CONDITIONAL_FORMAT_OPERATOR_PY2JS[
                serialized["operator"]
            ]
        for key in ("fill_color", "font_color"):
            if key in serialized:
                serialized[key] = _color_to_hex(serialized[key])
        self.parent.range.append_json_action(
            func="setConditionalFormat",
            args=[position, expected, serialized],
        )
        self._entry.update(serialized)

    def delete(self):
        try:
            position = self.parent.position(self._entry)
        except ValueError:
            raise XlwingsError(
                "This conditional-format rule is no longer in its collection."
            ) from None
        self.parent.range.append_json_action(
            func="deleteConditionalFormat",
            args=[position, self._snapshot()],
        )
        # Keep this loaded snapshot aligned with the actions already queued so
        # deleting several objects from it continues to target the right index.
        self.parent.remove(self._entry)


class ConditionalFormats(base_classes.ConditionalFormats):
    def __init__(self, range, entries=None):
        self.range = range
        self._api = entries
        self._pending = []

    @property
    def api(self):
        return self._api

    @property
    def parent(self):
        return self.range

    def _loaded(self):
        if self._api is None:
            raise NotImplementedError(
                "Inspecting conditional formats synchronously isn't supported on "
                "this engine. Use 'await myrange.get_conditional_formats()' to "
                "fetch them on demand."
            )
        return self._api

    def __call__(self, key):
        entries = self._loaded()
        if not isinstance(key, numbers.Number) or key < 1 or key > len(entries):
            raise KeyError(key)
        return ConditionalFormat(self, entries[key - 1])

    def __len__(self):
        return len(self._loaded())

    def __iter__(self):
        # Iterate over a copy so deleting the current rule doesn't skip the next.
        for entry in list(self._loaded()):
            yield ConditionalFormat(self, entry)

    def __contains__(self, key):
        return isinstance(key, numbers.Number) and 1 <= key <= len(self._loaded())

    def position(self, entry):
        entries = self._api if self._api is not None else self._pending
        return entries.index(entry)

    def remove(self, entry):
        entries = self._api if self._api is not None else self._pending
        entries.remove(entry)

    @staticmethod
    def _entry(rule_type, spec):
        entry = {
            "type": rule_type,
            "stop_if_true": spec.get("stop_if_true"),
            "fill_color": None,
            "font_color": None,
            "font_bold": None,
            "font_italic": None,
        }
        entry.update(spec)
        if "operator" in entry:
            entry["operator"] = _CONDITIONAL_FORMAT_OPERATOR_PY2JS[entry["operator"]]
        for key in ("fill_color", "font_color"):
            if entry.get(key) is not None:
                entry[key] = _color_to_hex(entry[key])
        if "colors" in entry:
            entry["colors"] = [_color_to_hex(value) for value in entry["colors"]]
        if entry.get("bar_color") is not None:
            entry["bar_color"] = _color_to_hex(entry["bar_color"])
        if "threshold_types" in entry:
            entry["threshold_types"] = [
                _CONDITIONAL_FORMAT_THRESHOLD_PY2JS[value]
                for value in entry["threshold_types"]
            ]
        if "thresholds" in entry:
            entry["thresholds"] = list(entry["thresholds"])
        if "icon_set" in entry:
            entry["icon_set"] = _CONDITIONAL_FORMAT_ICON_SET_PY2JS[entry["icon_set"]]
        return entry

    def _add(self, rule_type, spec):
        entry = self._entry(rule_type, spec)
        self.range.append_json_action(func="addConditionalFormat", args=[entry])
        entries = self._api if self._api is not None else self._pending
        entries.insert(0, entry)
        return ConditionalFormat(self, entry)

    def add_cell_value(self, spec):
        return self._add("CellValue", spec)

    def add_custom(self, spec):
        return self._add("Custom", spec)

    def add_color_scale(self, spec):
        return self._add("ColorScale", spec)

    def add_data_bar(self, spec):
        return self._add("DataBar", spec)

    def add_icon_set(self, spec):
        return self._add("IconSet", spec)

    def clear(self):
        self.range.append_json_action(func="clearConditionalFormats", args=[])
        if self._api is not None:
            self._api.clear()
        self._pending.clear()


class Comment(base_classes.Comment):
    def __init__(self, range, comment_id=None):
        self.range = range
        self.comment_id = comment_id

    @property
    def api(self):
        raise NotImplementedError("Comment.api is not available in xlwings Lite")

    async def _read(self, key):
        if sys.platform != "emscripten":
            raise NotImplementedError(f"get_{key}() is only supported in xlwings Lite")
        import js

        value = await js.xlwings.getCommentData(
            self.range.sheet.name, self.comment_id, self.range.address, key
        )
        return _normalize_jsnull(value.to_py() if hasattr(value, "to_py") else value)

    @property
    def text(self):
        raise NotImplementedError("Use 'await comment.get_text()' in xlwings Lite")

    @text.setter
    def text(self, value):
        self.range.append_json_action(
            func="setCommentText", args=[self.comment_id, self.range.address, value]
        )

    async def get_text(self):
        return await self._read("text")

    @property
    def author(self):
        raise NotImplementedError("Use 'await comment.get_author()' in xlwings Lite")

    async def get_author(self):
        return await self._read("author")

    @property
    def creation_date(self):
        raise NotImplementedError(
            "Use 'await comment.get_creation_date()' in xlwings Lite"
        )

    async def get_creation_date(self):
        value = await self._read("creation_date")
        return (
            dt.datetime.fromisoformat(value.replace("Z", "+00:00")) if value else None
        )

    @property
    def resolved(self):
        raise NotImplementedError("Use 'await comment.get_resolved()' in xlwings Lite")

    async def get_resolved(self):
        return await self._read("resolved")

    def set_resolved(self, value):
        self.range.append_json_action(
            func="setCommentResolved", args=[self.comment_id, self.range.address, value]
        )

    @property
    def location(self):
        raise NotImplementedError("Use 'await comment.get_location()' in xlwings Lite")

    async def get_location(self):
        value = await self._read("location")
        if value is None:
            return None
        sheet = self.range.sheet.book.sheets(value["sheet"])
        return Range(sheet=sheet, arg1=value["address"])

    @property
    def replies(self):
        raise NotImplementedError("Use 'await comment.get_replies()' in xlwings Lite")

    async def get_replies(self):
        if sys.platform != "emscripten":
            raise NotImplementedError("get_replies() is only supported in xlwings Lite")
        import js

        entries = _normalize_jsnull(
            (
                await js.xlwings.getCommentReplies(
                    self.range.sheet.name, self.comment_id, self.range.address
                )
            ).to_py()
        )
        return [CommentReply(self, entry["id"]) for entry in entries]

    def add_reply(self, text):
        self.range.append_json_action(
            func="addCommentReply", args=[self.comment_id, self.range.address, text]
        )

    def delete(self):
        self.range.append_json_action(
            func="deleteComment", args=[self.comment_id, self.range.address]
        )


class CommentReply(base_classes.CommentReply):
    def __init__(self, comment, reply_id):
        self.comment = comment
        self.reply_id = reply_id

    @property
    def text(self):
        raise NotImplementedError("Use 'await reply.get_text()' in xlwings Lite")

    async def get_text(self):
        if sys.platform != "emscripten":
            raise NotImplementedError("get_text() is only supported in xlwings Lite")
        import js

        return _normalize_jsnull(
            await js.xlwings.getCommentReplyText(
                self.comment.range.sheet.name,
                self.comment.comment_id,
                self.comment.range.address,
                self.reply_id,
            )
        )


class Note(base_classes.Note):
    def __init__(self, range):
        self.range = range

    @property
    def api(self):
        return self._entry

    @property
    def _entry(self):
        """This note's entry in the sheet's notes payload, keyed by address."""
        for note in self.range.sheet.api.get("notes") or []:
            if _address_key(note["address"]) == _address_key(self.range.address):
                return note
        return None

    @property
    def text(self):
        # The payload says which cells have a note, not what they say: note
        # text is unbounded, so sending it would put every note's full text in
        # every request.
        raise NotImplementedError(
            "Reading a note's text synchronously isn't supported on this "
            "engine. Use 'await mynote.get_text()' to fetch it on demand."
        )

    async def get_text(self):
        if sys.platform != "emscripten":
            raise NotImplementedError("get_text() is only supported in xlwings Lite")
        import js

        return _normalize_jsnull(
            await js.xlwings.getNoteText(self.range.sheet.name, self.range.address)
        )

    @property
    def author(self):
        raise NotImplementedError(
            "Reading a note's author synchronously isn't supported on this "
            "engine. Use 'await mynote.get_author()'."
        )

    async def get_author(self):
        if sys.platform != "emscripten":
            raise NotImplementedError("get_author() is only supported in xlwings Lite")
        import js

        return _normalize_jsnull(
            await js.xlwings.getNoteAuthor(self.range.sheet.name, self.range.address)
        )

    @property
    def location(self):
        raise NotImplementedError(
            "Reading a note's location synchronously isn't supported on this "
            "engine. Use 'await mynote.get_location()'."
        )

    async def get_location(self):
        if sys.platform != "emscripten":
            raise NotImplementedError(
                "get_location() is only supported in xlwings Lite"
            )
        import js

        address = _normalize_jsnull(
            await js.xlwings.getNoteLocation(self.range.sheet.name, self.range.address)
        )
        return Range(sheet=self.range.sheet, arg1=address) if address else None

    @text.setter
    def text(self, value):
        self.range.append_json_action(
            func="setNoteText", args=[self.range.address, value]
        )

    def delete(self):
        notes = self.range.sheet.api.get("notes") or []
        entry = self._entry
        if entry is not None:
            notes.remove(entry)
        self.range.append_json_action(func="deleteNote", args=[self.range.address])


class Shape(base_classes.Shape):
    def __init__(self, parent, key):
        self._parent = parent
        self._api = self.parent.api["shapes"][key - 1]
        self.key = key

    def append_json_action(self, **kwargs):
        self.parent.book.append_json_action(
            **{
                **kwargs,
                **{
                    "sheet_position": self.parent.index - 1,
                },
            }
        )

    @property
    def api(self):
        return self._api

    @property
    def parent(self):
        return self._parent

    @property
    def index(self):
        if isinstance(self.key, numbers.Number):
            return self.key
        for ix, obj in enumerate(self.parent.api["shapes"]):
            if obj["name"] == self.key:
                return ix + 1
        raise KeyError(self.key)

    @property
    def name(self):
        return self.api["name"]

    @name.setter
    def name(self, value):
        self.api["name"] = value
        self.append_json_action(func="setShapeName", args=[self.index - 1, value])

    @property
    def type(self):
        # Office.js' ShapeType is far coarser than the desktop engines' (5 vs
        # 32), but each of its members maps onto an xlwings name, so the
        # property speaks the same vocabulary everywhere. "Unsupported" has no
        # equivalent and passes through.
        shape_type = self.api["type"]
        return _SHAPE_TYPE_JS2PY.get(shape_type, shape_type)

    @property
    def left(self):
        return self.api["left"]

    @left.setter
    def left(self, value):
        self.api["left"] = value
        self.append_json_action(func="setShapeLeft", args=[self.index - 1, value])

    @property
    def top(self):
        return self.api["top"]

    @top.setter
    def top(self, value):
        self.api["top"] = value
        self.append_json_action(func="setShapeTop", args=[self.index - 1, value])

    @property
    def width(self):
        return self.api["width"]

    @width.setter
    def width(self, value):
        self.api["width"] = value
        self.append_json_action(func="setShapeWidth", args=[self.index - 1, value])

    @property
    def height(self):
        return self.api["height"]

    @height.setter
    def height(self, value):
        self.api["height"] = value
        self.append_json_action(func="setShapeHeight", args=[self.index - 1, value])

    async def _get_shape_data(self, key, start=None, length=None):
        """Fetch one on-demand property for this shape from the client.

        A shape's geometry is in the payload, but its text isn't: that lives
        on a separate text frame, so reading it costs its own round-trip
        whether or not anything asks for it. `start` and `length` narrow the
        read to a character slice, for `Characters`.
        """
        if sys.platform != "emscripten":
            raise NotImplementedError(f"get_{key}() is only supported in xlwings Lite")
        import js
        from pyodide.ffi import to_js

        options = to_js(
            {"start": start, "length": length},
            dict_converter=js.Object.fromEntries,
        )
        data_js = await js.xlwings.getShapeData(
            self.parent.name, self.index - 1, to_js([key]), options
        )
        return _normalize_jsnull(data_js.to_py())[key]

    async def get_text(self):
        return await self._get_shape_data("text")

    @property
    def text(self):
        # Not in the payload: a shape's text lives on a separate text frame,
        # so it's fetched on demand rather than costing a round-trip for
        # every shape on every request.
        raise NotImplementedError(
            "Reading a shape's text synchronously isn't supported on this "
            "engine. Use 'await myshape.get_text()' to fetch it on demand."
        )

    @text.setter
    def text(self, value):
        self.append_json_action(func="setShapeText", args=[self.index - 1, value])

    def delete(self):
        del self.parent.api["shapes"][self.index - 1]
        self.append_json_action(func="deleteShape", args=[self.index - 1])

    def activate(self):
        raise NotImplementedError(
            "Shape.activate() is not supported on this engine, which has no way "
            "to activate or select a shape."
        )

    def scale_height(self, factor, relative_to_original_size, scale):
        self._scale("height", factor, relative_to_original_size, scale)

    def scale_width(self, factor, relative_to_original_size, scale):
        self._scale("width", factor, relative_to_original_size, scale)

    def _scale(self, axis, factor, relative_to_original_size, scale):
        try:
            scale_from = _SHAPE_SCALE_FROM[scale]
        except KeyError:
            raise ValueError(
                f"Invalid scale: {scale!r}. Must be one of {sorted(_SHAPE_SCALE_FROM)}."
            ) from None
        scale_type = "OriginalSize" if relative_to_original_size else "CurrentSize"
        self.append_json_action(
            func="scaleShape",
            args=[self.index - 1, factor, scale_type, scale_from, axis],
        )
        # Keep width/height in step, so reading them back in the same script
        # reflects the scaling. Only the current-size case is predictable:
        # scaling from the original size needs a size this engine doesn't
        # know, so leave those to the next round-trip.
        if not relative_to_original_size and self.api.get(axis) is not None:
            self.api[axis] = self.api[axis] * factor

    @property
    def font(self):
        return Font(self, None)

    @property
    def characters(self):
        return Characters(self)


class Shapes(Collection):
    # base_classes has no Shapes; the desktop engines subclass their Collection
    # directly too.
    _attr = "shapes"
    _wrap = Shape


class PageSetup(base_classes.PageSetup):
    def __init__(self, sheet):
        self.sheet = sheet

    @property
    def api(self):
        return self.sheet.api

    @property
    def print_area(self):
        return self.sheet.api.get("print_area")

    @print_area.setter
    def print_area(self, value):
        self.sheet.api["print_area"] = value
        self.sheet.append_json_action(func="setPrintArea", args=[value])


class Border(base_classes.Border):
    """One side of a range's borders on this engine.

    Writes queue a `setBorderProperty` action with the side name, the
    attribute and its value; the client maps the snake_case names onto the
    Office.js `BorderIndex`/`BorderLineStyle`/`BorderWeight` strings. A
    `line_style` of None asks the client to remove the border.
    """

    def __init__(self, parent, side, api):
        self.parent = parent
        self.side = side
        self._api = api

    @property
    def api(self):
        return self._api

    def append_json_action(self, attribute, value):
        self.parent.append_json_action(
            func="setBorderProperty", args=[self.side, attribute, value]
        )

    async def _get_border(self):
        """All three attributes of this side, out of the single "borders" read.

        The client returns all eight sides in one round-trip, so there's
        nothing to gain from fetching them individually.

        No engine reports a side whose segments differ from cell to cell, so
        this can't return None for a mixed side either. Which value comes
        back differs from the desktop engines though (measured on Excel for
        Mac, 2026-09-07): an edge reports its first segment's value rather
        than the range's, and an inside border reads "none" as soon as the
        range's cells don't share the same border formatting, even though
        every cell's own borders are intact.
        """
        return (await self.parent._get_range_data("borders"))[self.side]

    async def get_line_style(self):
        return (await self._get_border())["line_style"]

    async def get_weight(self):
        return (await self._get_border())["weight"]

    async def get_color(self):
        color = (await self._get_border())["color"]
        return utils.hex_to_rgb(color) if color else None

    def _sync_read_error(self, getter):
        return NotImplementedError(
            "Reading border attributes synchronously isn't supported on this "
            f"engine. Use 'await myrange.borders[{self.side!r}].{getter}()' to "
            "fetch it on demand."
        )

    @property
    def line_style(self):
        raise self._sync_read_error("get_line_style")

    @line_style.setter
    def line_style(self, value):
        self.append_json_action("line_style", value)

    @property
    def weight(self):
        raise self._sync_read_error("get_weight")

    @weight.setter
    def weight(self, value):
        self.append_json_action("weight", value)

    @property
    def color(self):
        raise self._sync_read_error("get_color")

    @color.setter
    def color(self, color_or_rgb):
        if color_or_rgb is None:
            # Unlike a fill, a border has no "no color": removal is line_style
            raise ValueError(
                "Border color can't be None. To remove a border, set "
                "line_style=None or use clear()."
            )
        self.append_json_action("color", _color_to_hex(color_or_rgb))


class Borders(base_classes.Borders):
    def __init__(self, parent, api):
        self.parent = parent
        self._api = api

    @property
    def api(self):
        return self._api

    def __getitem__(self, side):
        return Border(self.parent, side, self._api)

    async def _common_value(self, attribute):
        """The value the existing grid sides share, or None if they differ.

        Only the sides are compared: Office.js doesn't report segments that
        differ within a side, see Border._get_border().
        """
        borders = await self.parent._get_range_data("borders")
        values = {borders[side][attribute] for side in self._grid_sides()}
        return values.pop() if len(values) == 1 else None

    async def get_line_style(self):
        return await self._common_value("line_style")

    async def get_weight(self):
        return await self._common_value("weight")

    async def get_color(self):
        color = await self._common_value("color")
        return utils.hex_to_rgb(color) if color else None

    def _sync_read_error(self, getter):
        return NotImplementedError(
            "Reading border attributes synchronously isn't supported on this "
            f"engine. Use 'await myrange.borders.{getter}()' to fetch it on demand."
        )

    @property
    def line_style(self):
        raise self._sync_read_error("get_line_style")

    @line_style.setter
    def line_style(self, value):
        self.set(base_classes.BORDER_GRID_SIDES, line_style=value)

    @property
    def weight(self):
        raise self._sync_read_error("get_weight")

    @weight.setter
    def weight(self, value):
        self.set(base_classes.BORDER_GRID_SIDES, weight=value)

    @property
    def color(self):
        raise self._sync_read_error("get_color")

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
        # `which` arrives validated and expanded by main.Borders. One action
        # per supplied attribute per side, in the documented order color,
        # weight, line style, which the client applies as-is.
        if color is not base_classes._UNSET:
            # Convert once, so an invalid color raises before anything queues
            color = _color_to_hex(color)
        for side in which:
            border = self[side]
            if color is not base_classes._UNSET:
                border.append_json_action("color", color)
            if weight is not base_classes._UNSET:
                border.weight = weight
            if line_style is not base_classes._UNSET:
                border.line_style = line_style

    def clear(self, which):
        self.set(which, line_style=None)


class Font(base_classes.Font):
    # TODO: support Shape (Shape.font and Characters.font need a parent that
    # isn't a Range; the setters below raise for those, see append_json_action)
    def __init__(self, parent, api):
        self.parent = parent
        self._api = api

    def _shape_and_slice(self):
        """The shape this font belongs to, plus its character slice if any.

        Returns (None, None, None) for a Range parent, which uses the range
        font action instead.
        """
        if isinstance(self.parent, Shape):
            return self.parent, None, None
        if isinstance(self.parent, Characters):
            return self.parent.parent, self.parent.start, self.parent.length
        return None, None, None

    def append_json_action(self, **kwargs):
        if isinstance(self.parent, Range):
            self.parent.append_json_action(
                **{
                    **kwargs,
                }
            )
            return
        shape, start, length = self._shape_and_slice()
        if shape is None:
            raise NotImplementedError(
                "Setting font attributes is only supported on a Range, a Shape "
                "or a Shape's characters on this engine."
            )
        # Shapes take their own action: the range one addresses cells.
        attribute, value = kwargs["args"]
        shape.append_json_action(
            func="setShapeFontProperty",
            args=[shape.index - 1, start, length, attribute, value],
        )

    @property
    def api(self):
        return self._api

    async def _get_font(self):
        """Fetch all five font attributes in one round-trip.

        They come from a single Office.js object, so there's nothing to gain
        from fetching them individually.
        """
        if isinstance(self.parent, Range):
            return await self.parent._get_range_data("font")
        shape, start, length = self._shape_and_slice()
        if shape is None:
            raise NotImplementedError(
                "Reading font attributes is only supported on a Range, a Shape "
                "or a Shape's characters on this engine."
            )
        font = await shape._get_shape_data("font", start=start, length=length)
        # A shape with no text has no font to report.
        return (
            font
            if font is not None
            else dict.fromkeys(["bold", "italic", "size", "color", "name"])
        )

    async def get_bold(self):
        return (await self._get_font())["bold"]

    async def get_italic(self):
        return (await self._get_font())["italic"]

    async def get_size(self):
        return (await self._get_font())["size"]

    async def get_name(self):
        return (await self._get_font())["name"]

    async def get_color(self):
        color = (await self._get_font())["color"]
        return utils.hex_to_rgb(color) if color else None

    @property
    def bold(self):
        raise NotImplementedError(
            "Reading font attributes synchronously isn't supported on this "
            "engine. Use 'await myrange.font.get_bold()' to fetch it on demand."
        )

    @bold.setter
    def bold(self, value):
        self.append_json_action(func="setFontProperty", args=["bold", value])

    @property
    def italic(self):
        raise NotImplementedError(
            "Reading font attributes synchronously isn't supported on this "
            "engine. Use 'await myrange.font.get_italic()' to fetch it on demand."
        )

    @italic.setter
    def italic(self, value):
        self.append_json_action(func="setFontProperty", args=["italic", value])

    @property
    def size(self):
        raise NotImplementedError(
            "Reading font attributes synchronously isn't supported on this "
            "engine. Use 'await myrange.font.get_size()' to fetch it on demand."
        )

    @size.setter
    def size(self, value):
        self.append_json_action(func="setFontProperty", args=["size", value])

    @property
    def color(self):
        raise NotImplementedError(
            "Reading font attributes synchronously isn't supported on this "
            "engine. Use 'await myrange.font.get_color()' to fetch it on demand."
        )

    @color.setter
    def color(self, color_or_rgb):
        self.append_json_action(
            func="setFontProperty", args=["color", _color_to_hex(color_or_rgb)]
        )

    @property
    def name(self):
        raise NotImplementedError(
            "Reading font attributes synchronously isn't supported on this "
            "engine. Use 'await myrange.font.get_name()' to fetch it on demand."
        )

    @name.setter
    def name(self, value):
        self.append_json_action(func="setFontProperty", args=["name", value])
