import datetime
import datetime as dt
import json
import math
from collections import OrderedDict
from typing import Any, Sequence

from .. import LicenseError, XlwingsError
from ..main import Range
from ..utils import chunk, is_jsnull, xlserial_to_datetime
from . import Accessor, Converter, Options, Pipeline, accessors

try:
    from ..pro import Markdown
    from ..pro.reports import markdown
except (ImportError, LicenseError, AttributeError):
    Markdown = None
try:
    import numpy as np
except ImportError:
    np = None


_date_handlers = {
    datetime.datetime: datetime.datetime,
    datetime.date: lambda year, month, day, **kwargs: datetime.date(year, month, day),
}

_number_handlers = {
    # https://github.com/xlwings/xlwings/issues/554
    int: lambda x: int(round(x)),
    "raw int": int,
}

# Distinguishes "chunksize not supplied" from an explicit opt-out
# (``chunksize=None`` / ``chunksize=0``), which must keep disabling chunking.
_MISSING = object()


def _budget_chunksize(budget, nrows, ncols):
    """Rows per chunk for an implicit cell budget, or None if the range fits."""
    if budget is None or nrows * ncols <= budget:
        return None
    return max(1, budget // ncols)


def _resolve_read_chunksize(options, rng):
    """Effective read chunksize: the user's explicit value wins (including falsy
    opt-outs); otherwise the engine's read budget applies only when the range is
    larger than the budget. Shape is only resolved when a budget is in effect."""
    chunksize = options.get("chunksize", _MISSING)
    if chunksize is not _MISSING:
        return chunksize
    budget = rng.impl.max_cells_per_read
    if budget is None:
        return None
    nrows, ncols = rng.shape
    return _budget_chunksize(budget, nrows, ncols)


def _resolve_write_chunksize(options, rng, value, scalar):
    """Effective write chunksize with the same precedence as reads. Non-scalar
    values are sized from the final converted matrix; scalar fills from the
    target range's shape (only resolved when a budget is in effect)."""
    chunksize = options.get("chunksize", _MISSING)
    if chunksize is not _MISSING:
        return chunksize
    budget = rng.impl.max_cells_per_write
    if budget is None:
        return None
    if scalar:
        nrows, ncols = rng.shape
    else:
        nrows, ncols = len(value), len(value[0])
    return _budget_chunksize(budget, nrows, ncols)


def _rows_2d(raw_value):
    """Normalize a chunk's raw value to a list of rows: COM and AppleScript return
    a scalar for a single cell and a flat list for a single row."""
    if not isinstance(raw_value, (list, tuple)):
        return [[raw_value]]
    if not isinstance(raw_value[0], (list, tuple)):
        return [raw_value]
    return raw_value


def _check_live_values(values_js, rng):
    """Office.js may return null instead of raising when a range get exceeds its
    5,000,000-cell limit. Catch it before ``.to_py()`` and say what to do."""
    if values_js is None or is_jsnull(values_js):
        raise XlwingsError(
            f"Excel returned no values for '{rng.sheet.name}'!{rng.address}. "
            "Office.js caps a single range read at 5,000,000 cells and may "
            "return null instead of raising. Read fewer rows or columns, "
            "or use a smaller explicit chunksize."
        )


class ExpandRangeStage:
    def __init__(self, options):
        self.expand = options.get("expand", None)

    def __call__(self, c):
        if c.range:
            # auto-expand the range
            if self.expand:
                c.range = c.range.expand(self.expand)


class AsyncExpandRangeStage:
    """Async expand stage that resolves expansion via JS global (xlwings Lite only)."""

    def __init__(self, options):
        self.expand = options.get("expand", None)

    async def __call__(self, c):
        if c.range and self.expand:
            import js

            expanded_address = await js.xlwings.getExpandedAddress(
                c.range.sheet.name, c.range.address, self.expand
            )
            expanded_address = str(expanded_address)
            # Strip sheet name prefix if present (e.g., "Sheet1!A1:B2" -> "A1:B2")
            if "!" in expanded_address:
                expanded_address = expanded_address.split("!", 1)[1]
            if expanded_address != c.range.address:
                c.range = c.range.sheet.range(expanded_address)


class WriteValueToRangeStage:
    def __init__(self, options, raw=False):
        self.raw = raw
        self.options = options

    def _write_value(self, rng, value, scalar):
        if rng.api and value:
            # it is assumed by this stage that value is a list of lists
            if scalar:
                value = value[0][0]
            else:
                rng = rng.resize(len(value), len(value[0]))

            chunksize = _resolve_write_chunksize(self.options, rng, value, scalar)
            if not chunksize:
                rng.raw_value = value
            elif scalar:
                # Keep the scalar a scalar: each row slice fills itself with it
                # (the engines expand a scalar to the target shape).
                nrows = rng.shape[0]
                for start in range(0, nrows, chunksize):
                    rng[start : start + chunksize, :].raw_value = value
            else:
                for ix, value_chunk in enumerate(chunk(value, chunksize)):
                    rng[
                        ix * chunksize : ix * chunksize + chunksize, :
                    ].raw_value = value_chunk

    def __call__(self, ctx):
        if ctx.range and ctx.value:
            if self.raw:
                ctx.range.raw_value = ctx.value
                return

            scalar = ctx.meta.get("scalar", False)
            if not scalar:
                ctx.range = ctx.range.resize(len(ctx.value), len(ctx.value[0]))

            self._write_value(ctx.range, ctx.value, scalar)


class ReadValueFromRangeStage:
    def __init__(self, options):
        self.options = options

    def __call__(self, c):
        if not c.range:
            # UDF arguments arrive pre-materialized via conversion.read(None, ...)
            return
        chunksize = _resolve_read_chunksize(self.options, c.range)
        if chunksize:
            parts = []
            for i in range(math.ceil(c.range.shape[0] / chunksize)):
                chunk_range = c.range[i * chunksize : (i * chunksize) + chunksize, :]
                # Slicing creates a fresh range without the read options, but
                # AppleScript and Calamine need e.g. err_to_str on the raw read.
                chunk_range = Range(impl=chunk_range.impl, **self.options)
                parts.extend(_rows_2d(chunk_range.raw_value))
            c.value = parts
        else:
            c.value = c.range.raw_value


class AsyncReadValueFromRangeStage:
    """Async read stage that fetches values on demand from Excel via JS global
    (xlwings Lite only)."""

    def __init__(self, options):
        self.options = options

    async def __call__(self, c):
        if not c.range:
            return
        import js

        chunksize = _resolve_read_chunksize(self.options, c.range)
        if chunksize:
            parts = []
            for i in range(math.ceil(c.range.shape[0] / chunksize)):
                chunk_range = c.range[i * chunksize : (i * chunksize) + chunksize, :]
                values_js = await js.xlwings.getRangeValues(
                    c.range.sheet.name, chunk_range.address
                )
                _check_live_values(values_js, chunk_range)
                parts.extend(_rows_2d(values_js.to_py()))
            c.value = parts
        else:
            values_js = await js.xlwings.getRangeValues(
                c.range.sheet.name, c.range.address
            )
            _check_live_values(values_js, c.range)
            c.value = values_js.to_py()


class CleanDataFromReadStage:
    def __init__(self, options):
        self.options = options
        dates_as = options.get("dates", datetime.datetime)
        self.empty_as = options.get("empty", None)
        self.dates_handler = _date_handlers.get(dates_as, dates_as)
        numbers_as = options.get("numbers", None)
        self.numbers_handler = _number_handlers.get(numbers_as, numbers_as)
        self.err_to_str = options.get("err_to_str", False)

    def __call__(self, c):
        c.value = c.engine.impl.clean_value_data(
            c.value,
            self.dates_handler,
            self.empty_as,
            self.numbers_handler,
            self.err_to_str,
        )


class CleanDataForWriteStage:
    def __init__(self, options):
        self.options = options

    def __call__(self, c):
        c.value = [
            [c.engine.impl.prepare_xl_data_element(x, self.options) for x in y]
            for y in c.value
        ]


class AdjustDimensionsStage:
    def __init__(self, options):
        self.ndim = options.get("ndim", None)

    def __call__(self, c):
        # the assumption is that value is 2-dimensional at this stage

        if self.ndim in (None, "squeeze"):
            # "squeeze" isn't documented yet, but could be used in case we want
            # to change the default to "natural" at some point
            if len(c.value) == 1:
                c.value = c.value[0][0] if len(c.value[0]) == 1 else c.value[0]
            elif len(c.value[0]) == 1:
                c.value = [x[0] for x in c.value]
            else:
                c.value = c.value

        elif self.ndim == 1:
            if len(c.value) == 1:
                c.value = c.value[0]
            elif len(c.value[0]) == 1:
                c.value = [x[0] for x in c.value]
            else:
                raise Exception("Range must be 1-by-n or n-by-1 when ndim=1.")

        elif self.ndim == "natural":
            # Single cell: return scalar
            # Horizontal range (1xN): return 1D array
            # Vertical range (Nx1) or 2D range (NxM): return 2D array
            if len(c.value) == 1 and len(c.value[0]) == 1:
                c.value = c.value[0][0]
            elif len(c.value) == 1:
                # Single row: return 1D array
                c.value = c.value[0]
            else:
                # Multiple rows: keep as 2D (even if single column)
                c.value = c.value

        # ndim = 2 is a no-op
        elif self.ndim != 2:
            raise ValueError("Invalid c.value ndim=%s" % self.ndim)


class Ensure2DStage:
    def __call__(self, c):
        if isinstance(c.value, (list, tuple)):
            if len(c.value) > 0:
                if not isinstance(c.value[0], (list, tuple)):
                    c.value = [c.value]
        else:
            c.meta["scalar"] = True
            c.value = [[c.value]]


class TransposeStage:
    def __call__(self, c):
        c.value = [
            [e[i] for e in c.value] for i in range(len(c.value[0]) if c.value else 0)
        ]


class FormatStage:
    def __init__(self, options):
        self.options = options

    def __call__(self, ctx):
        if Markdown and isinstance(ctx.source_value, Markdown):
            markdown.format_text(
                ctx.range, ctx.source_value.text, ctx.source_value.style
            )
        if "formatter" in self.options:
            self.options["formatter"](ctx.range, ctx.source_value)


class BaseAccessor(Accessor):
    @classmethod
    def reader(cls, options):
        return Pipeline().append_stage(
            ExpandRangeStage(options), only_if=options.get("expand", None)
        )


class RangeAccessor(Accessor):
    @staticmethod
    def copy_range_to_value(c):
        c.value = c.range

    @classmethod
    def reader(cls, options):
        return BaseAccessor.reader(options).append_stage(
            RangeAccessor.copy_range_to_value
        )


RangeAccessor.register("range", Range)


class RawValueAccessor(Accessor):
    @classmethod
    def reader(cls, options):
        return Accessor.reader(options).append_stage(ReadValueFromRangeStage(options))

    @classmethod
    def writer(cls, options):
        return Accessor.writer(options).prepend_stage(
            WriteValueToRangeStage(options, raw=True)
        )


RawValueAccessor.register("raw")


class ValueAccessor(Accessor):
    @staticmethod
    def reader(options):
        return (
            BaseAccessor.reader(options)
            .append_stage(ReadValueFromRangeStage(options))
            .append_stage(Ensure2DStage())
            .append_stage(CleanDataFromReadStage(options))
            .append_stage(TransposeStage(), only_if=options.get("transpose", False))
            .append_stage(AdjustDimensionsStage(options))
        )

    @staticmethod
    def writer(options):
        return (
            Pipeline()
            .prepend_stage(FormatStage(options))
            .prepend_stage(WriteValueToRangeStage(options))
            .prepend_stage(CleanDataForWriteStage(options))
            .prepend_stage(TransposeStage(), only_if=options.get("transpose", False))
            .prepend_stage(Ensure2DStage())
        )

    @classmethod
    def router(cls, value, rng, options):
        return accessors.get(type(value), cls)


ValueAccessor.register(
    None,
    "default",
    Any,
    Sequence,
    list,
    tuple,
    str,
    float,
    int,
    bool,
    dt.datetime,
)


class DictConverter(Converter):
    @classmethod
    def base_reader(cls, options):
        return super(DictConverter, cls).base_reader(Options(options).override(ndim=2))

    @classmethod
    def read_value(cls, value, options):
        assert not value or len(value[0]) == 2
        return dict(value)

    @classmethod
    def write_value(cls, value, options):
        return list(value.items())


DictConverter.register(dict)


class OrderedDictConverter(Converter):
    @classmethod
    def base_reader(cls, options):
        return super(OrderedDictConverter, cls).base_reader(
            Options(options).override(ndim=2)
        )

    @classmethod
    def read_value(cls, value, options):
        assert not value or len(value[0]) == 2
        return OrderedDict(value)

    @classmethod
    def write_value(cls, value, options):
        return list(value.items())


OrderedDictConverter.register(OrderedDict)


class DatetimeConverter(Converter):
    @classmethod
    def read_value(cls, value, options):
        return xlserial_to_datetime(value)

    @classmethod
    def write_value(cls, value, options):
        return value


DatetimeConverter.register(datetime.datetime)


class DateConverter(Converter):
    @classmethod
    def read_value(cls, value, options):
        return xlserial_to_datetime(value).date()

    @classmethod
    def write_value(cls, value, options):
        return value


DateConverter.register(datetime.date)


class TupleConverter(Converter):
    @classmethod
    def read_value(cls, value, options):
        if isinstance(value, list):
            if value and isinstance(value[0], list):
                # 2D list: convert each row to a tuple
                return tuple(tuple(row) for row in value)
            else:
                # 1D list: convert to a simple tuple
                return tuple(value)
        # Scalar
        return (value,)

    @classmethod
    def write_value(cls, value, options):
        # Don't do anything as the engines know how to write tuples
        return value


TupleConverter.register(tuple)


class JsonConverter(Converter):
    """Useful for sending context to LLMs"""

    @classmethod
    def read_value(cls, value, options):
        def serialize_datetime(obj):
            if isinstance(obj, (datetime.datetime, datetime.date)):
                return obj.isoformat()
            raise TypeError(f"Object of type {type(obj)} is not JSON serializable")

        return json.dumps(value, default=serialize_datetime)

    @classmethod
    def write_value(cls, value, options):
        def deserialize_datetime(obj):
            if isinstance(obj, list):
                return [deserialize_datetime(item) for item in obj]
            elif isinstance(obj, dict):
                return {key: deserialize_datetime(value) for key, value in obj.items()}
            elif isinstance(obj, str):
                try:
                    return dt.datetime.fromisoformat(obj)
                except ValueError:
                    return obj
            else:
                return obj

        def pad_jagged_array(values):
            if isinstance(values, list) and values and isinstance(values[0], list):
                max_length = max(len(row) for row in values)
                return [row + [None] * (max_length - len(row)) for row in values]
            else:
                return values

        def strip_markdown_code_block(text):
            """Remove markdown code block delimiters (```json```, etc.)"""
            if not isinstance(text, str):
                return text

            text = text.strip()
            # Remove opening code block marker (```json, ```JSON, or just ```)
            if text.startswith("```"):
                # Find the end of the first line (opening marker)
                first_newline = text.find("\n")
                if first_newline != -1:
                    text = text[first_newline + 1 :]
                else:
                    # Just ``` without newline, remove it
                    text = text[3:]

            # Remove closing code block marker
            if text.endswith("```"):
                text = text[:-3]

            return text.strip()

        # Strip potential markdown code blocks before parsing
        cleaned_value = strip_markdown_code_block(value)

        try:
            result = json.loads(cleaned_value)
        except json.JSONDecodeError:
            return value
        result = deserialize_datetime(result)
        # LLMs often give back things like this: [['a', 'b'], ['c']]
        # TODO: should be done in Ensure2DStage, see  see write() in __init__
        result = pad_jagged_array(result)
        return result


JsonConverter.register("json")
