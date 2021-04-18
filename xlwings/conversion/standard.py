import math
import datetime
from collections import OrderedDict

from . import Pipeline, Converter, Options, Accessor, accessors

from .. import xlplatform
from ..main import Range
from .. import LicenseError
from ..utils import chunk
try:
    from ..pro import Markdown
    from ..pro.reports import markdown
except (ImportError, LicenseError):
    Markdown = None
try:
    import numpy as np
except ImportError:
    np = None


_date_handlers = {
    datetime.datetime: datetime.datetime,
    datetime.date: lambda year, month, day, **kwargs: datetime.date(year, month, day)
}

_number_handlers = {
    int: lambda x: int(round(x)),
    'raw int': int,
}


class ExpandRangeStage:
    def __init__(self, options):
        self.expand = options.get('expand', None)

    def __call__(self, c):
        if c.range:
            # auto-expand the range
            if self.expand:
                c.range = c.range.expand(self.expand)


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

            chunksize = self.options.get('chunksize')
            if chunksize:
                for ix, value_chunk in enumerate(chunk(value, chunksize)):
                    rng[ix * chunksize:ix * chunksize + chunksize, :].raw_value = value_chunk
            else:
                rng.raw_value = value

    def __call__(self, ctx):
        if ctx.range and ctx.value:
            if self.raw:
                ctx.range.raw_value = ctx.value
                return

            scalar = ctx.meta.get('scalar', False)
            if not scalar:
                ctx.range = ctx.range.resize(len(ctx.value), len(ctx.value[0]))

            self._write_value(ctx.range, ctx.value, scalar)


class ReadValueFromRangeStage:

    def __init__(self, options):
        self.options = options

    def __call__(self, c):
        chunksize = self.options.get('chunksize')
        if c.range and chunksize:
            parts = []
            for i in range(math.ceil(c.range.shape[0] / chunksize)):
                raw_value = c.range[i * chunksize: (i * chunksize) + chunksize, :].raw_value
                if isinstance(raw_value[0], (list, tuple)):
                    parts.extend(raw_value)
                else:
                    # Turn a single row list into a 2d list
                    parts.extend([raw_value])

            c.value = parts
        elif c.range:
            c.value = c.range.raw_value


class CleanDataFromReadStage:

    def __init__(self, options):
        dates_as = options.get('dates', datetime.datetime)
        self.empty_as = options.get('empty', None)
        self.dates_handler = _date_handlers.get(dates_as, dates_as)
        numbers_as = options.get('numbers', None)
        self.numbers_handler = _number_handlers.get(numbers_as, numbers_as)

    def __call__(self, c):
        c.value = xlplatform.clean_value_data(c.value, self.dates_handler, self.empty_as, self.numbers_handler)


class CleanDataForWriteStage:

    def __call__(self, c):
        c.value = [
            [
                xlplatform.prepare_xl_data_element(x)
                for x in y
            ]
            for y in c.value
        ]


class AdjustDimensionsStage:

    def __init__(self, options):
        self.ndim = options.get('ndim', None)

    def __call__(self, c):

        # the assumption is that value is 2-dimensional at this stage

        if self.ndim is None:
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

        # ndim = 2 is a no-op
        elif self.ndim != 2:
            raise ValueError('Invalid c.value ndim=%s' % self.ndim)


class Ensure2DStage:

    def __call__(self, c):
        if isinstance(c.value, (list, tuple)):
            if len(c.value) > 0:
                if not isinstance(c.value[0], (list, tuple)):
                    c.value = [c.value]
        else:
            c.meta['scalar'] = True
            c.value = [[c.value]]


class TransposeStage:

    def __call__(self, c):
        c.value = [[e[i] for e in c.value] for i in range(len(c.value[0]) if c.value else 0)]


class FormatStage:
    def __init__(self, options):
        self.options = options

    def __call__(self, ctx):
        if Markdown and isinstance(ctx.source_value, Markdown):
            markdown.format_text(ctx.range, ctx.source_value.text, ctx.source_value.style)


class BaseAccessor(Accessor):

    @classmethod
    def reader(cls, options):
        return (
            Pipeline()
            .append_stage(ExpandRangeStage(options), only_if=options.get('expand', None))
        )


class RangeAccessor(Accessor):

    @staticmethod
    def copy_range_to_value(c):
        c.value = c.range

    @classmethod
    def reader(cls, options):
        return (
            BaseAccessor.reader(options)
            .append_stage(RangeAccessor.copy_range_to_value)
        )


RangeAccessor.register('range', Range)


class RawValueAccessor(Accessor):

    @classmethod
    def reader(cls, options):
        return (
            Accessor.reader(options)
            .append_stage(ReadValueFromRangeStage(options))
        )

    @classmethod
    def writer(cls, options):
        return (
            Accessor.writer(options)
            .prepend_stage(WriteValueToRangeStage(options, raw=True))
        )


RawValueAccessor.register('raw')


class ValueAccessor(Accessor):

    @staticmethod
    def reader(options):
        return (
            BaseAccessor.reader(options)
            .append_stage(ReadValueFromRangeStage(options))
            .append_stage(Ensure2DStage())
            .append_stage(CleanDataFromReadStage(options))
            .append_stage(TransposeStage(), only_if=options.get('transpose', False))
            .append_stage(AdjustDimensionsStage(options))
        )

    @staticmethod
    def writer(options):
        return (
            Pipeline()
            .prepend_stage(FormatStage(options))
            .prepend_stage(WriteValueToRangeStage(options))
            .prepend_stage(CleanDataForWriteStage())
            .prepend_stage(TransposeStage(), only_if=options.get('transpose', False))
            .prepend_stage(Ensure2DStage())
        )

    @classmethod
    def router(cls, value, rng, options):
        return accessors.get(type(value), cls)


ValueAccessor.register(None)


class DictConverter(Converter):

    writes_types = dict  # TODO: remove all these writes_types as this was long ago replaced by .register (?)

    @classmethod
    def base_reader(cls, options):
        return (
            super(DictConverter, cls).base_reader(
                Options(options)
                .override(ndim=2)
            )
        )

    @classmethod
    def read_value(cls, value, options):
        assert not value or len(value[0]) == 2
        return dict(value)

    @classmethod
    def write_value(cls, value, options):
        return list(value.items())


DictConverter.register(dict)


class OrderedDictConverter(Converter):

    writes_types = OrderedDict

    @classmethod
    def base_reader(cls, options):
        return (
            super(OrderedDictConverter, cls).base_reader(
                Options(options)
                .override(ndim=2)
            )
        )

    @classmethod
    def read_value(cls, value, options):
        assert not value or len(value[0]) == 2
        return OrderedDict(value)

    @classmethod
    def write_value(cls, value, options):
        return list(value.items())


OrderedDictConverter.register(OrderedDict)
