# -*- coding: utf-8 -*-

from . import Pipeline, Converter, Options, Accessor, accessors

from .. import xlplatform
from ..main import Range

import datetime

try:
    import numpy as np
except ImportError:
    np = None


_date_handlers = {
    datetime.datetime: datetime.datetime,
    datetime.date: lambda year, month, day, **kwargs: datetime.date(year, month, day)
}

_number_handlers = {
}


class ExpandRangeStage(object):
    def __init__(self, options):
        self.expand = options.get('expand', None)

    def __call__(self, c):
        if c.range:
            # auto-expand the range
            if self.expand:
                c.range = getattr(c.range, self.expand)


class WriteValueToRangeStage(object):
    def __init__(self, raw=False):
        self.raw = raw

    def __call__(self, ctx):
        if ctx.range:
            if self.raw:
                xlplatform.set_value(ctx.range.xl_range, ctx.value)
            else:
                # it is assumed by this stage that value is a list of lists
                if ctx.meta.get('scalar', False):
                    # transform scalars back from list of list so you can do:
                    # Range('A1:B2').value = scalar
                    ctx.value = ctx.value[0][0]
                    row2 = ctx.range.row2
                    col2 = ctx.range.col2
                else:
                    row2 = ctx.range.row1 + len(ctx.value) - 1
                    col2 = ctx.range.col1 + len(ctx.value[0]) - 1

                xlplatform.set_value(xlplatform.get_range_from_indices(
                    ctx.range.xl_sheet,
                    ctx.range.row1,
                    ctx.range.col1,
                    row2,
                    col2
                ), ctx.value)


class ReadValueFromRangeStage(object):

    def __call__(self, c):
        if c.range:
            c.value = xlplatform.get_value_from_range(c.range.xl_range)


class CleanDataFromReadStage(object):

    def __init__(self, options):
        dates_as = options.get('dates', datetime.datetime)
        self.empty_as = options.get('empty', None)
        self.dates_handler = _date_handlers.get(dates_as, dates_as)
        numbers_as = options.get('numbers', None)
        self.numbers_handler = _number_handlers.get(numbers_as, numbers_as)

    def __call__(self, c):
        c.value = xlplatform.clean_value_data(c.value, self.dates_handler, self.empty_as, self.numbers_handler)


class CleanDataForWriteStage(object):

    def __call__(self, c):
        c.value = [
            [
                xlplatform.prepare_xl_data_element(x)
                for x in y
            ]
            for y in c.value
        ]


class AdjustDimensionsStage(object):

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


class Ensure2DStage(object):

    def __call__(self, c):
        if isinstance(c.value, (list, tuple)):
            if len(c.value) > 0:
                if not isinstance(c.value[0], (list, tuple)):
                    c.value = [c.value]
        else:
            c.meta['scalar'] = True
            c.value = [[c.value]]


class TransposeStage(object):

    def __call__(self, c):
        c.value = [[e[i] for e in c.value] for i in range(len(c.value[0]) if c.value else 0)]


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


RangeAccessor.register(Range)


class RawValueAccessor(Accessor):

    @classmethod
    def reader(cls, options):
        return (
            Accessor.reader(options)
            .append_stage(ReadValueFromRangeStage())
        )

    @classmethod
    def writer(cls, options):
        return (
            Accessor.writer(options)
            .prepend_stage(WriteValueToRangeStage(raw=True))
        )

RawValueAccessor.register('raw')


class ValueAccessor(Accessor):

    @staticmethod
    def reader(options):
        return (
            BaseAccessor.reader(options)
            .append_stage(ReadValueFromRangeStage())
            .append_stage(Ensure2DStage())
            .append_stage(CleanDataFromReadStage(options))
            .append_stage(TransposeStage(), only_if=options.get('transpose', False))
            .append_stage(AdjustDimensionsStage(options))
        )

    @staticmethod
    def writer(options):
        return (
            Pipeline()
            .prepend_stage(WriteValueToRangeStage())
            .prepend_stage(CleanDataForWriteStage())
            .prepend_stage(TransposeStage(), only_if=options.get('transpose', False))
            .prepend_stage(Ensure2DStage())
        )

    @classmethod
    def router(cls, value, rng, options):
        return accessors.get(type(value), cls)


ValueAccessor.register(None)


class DictConverter(Converter):

    writes_types = dict

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