# -*- coding: utf-8 -*-

from .. import xlplatform

from ..main import Range

import datetime


# Optional imports
try:
    import numpy as np
except ImportError:
    np = None

try:
    import pandas as pd
except ImportError:
    pd = None


converters = {}


_date_handlers = {
    datetime.datetime: datetime.datetime,
    datetime.date: lambda year, month, day, **kwargs: datetime.date(year, month, day)
}


class ConversionContext(object):
    __slots__ = ['range', 'value']

    def __init__(self, range=None, value=None):
        self.range = range
        self.value = value


class Options(dict):

    def __init__(self, original):
        super(Options, self).__init__(original)

    def override(self, **overrides):
        self.update(overrides)
        return self

    def erase(self, keys):
        for key in keys:
            self.pop(key, None)
        return self

    def defaults(self, **defaults):
        for k, v in defaults.items():
            self.setdefault(k, v)
        return self


class Pipeline(list):

    def prepend_stage(self, stage, only_if=True):
        if only_if:
            self.insert(0, stage)
        return self

    def append_stage(self, stage, only_if=True):
        if only_if:
            self.append(stage)
        return self

    def insert_stage(self, stage, index=None, after=None, before=None, replace=None, only_if=True):
        if only_if:
            if sum(x is not None for x in (index, after, before, replace)) != 1:
                raise ValueError("Must specify exactly one of arguments: index, after, before, replace")
            if index is not None:
                indices = (index,)
            elif after is not None:
                indices = tuple(i+1 for i, x in enumerate(self) if isinstance(x, after))
            elif before is not None:
                indices = tuple(i for i, x in enumerate(self) if isinstance(x, before))
            elif replace is not None:
                for i, x in enumerate(self):
                    if isinstance(x, replace):
                        self[i] = stage
                return self
            for i in reversed(indices):
                self.insert(i, stage)
        return self


def execute_write_pipeline(pipeline, value, range):
    ctx = ConversionContext(range=range, value=value)
    for stage in pipeline:
        stage.write(ctx)


def execute_read_pipeline(pipeline, range):
    ctx = ConversionContext(range=range)
    for stage in pipeline:
        stage.read(ctx)
    return ctx.value


class ResizeRange(object):
    def __init__(self, options):
        self.expand = options.get('expand', None)

    def read(self, ctx):
        if ctx.range:
            # auto-expand the range
            if self.expand:
                ctx.range = getattr(ctx.range, self.expand)


class WriteValueToRange(object):

    def write(self, ctx):
        if ctx.range:
            # it is assumed by this stage that value is a list of lists
            row2 = ctx.range.row1 + len(ctx.value) - 1
            col2 = ctx.range.col1 + len(ctx.value[0]) - 1
            xlplatform.set_value(xlplatform.get_range_from_indices(
                ctx.range.xl_sheet,
                ctx.range.row1,
                ctx.range.col1,
                row2,
                col2
            ), ctx.value)


class ExtractValue(object):

    def read(self, c):
        c.value = xlplatform.get_value_from_range(c.range.xl_range)
        if not isinstance(c.value, (list, tuple)):
            c.value = [[c.value]]


class CleanValueData(object):

    def __init__(self, options):
        dates_as = options.get('dates_as', datetime.datetime)
        self.empty_as = options.get('empty_as', None)
        self.dates_handler = _date_handlers.get(dates_as, dates_as)

    def read(self, c):
        c.value = xlplatform.clean_value_data(c.value, self.dates_handler, self.empty_as)

    def write(self, c):
        c.value = [
            [
                None
                if np and isinstance(x, float) and np.isnan(x)
                else xlplatform.prepare_xl_data_element(x)
                for x in y
            ]
            for y in c.value
        ]


class AdjustDimensionality(object):

    def __init__(self, options):
        self.ndim = options.get('ndim', None)

    def read(self, c):

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
                c.value = [ x[0] for x in c.value ]
            else:
                raise Exception("Range must be 1-by-n or n-by-1 when ndim=1.")

        # ndim = 2 is a no-op
        elif self.ndim != 2:
            raise ValueError('Invalid c.value ndim=%s' % self.ndim)

    def write(self, c):
        if isinstance(c.value, (list, tuple)):
            if len(c.value) > 0:
                if not isinstance(c.value[0], (list, tuple)):
                    c.value = [c.value]
        else:
            c.value = [[c.value]]


class Transpose(object):

    def read(self, c):
        return [[e[i] for e in c.value] for i in range(len(c.value[0]) if c.value else 0)]

    def write(self, c):
        return [[e[i] for e in c.value] for i in range(len(c.value[0]) if c.value else 0)]


def default_router(value, rng, options):
    return converters.get(type(value), converters[None])


class RangeAccessor(object):

    def reader(self, options):
        return (
            Pipeline()
            .append_stage(ResizeRange(), only_if=options.get('expand', None))
        )

    router = default_router


converters[Range] = RangeAccessor()


class ValueAccessor:

    @staticmethod
    def reader(options):
        return (
            Pipeline()
            .append_stage(ResizeRange(options))
            .append_stage(ExtractValue())
            .append_stage(CleanValueData(options))
            .append_stage(Transpose(), only_if=options.get('transpose', False))
            .append_stage(AdjustDimensionality(options))
        )

    @staticmethod
    def writer(options):
        return (
            Pipeline()
            .prepend_stage(WriteValueToRange())
            .prepend_stage(CleanValueData(options))
            .prepend_stage(Transpose(), only_if=options.get('transpose', False))
            .prepend_stage(AdjustDimensionality(options))
        )

    @classmethod
    def router(cls, value, rng, options):
        if isinstance(value, (int, float, list, tuple, str, bool)):
            return cls
        else:
            return default_router(value, rng, options)


converters[None] = ValueAccessor


if np:

    class ConvertNumpyArray(object):

        def __init__(self, options):
            self.dtype = options.get('dtype', None)
            self.ndim = options.get('ndim', 0)

        def read(self, c):
            dtype = self.dtype
            ndim = self.ndim

            c.value = np.array(c.value, dtype=dtype, ndmin=ndim)

        def write(self, c):
            try:
                c.value = np.where(np.isnan(c.value), None, c.value)
                c.value = c.value.tolist()
            except TypeError:
                # isnan doesn't work on arrays of dtype=object
                if pd:
                    c.value[pd.isnull(c.value)] = None
                    c.value = c.value.tolist()
                else:
                    # expensive way of replacing nan with None in object arrays in case Pandas is not available
                    c.value = [[None if isinstance(c, float) and np.isnan(c) else c for c in row] for row in c.value]


    class NumpyArrayAccessor(object):

        @staticmethod
        def reader(options):
            return (
                converters[None].reader(
                    Options(options)
                    # .override(ndim=2)
                    # .defaults(expand='table')
                    .defaults(empty_as=np.nan)
                )
                .append_stage(ConvertNumpyArray(options))
            )

        @staticmethod
        def writer(options):
            return (
                converters[None].writer(options)
                .prepend_stage(ConvertNumpyArray(options))
            )

        @classmethod
        def router(cls, value, rng, options):
            if isinstance(value, np.ndarray):
                return cls
            else:
                return default_router(value, rng, options)


    converters[np.array] = converters[np.ndarray] = NumpyArrayAccessor


if pd:

    class ConvertPandasDataFrame(object):

        def __init__(self, options):
            self.index = options.get('index', True)
            self.header = options.get('header', True)

        def read(self, c):
            c.value = pd.DataFrame(c.value[1:], columns=c.value[0])

        def write(self, c):
            if self.index:
                if c.value.index.name in c.value.columns:
                    # Prevents column name collision when resetting the index
                    c.value.index.rename(None, inplace=True)
                c.value = c.value.reset_index()

            if self.header:
                if isinstance(c.value.columns, pd.MultiIndex):
                    columns = list(zip(*c.value.columns.tolist()))
                else:
                    columns = [c.value.columns.tolist()]
                c.value = columns + c.value.values.tolist()
            else:
                c.value = c.value.values.tolist()


    class PandasDataFrameAccessor(object):

        @staticmethod
        def reader(options):
            return (
                converters[None].reader(
                    Options(options)
                    .override(ndim=2)
                    # .defaults(expand='table')
                )
                .append_stage(ConvertPandasDataFrame(options))
            )

        @staticmethod
        def writer(options):
            return (
                converters[None].writer(
                    options
                )
                .prepend_stage(ConvertPandasDataFrame(options))
            )

        @classmethod
        def router(cls, value, rng, options):
            if isinstance(value, pd.DataFrame):
                return cls
            else:
                return default_router(value, rng, options)


    converters[pd.DataFrame] = PandasDataFrameAccessor


    class ConvertPandasDataSeries(object):

        types = (pd.Series,)

        def __init__(self, options):
            self.index = options.get('index', True)

        def read(self, c):
            c.value = pd.Series(c.value[1:])

        def write(self, c):
            if self.index:
                c.value = c.value.reset_index().values.tolist()
            else:
                c.value = c.value.values[:, np.newaxis].tolist()


    class PandasDataSeriesAccessor(object):

        @staticmethod
        def reader(options):
            return (
                converters[None].reader(
                    Options(options)
                    .override(ndim=1)
                    # .defaults(expand='table')
                )
                .append_stage(ConvertPandasDataSeries(options))
            )

        @staticmethod
        def writer(options):
            return (
                converters[None].writer(
                    options
                )
                .prepend_stage(ConvertPandasDataSeries(options))
            )

        @classmethod
        def router(cls, value, rng, options):
            if isinstance(value, pd.Series):
                return cls
            else:
                return default_router(value, rng, options)


    converters[pd.Series] = PandasDataSeriesAccessor


class AbstractConverter(object):

    class ConversionStage(object):

        def __init__(self, read_value, write_value, options):
            self.read_value = read_value
            self.write_value = write_value

        def read(self, c):
            c.value = self.read_value(c.value, self.options)

        def write(self, c):
            c.value = self.write_value(c.value, self.options)

    @classmethod
    def base_reader(cls, options):
        return converters[None].reader(options)

    @classmethod
    def base_writer(cls, options):
        return converters[None].writer(options)

    @classmethod
    def reader(cls, options):
        return (
            cls.base_reader(options)
            .append_stage(AbstractConverter.ConversionStage(cls, options))
        )

    @classmethod
    def writer(cls, options):
        return (
            cls.base_writer(options)
            .prepend_stage(AbstractConverter.ConversionStage(cls, options))
        )

    @classmethod
    def router(cls, value, rng, options):
        if isinstance(value, cls.writes_types):
            return cls
        else:
            return default_router(value, rng, options)


class DictConverter(AbstractConverter):

    writes_types = dict

    @classmethod
    def base_reader(cls, options):
        return (
            super(cls).base_reader(
                Options(options)
                .ovveride(ndim=2)
                .defaults(expand='table')
            )
        )

    @classmethod
    def read_value(cls, value, options):
        assert not value or len(value[0]) == 2
        return dict(value)

    @classmethod
    def write_value(cls, value, options):
        return list(value.items())


converters[dict] = DictConverter
