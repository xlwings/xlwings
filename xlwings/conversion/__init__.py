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
        for k, v in defaults:
            if k not in self:
                self[k] = v
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


class ResizeRange(object):
    def __init__(self, options):
        self.expand = options.get('expand', None)

    def read(self, value):
        if isinstance(value, Range):
            # auto-expand the range
            if self.expand:
                value = getattr(value, self.expand)
            return value
        else:
            raise ValueError("Expected Range object")


class WriteValueToRange(object):

    def write(self, value, rng):
        if rng is not None:
            # it is assumed by this stage that value is a list of lists
            row2 = rng.row1 + len(value) - 1
            col2 = rng.col1 + len(value[0]) - 1
            xlplatform.set_value(xlplatform.get_range_from_indices(rng.xl_sheet, rng.row1, rng.col1, row2, col2), value)

        return value, rng


class ExtractValue(object):

    def read(self, value):
        value = xlplatform.get_value_from_range(value.xl_range)
        if type(value) is not list:
            value = [[value]]
        return value


class CleanValueData(object):

    def __init__(self, options):
        dates_as = options.get('dates_as', datetime.datetime)
        self.dates_handler = _date_handlers.get(dates_as, dates_as)

    def read(self, value):
        return xlplatform.clean_value_data(value, self.dates_handler)

    def write(self, value, range):
        return [
            [
                None
                if np and isinstance(x, float) and np.isnan(x)
                else xlplatform.prepare_xl_data_element(x)
                for x in y
            ]
            for y in value
        ]


class AdjustDimensionality(object):

    def __init__(self, options):
        self.ndim = options.get('ndim', None)

    def read(self, value):

        # the assumption is that value is 2-dimensional at this stage

        if self.ndim is None:
            if len(value) == 1:
                return value[0][0] if len(value) == 0 else value[0]
            elif len(value[0]) == 1:
                return [ x[0] for x in value ]
            else:
                return value

        elif self.ndim == 1:
            if len(value) == 1:
                return value[0]
            elif len(value[0]) == 1:
                return [ x[0] for x in value ]
            else:
                raise Exception("Range must be 1-by-n or n-by-1 when ndim=1.")

        elif self.ndim == 2:
            return value

        else:
            raise ValueError('Invalid value ndim=%s' % self.ndim)

    def write(self, value, rng):
        if isinstance(value, (list, tuple)):
            if len(value) > 0:
                if not isinstance(value[0], (list, tuple)):
                    value = [value]
        else:
            value = [[value]]
        return value, rng


class Transpose(object):

    def read(self, value):
        return [[e[i] for e in value] for i in range(len(value[0]) if value else 0)]

    def write(self, value, rng):
        return [[e[i] for e in value] for i in range(len(value[0]) if value else 0)], rng


def default_router(value, rng, options):
    return converters[type(value), converters[None]]


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
            .append_stage(ExtractValue(options))
            .append_stage(CleanValueData(options))
            .append_stage(Transpose(), only_if=options.get('transpose', False))
            .append_stage(AdjustDimensionality(options))
        )

    @staticmethod
    def writer(options):
        return (
            Pipeline()
            .prepend_stage(WriteValueToRange(options))
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

        def read(self, value):
            value = [[np.nan if x is None else x for x in i] for i in value]

            dtype = self.dtype
            ndim = self.ndim

            return np.array(value, dtype=dtype, ndmin=ndim)

        def write(self, value, rng):
            try:
                value = np.where(np.isnan(value), None, value)
                value = value.tolist()
            except TypeError:
                # isnan doesn't work on arrays of dtype=object
                if pd:
                    value[pd.isnull(value)] = None
                    value = value.tolist()
                else:
                    # expensive way of replacing nan with None in object arrays in case Pandas is not available
                    value = [[None if isinstance(c, float) and np.isnan(c) else c for c in row] for row in value]

            return value


    class NumpyArrayAccessor(object):

        @staticmethod
        def reader(options):
            return (
                converters[None].reader(
                    Options(options)
                    .override(ndim=2)
                    .defaults(expand='table')
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

        def read(self, value):
            return pd.DataFrame(value[1:], columns=value[0])

        def write(self, value, rng):
            if self.index:
                if value.index.name in value.columns:
                    # Prevents column name collision when resetting the index
                    value.index.rename(None, inplace=True)
                value = value.reset_index()

            if self.header:
                if isinstance(value.columns, pd.MultiIndex):
                    columns = list(zip(*value.columns.tolist()))
                else:
                    columns = [value.columns.tolist()]
                value = columns + value.values.tolist()
            else:
                value = value.values.tolist()

            return value, rng


    class PandasDataFrameAccessor(object):

        @staticmethod
        def reader(options):
            return (
                converters[None].reader(
                    Options(options)
                    .override(ndim=2)
                    .defaults(expand='table')
                )
                .append_stage(ConvertPandasDataFrame(options))
            )

        @staticmethod
        def writer(options):
            return (
                converters[None].writer(
                    options
                )
                .append_stage(ConvertPandasDataFrame(options))
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

        def read(self, value, options):
            return pd.Series(value[1:])

        def write(self, value, rng, options):
            if self.index:
                value = value.reset_index().values.tolist()
            else:
                value = value.values[:, np.newaxis].tolist()

            return value, rng


    class PandasDataSeriesAccessor(object):

        @staticmethod
        def reader(options):
            return (
                converters[None].reader(
                    Options(options)
                    .override(ndim=1)
                    .defaults(expand='table')
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

        def read(self, value):
            return self.read_value(value, self.options)

        def write(self, value, rng):
            return self.write_value(value, self.options), rng

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
    def read_value(cls, value):
        assert not value or len(value[0]) == 2
        return dict(value)

    @classmethod
    def write_value(cls, value):
        return list(value.items())


converters[dict] = DictConverter
