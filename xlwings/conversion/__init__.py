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


class Transpose(object):

    def read(self, value):
        return [[e[i] for e in value] for i in range(len(value[0]) if value else 0)]

    def write(self, value, rng):
        return [[e[i] for e in value] for i in range(len(value[0]) if value else 0)]



class Pipeline(list):

    def add_stage(self, stage, only_if=True):
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


class RangeAccessor(object):

    types = ()

    def vba_read(self, vba, argname, options):
        # auto-expand the range
        expand = options.get('expand', None)
        if expand == 'vertical':
            with vba.block("If TypeOf {argname} Is Range", argname=argname):
                with vba.block("If IsEmpty({argname}) Then", argname=argname):
                    vba.writeln("{argname} = Array()", argname=argname)
                with vba.block("ElseIf Not IsEmpty({argname}.Offset(1, 0).Value", argname=argname):
                    vba.writeln("Set {argname} = Range({argname}, {argname}.End(xlDown))", argname=argname)
                vba.writeln("End If")
            vba.writeln("End If")
        elif expand == 'horizontal':
            with vba.block("If TypeOf {argname} Is Range", argname=argname):
                with vba.block("If IsEmpty({argname}) Then", argname=argname):
                    vba.writeln("{argname} = Array()", argname=argname)
                with vba.block("ElseIf Not IsEmpty({argname}.Offset(0, 1).Value", argname=argname):
                    vba.writeln("Set {argname} = Range({argname}, {argname}.End(xlRight))", argname=argname)
                vba.writeln("End If")
            vba.writeln("End If")
        elif expand == 'table':
            with vba.block("If TypeOf {argname} Is Range", argname=argname):
                with vba.block("If IsEmpty({argname}) Then", argname=argname):
                    vba.writeln("{argname} = Array()", argname=argname)
                with vba.block("ElseIf Not IsEmpty({argname}.Offset(0, 1).Value", argname=argname):
                    with vba.block("If Not IsEmpty({argname}.Offset(1, 0).Value", argname=argname):
                        vba.writeln("Set {argname} = Range({argname}, {argname}.End(xlRight).End(xlDown))", argname=argname)
                    with vba.block("Else"):
                        vba.writeln("Set {argname} = Range({argname}, {argname}.End(xlRight))", argname=argname)
                    vba.writeln("End If")
                vba.writeln("End If")
            vba.writeln("End If")

    def read(self, value, options):
        if isinstance(value, Range):
            # auto-expand the range
            expand = options.get('expand', None)
            if expand:
                value = getattr(value, expand)
            return value
        else:
            raise ValueError("Expected Range object")

    def write_any(self, value, rng, options):
        if isinstance(value, self.types):
            self.write(value, rng, options)
        else:
            return converters.get(type(value), converters[None]).write(value, rng, options)

    def write(self, value, rng, options):

        if rng is not None:
            # it is assumed by this stage that value is a list of lists
            row2 = rng.row1 + len(value) - 1
            col2 = rng.col1 + len(value[0]) - 1

            xlplatform.set_value(xlplatform.get_range_from_indices(rng.xl_sheet, rng.row1, rng.col1, row2, col2), value)

        return value


converters[Range] = RangeAccessor()


class ValueAccessor(RangeAccessor):

    types = (int, float, list, tuple, str, bool)

    def reader(self, options):
        return (
            Pipeline()
            .add_stage(ResizeRange(options))
            .add_stage(ExtractValue(options))
            .add_stage(CleanValueData(options))
            .add_stage(Transpose(), only_if=options.get('transpose', False))
            .add_stage(AdjustDimensionality(options))
        )

    def writer(self, options):
        return (
            .add_stage()
            .add_stage(ResizeRange(options))
        )

    def vba_read(self, vba, argname, options):
        RangeAccessor.vba_read(self, vba, argname, options)
        vba.write("If TypeOf {arg} Is Range Then {arg} = {arg}.Value", arg=argname)

    def read(self, value, options):
        value = RangeAccessor.read(self, value, options)
        value = xlplatform.get_value_from_range(value.xl_range)
        value = xlplatform.clean_value_data(value, _date_handlers[options.get('dates_as', datetime.datetime)])

        ndim = options.get('ndim', None)
        value = self._ensure_dimensionality(value, ndim)

        if options.get('transpose', False):
            if value and isinstance(value, (list, tuple)) and isinstance(value[0], (list, tuple)):
                value = [[e[i] for e in value] for i in range(len(value[0]) if value else 0)]

        return value

    def _write_element(self, value):
        if np and isinstance(value, float) and np.isnan(value):
            return None
        return xlplatform.prepare_xl_data_element(value)

    def write(self, value, rng, options):

        if isinstance(value, (tuple, list)):
            if len(value) == 0:
                return []
            if isinstance(value[0], (tuple, list)):
                value = [
                    [self._write_element(y) for y in x]
                    for x in value
                ]
            else:
                value = [
                    [self._write_element(x) for x in value]
                ]
        else:
            value = [[self._write_element(value)]]

        if options.get('transpose', False):
            value = [[e[i] for e in value] for i in range(len(value[0]) if value else 0)]

        return RangeAccessor.write(self, value, rng, options)


converters[None] = ValueAccessor()


if np:
    class NumpyArrayAccessor(ValueAccessor):

        types = (np.ndarray,)

        def read(self, value, options):
            value = ValueAccessor.read(self, value, Options(options).erase('ndim'))
            if value is None:
                value = np.nan
            elif isinstance(value, list):
                if isinstance(value[0], list):
                    value = [[np.nan if x is None else x for x in i] for i in value]
                else:
                    value = [np.nan if x is None else x for x in value]
            dtype = options.get('dtype', None)
            ndim = options.get('ndim', 0)

            return np.array(value, dtype=dtype, ndmin=ndim)

        def write(self, value, rng, options):
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

            return ValueAccessor.write(self, value, rng, options)

    converters[np.array] = converters[np.ndarray] = NumpyArrayAccessor()


if pd:
    class PandasDataFrameAccessor(ValueAccessor):

        types = (pd.DataFrame,)

        def read(self, value, options):
            value = ValueAccessor.read(self, value, Options(options, ndim=2))
            return pd.DataFrame(value[1:], columns=value[0])

        def write(self, value, rng, options):
            if options.get('index', True):
                if value.index.name in value.columns:
                    # Prevents column name collision when resetting the index
                    value.index.rename(None, inplace=True)
                value = value.reset_index()

            if options.get('header', True):
                if isinstance(value.columns, pd.MultiIndex):
                    columns = list(zip(*value.columns.tolist()))
                else:
                    columns = [value.columns.tolist()]
                value = columns + value.values.tolist()
            else:
                value = value.values.tolist()

            return ValueAccessor.write(self, value, rng, options)

    converters[pd.DataFrame] = PandasDataFrameAccessor()

    class PandaSeriesStage(ValueAccessor):

        types = (pd.Series,)

        def read(self, value, options):
            value = ValueAccessor.read(self, value, Options(options, ndim=1, expand='table'))
            return pd.Series(value[1:])

        def write(self, value, rng, options):
            if options.get('index', True):
                value = value.reset_index().values.tolist()
            else:
                value = value.values[:, np.newaxis].tolist()

            return ValueAccessor.write(self, value, rng, options)

    converters[pd.Series] = PandasSeriesAccessor()


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

    class PandasDataFrameAccessor:

        @staticmethod
        def reader(options):
            return (
                converters[None].reader(
                    Options(options)
                    .override(ndim=2)
                    .defaults(expand='table')
                )
                .add_stage(ConvertPandasDataFrame(options))
            )

        @staticmethod
        def writer(options):
            return (
                converters[None].writer(
                    options
                )
                .add_stage(ConvertPandasDataFrame(options))
            )




converters[pd.Series] = lambda options: (

)


    base = default_pipeline(options, )

    base.add(PandaSeriesStage)



class DictAccessor(ValueAccessor):

    types = (dict,)

    def write(self, value, rng, options):
        return super(DictAccessor, self).write(list(value.items()), rng, options)

    def read(self, value, options):
        value = super(DictAccessor, self).read(value, WithOverrides(options, ndim=2))
        assert len(value[0]) == 2
        return {x[0]: x[1] for x in value}

converters[dict] = DictAccessor()
