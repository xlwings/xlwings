from .. import xlplatform
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
    datetime.datetime: datetime,
    datetime.date: lambda year, month, day, **kwargs: datetime.date(year, month, day)
}


class DefaultAccessor(object):

    @classmethod
    def read_range(cls, rng, options):
        value = xlplatform.get_value_from_range(rng.xl_range)
        return cls.read_value(value, options)

    @classmethod
    def read_value(cls, value, options):
        value = xlplatform.clean_value_data(value, _date_handlers[options.get('dates_as', datetime.datetime)])
        converter = converters[options.get('read_as', None)]
        return converter.read(value, options)

    @classmethod
    def write_range(cls, rng, value, options):
        value = cls.write_value(value, options)

        if isinstance(value, (tuple, list)):
            if len(value) == 0:
                return
            if isinstance(value[0], (tuple, list)):
                row2 = rng.row1 + len(value) - 1
                col2 = rng.col1 + len(value[0]) - 1
            else:
                row2 = rng.row1
                col2 = rng.col1 + len(value) - 1
                value = [value]
        else:
            row2 = rng.row2
            col2 = rng.col2

        xlplatform.set_value(xlplatform.get_range_from_indices(rng.xl_sheet, rng.row1, rng.col1, row2, col2), value)

    @classmethod
    def write_value(cls, value, options):
        converter = converters.get(type(value), converters[None])
        value = xlplatform.prepare_xl_data(value)
        return converter.write(value, options)


class DefaultConverter(object):

    def _ensure_dimensionality(self, value, ndim):

        if ndim is None:
            if isinstance(value, (list, tuple)):
                if value and isinstance(value[0], (list, tuple)):
                    if len(value) == 1:
                        return value[0]
                    elif len(value[0]) == 1:
                        return [x[0] for x in value]
            return value

        if ndim == 1:
            if isinstance(value, (list, tuple)):
                if len(value) > 0 and isinstance(value[0], (list, tuple)):
                    if len(value) == 1:
                        return value[0]
                    elif len(value[0]) == 1:
                        return [x[0] for x in value]
                    else:
                        raise Exception("Range must be 1-by-n or n-by-1 when ndim=1.")
                else:
                    return value
            else:
                return [value]

        if ndim == 2:
            if isinstance(value, (list, tuple)):
                if len(value) > 0 and isinstance(value[0], (list, tuple)):
                    return value
                else:
                    return [value]
            else:
                return [[value]]

        raise ValueError('Invalid value ndim=%s' % ndim)

    def read(self, value, options):

        ndim = getattr(self, 'ndim', None) or options.get('ndim', None)
        value = self._ensure_dimensionality(value, ndim)

        if options.get('transpose', False):
            if value and isinstance(value, (list, tuple)) and isinstance(value[0], (list, tuple)):
                value = [[e[i] for e in value] for i in range(len(value[0]) if value else 0)]

        return value

    def write(self, value, options):

        if options.get('transpose', False):
            value = [[e[i] for e in value] for i in range(len(value[0]) if value else 0)]

        return value

converters[None] = DefaultConverter()


if np:
    class NumpyArrayConverter(DefaultConverter):

        def read(self, value, options):
            return np.array(DefaultConverter.read(self, value, options))

        def write(self, value, options):
            return DefaultConverter.write(self, value.tolist(), options)

    converters[np.array] = NumpyArrayConverter()


if pd:
    class PandasDataFrameConverter(DefaultConverter):

        def read(self, value, options):
            return np.array(DefaultConverter.read(self, value, options))

        def write(self, value, options):
            return DefaultConverter.write(self, value.tolist(), options)

    converters[pd.DataFrame] = PandasDataFrameConverter()