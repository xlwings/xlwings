# -*- coding: utf-8 -*-

try:
    import numpy as np
except ImportError:
    np = None


if np:

    try:
        import pandas as pd
    except ImportError:
        pd = None

    from . import ConverterAccessor, Options

    class NumpyArrayConverter(ConverterAccessor):

        writes_types = np.ndarray

        @classmethod
        def base_reader(cls, options):
            return (
                super(NumpyArrayConverter, cls).base_reader(
                    Options(options)
                    .defaults(empty_as=np.nan)
                )
            )

        @classmethod
        def read_value(cls, value, options):
            dtype = options.get('dtype', None)
            ndim = options.get('ndim', None) or 0
            return np.array(value, dtype=dtype, ndmin=ndim)

        @classmethod
        def write_value(cls, value, options):
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


    NumpyArrayConverter.install_for(np.array, np.ndarray)
