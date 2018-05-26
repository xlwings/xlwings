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

    from . import Converter, Options
    from ..utils import np_datetime_to_datetime

    NAN = ""


    class NumpyArrayConverter(Converter):

        writes_types = np.ndarray

        @classmethod
        def base_reader(cls, options):
            return (
                super(NumpyArrayConverter, cls).base_reader(
                    Options(options)
                    .defaults(empty=np.nan)
                )
            )

        @classmethod
        def read_value(cls, value, options):
            dtype = options.get('dtype', None)
            copy = options.get('copy', True)
            order = options.get('order', None)
            ndim = options.get('ndim', None) or 0
            return np.array(value, dtype=dtype, copy=copy, order=order, ndmin=ndim)

        @classmethod
        def write_value(cls, value, options):
            return value.tolist()


    NumpyArrayConverter.register(np.array, np.ndarray)


    class NumpyDateConverter(Converter):

        write_types = np.datetime64

        @classmethod
        def write_value(cls, value, options):
            return np_datetime_to_datetime(value)


    NumpyDateConverter.register(np.datetime64)


    class NumpyFloatConverter(Converter):
        write_types = float

        @classmethod
        def write_value(cls, value, options):
            if np.isnan(value):
                return NAN
            else:
                return value


    NumpyFloatConverter.register(float)


    class NumpyNumberConverter(Converter):

        base_type = float

        @classmethod
        def read_value(cls, value, options):
            if value == NAN:
                return np.NaN
            else:
                return value

        @classmethod
        def write_value(cls, value, options):
            return super(NumpyNumberConverter, cls).write_value(float(value), options)


    NumpyNumberConverter.register(np.number)
