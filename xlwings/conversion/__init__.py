try:
    import pandas as pd
except ImportError:
    pd = None
try:
    import numpy as np
except ImportError:
    np = None

from .framework import ConversionContext, Options, Pipeline, Converter, accessors, Accessor

from .standard import (DictConverter, Accessor, RangeAccessor, RawValueAccessor, ValueAccessor,
                       AdjustDimensionsStage, CleanDataForWriteStage, CleanDataFromReadStage, Ensure2DStage,
                       ExpandRangeStage, ReadValueFromRangeStage, TransposeStage, WriteValueToRangeStage,
                       Options, Pipeline)
if np:
    from .numpy_conv import NumpyArrayConverter
if pd:
    from .pandas_conv import PandasDataFrameConverter, PandasSeriesConverter


def read(rng, value, options):
    convert = options.get('convert', None)
    pipeline = accessors.get(convert, convert).reader(options)
    ctx = ConversionContext(rng=rng, value=value)
    pipeline(ctx)
    return ctx.value


def write(value, rng, options):
    # Don't allow to write lists and tuples as jagged arrays as appscript and pywin32 don't handle that properly
    # This should really be handled in Ensure2DStage, but we'd have to set the original format in the conversion
    # ctx meta as the check should only run for list and tuples
    if isinstance(value, (list, tuple)) and len(value) > 0 and isinstance(value[0], (list, tuple)):
        first_row = value[0]
        for row in value:
            if len(first_row) != len(row):
                raise Exception('All elements of a 2d list or tuple must be of the same length')
    convert = options.get('convert', None)
    pipeline = accessors.get(convert, convert).router(value, rng, options).writer(options)
    ctx = ConversionContext(rng=rng, value=value)
    pipeline(ctx)
    return ctx.value
