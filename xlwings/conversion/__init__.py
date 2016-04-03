# -*- coding: utf-8 -*-
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
    convert = options.get('convert', None)
    pipeline = accessors.get(convert, convert).router(value, rng, options).writer(options)
    ctx = ConversionContext(rng=rng, value=value)
    pipeline(ctx)
    return ctx.value
