# -*- coding: utf-8 -*-

from .framework import ConversionContext, Options, Pipeline, ConverterAccessor, accessors, Accessor

from .standard import (DictConverter, ConverterAccessor, Accessor, RangeAccessor, RawValueAccessor, ValueAccessor,
                       AdjustDimensionsStage, CleanDataForWriteStage, CleanDataFromReadStage, Ensure2DStage,
                       ExpandRangeStage, ReadValueFromRangeStage, TransposeStage, WriteValueToRangeStage,
                       Options, Pipeline)
from .numpy_conv import NumpyArrayConverter
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
