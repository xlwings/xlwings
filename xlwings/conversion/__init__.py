# -*- coding: utf-8 -*-

from .framework import ConversionContext, Options, Pipeline, ConverterAccessor, converters, Accessor

from . import standard
from . import numpy_conv
from . import pandas_conv


def read_from_range(rng, options):
    as_ = options.get('as_', None)
    pipeline = converters.get(as_, as_).reader(options)
    ctx = ConversionContext(range=rng, value=None)
    pipeline(ctx)
    return ctx.value


def write_to_range(value, rng, options):
    as_ = options.get('as_', None)
    pipeline = converters.get(as_, as_).router(value, rng, options).writer(options)
    pipeline(ConversionContext(range=rng, value=value))
