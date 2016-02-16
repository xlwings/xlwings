# -*- coding: utf-8 -*-

from .framework import ConversionContext, Options, Pipeline, ConverterAccessor, accessors, Accessor

from . import standard
from . import numpy_conv
from . import pandas_conv


def read(rng, value, options):
    as_ = options.get('as_', None)
    pipeline = accessors.get(as_, as_).reader(options)
    ctx = ConversionContext(rng=rng, value=value)
    pipeline(ctx)
    return ctx.value


def write(value, rng, options):
    as_ = options.get('as_', None)
    pipeline = accessors.get(as_, as_).router(value, rng, options).writer(options)
    ctx = ConversionContext(rng=rng, value=value)
    pipeline(ctx)
    return ctx.value
