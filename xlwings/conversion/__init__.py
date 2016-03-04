# -*- coding: utf-8 -*-

from .framework import ConversionContext, Options, Pipeline, ConverterAccessor, accessors, Accessor

from . import standard
from . import numpy_conv
from . import pandas_conv


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
