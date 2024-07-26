try:
    import pandas as pd
except ImportError:
    pd = None
try:
    import numpy as np
except ImportError:
    np = None

from .framework import (
    Accessor,
    ConversionContext,
    Converter,
    Options,
    Pipeline,
    accessors,
)
from .standard import (
    AdjustDimensionsStage,
    CleanDataForWriteStage,
    CleanDataFromReadStage,
    DictConverter,
    Ensure2DStage,
    ExpandRangeStage,
    RangeAccessor,
    RawValueAccessor,
    ReadValueFromRangeStage,
    TransposeStage,
    ValueAccessor,
    WriteValueToRangeStage,
)

if np:
    from .numpy_conv import NumpyArrayConverter
if pd:
    from .pandas_conv import PandasDataFrameConverter, PandasSeriesConverter

from .. import LicenseError
from ..utils import await_me_maybe

try:
    from ..pro.reports.markdown import Markdown, MarkdownConverter

    MarkdownConverter.register(Markdown)
except (ImportError, LicenseError, AttributeError):
    pass


__all__ = (
    "Accessor",
    "ConversionContext",
    "Converter",
    "Options",
    "Pipeline",
    "accessors",
    "AdjustDimensionsStage",
    "CleanDataForWriteStage",
    "CleanDataFromReadStage",
    "DictConverter",
    "Ensure2DStage",
    "ExpandRangeStage",
    "RangeAccessor",
    "RawValueAccessor",
    "ReadValueFromRangeStage",
    "TransposeStage",
    "ValueAccessor",
    "WriteValueToRangeStage",
    "NumpyArrayConverter",
    "PandasDataFrameConverter",
    "PandasSeriesConverter",
)

import asyncio
from functools import wraps


def dual_mode(func):
    @wraps(func)
    def sync_wrapper(*args, **kwargs):
        # Check if there is an active event loop
        try:
            loop = asyncio.get_running_loop()
            # If there's a running loop, schedule the coroutine in it
            return loop.run_until_complete(func(*args, **kwargs))
        except RuntimeError:
            # No running loop, safe to use asyncio.run
            return asyncio.run(func(*args, **kwargs))

    @wraps(func)
    async def async_wrapper(*args, **kwargs):
        return await func(*args, **kwargs)

    sync_wrapper.async_version = async_wrapper
    return sync_wrapper


@dual_mode
async def read(rng, value, options, engine_name=None):
    convert = options.get("convert", None)
    pipeline = accessors.get(convert, convert).reader(options)
    ctx = ConversionContext(rng=rng, value=value, engine_name=engine_name)
    await await_me_maybe(pipeline, ctx)
    return ctx.value


async def write(value, rng, options, engine_name=None):
    # Don't allow to write lists and tuples as jagged arrays as appscript and pywin32
    # don't handle that properly. This should really be handled in Ensure2DStage, but
    # we'd have to set the original format in the conversion ctx meta as the check
    # should only run for list and tuples.
    if (
        isinstance(value, (list, tuple))
        and len(value) > 0
        and isinstance(value[0], (list, tuple))
    ):
        first_row = value[0]
        for row in value:
            if len(first_row) != len(row):
                raise Exception(
                    "All elements of a 2d list or tuple must be of the same length"
                )
    convert = options.get("convert", None)
    pipeline = (
        accessors.get(convert, convert).router(value, rng, options).writer(options)
    )
    ctx = ConversionContext(rng=rng, value=value, engine_name=engine_name)
    await await_me_maybe(pipeline, ctx)
    return ctx.value
