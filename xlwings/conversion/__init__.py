from contextlib import suppress
from typing import Any, Mapping, Type

try:
    import pandas as pd
except ImportError:
    pd = None
try:
    import numpy as np
except ImportError:
    np = None
try:
    import polars as pl
except ImportError:
    pl = None

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
if pl:
    from .polars_conv import PolarsDataFrameConverter, PolarsSeriesConverter

from .. import LicenseError

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
    "PolarsDataFrameConverter",
    "PolarsSeriesConverter",
)


def _get_accessor(
    *,
    convert: Any,
    default: Type[Accessor],
    registered: Mapping[Any, Type[Accessor]],
) -> Type[Accessor]:
    """
    Get an Accessor class based on the user-provided `convert` option.

    Args:
        convert:
            The user-provided `convert` range option.
            This may be one of:
                - A registered name that maps to an Accessor class.
                - A registered type that maps to an Accessor class.
                - An accessor class (which will be returned directly).
                - None, in which case the default accessor is used.
        default:
            The default Accessor class to use if `convert` is None or not found.
        registered:
            A mapping of registered Accessor classes.
            The keys may be accessor names or classes, and the values should be the corresponding Accessor classes.

    Returns:
        A registered accessor instance.
    """
    if convert is None:
        return default

    with suppress(TypeError):
        if issubclass(convert, Accessor):
            return convert

    accessor = registered.get(convert, default)

    if not issubclass(accessor, Accessor):
        raise RuntimeError(f"Accessor `{convert}` is not a subclass of `Accessor`.")

    return accessor


def read(rng, value, options, engine_name=None):
    accessor = _get_accessor(
        convert=options.get("convert"),
        default=ValueAccessor,
        registered=accessors,
    )
    pipeline = accessor.reader(options)
    ctx = ConversionContext(rng=rng, value=value, engine_name=engine_name)
    pipeline(ctx)
    return ctx.value


def write(value, rng, options, engine_name=None):
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
    accessor = _get_accessor(
        convert=options.get("convert"),
        default=ValueAccessor,
        registered=accessors,
    )
    pipeline = accessor.router(value, rng, options).writer(options)
    ctx = ConversionContext(rng=rng, value=value, engine_name=engine_name)
    pipeline(ctx)
    return ctx.value
