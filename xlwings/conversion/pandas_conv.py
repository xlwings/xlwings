# -*- coding: utf-8 -*-

try:
    import pandas as pd
except ImportError:
    pd = None


if pd:
    import numpy as np
    from . import ConverterAccessor, Options

    class PandasDataFrameConverter(ConverterAccessor):

        writes_types = pd.DataFrame

        @classmethod
        def base_reader(cls, options):
            return (
                super(PandasDataFrameConverter, cls).base_reader(
                    Options(options)
                    .override(ndim=2)
                )
            )

        @classmethod
        def read_value(cls, value, options):
            return pd.DataFrame(value[1:], columns=value[0])

        @classmethod
        def write_value(cls, value, options):
            index = options.get('index', True)
            header = options.get('header', True)

            if index:
                if value.index.name in value.columns:
                    # Prevents column name collision when resetting the index
                    value.index.rename(None, inplace=True)
                value = value.reset_index()

            if header:
                if isinstance(value.columns, pd.MultiIndex):
                    columns = list(zip(*value.columns.tolist()))
                else:
                    columns = [value.columns.tolist()]
                value = columns + value.values.tolist()
            else:
                value = value.values.tolist()

            return value


    PandasDataFrameConverter.install_for(pd.DataFrame)


    class PandasSeriesConverter(ConverterAccessor):

        writes_types = pd.DataFrame

        @classmethod
        def base_reader(cls, options):
            return (
                super(PandasSeriesConverter, cls).base_reader(
                    Options(options)
                    .override(ndim=1)
                )
            )

        @classmethod
        def read_value(cls, value, options):
            return pd.Series(value[1:])

        @classmethod
        def write_value(cls, value, options):
            index = options.get('index', True)

            if index:
                value = value.reset_index().values.tolist()
            else:
                value = value.values[:, np.newaxis].tolist()

            return value


    PandasSeriesConverter.install_for(pd.Series)
