from ..utils import xlserial_to_datetime

try:
    import pandas as pd
except ImportError:
    pd = None


if pd:
    from . import Converter, Options

    def _parse_dates(df, parse_dates):
        # Office.js UDFs don't have the info whether the cell is in date format
        if parse_dates is True:
            parse_dates = [0]
        elif not isinstance(parse_dates, list):
            parse_dates = [parse_dates]
        for col in parse_dates:
            if isinstance(col, str):
                df.loc[:, col] = df.loc[:, col].apply(xlserial_to_datetime)
            else:
                df.iloc[:, col] = df.iloc[:, col].apply(xlserial_to_datetime)
        return df

    def write_value(cls, value, options):
        index = options.get("index", True)
        header = options.get("header", True)
        assign_empty_index_names = options.get("assign_empty_index_names", False)

        index_names = value.index.names
        if assign_empty_index_names:
            # Useful when you want to have your DataFrame formatted as an Excel table
            # which requires column header names. Since Excel tables only allow an empty
            # space once, we'll generate multiple empty spaces for each column.
            index_names = [
                " " * (i + 1) if name is None else name
                for i, name in enumerate(index_names)
            ]
        else:
            index_names = ["" if name is None else name for name in index_names]
        index_levels = len(index_names)

        if index:
            if value.index.name in value.columns:
                # Prevents column name collision when resetting the index
                value.index = value.index.rename(None)
            value = value.reset_index()

        # Convert pandas-specific types without corresponding Excel type to strings
        for ix, col in enumerate(value.columns):
            if (
                isinstance(value.iloc[:, ix].dtype, pd.PeriodDtype)
                or value.iloc[:, ix].dtype == "timedelta64[ns]"
            ):
                value.iloc[:, ix] = value.iloc[:, ix].astype(str)

        if header:
            if isinstance(value.columns, pd.MultiIndex):
                columns = list(zip(*value.columns.tolist()))
                columns = [list(i) for i in columns]
                # Move index names right above the index
                if index:
                    for c in columns[:-1]:
                        c[:index_levels] = [""] * index_levels
                    columns[-1][:index_levels] = index_names
            else:
                columns = [value.columns.tolist()]
                if index:
                    columns[0][:index_levels] = index_names
            value = columns + value.values.tolist()
        else:
            value = value.values.tolist()

        return value

    class PandasDataFrameConverter(Converter):
        @classmethod
        def base_reader(cls, options):
            return super(PandasDataFrameConverter, cls).base_reader(
                Options(options).override(ndim=2)
            )

        @classmethod
        def read_value(cls, value, options):
            index = options.get("index", 1)
            header = options.get("header", 1)
            dtype = options.get("dtype", None)
            copy = options.get("copy", False)
            parse_dates = options.get("parse_dates", None)

            # build dataframe with only columns (no index) but correct header
            if header == 1:
                columns = pd.Index(value[0])
            elif header > 1:
                columns = pd.MultiIndex.from_arrays(value[:header])
            else:
                columns = None

            df = pd.DataFrame(value[header:], columns=columns, dtype=dtype, copy=copy)

            if parse_dates is not None:
                df = _parse_dates(df, parse_dates)

            # handle index by resetting the index to the index first columns
            # and renaming the index according to the name in the last row
            if index > 0:
                # rename uniquely the index columns to some never used name for column
                # we do not use the column name directly as it would cause issues if
                # several columns have the same name
                df.columns = pd.Index(range(len(df.columns)))
                df = df.set_index(list(df.columns)[:index])

                df.index.names = pd.Index(
                    value[header - 1][:index] if header else [None] * index
                )

                if header:
                    df.columns = columns[index:]
                else:
                    df.columns = pd.Index(range(len(df.columns)))

            return df

        @classmethod
        def write_value(cls, value, options):
            return write_value(cls, value, options)

    PandasDataFrameConverter.register(pd.DataFrame, "df")

    class PandasSeriesConverter(Converter):
        @classmethod
        def read_value(cls, value, options):
            index = options.get("index", 1)
            header = options.get("header", True)
            dtype = options.get("dtype", None)
            copy = options.get("copy", False)
            parse_dates = options.get("parse_dates", None)

            if header:
                columns = value[0]
                if not isinstance(columns, list):
                    columns = [columns]
                data = value[1:]
            else:
                columns = None
                data = value

            df = pd.DataFrame(data, columns=columns, dtype=dtype, copy=copy)

            if parse_dates is not None:
                df = _parse_dates(df, parse_dates)

            if index:
                df.columns = pd.Index(range(len(df.columns)))
                df = df.set_index(list(df.columns)[:index])
                df.index.names = pd.Index(
                    value[header - 1][:index] if header else [None] * index
                )

            if header:
                df.columns = columns[index:]
            else:
                df.columns = pd.Index(range(len(df.columns)))

            series = df.squeeze()

            if not header:
                series.name = None
                series.index.name = None

            return series

        @classmethod
        def write_value(cls, value, options):
            if all(v is None for v in value.index.names) and value.name is None:
                default_header = False
            else:
                default_header = True

            options["header"] = options.get("header", default_header)
            values = write_value(cls, value.to_frame(), options)
            return values

    PandasSeriesConverter.register(pd.Series)
