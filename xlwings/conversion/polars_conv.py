from ..utils import xlserial_to_datetime

try:
    import polars as pl
except ImportError:
    pl = None


if pl:
    from . import Converter, Options

    def _parse_dates(df, parse_dates):
        # Office.js UDFs don't have the info whether the cell is in date format
        if not isinstance(parse_dates, list):
            parse_dates = [parse_dates]

        result = df  # Initialize result with the input dataframe

        for col in parse_dates:
            if isinstance(col, str):
                # For named columns
                result = result.with_columns(
                    pl.col(col).map_elements(
                        xlserial_to_datetime, return_dtype=pl.Datetime
                    )
                )
            else:
                # For integer column index
                col_name = result.columns[col]
                result = result.with_columns(
                    pl.col(col_name).map_elements(
                        xlserial_to_datetime, return_dtype=pl.Datetime
                    )
                )

        return result

    class PolarsDataFrameConverter(Converter):
        @classmethod
        def base_reader(cls, options):
            return super(PolarsDataFrameConverter, cls).base_reader(
                Options(options).override(ndim=2)
            )

        @classmethod
        def read_value(cls, value, options):
            has_header = options.get("has_header", options.get("header", True))
            schema = options.get("schema")
            schema_overrides = options.get("schema_overrides")
            strict = options.get("strict", True)
            infer_schema_length = options.get("infer_schema_length", 100)
            nan_to_null = options.get("nan_to_null", False)
            parse_dates = options.get("parse_dates")

            if has_header:
                df = pl.DataFrame(
                    data=value[1:],
                    schema=value[0],
                    orient="row",
                    schema_overrides=schema_overrides,
                    strict=strict,
                    infer_schema_length=infer_schema_length,
                    nan_to_null=nan_to_null,
                )
            else:
                df = pl.DataFrame(
                    data=value,
                    schema=schema,
                    orient="row",
                    schema_overrides=schema_overrides,
                    strict=strict,
                    infer_schema_length=infer_schema_length,
                    nan_to_null=nan_to_null,
                )

            if parse_dates is not None:
                df = _parse_dates(df, parse_dates)
            return df

        @classmethod
        def write_value(cls, value, options):
            df = value
            result = [df.columns]
            result.extend([list(row) for row in df.rows()])
            return result

    PolarsDataFrameConverter.register(pl.DataFrame)

    class PolarsSeriesConverter(Converter):
        @classmethod
        def read_value(cls, value, options):
            has_header = options.get("has_header", options.get("header", True))
            dtype = options.get("dtype")
            strict = options.get("strict", True)
            nan_to_null = options.get("nan_to_null", False)
            parse_dates = options.get("parse_dates")

            if has_header:
                series = pl.Series(
                    name=value[0],
                    values=value[1:],
                    dtype=dtype,
                    strict=strict,
                    nan_to_null=nan_to_null,
                )
            else:
                series = pl.Series(
                    values=value,
                    dtype=dtype,
                    strict=strict,
                    nan_to_null=nan_to_null,
                )
            if parse_dates is not None:
                df = series.to_frame()
                df = _parse_dates(df, parse_dates)
                series = df.to_series()

            return series

        @classmethod
        def write_value(cls, value, options):
            has_header = options.get("has_header", options.get("header", True))
            series = value
            if series.name and has_header:
                values = [series.name] + series.to_list()
            else:
                values = series.to_list()
            return [[item] for item in values]

    PolarsSeriesConverter.register(pl.Series)
