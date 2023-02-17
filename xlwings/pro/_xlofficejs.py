"""
Required Notice: Copyright (C) Zoomer Analytics GmbH.

xlwings PRO is dual-licensed under one of the following licenses:

* PolyForm Noncommercial License 1.0.0 (for noncommercial use):
  https://polyformproject.org/licenses/noncommercial/1.0.0
* xlwings PRO License (for commercial use):
  https://github.com/xlwings/xlwings/blob/main/LICENSE_PRO.txt

Commercial licenses can be purchased at https://www.xlwings.org
"""

"""
This engine is only used in connection with Office.js UDFs, not with runPython.
"""

import datetime as dt

try:
    import numpy as np
except ImportError:
    np = None
try:
    import pandas as pd
except ImportError:
    pd = None

from .. import utils


def datetime_to_formatted_number(datetime_object, date_format):
    return {
        "type": "FormattedNumber",
        "basicValue": utils.datetime_to_xlserial(datetime_object),
        "numberFormat": date_format,
    }


def _clean_value_data_element(
    value,
    datetime_builder,
    empty_as,
    number_builder,
    err_to_str,
):
    if value == "":
        return empty_as
    if isinstance(value, str):
        if not err_to_str and value in [
            "#DIV/0!",
            "#N/A",
            "#NAME?",
            "#NULL!",
            "#NUM!",
            "#REF!",
            "#VALUE!",
            "#DATA!",
        ]:
            value = None
        else:
            value = value
    if isinstance(value, dt.datetime) and datetime_builder is not dt.datetime:
        value = datetime_builder(
            month=value.month,
            day=value.day,
            year=value.year,
            hour=value.hour,
            minute=value.minute,
            second=value.second,
            microsecond=value.microsecond,
            tzinfo=None,
        )
    elif number_builder is not None and type(value) == float:
        value = number_builder(value)
    return value


class Engine:
    @staticmethod
    def clean_value_data(data, datetime_builder, empty_as, number_builder, err_to_str):
        # TODO
        return [
            [
                _clean_value_data_element(
                    c, datetime_builder, empty_as, number_builder, err_to_str
                )
                for c in row
            ]
            for row in data
        ]

    @staticmethod
    def prepare_xl_data_element(x, date_format):
        if x is None:
            return ""
        elif pd and pd.isna(x):
            return ""
        elif np and isinstance(x, (np.floating, float)) and np.isnan(x):
            return ""
        elif np and isinstance(x, np.number):
            return float(x)
        elif np and isinstance(x, np.datetime64):
            return datetime_to_formatted_number(
                utils.np_datetime_to_datetime(x), date_format
            )
        elif pd and isinstance(x, pd.Timestamp):
            return datetime_to_formatted_number(x.to_pydatetime(), date_format)
        elif pd and isinstance(x, type(pd.NaT)):
            return None
        elif isinstance(x, (dt.date, dt.datetime)):
            return datetime_to_formatted_number(x, date_format)
        return x

    @property
    def name(self):
        return "officejs"

    @property
    def type(self):
        return "remote"


engine = Engine()
