"""
Required Notice: Copyright (C) Zoomer Analytics GmbH.

xlwings PRO is dual-licensed under one of the following licenses:

* PolyForm Noncommercial License 1.0.0 (for noncommercial use):
  https://polyformproject.org/licenses/noncommercial/1.0.0
* xlwings PRO License (for commercial use):
  https://github.com/xlwings/xlwings/blob/main/LICENSE_PRO.txt

Commercial licenses can be purchased at https://www.xlwings.org
"""

import sys
from itertools import groupby

try:
    import numpy as np
except ImportError:
    np = None
try:
    import pandas as pd
except ImportError:
    pd = None


def _get_filter_value(filter_list, filter_name, default=None):
    for f in filter_list:
        for k, v in f.items():
            if k == filter_name:
                return v[0].as_const()
    return default


# Standard Jinja custom filters
def datetime(value, format=None):
    # Custom Jinja filter that can be used by strings/Markdown
    if format is None:
        # Default format: July 1, 2020
        format = f"%B %{'#' if sys.platform.startswith('win') else '-'}d, %Y"
    return value.strftime(format)


def string_format(value, format):
    return f"{value:{format}}"


# Image filters
def width(filter_list):
    return _get_filter_value(filter_list, "width")


def height(filter_list):
    return _get_filter_value(filter_list, "height")


def scale(filter_list):
    return _get_filter_value(filter_list, "scale")


def image_format(filter_list):
    return _get_filter_value(filter_list, "format", "png")


def top(filter_list):
    return _get_filter_value(filter_list, "top", 0)


def left(filter_list):
    return _get_filter_value(filter_list, "left", 0)


# Font filters
def fontcolor(value=None, filter_list=None):
    # Useful for working with white fonts without them disappearing on screen
    if value:
        # Ignore if called as standard Jinja filter
        return value
    elif filter_list:
        # If called from a single cell/shape placeholder
        color = _get_filter_value(filter_list, "fontcolor")
        colors = {"white": "#ffffff", "black": "#000000"}
        if color.lower() in colors:
            return colors[color.lower()]
        else:
            return color


# DataFrame filters
def sortasc(df, filter_args):
    if df.empty:
        return df
    columns = [arg.as_const() for arg in filter_args]
    return df.sort_values(list(df.columns[columns]), ascending=True)


def sortdesc(df, filter_args):
    if df.empty:
        return df
    columns = [arg.as_const() for arg in filter_args]
    return df.sort_values(list(df.columns[columns]), ascending=False)


def mul(df, filter_args):
    if df.empty:
        return df
    value, col_ix = filter_args[0].as_const(), filter_args[1].as_const()
    fill_value = filter_args[2].as_const() if len(filter_args) > 2 else None
    df.iloc[:, col_ix] = df.iloc[:, col_ix].mul(value, fill_value=fill_value)
    return df


def div(df, filter_args):
    if df.empty:
        return df
    value, col_ix = filter_args[0].as_const(), filter_args[1].as_const()
    fill_value = filter_args[2].as_const() if len(filter_args) > 2 else None
    df.iloc[:, col_ix] = df.iloc[:, col_ix].div(value, fill_value=fill_value)
    return df


def add(df, filter_args):
    if df.empty:
        return df
    value, col_ix = filter_args[0].as_const(), filter_args[1].as_const()
    fill_value = filter_args[2].as_const() if len(filter_args) > 2 else None
    df.iloc[:, col_ix] = df.iloc[:, col_ix].add(value, fill_value=fill_value)
    return df


def sub(df, filter_args):
    if df.empty:
        return df
    value, col_ix = filter_args[0].as_const(), filter_args[1].as_const()
    fill_value = filter_args[2].as_const() if len(filter_args) > 2 else None
    df.iloc[:, col_ix] = df.iloc[:, col_ix].sub(value, fill_value=fill_value)
    return df


def maxrows(df, filter_args):
    if df.empty:
        return df
    if len(df) > filter_args[0].as_const():
        splitrow = filter_args[0].as_const() - 1
        other = df.iloc[splitrow:, :].sum(numeric_only=True)
        other_name = filter_args[1].as_const()
        other.name = other_name
        df = pd.concat([df.iloc[:splitrow, :], other.to_frame().transpose()])
        col_ix = filter_args[2].as_const() if len(filter_args) > 2 else 0
        df.iloc[-1, col_ix] = other_name
    return df


def aggsmall(df, filter_args):
    if df.empty:
        return df
    threshold = filter_args[0].as_const()
    col_ix = filter_args[1].as_const()
    other_name = filter_args[2].as_const()
    other_ix = filter_args[3].as_const() if len(filter_args) > 3 else 0
    min_rows = filter_args[4].as_const() if len(filter_args) > 4 else 1
    df.loc[:, "__is_small__"] = df.iloc[:, col_ix] < threshold
    if min_rows >= len(df):
        df.loc[:, "__is_over_min__"] = False
    else:
        df.loc[:, "__is_over_min__"] = [False] * (min_rows - 1) + [True] * (
            len(df) - min_rows + 1
        )
    df.loc[:, "__total__"] = df["__is_small__"] & df["__is_over_min__"]
    if True in df["__total__"].unique():
        other = df.groupby("__total__").sum(numeric_only=True).loc[True, :]
        other.name = other_name
        df = pd.concat([df.loc[~df["__total__"], :], other.to_frame().transpose()])
        df.iloc[-1, other_ix] = other_name
    df = df.drop(columns=["__is_small__", "__is_over_min__", "__total__"])
    return df


def head(df, filter_args):
    if df.empty:
        return df
    return df.head(filter_args[0].as_const())


def tail(df, filter_args):
    if df.empty:
        return df
    return df.tail(filter_args[0].as_const())


def rowslice(df, filter_args):
    if df.empty:
        return df
    args = [arg.as_const() for arg in filter_args]
    if len(args) == 1:
        args.append(None)
    df = df.iloc[args[0] : args[1], :]
    return df


def colslice(df, filter_args):
    args = [arg.as_const() for arg in filter_args]
    if len(args) == 1:
        args.append(None)
    df = df.iloc[:, args[0] : args[1]]
    return df


def columns(df, filter_args):
    if df.empty:
        return df
    cols = [arg.as_const() for arg in filter_args]
    df = df.iloc[:, [col for col in cols if col is not None]]
    empty_col_indices = [i for i, v in enumerate(cols) if v is None]
    for n, col_ix in enumerate(empty_col_indices):
        # insert() method is inplace!
        # Since Excel tables only allow an empty space once, we'll generate multiple
        # empty spaces for each column.
        df.insert(loc=col_ix, column=" " * (n + 1), value=np.nan)
    return df


def header(df, filter_args):
    # Replace the spaces introduced by a potential previous call of columns()
    # as headers alone can't be used in Excel tables
    return [None if i.isspace() else i for i in df.columns]


# DataFrame formatting filters
def vmerge(df, filter_args, top_left_cell, header):
    """
    vmerge(0, 1) is hierarchical: i.e., it only merges the cells with the same content
    in col 1 as long as they are within a merged cell in col 0.

    To merge columns independently, the filter can be used like this:
    vmerge(0) | vmerge(1)

    vmerge without args merges hierarchically over all columns
    """
    if df.empty:
        return []
    if not filter_args:
        # Default merges hierarchically over all columns
        cols = list(range(len(df.columns)))
    else:
        cols = [arg.as_const() for arg in filter_args]

    merged_cells_count_all = []
    for ix, col in enumerate(cols):
        if ix == 0:
            # ['a', 'a', 'b', 'c', 'c'] = > [2, 1, 2]
            merged_cells_count_origin = [
                sum(1 for _ in group) for _, group in groupby(df.iloc[:, col])
            ]
            merged_cells_count_all.append(merged_cells_count_origin)
        else:
            merged_cells_count_subsection_all = []
            cum_section_len = 0
            for section_len in merged_cells_count_all[ix - 1]:
                merged_cells_count_subsection = [
                    sum(1 for _ in group)
                    for _, group in groupby(
                        df.iloc[cum_section_len : cum_section_len + section_len, col]
                    )
                ]
                merged_cells_count_subsection_all.extend(merged_cells_count_subsection)
                cum_section_len += section_len
            merged_cells_count_all.append(merged_cells_count_subsection_all)

    ranges_to_merge = []
    for ix, col in enumerate(cols):
        ranges_to_merge_col = []
        row_offset = 1 if header else 0
        for count in merged_cells_count_all[ix]:
            if count > 1:
                ranges_to_merge_col.append(
                    top_left_cell.offset(
                        row_offset=row_offset, column_offset=col
                    ).resize(row_size=count)
                )
            ranges_to_merge.extend(ranges_to_merge_col)
            row_offset += count
    return ranges_to_merge


def df_formatter(filter_list):
    from . import format_callbacks

    fname = _get_filter_value(filter_list, "formatter")
    return format_callbacks.get(fname)
