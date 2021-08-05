import sys

try:
    import numpy as np
except ImportError:
    np = None


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


def fmt(value, format):
    return f'{value:{format}}'


# Image filters
def width(filter_list):
    return _get_filter_value(filter_list, 'width')


def height(filter_list):
    return _get_filter_value(filter_list, 'height')


def scale(filter_list):
    return _get_filter_value(filter_list, 'scale')


def image_format(filter_list):
    return _get_filter_value(filter_list, 'format', 'png')


def top(filter_list):
    return _get_filter_value(filter_list, 'top', 0)


def left(filter_list):
    return _get_filter_value(filter_list, 'left', 0)


# Font filters
def fontcolor(value=None, filter_list=None):
    # Useful for working with white fonts without them disappearing on screen
    if value:
        # Ignore if called as standard Jinja filter
        return value
    elif filter_list:
        # If called from a single cell/shape placeholder
        color = _get_filter_value(filter_list, 'fontcolor')
        colors = {'white': '#ffffff', 'black': '#000000'}
        if color.lower() in colors:
            return colors[color.lower()]
        else:
            return color


# DataFrame filters
def sortasc(df, filter_args):
    columns = [arg.as_const() for arg in filter_args]
    return df.sort_values(list(df.columns[columns]), ascending=True)


def sortdesc(df, filter_args):
    columns = [arg.as_const() for arg in filter_args]
    return df.sort_values(list(df.columns[columns]), ascending=False)


def mul(df, filter_args):
    value, col_ix = filter_args[0].as_const(), filter_args[1].as_const()
    fill_value = filter_args[2].as_const() if len(filter_args) > 2 else None
    df.iloc[:, col_ix] = df.iloc[:, col_ix].mul(value, fill_value=fill_value)
    return df


def div(df, filter_args):
    value, col_ix = filter_args[0].as_const(), filter_args[1].as_const()
    fill_value = filter_args[2].as_const() if len(filter_args) > 2 else None
    df.iloc[:, col_ix] = df.iloc[:, col_ix].div(value, fill_value=fill_value)
    return df


def add(df, filter_args):
    value, col_ix = filter_args[0].as_const(), filter_args[1].as_const()
    fill_value = filter_args[2].as_const() if len(filter_args) > 2 else None
    df.iloc[:, col_ix] = df.iloc[:, col_ix].add(value, fill_value=fill_value)
    return df


def sub(df, filter_args):
    value, col_ix = filter_args[0].as_const(), filter_args[1].as_const()
    fill_value = filter_args[2].as_const() if len(filter_args) > 2 else None
    df.iloc[:, col_ix] = df.iloc[:, col_ix].sub(value, fill_value=fill_value)
    return df


def maxrows(df, filter_args):
    if len(df) > filter_args[0].as_const():
        splitrow = filter_args[0].as_const() - 1
        other = df.iloc[splitrow:, :].sum(numeric_only=True)
        other_name = filter_args[1].as_const()
        other.name = other_name
        df = df.iloc[:splitrow, :].append(other)
        col_ix = filter_args[2].as_const() if len(filter_args) > 2 else 0
        df.iloc[-1, col_ix] = other_name
    return df


def aggsmall(df, filter_args):
    threshold = filter_args[0].as_const()
    col_ix = filter_args[1].as_const()
    dummy_col = '__aggregate__'
    df.loc[:, dummy_col] = df.iloc[:, col_ix] < threshold
    if True in df[dummy_col].unique():
        # unlike aggregate, groupby conveniently drops non-numeric values
        other = df.groupby(dummy_col).sum().loc[True, :]
        other_name = filter_args[2].as_const()
        other.name = other_name
        df = df.loc[df.iloc[:, col_ix] >= threshold, :].append(other)
        other_ix = filter_args[3].as_const() if len(filter_args) > 3 else 0
        df.iloc[-1, other_ix] = other_name
    df = df.drop(columns=dummy_col)
    return df


def head(df, filter_args):
    return df.head(filter_args[0].as_const())


def tail(df, filter_args):
    return df.tail(filter_args[0].as_const())


def rowslice(df, filter_args):
    args = [arg.as_const() for arg in filter_args]
    if len(args) == 1:
        args.append(None)
    df = df.iloc[args[0]:args[1], :]
    return df


def colslice(df, filter_args):
    args = [arg.as_const() for arg in filter_args]
    if len(args) == 1:
        args.append(None)
    df = df.iloc[:, args[0]:args[1]]
    return df


def columns(df, filter_args):
    columns = [arg.as_const() for arg in filter_args]
    df = df.iloc[:, [col for col in columns if col is not None]]
    empty_col_indices = [i for i, v in enumerate(columns) if v is None]
    for n, col_ix in enumerate(empty_col_indices):
        # insert() method is inplace!
        # Since Excel tables only allow an empty space once, we'll generate multiple
        # empty spaces for each column.
        df.insert(loc=col_ix, column=' ' * (n + 1), value=np.nan)
    return df


def header(df, filter_args):
    # Replace the spaces introduced by a potential previous call of columns()
    # as headers alone can't be used in Excel tables
    return [None if i.isspace() else i for i in df.columns]
