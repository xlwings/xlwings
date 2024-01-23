"""
Required Notice: Copyright (C) Zoomer Analytics GmbH.

xlwings PRO is dual-licensed under one of the following licenses:

* PolyForm Noncommercial License 1.0.0 (for noncommercial use):
  https://polyformproject.org/licenses/noncommercial/1.0.0
* xlwings PRO License (for commercial use):
  https://github.com/xlwings/xlwings/blob/main/LICENSE_PRO.txt

Commercial licenses can be purchased at https://www.xlwings.org
"""

import datetime as dt
import numbers
import shutil
import sys
import warnings

try:
    from jinja2 import Environment, nodes
except ImportError:
    Environment, nodes = None, None

from ...main import Book, XlwingsError
from ..utils import LicenseHandler
from . import filters
from .image import Image
from .markdown import Markdown

try:
    import PIL
    import PIL.Image
except ImportError:
    PIL = None

try:
    from matplotlib.figure import Figure
except ImportError:
    Figure = None

try:
    import numpy as np
except ImportError:
    np = None

try:
    import pandas as pd
except ImportError:
    pd = None

try:
    import plotly
except ImportError:
    plotly = None

LicenseHandler.validate_license("reports")


def parse_single_placeholder(value, env):
    """This is only for cells that contain a single placeholder.
    Text with multiple placeholders is handled by Jinja's native (custom) filter system.
    Returns var, filter_list with filter_name:filter_args (list of dicts)
    """
    ast = env.parse(value)
    found_nodes = list(ast.find_all(node_type=nodes.Filter))
    if found_nodes:
        node = found_nodes[0]
        f = node
        filter_list = [{f.name: f.args}]
        while isinstance(f.node, nodes.Filter):
            f = f.node
            filter_list.insert(0, {f.name: f.args})
        return f.node.name, filter_list
    else:
        return value.replace("{{", "").replace("}}", "").strip(), []


def render_sheet(sheet, **data):
    """
    Replaces the Jinja2 placeholders in a given sheet
    """
    if sheet.name.startswith("##"):
        return
    book = sheet.book

    # Shapes aren't properly moved otherwise
    if sheet.visible:
        # Select fails on Windows for hidden sheets
        sheet.select()

    # Inserting rows with Frames changes the print area.
    # Get it here so we can revert at the end.
    print_area = sheet.page_setup.print_area

    # A Jinja env defines the placeholder markers and allows to register custom filters
    if not Environment:
        raise XlwingsError(
            "Couldn't find the 'jinja2' package, which is required for xlwings Reports."
        )
    env = Environment()
    env.filters["datetime"] = filters.datetime
    env.filters["format"] = filters.string_format  # Overrides Jinja's built-in filter
    env.filters["fontcolor"] = filters.fontcolor

    # used_range doesn't start automatically in A1
    last_cell = sheet.used_range.last_cell
    values_all = (
        sheet.range((1, 1), (last_cell.row, last_cell.column)).options(ndim=2).value
        if sheet.used_range.value
        else []
    )

    # Frames
    uses_frames = False
    frame_indices = []
    for ix, cell in enumerate(sheet.range((1, 1), (1, last_cell.column))):
        if cell.note:
            if "<frame>" in cell.note.text:
                frame_indices.append(ix)
                uses_frames = True
    is_single_frame = True if len(frame_indices) == 1 else False
    frame_indices += [0, last_cell.column]
    frame_indices = list(sorted(set(frame_indices)))
    values_per_frame = []
    for ix in range(len(frame_indices) - 1):
        values_per_frame.append(
            [i[frame_indices[ix] : frame_indices[ix + 1]] for i in values_all]
        )

    # Loop through every cell for each frame
    for ix, values in enumerate(values_per_frame):
        row_shift = 0
        for i, row in enumerate(values):
            for j, value in enumerate(row):
                cell = sheet[i + row_shift, j + frame_indices[ix]]
                if isinstance(value, str):
                    if (
                        value.count("{{") == 1
                        and value.startswith("{{")
                        and value.endswith("}}")
                    ):
                        # Cell contains single Jinja variable
                        var, filter_list = parse_single_placeholder(value, env)
                        result = env.compile_expression(var)(**data)
                        if (
                            isinstance(result, Image)
                            or (PIL and isinstance(result, PIL.Image.Image))
                            or (Figure and isinstance(result, Figure))
                            or (plotly and isinstance(result, plotly.graph_objs.Figure))
                        ):
                            # Image filters: these filters can only be used once. If
                            # supplied multiple times, the first one will be used.
                            width = filters.width(filter_list)
                            height = filters.height(filter_list)
                            scale = filters.scale(filter_list)
                            format_ = filters.image_format(filter_list)
                            top = filters.top(filter_list)
                            left = filters.left(filter_list)

                            image_types = (Image, PIL.Image.Image) if PIL else (Image,)
                            image = (
                                result.filename
                                if isinstance(result, image_types)
                                else result
                            )
                            sheet.pictures.add(
                                image,
                                top=top + cell.top,
                                left=left + cell.left,
                                width=width,
                                height=height,
                                scale=scale,
                                format=format_,
                            )
                            cell.value = None
                        elif isinstance(result, (str, numbers.Number)):
                            if any(["fontcolor" in f for f in filter_list]):
                                cell.font.color = filters.fontcolor(
                                    filter_list=filter_list
                                )
                            cell.value = env.from_string(value).render(**data)
                        elif isinstance(result, Markdown):
                            # This will conveniently render placeholders
                            # within Markdown instances
                            cell.value = Markdown(
                                text=env.from_string(result.text).render(**data),
                                style=result.style,
                            )
                        elif isinstance(result, dt.datetime):
                            cell.value = env.from_string(value).render(**data)
                        else:
                            # Arrays
                            options = {"index": False, "header": True}  # defaults
                            if isinstance(result, (list, tuple)) and isinstance(
                                result[0], (list, tuple)
                            ):
                                result_len = len(result)
                            elif np and isinstance(result, np.ndarray):
                                result_len = len(result)
                            elif pd and isinstance(result, pd.DataFrame):
                                result = (
                                    result.copy()
                                )  # prevents manipulation of the df in the data dict
                                # DataFrame Filters
                                for filter_item in filter_list:
                                    for filter_name, filter_args in filter_item.items():
                                        if filter_name in (
                                            "showindex",
                                            "noheader",
                                            "vmerge",
                                            "formatter",
                                        ):
                                            continue  # handled below
                                        func = getattr(filters, filter_name)
                                        result = func(result, filter_args)
                                # showindex is undocumented
                                # as df.reset_index() is preferred
                                options = {
                                    "index": any(
                                        ["showindex" in f for f in filter_list]
                                    ),
                                    "header": not any(
                                        ["noheader" in f for f in filter_list]
                                    ),
                                }

                                # Assumes 1 header row,
                                # MultiIndex headers aren't supported
                                if any(["header" in f for f in filter_list]):
                                    # Hack for the 'header' filter
                                    result_len = 1
                                else:
                                    result_len = (
                                        len(result) + 1
                                        if options["header"]
                                        else len(result)
                                    )
                            else:
                                result_len = 1
                            # Insert rows if within <frame>
                            # and 'result' is multiple rows high
                            rows_to_be_inserted = 0
                            if uses_frames and result_len > 1:
                                # Deduct header and first data row
                                rows_to_be_inserted = result_len - (
                                    2 if options["header"] else 1
                                )
                                if rows_to_be_inserted > 0:
                                    properties = (
                                        {"screen_updating": True}
                                        if sys.platform.startswith("win")
                                        else {}
                                    )
                                    with book.app.properties(**properties):
                                        # Windows doesn't move objects properly with
                                        # screen_updating=False. Since CopyOrigin is
                                        # not supported on Mac, we start copying two
                                        # rows below the header so the data row
                                        # formatting gets carried over.
                                        start_row = (
                                            i
                                            + row_shift
                                            + (3 if options["header"] else 2)
                                        )
                                        start_col = j + frame_indices[ix] + 1
                                        end_row = (
                                            i
                                            + row_shift
                                            + rows_to_be_inserted
                                            + (2 if options["header"] else 1)
                                        )
                                        end_col = frame_indices[ix] + len(values[0])
                                        if is_single_frame:
                                            # This will preserve the row height of rows
                                            # below the inserted ones
                                            sheet.range(
                                                f"{start_row}:{end_row}"
                                            ).insert("down")
                                        else:
                                            sheet.range(
                                                (start_row, start_col),
                                                (end_row, end_col),
                                            ).insert("down")
                                        # Inserting does not take over borders and
                                        # copy/paste format can cause conflicts with
                                        # other processes that use the clipboard
                                        origin = sheet.range(
                                            (start_row - 1, start_col),
                                            (start_row - 1, end_col),
                                        )
                                        destination = sheet.range(
                                            (start_row - 1, start_col),
                                            (end_row, end_col),
                                        )
                                        if not origin.table:
                                            origin.autofill(destination, "fill_formats")
                            # Write the array to Excel
                            if cell.table:
                                cell.table.update(result, index=options["index"])
                            else:
                                df_formatter = filters.df_formatter(filter_list)
                                if df_formatter:
                                    options["formatter"] = df_formatter
                                cell.options(chunksize=10000, **options).value = result
                            # DataFrame formatting filters
                            for filter_item in filter_list:
                                for filter_name, filter_args in filter_item.items():
                                    if filter_name == "vmerge":
                                        merge_ranges = filters.vmerge(
                                            result, filter_args, cell, options["header"]
                                        )
                                        for merge_range in merge_ranges:
                                            merge_range.merge()
                            row_shift += rows_to_be_inserted
                    elif "{{" in value:
                        # These are strings with (multiple) Jinja variables so apply
                        # standard text rendering here
                        template = env.from_string(value)
                        cell.value = template.render(data)
                    else:
                        # Don't do anything with cells that don't contain any templating
                        # so we don't lose the formatting
                        pass

    # Loop through all shapes of interest with a template text
    for shape in [
        shape for shape in sheet.shapes if shape.type in ("auto_shape", "text_box")
    ]:
        shapetext = shape.text
        if shapetext and "{{" in shapetext:
            if (
                shapetext.count("{{") == 1
                and shapetext.startswith("{{")
                and shapetext.endswith("}}")
            ):
                # Single Jinja variable case, the only case we support with Markdown
                var, filter_list = parse_single_placeholder(shapetext, env)
                result = env.compile_expression(var)(**data)
                if isinstance(result, Markdown):
                    # This will conveniently render placeholders within Markdown text
                    shape.text = Markdown(
                        text=env.from_string(result.text).render(**data),
                        style=result.style,
                    )
                else:
                    # Single Jinja var but no Markdown
                    if any(["fontcolor" in f for f in filter_list]):
                        shape.font.color = filters.fontcolor(filter_list=filter_list)
                    template = env.from_string(shapetext)
                    shape.text = template.render(data)
            else:
                # Multiple Jinja vars and no Markdown
                template = env.from_string(shapetext)
                shape.text = template.render(data)

    # Copy/pasting the formatting leaves ranges selected.
    book.app.cut_copy_mode = False

    # Reset print area
    if print_area:
        sheet.page_setup.print_area = print_area

    try:
        sheet["A1"].select()
    except:  # noqa: E722
        pass


def create_report(template=None, output=None, book_settings=None, **data):
    warnings.warn("Deprecated. Use render_template instead.")
    return render_template(
        template=template, output=output, book_settings=book_settings, **data
    )


def render_template(template, output, book_settings=None, app=None, **data):
    """
    This function requires xlwings :guilabel:`PRO`.

    This is a convenience wrapper around
    :meth:`mysheet.render_template <xlwings.Sheet.render_template>`

    Writes the values of all key word arguments to the ``output`` file according to the
    ``template`` and the variables contained in there (Jinja variable syntax).
    Following variable types are supported:

    strings, numbers, lists, simple dicts, NumPy arrays, Pandas DataFrames, pictures and
    Matplotlib/Plotly figures.

    Parameters
    ----------
    template: str or path-like
        Path to your Excel template, e.g. ``r'C:\\Path\\to\\my_template.xlsx'``

    output: str or path-like
        Path to your Report, e.g. ``r'C:\\Path\\to\\my_report.xlsx'``

    book_settings: dict, default None
        A dict of ``xlwings.Book`` parameters, for details see: :attr:`xlwings.Book`.
        For example: ``book_settings={'update_links': False}``.

    app: xlwings App, default None
        By passing in an xlwings App instance, you can control where your report runs
        and configure things like ``visible=False``. For details see
        :attr:`xlwings.App`. By default, it creates the report in the currently active
        instance of Excel.

    data: kwargs
        All key/value pairs that are used in the template.

    Returns
    -------
    xlwings Book


    Examples
    --------
    In ``my_template.xlsx``, put the following Jinja variables in two cells:
    ``{{ title }}`` and ``{{ df }}``

    >>> from xlwings.reports import render_template
    >>> import pandas as pd
    >>> df = pd.DataFrame(data=[[1,2],[3,4]])
    >>> mybook = render_template('my_template.xlsx', 'my_report.xlsx',
                                 title='MyTitle', df=df)

    With many template variables it may be useful to collect the data first:

    >>> data = dict(title='MyTitle', df=df)
    >>> mybook = render_template('my_template.xlsx', 'my_report.xlsx', **data)

    If you need to handle external links or a password, use it like so:

    >>> mybook = render_template('my_template.xlsx', 'my_report.xlsx',
                                 book_settings={'update_links': True,
                                 'password': 'mypassword'}, **data)
    """
    shutil.copyfile(template, output)
    if app:
        if book_settings:
            wb = app.books.open(output, **book_settings)
        else:
            wb = app.books.open(output)
    else:
        # Use existing Excel instance or create a new one if there is none
        if book_settings:
            wb = Book(output, **book_settings)
        else:
            wb = Book(output)

    for sheet in reversed(wb.sheets):
        render_sheet(sheet, **data)

    wb.save()
    return wb
