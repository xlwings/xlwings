import sys
import shutil
import datetime as dt

try:
    from jinja2 import Environment, nodes
except ImportError:
    pass

from .markdown import Markdown
from .image import Image
from ..utils import LicenseHandler
from ...main import Book

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

LicenseHandler.validate_license('reports')


def get_filters(ast):
    """This is only for cells that contain a single placeholder.
    Normal text with multiple placeholders could be handled by Jinja's native (custom) filter system.
    Returns var, filter_names (list), arguments (dict)
    """
    found_nodes = list(ast.find_all(node_type=nodes.Filter))
    if found_nodes:
        node = found_nodes[0]
        filters = []
        args = {}
        f = node
        filters.append(f.name)
        args[f.name] = f.args
        while isinstance(f.node, nodes.Filter):
            f = f.node
            filters.append(f.name)
            args[f.name] = f.args
        return f.node.name, list(reversed(filters)), args
    else:
        return None, [], {}


def filter_datetime(value, format=None):
    # Custom Jinja filter that can be used by strings/Markdown
    if format is None:
        # Default format: July 1, 2020
        format = f"%B %{'#' if sys.platform.startswith('win') else '-'}d, %Y"
    return value.strftime(format)


def render_template(sheet, **data):
    """
    Replaces the Jinja2 placeholders in a given sheet
    """
    # On Windows, Excel will not move objects correctly with screen_updating = False during row insert/delete operations
    # So we'll need to set it to True before any such operations. Getting origin state here to revert to.
    book = sheet.book
    screen_updating_original_state = book.app.screen_updating

    # Inserting rows with Frames changes the print area. Get it here so we can revert at the end.
    print_area = sheet.page_setup.print_area

    # A Jinja env defines the placeholder markers and allows to register custom filters
    env = Environment()
    env.filters["datetime"] = filter_datetime

    # used_range doesn't start automatically in A1
    last_cell = sheet.used_range.last_cell
    values_all = sheet.range((1, 1), (last_cell.row, last_cell.column)).options(
        ndim=2).value if sheet.used_range.value else []

    # Frames
    uses_frames = False
    frame_indices = []
    for ix, cell in enumerate(sheet.range((1, 1), (1, last_cell.column))):
        if cell.note:
            if cell.note.text.strip() == '<frame>':
                frame_indices.append(ix)
                uses_frames = True
    frame_indices += [0, last_cell.column]
    frame_indices = list(sorted(set(frame_indices)))
    values_per_frame = []
    for ix in range(len(frame_indices) - 1):
        values_per_frame.append([i[frame_indices[ix]:frame_indices[ix + 1]] for i in values_all])

    # Loop through every cell for each frame
    for ix, values in enumerate(values_per_frame):
        row_shift = 0
        for i, row in enumerate(values):
            for j, value in enumerate(row):
                if isinstance(value, str):
                    if value.count('{{') == 1 and value.startswith('{{') and value.endswith('}}'):
                        # Cell contains single Jinja variable
                        # Handle filters
                        ast = env.parse(value)
                        var, filter_names, filter_args = get_filters(ast)
                        if filter_names:
                            result = env.compile_expression(var)(**data)
                        else:
                            result = env.compile_expression(value.replace('{{', '').replace('}}', '').strip())(**data)
                        if (isinstance(result, Image)
                                or (PIL and isinstance(result, PIL.Image.Image))
                                or (Figure and isinstance(result, Figure))
                                or (plotly and isinstance(result, plotly.graph_objs.Figure))):
                            width = filter_args['width'][0].as_const() if 'width' in filter_names else None
                            height = filter_args['height'][0].as_const() if 'height' in filter_names else None
                            scale = filter_args['scale'][0].as_const() if 'scale' in filter_names else None
                            format_ = filter_args['format'][0].as_const() if 'format' in filter_names else 'png'
                            top = filter_args['top'][0].as_const() if 'top' in filter_names else 0
                            left = filter_args['left'][0].as_const() if 'left' in filter_names else 0
                            image = result.filename if isinstance(result, (Image, PIL.Image.Image)) else result
                            sheet.pictures.add(image,
                                               top=top + sheet[i + row_shift, j + frame_indices[ix]].top,
                                               left=left + sheet[i + row_shift, j + frame_indices[ix]].left,
                                               width=width, height=height, scale=scale, format=format_)
                            sheet[i + row_shift, j + frame_indices[ix]].value = None
                        elif isinstance(result, Markdown):
                            # This will conveniently render placeholders within Markdown instances
                            sheet[i + row_shift,
                                  j + frame_indices[ix]].value = Markdown(text=env.from_string(result.text).render(**data),
                                                                          style=result.style)
                        elif isinstance(result, dt.datetime):
                            # Hack for single cell datetime
                            # Since compile_expression has already run, we need to use value instead of result
                            sheet[i + row_shift, j + frame_indices[ix]].value = env.from_string(value).render(**data)
                        else:
                            # Simple Jinja variables
                            # Check for height of 2d array
                            options = {'index': True, 'header': True}  # defaults
                            if isinstance(result, (list, tuple)) and isinstance(result[0], (list, tuple)):
                                result_len = len(result)
                            elif np and isinstance(result, np.ndarray):
                                result_len = len(result)
                            elif pd and isinstance(result, pd.DataFrame):
                                result = result.copy()  # prevents manipulation of the df in the data dict
                                if 'body' in filter_names:
                                    options = {'index': False, 'header': False}
                                else:
                                    options = {'index': 'noindex' not in filter_names,
                                               'header': 'noheader' not in filter_names}
                                if 'sortasc' in filter_names:
                                    columns = [arg.as_const() for arg in filter_args['sortasc']]
                                    result = result.sort_values(list(result.columns[columns]), ascending=True)
                                if 'sortdesc' in filter_names:
                                    columns = [arg.as_const() for arg in filter_args['sortdesc']]
                                    result = result.sort_values(list(result.columns[columns]), ascending=False)
                                if 'multiply' in filter_names:
                                    multiply_col = filter_args['multiply'][0].as_const()
                                    multiply_val = filter_args['multiply'][1].as_const()
                                    result.iloc[:, multiply_col] = result.iloc[:, multiply_col] * multiply_val
                                if 'divide' in filter_names:
                                    divide_col = filter_args['divide'][0].as_const()
                                    divide_val = filter_args['divide'][1].as_const()
                                    result.iloc[:, divide_col] = result.iloc[:, divide_col] / divide_val
                                if 'add' in filter_names:
                                    add_col = filter_args['add'][0].as_const()
                                    add_val = filter_args['add'][1].as_const()
                                    result.iloc[:, add_col] = result.iloc[:, add_col] + add_val
                                if 'subtract' in filter_names:
                                    subtract_col = filter_args['subtract'][0].as_const()
                                    subtract_val = filter_args['subtract'][1].as_const()
                                    result.iloc[:, subtract_col] = result.iloc[:, subtract_col] - subtract_val
                                if 'maxrows' in filter_names and len(result) > filter_args['maxrows'][0].as_const():
                                    splitrow = filter_args['maxrows'][0].as_const() - 1
                                    other = result.iloc[splitrow:, :].sum(numeric_only=True)
                                    other_name = filter_args['maxrows'][1].as_const()
                                    other.name = other_name
                                    result = result.iloc[:splitrow, :].append(other)
                                    if len(filter_args['maxrows']) > 2:
                                        result.iloc[-1, filter_args['maxrows'][2].as_const()] = other_name
                                if 'aggsmall' in filter_names:
                                    threshold = filter_args['aggsmall'][0].as_const()
                                    col_ix = filter_args['aggsmall'][1].as_const()
                                    dummy_col = '__aggregate__'
                                    result.loc[:, dummy_col] = result.iloc[:, col_ix] < threshold
                                    if True in result[dummy_col].unique():
                                        # unlike aggregate, groupby conveniently drops non-numeric values
                                        other = result.groupby(dummy_col).sum().loc[True, :]
                                        other_name = filter_args['aggsmall'][2].as_const()
                                        other.name = other_name
                                        result = result.loc[result.iloc[:, col_ix] >= threshold, :].append(other)
                                        if len(filter_args['aggsmall']) > 3:
                                            result.iloc[-1, filter_args['aggsmall'][3].as_const()] = other_name
                                    result = result.drop(columns=dummy_col)
                                if 'head' in filter_names:
                                    result = result.head(filter_args['head'][0].as_const())
                                if 'tail' in filter_names:
                                    result = result.tail(filter_args['tail'][0].as_const())
                                if 'rowslice' in filter_names:
                                    args = [arg.as_const() for arg in filter_args['rowslice']]
                                    if len(args) == 1:
                                        args.append(None)
                                    result = result.iloc[args[0]:args[1], :]
                                if 'colslice' in filter_names:
                                    args = [arg.as_const() for arg in filter_args['colslice']]
                                    if len(args) == 1:
                                        args.append(None)
                                    result = result.iloc[:, args[0]:args[1]]
                                if 'columns' in filter_names:
                                    # Must come after maxrows/aggsmall as the duplicate column names would cause issues
                                    columns = [arg.as_const() for arg in filter_args['columns']]
                                    result = result.iloc[:, [col for col in columns if col is not None]]
                                    empty_col_indices = [i for i, v in enumerate(columns) if v is None]
                                    for n, col_ix in enumerate(empty_col_indices):
                                        # insert() method is inplace!
                                        # Since Excel tables only allow an empty space once, we'll generate multiple
                                        # empty spaces for each column.
                                        result.insert(loc=col_ix, column=' ' * (n + 1), value=np.nan)

                                # TODO: handle MultiIndex headers
                                result_len = len(result) + 1 if options['header'] else len(result)
                            else:
                                result_len = 1
                            # Insert rows if within <frame> and 'result' is multiple rows high
                            rows_to_be_inserted = 0
                            if uses_frames and result_len > 1:
                                # Deduct header and first data row that are part of template
                                rows_to_be_inserted = result_len - (2 if options['header'] else 1)
                                if rows_to_be_inserted > 0:
                                    if sys.platform.startswith('win'):
                                        book.app.screen_updating = True
                                    # Since CopyOrigin is not supported on Mac, we start copying two rows
                                    # below the header so the data row formatting gets carried over
                                    start_row = i + row_shift + (3 if options['header'] else 2)
                                    start_col = j + frame_indices[ix] + 1
                                    end_row = i + row_shift + rows_to_be_inserted + (2 if options['header'] else 1)
                                    end_col = frame_indices[ix] + len(values[0])
                                    sheet.range((start_row, start_col),
                                                (end_row, end_col)).insert('down')
                                    # Inserting does not take over borders
                                    sheet.range((start_row - 1, start_col),
                                                (start_row - 1, end_col)).copy()
                                    sheet.range((start_row - 1, start_col),
                                                (end_row, end_col)).paste(paste='formats')
                                    book.app.screen_updating = screen_updating_original_state
                            # Write the 2d array to Excel
                            if sheet[i + row_shift, j + frame_indices[ix]].table:
                                sheet[i + row_shift, j + frame_indices[ix]].table.update(result, index=options['index'])
                            else:
                                sheet[i + row_shift,
                                      j + frame_indices[ix]].options(**options).value = result
                            row_shift += rows_to_be_inserted
                    elif '{{' in value:
                        # These are strings with (multiple) Jinja variables so apply standard text rendering here
                        template = env.from_string(value)
                        sheet[i + row_shift, j + frame_indices[ix]].value = template.render(data)
                    else:
                        # Don't do anything with cells that don't contain any templating so we don't lose the formatting
                        pass

    # Loop through all shapes of interest with a template text
    for shape in [shape for shape in sheet.shapes if shape.type in ('auto_shape', 'text_box')]:
        shapetext = shape.text
        if shapetext and '{{' in shapetext:
            # Single Jinja variable case, the only case we support with Markdown
            if shapetext.count('{{') == 1 and shapetext.startswith('{{') and shapetext.endswith('}}'):
                ast = env.parse(shapetext)
                var, filter_names, filter_args = get_filters(ast)
                if filter_names:
                    result = env.compile_expression(var)(**data)
                else:
                    result = env.compile_expression(shapetext.replace('{{', '').replace('}}', '').strip())(**data)
                if isinstance(result, Markdown):
                    # This will conveniently render placeholders within Markdown text
                    shape.text = Markdown(text=env.from_string(result.text).render(**data),
                                          style=result.style)
                else:
                    # Single Jinja var but no Markdown
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


def create_report(template, output, book_settings=None, app=None, **data):
    """
    This function requires xlwings :guilabel:`PRO`.

    This is a convenience wrapper around :meth:`mysheet.render_template <xlwings.Sheet.render_template>`

    Writes the values of all key word arguments to the ``output`` file according to the ``template`` and the variables
    contained in there (Jinja variable syntax).
    Following variable types are supported:

    strings, numbers, lists, simple dicts, NumPy arrays, Pandas DataFrames, pictures and
    Matplotlib/Plotly figures.

    Parameters
    ----------
    template: str
        Path to your Excel template, e.g. ``r'C:\\Path\\to\\my_template.xlsx'``

    output: str
        Path to your Report, e.g. ``r'C:\\Path\\to\\my_report.xlsx'``

    book_settings: dict, default None
        A dictionary of ``xlwings.Book`` parameters, for details see: :attr:`xlwings.Book`.
        For example: ``book_settings={'update_links': False}``.

    app: xlwings App, default None
        By passing in an xlwings App instance, you can control where your report runs and configure things like ``visible=False``.
        For details see :attr:`xlwings.App`. By default, it creates the
        report in the currently active instance of Excel.

    data: kwargs
        All key/value pairs that are used in the template.

    Returns
    -------
    wb: xlwings Book


    Examples
    --------
    In ``my_template.xlsx``, put the following Jinja variables in two cells: ``{{ title }}`` and ``{{ df }}``

    >>> from xlwings.pro.reports import create_report
    >>> import pandas as pd
    >>> df = pd.DataFrame(data=[[1,2],[3,4]])
    >>> wb = create_report('my_template.xlsx', 'my_report.xlsx', title='MyTitle', df=df)

    With many template variables it may be useful to collect the data first:

    >>> data = dict(title='MyTitle', df=df)
    >>> wb = create_report('my_template.xlsx', 'my_report.xlsx', **data)

    If you need to handle external links or a password, use it like so:

    >>> wb = create_report('my_template.xlsx', 'my_report.xlsx',
                           book_settings={'update_links': True, 'password': 'mypassword'},
                           **data)

    You can control the Excel instance by passing in an xlwings App instance. For example, to
    run the report in a separate and hidden instance of Excel, do the following:

    >>> import xlwings as xw
    >>> from xlwings.pro.reports import create_report
    >>> app = xw.App(visible=False)  # Separate and hidden Excel instance
    >>> wb = create_report('my_template.xlsx', 'my_report.xlsx', app=app, **data)
    >>> app.quit()  # Close the wb and quit the Excel instance
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

    for sheet in wb.sheets:
        render_template(sheet, **data)

    wb.save()
    return wb
