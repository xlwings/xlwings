import sys
import shutil

try:
    from jinja2 import Environment
except ImportError:
    pass

from .markdown import Markdown
from ..utils import LicenseHandler
from ...main import Book

try:
    import PIL
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

LicenseHandler.validate_license('reports')


def render_template(sheet, **data):
    """
    Replaces the Jinja2 placeholders in a given sheet
    """
    # On Windows, Excel will not move objects correctly with screen_updating = False during row insert/delete operations
    # So we'll need to set it to True before any such operations. Getting origin state here to revert to.
    book = sheet.book
    screen_updating_original_state = book.app.screen_updating

    env = Environment()
    locals().update(data)

    # used_range doesn't start automatically in A1
    last_cell = sheet.used_range.last_cell
    values_all = sheet.range((1, 1), (last_cell.row, last_cell.column)).options(
        ndim=2).value if sheet.used_range.value else []
    # Frame markers
    frame_markers = []
    if values_all and '<frame>' in values_all[0]:
        frame_markers = values_all[0]
        values = values_all[1:]
        if sys.platform.startswith('win'):
            book.app.screen_updating = True
        sheet['1:1'].delete('up')
        book.app.screen_updating = screen_updating_original_state
        frame_indices = [i for i, val in enumerate(frame_markers) if val == '<frame>']
        frame_indices += [0, last_cell.column]
        frame_indices = list(sorted(set(frame_indices)))
    else:
        values = values_all
        frame_indices = [0, last_cell.column]
    values_per_frame = []
    for ix in range(len(frame_indices) - 1):
        values_per_frame.append([i[frame_indices[ix]:frame_indices[ix + 1]] for i in values])
    # Loop through every cell for each frame
    for ix, values in enumerate(values_per_frame):
        row_shift = 0
        for i, row in enumerate(values):
            for j, value in enumerate(row):
                if isinstance(value, str):
                    tokens = list(env.lex(value))
                    if value.count('{{') == 1 and tokens[0][1] == 'variable_begin' and tokens[-1][1] == 'variable_end':
                        # Cell contains single Jinja variable
                        var = None
                        for _, token_type, value in tokens:
                            if token_type == 'variable_begin':
                                var = ''
                            elif token_type == 'variable_end':
                                result = eval(var)
                                if PIL and isinstance(result, PIL.Image.Image):
                                    # TODO: properly support Image objects in xlwings
                                    sheet.pictures.add(result.filename,
                                                       top=sheet[i + row_shift, j + frame_indices[ix]].top,
                                                       left=sheet[i + row_shift, j + frame_indices[ix]].left,
                                                       width=result.width, height=result.height)
                                    sheet[i + row_shift, j + frame_indices[ix]].value = None
                                elif Figure and isinstance(result, Figure):
                                    # Matplotlib figures
                                    sheet.pictures.add(result,
                                                       top=sheet[i + row_shift, j + frame_indices[ix]].top,
                                                       left=sheet[i + row_shift, j + frame_indices[ix]].left)
                                    sheet[i + row_shift, j + frame_indices[ix]].value = None
                                elif isinstance(result, Markdown):
                                    sheet[i + row_shift,
                                          j + frame_indices[ix]].value = result
                                else:
                                    # Simple Jinja variables
                                    # Check for height of 2d array
                                    if isinstance(result, (list, tuple)) and isinstance(result[0], (list, tuple)):
                                        result_len = len(result)
                                    elif np and isinstance(result, np.ndarray):
                                        result_len = len(result)
                                    elif pd and isinstance(result, pd.DataFrame):
                                        # TODO: handle MultiIndex headers
                                        result_len = len(result) + 1
                                    else:
                                        result_len = 1
                                    # Insert rows if within <frame> and 'result' is multiple rows high
                                    rows_to_be_inserted = 0
                                    if frame_markers and result_len > 1:
                                        # Deduct header and first data row that are part of template
                                        rows_to_be_inserted = result_len - 2
                                        if rows_to_be_inserted > 0:
                                            if sys.platform.startswith('win'):
                                                book.app.screen_updating = True
                                            # Since CopyOrigin is not supported on Mac, we start copying two rows
                                            # below the header so the data row formatting gets carried over
                                            end_column = frame_indices[ix] + len(values[0])
                                            sheet.range((i + row_shift + 3, j + frame_indices[ix] + 1),
                                                        (i + row_shift + rows_to_be_inserted + 2, end_column)).insert(
                                                'down')
                                            # Inserting does not take over borders
                                            sheet.range((i + row_shift + 2, j + frame_indices[ix] + 1),
                                                        (i + row_shift + 2, end_column)).copy()
                                            sheet.range((i + row_shift + 2, j + frame_indices[ix] + 1),
                                                        (i + row_shift + rows_to_be_inserted + 2, end_column)).paste(
                                                paste='formats')
                                            book.app.screen_updating = screen_updating_original_state
                                    if sheet[i + row_shift, j + frame_indices[ix]].table:
                                        sheet[i + row_shift, j + frame_indices[ix]].table.update(result)
                                    else:
                                        sheet[i + row_shift, j + frame_indices[ix]].value = result
                                    row_shift += rows_to_be_inserted
                            elif var is not None and token_type not in ('whitespace',):
                                var += value
                    elif '{{' in value:
                        # These are strings with (multiple) Jinja variables so apply standard text rendering here
                        template = env.from_string(value)
                        sheet[i + row_shift, j + frame_indices[ix]].value = template.render(data)
                    else:
                        # Don't do anything with cells that don't contain any templating so we don't lose the formatting
                        pass

    # Loop through all shapes with a template text
    for shape in sheet.shapes:
        shapetext = shape.text
        if shapetext and '{{' in shapetext:
            tokens = list(env.lex(shapetext))
            # Single Jinja variable case, the only case we support with Markdown
            if shapetext.count('{{') == 1 and tokens[0][1] == 'variable_begin' and tokens[-1][1] == 'variable_end':
                for _, token_type, token_value in tokens:
                    if token_type == 'name':
                        if isinstance(data[token_value], Markdown):
                            shape.text = data[token_value]
                        else:
                            # Single Jinja var but no Markdown
                            template = env.from_string(shapetext)
                            shape.text = template.render(data)
            else:
                # Multiple Jinja vars and no Markdown
                template = env.from_string(shapetext)
                shape.text = template.render(data)


def create_report(template, output, book_settings=None, app=None, **data):
    """
    This function requires xlwings :guilabel:`PRO`.

    This is a convenience wrapper around :meth:`mysheet.render_template <xlwings.Sheet.render_template>`

    Writes the values of all key word arguments to the ``output`` file according to the ``template`` and the variables
    contained in there (Jinja variable syntax).
    Following variable types are supported:

    strings, numbers, lists, simple dicts, NumPy arrays, Pandas DataFrames, PIL Image objects that have a filename and
    Matplotlib figures.

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


