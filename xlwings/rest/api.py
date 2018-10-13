# See also: https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/excel
import sys
import logging
from werkzeug.routing import PathConverter
import xlwings as xw

try:
    import flask
    from flask import Flask, jsonify, request
except ImportError:
    raise Exception("To use the xlwings REST API server, you need Flask>=1.0.0 installed.")


api = Flask(__name__)
logger = logging.getLogger(__name__)


class EverythingConverter(PathConverter):
    regex = '.*?'


if sys.platform.startswith('darwin'):
    # Hack to allow leading slashes on Mac
    api.url_map.converters['path'] = EverythingConverter


def serialize_app(app):
    return {
        'version': str(app.version),
        'visible': app.visible,
        'screen_updating': app.screen_updating,
        'display_alerts': app.display_alerts,
        'calculation': app.calculation,
        'selection': app.selection.address,
        'books': [book.name for book in app.books],
        'hwnd': app.hwnd,
        'pid': app.pid
    }


def serialize_book(book):
    return {
        'name': book.name,
        'sheets': [sheet.name for sheet in book.sheets],
        'fullname': book.fullname,
        'names': [name for name in book.names],
        'selection': book.selection.address
    }


def serialize_sheet(sheet):
    return {
        'name': sheet.name,
        'names': [name.name for name in sheet.names],
        'index': sheet.index,
        'charts': [chart.name for chart in sheet.charts],
        'shapes': [shape.name for shape in sheet.shapes],
        'pictures': [picture.name for picture in sheet.pictures]
    }


def serialize_chart(chart):
    return {
        'name': chart.name
    }


def serialize_picture(picture):
    return {
        'name': picture.name
    }


def serialize_shapes(shape):
    return {
        'name': shape.name
    }


def serialize_names(name):
    return {
        'name': name.name
    }


def serialize_range(rng):
    return {
        'value': rng.value,
        'count': rng.count,
        'row': rng.row,
        'column': rng.column,
        'formula': rng.formula,
        'formula_array': rng.formula_array,
        'column_width': rng.column_width,
        'row_height': rng.row_height,
        'address': rng.address,
        'color': rng.color,
        'current_region': rng.current_region.address,
        'height': rng.height,
        'last_cell': rng.last_cell.address,
        'left': rng.left,
        'name': rng.name,
        'number_format': rng.number_format,
        'shape': rng.shape,
        'size': rng.size,
        'top': rng.top,
        'width': rng.width
    }


def get_book(fullname_or_ix, app_ix=None):
    if fullname_or_ix.isdigit():
        fullname_or_ix = int(fullname_or_ix)
    app = xw.apps[int(app_ix)] if app_ix else xw.apps.active
    return app.books[fullname_or_ix]


def get_sheet(book, name_or_id):
    if name_or_id.isdigit():
        name_or_id = int(name_or_id)
    return book.sheets[name_or_id]


@api.route('/apps', methods=['GET'])
def apps():
    return jsonify(apps=[serialize_app(app)
                         for app in xw.apps])


@api.route('/apps/<pid>/', methods=['GET'])
def app_(pid):
    return jsonify(serialize_app(xw.apps[int(pid)]))


@api.route('/apps/<pid>/books', methods=['GET'])
@api.route('/books', methods=['GET'])
def books_(pid=None):
    books = xw.apps[int(pid)].books if pid else xw.books
    return jsonify(books=[serialize_book(book)
                          for book in books])


@api.route('/apps/<pid>/books/<book_name_or_ix>', methods=['GET'])
@api.route('/books/<book_name_or_ix>', methods=['GET'])
def book_(book_name_or_ix, pid=None):
    book = get_book(book_name_or_ix, pid)
    return jsonify(serialize_book(book))


@api.route('/apps/<pid>/books/<book_name_or_ix>/sheets', methods=['GET'])
@api.route('/books/<book_name_or_ix>/sheets', methods=['GET'])
def sheets(book_name_or_ix, pid=None):
    book = get_book(book_name_or_ix, pid)
    return jsonify(sheets=[serialize_sheet(sheet)
                           for sheet in book.sheets])


@api.route('/apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_index>', methods=['GET'])
@api.route('/books/<book_name_or_ix>/sheets/<sheet_name_or_index>', methods=['GET'])
def sheet_(sheet_name_or_index, book_name_or_ix, pid=None):
    book = get_book(book_name_or_ix, pid)
    sheet = get_sheet(book, sheet_name_or_index)
    return jsonify(value=serialize_sheet(sheet))


@api.route('/apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_index>/range', methods=['GET'])
@api.route('/books/<book_name_or_ix>/sheets/<sheet_name_or_index>/range', methods=['GET'])
def range_(sheet_name_or_index, book_name_or_ix, pid=None):
    book = get_book(book_name_or_ix, pid)
    sheet = get_sheet(book, sheet_name_or_index)
    return jsonify(serialize_range(sheet.used_range))


@api.route('/apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_index>/range/<rng>', methods=['GET'])
@api.route('/books/<book_name_or_ix>/sheets/<sheet_name_or_index>/range/<rng>', methods=['GET'])
def range_(rng, sheet_name_or_index, book_name_or_ix, pid=None):
    book = get_book(book_name_or_ix, pid)
    sheet = get_sheet(book, sheet_name_or_index)
    return jsonify(serialize_range(sheet.range(rng)))





@api.route('/book/<path:path>/names', methods=['GET'])
def names(path):
    wb = xw.Book(path)
    return jsonify(value=[
        {
            'name': name.name,
            'type': "Range",
            'value': name.refers_to[1:],
            'visible': True

        }
        for name in wb.names
    ])

# @api.route('/book/<path:fullname>/sheets', methods=['GET'])
# def sheets(fullname):
#     wb = xw.Book(fullname)
#     return jsonify(value=[serialize_sheet(sheet)
#                           for sheet in wb.sheets])


# @api.route('/book/<path:path>/sheets/<string:name_or_id>', methods=['GET'])
# def sheet_(path, name_or_id):
#     wb = xw.Book(path)
#     if name_or_id.isdigit():
#         name_or_id = int(name_or_id) - 1
#     sheet = wb.sheets[name_or_id]
#     return jsonify(serialize_sheet(sheet))
# 
# @api.route('/<path:path>/sheets/<string:name_or_id>/range(address=<string:address>)', methods=['GET'])
# @api.route('/<path:path>/sheets/<string:name_or_id>/range', defaults={'address': None}, methods=['GET'])
# def rng(path, name_or_id, address=None):
#     wb = xw.Book(path)
#     sheet = get_sheet(wb, name_or_id)
#     if not address:
#         address = sheet.api.UsedRange.Address
#     rng = sheet[address]
#     rows_count = rng.rows.count
#     columns_count = rng.columns.count
#     return jsonify({
#         "address": sheet.name + '!' + rng.address,
#         "addressLocal": sheet.name + '!' + rng.address,
#         "cellCount": rng.count,
#         "columnCount": columns_count,
#         "columnHidden": False,
#         "columnIndex": rng.columns[0].column,
#         "formulas": rng.formula,
#         "formulasLocal": rng.formula,  # TODO!
#         "formulasR1C1": rng.formula,  # TODO!
#         "hidden": False,
#         "numberFormat": [[None for i in range(columns_count)] for i in range(rows_count)], #[[cell.number_format for cell in row] for row in rng.rows],  # TODO: is this fast enough for big ranges?
#         "rowCount": rows_count,
#         "rowHidden": False,
#         "rowIndex": rng.rows[0].row,
#         "text": rng.value,  # TODO!
#         "values": rng.value,
#         "valueTypes": [[None for i in range(columns_count)] for i in range(rows_count)] # TODO: just prints None at the moment + is this fast enough for big ranges?
#     })


# @api.route('/<path:path>/sheets/<string:name_or_id>/cell(row=<int:row>,column=<int:column>)', methods=['GET'])
# def cell(path, name_or_id, row, column):
#     wb = xw.Book(path)
#     sheet = get_sheet(wb, name_or_id)
#     rng = sheet.cells[row, column]
#     return jsonify({
#         "address": rng.address,
#         "addressLocal": rng.address,
#         "cellCount": rng.count,
#         "columnCount": rng.columns.count,
#         "columnIndex": rng.columns[0].column,
#         "valueTypes": None
#     })


def run_server(port=5000,
               debug=False,
               **flask_run_options):
    """
    Run Flask development server
    """
    api.run(port=port, debug=debug,
            **flask_run_options)


if __name__ == '__main__':
    run_server(debug=True)
