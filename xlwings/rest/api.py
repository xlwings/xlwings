# See also:
# https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/excel
# TODO: proper exception handling in xlwings base package so it can be used here
import logging
import sys

from werkzeug.routing import PathConverter

import xlwings as xw
from xlwings.rest.serializers import (
    serialize_app,
    serialize_book,
    serialize_chart,
    serialize_name,
    serialize_picture,
    serialize_range,
    serialize_shape,
    serialize_sheet,
)

try:
    from flask import Flask, abort, jsonify, request
except ImportError:
    raise Exception(
        "To use the xlwings REST API server, you need Flask>=1.0.0 installed."
    )


api = Flask(__name__)
logger = logging.getLogger(__name__)


class EverythingConverter(PathConverter):
    regex = ".*?"


if sys.platform.startswith("darwin"):
    # Hack to allow leading slashes on Mac
    api.url_map.converters["path"] = EverythingConverter


def get_book_object(fullname=None, name_or_ix=None, app_ix=None):
    assert fullname is None or name_or_ix is None
    if fullname:
        try:
            return xw.Book(fullname)
        except Exception as e:
            logger.exception(str(e))
            abort(500, str(e))
    elif name_or_ix:
        if name_or_ix.isdigit():
            name_or_ix = int(name_or_ix)
        app = xw.apps[int(app_ix)] if app_ix else xw.apps.active
        try:
            return app.books[name_or_ix]
        except KeyError as e:
            logger.exception(str(e))
            abort(500, "Couldn't find Book: " + str(e))
        except Exception as e:
            logger.exception(str(e))
            abort(500, str(e))


def get_sheet_object(book, name_or_id):
    if name_or_id.isdigit():
        name_or_id = int(name_or_id)
    return book.sheets[name_or_id]


@api.route("/apps", methods=["GET"])
def get_apps():
    return jsonify(apps=[serialize_app(app) for app in xw.apps])


@api.route("/apps/<pid>", methods=["GET"])
def get_app(pid):
    return jsonify(serialize_app(xw.apps[int(pid)]))


@api.route("/apps/<pid>/books", methods=["GET"])
@api.route("/books", methods=["GET"])
def get_books(pid=None):
    books = xw.apps[int(pid)].books if pid else xw.books
    return jsonify(books=[serialize_book(book) for book in books])


@api.route("/apps/<pid>/books/<book_name_or_ix>", methods=["GET"])
@api.route("/books/<book_name_or_ix>", methods=["GET"])
@api.route("/book/<path:fullname_or_name>", methods=["GET"])
def get_book(book_name_or_ix=None, fullname_or_name=None, pid=None):
    book = get_book_object(fullname_or_name, book_name_or_ix, pid)
    return jsonify(serialize_book(book))


@api.route("/apps/<pid>/books/<book_name_or_ix>/sheets", methods=["GET"])
@api.route("/books/<book_name_or_ix>/sheets", methods=["GET"])
@api.route("/book/<path:fullname_or_name>/sheets", methods=["GET"])
def get_sheets(book_name_or_ix=None, fullname_or_name=None, pid=None):
    book = get_book_object(fullname_or_name, book_name_or_ix, pid)
    return jsonify(sheets=[serialize_sheet(sheet) for sheet in book.sheets])


@api.route(
    "/apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>", methods=["GET"]
)
@api.route("/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>", methods=["GET"])
@api.route("/book/<path:fullname_or_name>/sheets/<sheet_name_or_ix>", methods=["GET"])
def get_sheet(sheet_name_or_ix, book_name_or_ix=None, fullname_or_name=None, pid=None):
    book = get_book_object(fullname_or_name, book_name_or_ix, pid)
    sheet = get_sheet_object(book, sheet_name_or_ix)
    return jsonify(serialize_sheet(sheet))


@api.route(
    "/apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/range",
    methods=["GET"],
)
@api.route("/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/range", methods=["GET"])
@api.route(
    "/book/<path:fullname_or_name>/sheets/<sheet_name_or_ix>/range", methods=["GET"]
)
def get_range(sheet_name_or_ix, book_name_or_ix=None, fullname_or_name=None, pid=None):
    book = get_book_object(fullname_or_name, book_name_or_ix, pid)
    sheet = get_sheet_object(book, sheet_name_or_ix)
    return jsonify(serialize_range(sheet.used_range))


@api.route(
    "/apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/range/<address>",
    methods=["GET"],
)
@api.route(
    "/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/range/<address>",
    methods=["GET"],
)
@api.route(
    "/book/<path:fullname_or_name>/sheets/<sheet_name_or_ix>/range/<address>",
    methods=["GET"],
)
def get_range_address(
    address, sheet_name_or_ix, book_name_or_ix=None, fullname_or_name=None, pid=None
):
    book = get_book_object(fullname_or_name, book_name_or_ix, pid)
    sheet = get_sheet_object(book, sheet_name_or_ix)
    options = {k: v for k, v in request.args.items()}
    return jsonify(serialize_range(sheet.range(address).options(**options)))


@api.route("/apps/<pid>/books/<book_name_or_ix>/names", methods=["GET"])
@api.route("/books/<book_name_or_ix>/names", methods=["GET"])
@api.route("/book/<path:fullname_or_name>/names", methods=["GET"])
def get_book_names(book_name_or_ix=None, fullname_or_name=None, pid=None):
    book = get_book_object(fullname_or_name, book_name_or_ix, pid)
    return jsonify(names=[serialize_name(name) for name in book.names])


@api.route("/apps/<pid>/books/<book_name_or_ix>/names/<name>", methods=["GET"])
@api.route("/books/<book_name_or_ix>/names/<name>", methods=["GET"])
@api.route("/book/<path:fullname_or_name>/names/<name>", methods=["GET"])
def get_book_name(name, book_name_or_ix=None, fullname_or_name=None, pid=None):
    book = get_book_object(fullname_or_name, book_name_or_ix, pid)
    return jsonify(serialize_name(book.names[name]))


@api.route("/apps/<pid>/books/<book_name_or_ix>/names/<name>/range", methods=["GET"])
@api.route("/books/<book_name_or_ix>/names/<name>/range", methods=["GET"])
@api.route("/book/<path:fullname_or_name>/names/<name>/range", methods=["GET"])
def get_book_name_range(name, book_name_or_ix=None, fullname_or_name=None, pid=None):
    book = get_book_object(fullname_or_name, book_name_or_ix, pid)
    return jsonify(serialize_range(book.names[name].refers_to_range))


@api.route(
    "/apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/names",
    methods=["GET"],
)
@api.route("/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/names", methods=["GET"])
@api.route(
    "/book/<path:fullname_or_name>/sheets/<sheet_name_or_ix>/names", methods=["GET"]
)
def get_sheet_names(
    sheet_name_or_ix, book_name_or_ix=None, fullname_or_name=None, pid=None
):
    book = get_book_object(fullname_or_name, book_name_or_ix, pid)
    sheet = get_sheet_object(book, sheet_name_or_ix)
    return jsonify(names=[serialize_name(name) for name in sheet.names])


@api.route(
    "/apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/names/<sheet_scope_name>",
    methods=["GET"],
)
@api.route(
    "/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/names/<sheet_scope_name>",
    methods=["GET"],
)
@api.route(
    "/book/<path:fullname_or_name>/sheets/<sheet_name_or_ix>/names/<sheet_scope_name>",
    methods=["GET"],
)
def get_sheet_name(
    sheet_name_or_ix,
    sheet_scope_name,
    book_name_or_ix=None,
    fullname_or_name=None,
    pid=None,
):
    book = get_book_object(fullname_or_name, book_name_or_ix, pid)
    sheet = get_sheet_object(book, sheet_name_or_ix)
    return jsonify(serialize_name(sheet.names[sheet_scope_name]))


@api.route(
    "/apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/names/<sheet_scope_name>/range",
    methods=["GET"],
)
@api.route(
    "/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/names/<sheet_scope_name>/range",
    methods=["GET"],
)
@api.route(
    "/book/<path:fullname_or_name>/sheets/<sheet_name_or_ix>/names/<sheet_scope_name>/range",
    methods=["GET"],
)
def get_sheet_name_range(
    sheet_name_or_ix,
    sheet_scope_name,
    book_name_or_ix=None,
    fullname_or_name=None,
    pid=None,
):
    book = get_book_object(fullname_or_name, book_name_or_ix, pid)
    sheet = get_sheet_object(book, sheet_name_or_ix)
    return jsonify(serialize_range(sheet.names[sheet_scope_name].refers_to_range))


@api.route(
    "/apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/charts",
    methods=["GET"],
)
@api.route("/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/charts", methods=["GET"])
@api.route(
    "/book/<path:fullname_or_name>/sheets/<sheet_name_or_ix>/charts", methods=["GET"]
)
def get_charts(sheet_name_or_ix, book_name_or_ix=None, fullname_or_name=None, pid=None):
    book = get_book_object(fullname_or_name, book_name_or_ix, pid)
    sheet = get_sheet_object(book, sheet_name_or_ix)
    return jsonify(charts=[serialize_chart(chart) for chart in sheet.charts])


@api.route(
    "/apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/charts/<chart_name_or_ix>",
    methods=["GET"],
)
@api.route(
    "/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/charts/<chart_name_or_ix>",
    methods=["GET"],
)
@api.route(
    "/book/<path:fullname_or_name>/sheets/<sheet_name_or_ix>/charts/<chart_name_or_ix>",
    methods=["GET"],
)
def get_chart(
    sheet_name_or_ix,
    chart_name_or_ix,
    book_name_or_ix=None,
    fullname_or_name=None,
    pid=None,
):
    book = get_book_object(fullname_or_name, book_name_or_ix, pid)
    sheet = get_sheet_object(book, sheet_name_or_ix)
    chart = int(chart_name_or_ix) if chart_name_or_ix.isdigit() else chart_name_or_ix
    return jsonify(serialize_chart(sheet.charts[chart]))


@api.route(
    "/apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/shapes",
    methods=["GET"],
)
@api.route("/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/shapes", methods=["GET"])
@api.route(
    "/book/<path:fullname_or_name>/sheets/<sheet_name_or_ix>/shapes", methods=["GET"]
)
def get_shapes(sheet_name_or_ix, book_name_or_ix=None, fullname_or_name=None, pid=None):
    book = get_book_object(fullname_or_name, book_name_or_ix, pid)
    sheet = get_sheet_object(book, sheet_name_or_ix)
    return jsonify(shapes=[serialize_shape(shp) for shp in sheet.shapes])


@api.route(
    "/apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/shapes/<shape_name_or_ix>",
    methods=["GET"],
)
@api.route(
    "/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/shapes/<shape_name_or_ix>",
    methods=["GET"],
)
@api.route(
    "/book/<path:fullname_or_name>/sheets/<sheet_name_or_ix>/shapes/<shape_name_or_ix>",
    methods=["GET"],
)
def get_shape(
    sheet_name_or_ix,
    shape_name_or_ix,
    book_name_or_ix=None,
    fullname_or_name=None,
    pid=None,
):
    book = get_book_object(fullname_or_name, book_name_or_ix, pid)
    sheet = get_sheet_object(book, sheet_name_or_ix)
    shape = int(shape_name_or_ix) if shape_name_or_ix.isdigit() else shape_name_or_ix
    return jsonify(serialize_shape(sheet.shapes[shape]))


@api.route(
    "/apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/pictures",
    methods=["GET"],
)
@api.route(
    "/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/pictures", methods=["GET"]
)
@api.route(
    "/book/<path:fullname_or_name>/sheets/<sheet_name_or_ix>/pictures", methods=["GET"]
)
def get_pictures(
    sheet_name_or_ix, book_name_or_ix=None, fullname_or_name=None, pid=None
):
    book = get_book_object(fullname_or_name, book_name_or_ix, pid)
    sheet = get_sheet_object(book, sheet_name_or_ix)
    return jsonify(pictures=[serialize_picture(pic) for pic in sheet.pictures])


@api.route(
    "/apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/pictures/<picture_name_or_ix>",
    methods=["GET"],
)
@api.route(
    "/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/pictures/<picture_name_or_ix>",
    methods=["GET"],
)
@api.route(
    "/book/<path:fullname_or_name>/sheets/<sheet_name_or_ix>/pictures/<picture_name_or_ix>",
    methods=["GET"],
)
def get_picture(
    sheet_name_or_ix,
    picture_name_or_ix,
    book_name_or_ix=None,
    fullname_or_name=None,
    pid=None,
):
    book = get_book_object(fullname_or_name, book_name_or_ix, pid)
    sheet = get_sheet_object(book, sheet_name_or_ix)
    pic = (
        int(picture_name_or_ix) if picture_name_or_ix.isdigit() else picture_name_or_ix
    )
    return jsonify(serialize_picture(sheet.pictures[pic]))


def run(host=None, port=None, debug=None, **options):
    """
    Run Flask development server
    """
    api.run(host=host, port=port, debug=debug, **options)


if __name__ == "__main__":
    run(debug=True)
