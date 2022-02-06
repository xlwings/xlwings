import os
import datetime as dt
import requests
import xlwings as xw
from xlwings.rest.api import api

# NOTE: the API server must run before running this script
BASE_URL = "http://localhost:5000"

# Setup workbooks for generating the sample responses
app1 = xw.App()
app2 = xw.App()

wb1 = app1.books.add()
wb1b = app1.books.add()
wb2 = app2.books.add()
for wb in [wb1, wb2]:
    if len(wb.sheets) == 1:
        wb.sheets.add(after=1)

sheet1 = wb1.sheets[0]

sheet1["A1"].value = [[1.1, "a string"], [dt.datetime.now(), None]]
sheet1["A1"].formula = "=1+1.1"
chart = sheet1.charts.add()
chart.set_source_data(sheet1["A1"])
chart.chart_type = "line"


pic = os.path.abspath(os.path.join("..", "xlwings", "tests", "sample_picture.png"))
sheet1.pictures.add(pic)
wb1.sheets[0].range("B2:C3").name = "Sheet1!myname1"
wb1.sheets[0].range("A1").name = "myname2"
wb1.save("Book1.xlsx")
wb1 = xw.Book("Book1.xlsx")  # hack as save doesn't return the wb properly
app1.activate()


def generate_get_endpoint(endpoint):
    docs = []
    docs.append(".. http:get:: " + endpoint.replace("path:", ""))
    docs.append("")
    url = BASE_URL + endpoint.replace("<pid>", str(wb1.app.pid)).replace(
        "<book_name_or_ix>", wb1.name
    ).replace("<chart_name_or_ix>", "0").replace("<name>", "myname2").replace(
        "<sheet_scope_name>", "myname1"
    ).replace(
        "<sheet_name_or_ix>", "sheet1"
    ).replace(
        "<shape_name_or_ix>", "0"
    ).replace(
        "<path:fullname_or_name>", wb1.name
    ).replace(
        "<picture_name_or_ix>", "0"
    ).replace(
        "<address>", "A1:B2"
    )
    rv = requests.get(url)
    docs.append("**Example response**:")
    docs.append("")
    docs.append(".. sourcecode:: json")
    docs.append("")
    for i in rv.text.splitlines():
        docs.append("    " + i)
    docs.append("")
    return docs


# Get all routes
get_apps_urls = []
get_books_urls = []
get_book_urls = []
for rule in api.url_map.iter_rules():
    if "GET" in rule.methods:
        if rule.rule.startswith("/book/"):
            get_book_urls.append(rule.rule)
        elif rule.rule.startswith("/books"):
            get_books_urls.append(rule.rule)
        elif rule.rule.startswith("/apps"):
            get_apps_urls.append(rule.rule)

get_apps_urls = sorted(get_apps_urls)
get_books_urls = sorted(get_books_urls)
get_book_urls = sorted(get_book_urls)

text = []
intro = """.. _rest_api:

REST API
========

.. versionadded:: 0.13.0

Quickstart
----------

xlwings offers an easy way to expose an Excel workbook via REST API both on Windows and 
macOS. This can be useful when you have a workbook running on a single computer and want
to access it from another computer. Or you can build a Linux based web app that can 
interact with a legacy Excel application while you are in the progress of migrating the
Excel functionality into your web app (if you need help with that, `give us a shout 
<https://www.zoomeranalytics.com/contact>`_).

You can run the REST API server from a command prompt or terminal as follows (this 
requires Flask>=1.0, so make sure to ``pip install Flask``)::

    xlwings restapi run

Then perform a GET request e.g. via PowerShell on Windows or Terminal on Mac (while 
having an unsaved "Book1" open). Note that you need to run the server and the GET
request from two separate terminals (or you can use something more convenient like 
`Postman <https://www.getpostman.com/>`_ or `Insomnia <https://insomnia.rest/>`_ for 
testing the API)::

    $ curl "http://127.0.0.1:5000/book/book1/sheets/0/range/A1:B2"
    {
      "address": "$A$1:$B$2",
      "color": null,
      "column": 1,
      "column_width": 10.0,
      "count": 4,
      "current_region": "$A$1:$B$2",
      "formula": [
        [
          "1",
          "2"
        ],
        [
          "3",
          "4"
        ]
      ],
      "formula_array": null,
      "height": 32.0,
      "last_cell": "$B$2",
      "left": 0.0,
      "name": null,
      "number_format": "General",
      "row": 1,
      "row_height": 16.0,
      "shape": [
        2,
        2
      ],
      "size": 4,
      "top": 0.0,
      "value": [
        [
          1.0,
          2.0
        ],
        [
          3.0,
          4.0
        ]
      ],
      "width": 130.0
    }

In the command prompt where your server is running, press ``Ctrl-C`` to shut it down 
again.

The xlwings REST API is a thin wrapper around the :ref:`Python API <api>` which makes it
very easy if you have worked previously with xlwings. It also means that the REST API
does require the Excel application to be up and running which makes it a great choice if
the data in your Excel workbook is constantly changing as the REST API will always
deliver the current state of the workbook without the need of saving it first.

.. note::
    Currently, we only provide the GET methods to read the workbook. If you are also
    interested in the POST methods to edit the workbook, let us know via GitHub issues.
    Some other things will also need improvement, most notably exception handling.

Run the server
--------------

``xlwings restapi run`` will run a Flask development server on http://127.0.0.1:5000.
You can provide ``--host`` and ``--port`` as command line args and it also respects the
Flask environment variables like ``FLASK_ENV=development``.

If you want to have more control, you can run the server directly with Flask, see the
`Flask docs <http://flask.pocoo.org/docs/1.0/quickstart/>`_ for more details::

    set FLASK_APP=xlwings.rest.api
    flask run

If you are on Mac, use ``export FLASK_APP=xlwings.rest.api`` instead of ``set
FLASK_APP=xlwings.rest.api``.

For production, you can use any WSGI HTTP Server like 
`gunicorn <https://gunicorn.org/>`_ (on Mac) or 
`waitress <https://docs.pylonsproject.org/projects/waitress/en/latest/>`_ 
(on Mac/Windows) to serve the API. For example, with gunicorn you would do: 
``gunicorn xlwings.rest.api:api``. Or with waitress (adjust the host accordingly if you
want to make the api accessible from outside of localhost)::

    from xlwings.rest.api import api
    from waitress import serve
    serve(wsgiapp, host='127.0.0.1', port=5000)

Indexing
--------

While the Python API offers Python's 0-based indexing (e.g. ``xw.books[0]``) as well as
Excel's 1-based indexing (e.g. ``xw.books(1)``), the REST API only offers 0-based
indexing, e.g. ``/books/0``.

Range Options
-------------

The REST API accepts Range options as query parameters, see 
:meth:`xlwings.Range.options` e.g.,

``/book/book1/sheets/0/range/A1?expand=table&transpose=true``

Remember that ``options`` only affect the ``value`` property.

Endpoint overview
-----------------

+----------------+---------------------+----------------------------------------------------------------------------------------------+
| Endpoint       | Corresponds to      | Short Description                                                                            |
+================+=====================+==============================================================================================+
| :ref:`book`    | :ref:`python_book`  | Finds your workbook across all open instances of Excel and will open it if it can't find it  |
+----------------+---------------------+----------------------------------------------------------------------------------------------+
| :ref:`books`   | :ref:`python_books` | Books collection of the active Excel instance                                                |
+----------------+---------------------+----------------------------------------------------------------------------------------------+
| :ref:`apps`    | :ref:`python_apps`  | This allows you to specify the Excel instance you want to work with                          |
+----------------+---------------------+----------------------------------------------------------------------------------------------+

Endpoint details
----------------

"""
text.append(intro)
text.append("")

text.append(".. _book:")
text.append("")
text.append("/book")
text.append("*****")
text.append("")
for url in get_book_urls:
    text.extend(generate_get_endpoint(url))

text.append(".. _books:")
text.append("")
text.append("/books")
text.append("******")
text.append("")
for url in get_books_urls:
    text.extend(generate_get_endpoint(url))

text.append(".. _apps:")
text.append("")
text.append("/apps")
text.append("*****")
text.append("")
for url in get_apps_urls:
    text.extend(generate_get_endpoint(url))

with open("../docs/rest_api.rst", "w") as f:
    for line in text:
        f.write(f"{line}\n")

wb_path = wb1.fullname
app1.kill()
app2.kill()
os.remove(wb_path)
