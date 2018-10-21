.. _rest_api:

REST API
========

.. versionadded:: 0.13.0

Quickstart
----------

xlwings offers an easy way to expose an Excel workbook via REST API both on Windows and macOS. This can be useful
when you have a workbook running on a single computer and want to access it from another computer. Or you can
build a Linux based web app that can interact with a legacy Excel application while you are in the progress
of migrating the Excel functionality into your web app (if you need help with that, `give us a shout <https://www.zoomeranalytics.com/contact>`_).

You can run the REST API server from a command prompt or terminal as follows (this requires Flask>=1.0, so make sure to ``pip install Flask``)::

    xlwings restapi run

Then perform a GET request e.g. via PowerShell on Windows or Terminal on Mac (while having an unsaved "Book1" open). Note
that you need to run the server and the GET request from two separate terminals (or you can use something
more convenient like `Postman <https://www.getpostman.com/>`_ or `Insomnia <https://insomnia.rest/>`_ for testing the API)::

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

In the command prompt where your server is running, press ``Ctrl-C`` to shut it down again.

The xlwings REST API is a thin wrapper around the :ref:`Python API <api>` which makes it very easy if
you have worked previously with xlwings. It also means that the REST API does require the Excel application to be up and
running which makes it a great choice if the data in your Excel workbook is constantly changing as the REST API will
always deliver the current state of the workbook without the need of saving it first.

.. note::
    Currently, we only provide the GET methods to read the workbook. If you are also interested in the POST methods
    to edit the workbook, let us know via GitHub issues. Some other things will also need improvement, most notably
    exception handling.

Run the server
--------------

``xlwings restapi run`` will run a Flask development server on http://127.0.0.1:5000. You can provide ``--host`` and ``--port`` as
command line args and it also respects the Flask environment variables like ``FLASK_ENV=development``.

If you want to have more control, you can run the server directly with Flask, see the
`Flask docs <http://flask.pocoo.org/docs/1.0/quickstart/>`_ for more details::

    set FLASK_APP=xlwings.rest.api
    flask run

If you are on Mac, use ``export FLASK_APP=xlwings.rest.api`` instead of ``set FLASK_APP=xlwings.rest.api``.

For production, you can use any WSGI HTTP Server like `gunicorn <https://gunicorn.org/>`_ (on Mac) or `waitress
<https://docs.pylonsproject.org/projects/waitress/en/latest/>`_ (on Mac/Windows) to serve the API. For example,
with gunicorn you would do: ``gunicorn xlwings.rest.api:api``. Or with waitress (adjust the host accordingly if
you want to make the api accessible from outside of localhost)::

    from xlwings.rest.api import api
    from waitress import serve
    serve(wsgiapp, host='127.0.0.1', port=5000)

Indexing
--------

While the Python API offers Python's 0-based indexing (e.g. ``xw.books[0]``) as well as Excel's 1-based indexing (e.g. ``xw.books(1)``),
the REST API only offers 0-based indexing, e.g. ``/books/0``.

Range Options
-------------

The REST API accepts Range options as query parameters, see :meth:`xlwings.Range.options` e.g.

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



.. _book:

/book
*****

.. http:get:: /book/<fullname_or_name>

**Example response**:

.. sourcecode:: json

    {
      "app": 1104, 
      "fullname": "C:\\Users\\felix\\DEV\\xlwings\\scripts\\Book1.xlsx", 
      "name": "Book1.xlsx", 
      "names": [
        "Sheet1!myname1", 
        "myname2"
      ], 
      "selection": "Sheet2!$A$1", 
      "sheets": [
        "Sheet1", 
        "Sheet2"
      ]
    }

.. http:get:: /book/<fullname_or_name>/names

**Example response**:

.. sourcecode:: json

    {
      "names": [
        {
          "name": "Sheet1!myname1", 
          "refers_to": "=Sheet1!$B$2:$C$3"
        }, 
        {
          "name": "myname2", 
          "refers_to": "=Sheet1!$A$1"
        }
      ]
    }

.. http:get:: /book/<fullname_or_name>/names/<name>

**Example response**:

.. sourcecode:: json

    {
      "name": "myname2", 
      "refers_to": "=Sheet1!$A$1"
    }

.. http:get:: /book/<fullname_or_name>/names/<name>/range

**Example response**:

.. sourcecode:: json

    {
      "address": "$A$1", 
      "color": null, 
      "column": 1, 
      "column_width": 8.47, 
      "count": 1, 
      "current_region": "$A$1:$B$2", 
      "formula": "=1+1.1", 
      "formula_array": "=1+1,1", 
      "height": 14.25, 
      "last_cell": "$A$1", 
      "left": 0.0, 
      "name": "myname2", 
      "number_format": "General", 
      "row": 1, 
      "row_height": 14.3, 
      "shape": [
        1, 
        1
      ], 
      "size": 1, 
      "top": 0.0, 
      "value": 2.1, 
      "width": 51.0
    }

.. http:get:: /book/<fullname_or_name>/sheets

**Example response**:

.. sourcecode:: json

    {
      "sheets": [
        {
          "charts": [
            "Chart 1"
          ], 
          "name": "Sheet1", 
          "names": [
            "Sheet1!myname1"
          ], 
          "pictures": [
            "Picture 3"
          ], 
          "shapes": [
            "Chart 1", 
            "Picture 3"
          ], 
          "used_range": "$A$1:$B$2"
        }, 
        {
          "charts": [], 
          "name": "Sheet2", 
          "names": [], 
          "pictures": [], 
          "shapes": [], 
          "used_range": "$A$1"
        }
      ]
    }

.. http:get:: /book/<fullname_or_name>/sheets/<sheet_name_or_ix>

**Example response**:

.. sourcecode:: json

    {
      "charts": [
        "Chart 1"
      ], 
      "name": "Sheet1", 
      "names": [
        "Sheet1!myname1"
      ], 
      "pictures": [
        "Picture 3"
      ], 
      "shapes": [
        "Chart 1", 
        "Picture 3"
      ], 
      "used_range": "$A$1:$B$2"
    }

.. http:get:: /book/<fullname_or_name>/sheets/<sheet_name_or_ix>/charts

**Example response**:

.. sourcecode:: json

    {
      "charts": [
        {
          "chart_type": "line", 
          "height": 211.0, 
          "left": 0.0, 
          "name": "Chart 1", 
          "top": 0.0, 
          "width": 355.0
        }
      ]
    }

.. http:get:: /book/<fullname_or_name>/sheets/<sheet_name_or_ix>/charts/<chart_name_or_ix>

**Example response**:

.. sourcecode:: json

    {
      "chart_type": "line", 
      "height": 211.0, 
      "left": 0.0, 
      "name": "Chart 1", 
      "top": 0.0, 
      "width": 355.0
    }

.. http:get:: /book/<fullname_or_name>/sheets/<sheet_name_or_ix>/names

**Example response**:

.. sourcecode:: json

    {
      "names": [
        {
          "name": "Sheet1!myname1", 
          "refers_to": "=Sheet1!$B$2:$C$3"
        }
      ]
    }

.. http:get:: /book/<fullname_or_name>/sheets/<sheet_name_or_ix>/names/<sheet_scope_name>

**Example response**:

.. sourcecode:: json

    {
      "name": "Sheet1!myname1", 
      "refers_to": "=Sheet1!$B$2:$C$3"
    }

.. http:get:: /book/<fullname_or_name>/sheets/<sheet_name_or_ix>/names/<sheet_scope_name>/range

**Example response**:

.. sourcecode:: json

    {
      "address": "$B$2:$C$3", 
      "color": null, 
      "column": 2, 
      "column_width": 8.47, 
      "count": 4, 
      "current_region": "$A$1:$B$2", 
      "formula": [
        [
          "", 
          ""
        ], 
        [
          "", 
          ""
        ]
      ], 
      "formula_array": "", 
      "height": 28.5, 
      "last_cell": "$C$3", 
      "left": 51.0, 
      "name": "Sheet1!myname1", 
      "number_format": "General", 
      "row": 2, 
      "row_height": 14.3, 
      "shape": [
        2, 
        2
      ], 
      "size": 4, 
      "top": 14.25, 
      "value": [
        [
          null, 
          null
        ], 
        [
          null, 
          null
        ]
      ], 
      "width": 102.0
    }

.. http:get:: /book/<fullname_or_name>/sheets/<sheet_name_or_ix>/pictures

**Example response**:

.. sourcecode:: json

    {
      "pictures": [
        {
          "height": 100.0, 
          "left": 0.0, 
          "name": "Picture 3", 
          "top": 0.0, 
          "width": 100.0
        }
      ]
    }

.. http:get:: /book/<fullname_or_name>/sheets/<sheet_name_or_ix>/pictures/<picture_name_or_ix>

**Example response**:

.. sourcecode:: json

    {
      "height": 100.0, 
      "left": 0.0, 
      "name": "Picture 3", 
      "top": 0.0, 
      "width": 100.0
    }

.. http:get:: /book/<fullname_or_name>/sheets/<sheet_name_or_ix>/range

**Example response**:

.. sourcecode:: json

    {
      "address": "$A$1:$B$2", 
      "color": null, 
      "column": 1, 
      "column_width": 8.47, 
      "count": 4, 
      "current_region": "$A$1:$B$2", 
      "formula": [
        [
          "=1+1.1", 
          "a string"
        ], 
        [
          "43395.0064583333", 
          ""
        ]
      ], 
      "formula_array": null, 
      "height": 28.5, 
      "last_cell": "$B$2", 
      "left": 0.0, 
      "name": null, 
      "number_format": null, 
      "row": 1, 
      "row_height": 14.3, 
      "shape": [
        2, 
        2
      ], 
      "size": 4, 
      "top": 0.0, 
      "value": [
        [
          2.1, 
          "a string"
        ], 
        [
          "Mon, 22 Oct 2018 00:09:18 GMT", 
          null
        ]
      ], 
      "width": 102.0
    }

.. http:get:: /book/<fullname_or_name>/sheets/<sheet_name_or_ix>/range/<address>

**Example response**:

.. sourcecode:: json

    {
      "address": "$A$1:$B$2", 
      "color": null, 
      "column": 1, 
      "column_width": 8.47, 
      "count": 4, 
      "current_region": "$A$1:$B$2", 
      "formula": [
        [
          "=1+1.1", 
          "a string"
        ], 
        [
          "43395.0064583333", 
          ""
        ]
      ], 
      "formula_array": null, 
      "height": 28.5, 
      "last_cell": "$B$2", 
      "left": 0.0, 
      "name": null, 
      "number_format": null, 
      "row": 1, 
      "row_height": 14.3, 
      "shape": [
        2, 
        2
      ], 
      "size": 4, 
      "top": 0.0, 
      "value": [
        [
          2.1, 
          "a string"
        ], 
        [
          "Mon, 22 Oct 2018 00:09:18 GMT", 
          null
        ]
      ], 
      "width": 102.0
    }

.. http:get:: /book/<fullname_or_name>/sheets/<sheet_name_or_ix>/shapes

**Example response**:

.. sourcecode:: json

    {
      "shapes": [
        {
          "height": 211.0, 
          "left": 0.0, 
          "name": "Chart 1", 
          "top": 0.0, 
          "type": "chart", 
          "width": 355.0
        }, 
        {
          "height": 100.0, 
          "left": 0.0, 
          "name": "Picture 3", 
          "top": 0.0, 
          "type": "picture", 
          "width": 100.0
        }
      ]
    }

.. http:get:: /book/<fullname_or_name>/sheets/<sheet_name_or_ix>/shapes/<shape_name_or_ix>

**Example response**:

.. sourcecode:: json

    {
      "height": 211.0, 
      "left": 0.0, 
      "name": "Chart 1", 
      "top": 0.0, 
      "type": "chart", 
      "width": 355.0
    }

.. _books:

/books
******

.. http:get:: /books

**Example response**:

.. sourcecode:: json

    {
      "books": [
        {
          "app": 1104, 
          "fullname": "Book1", 
          "name": "Book1", 
          "names": [], 
          "selection": "Sheet2!$A$1", 
          "sheets": [
            "Sheet1"
          ]
        }, 
        {
          "app": 1104, 
          "fullname": "C:\\Users\\felix\\DEV\\xlwings\\scripts\\Book1.xlsx", 
          "name": "Book1.xlsx", 
          "names": [
            "Sheet1!myname1", 
            "myname2"
          ], 
          "selection": "Sheet2!$A$1", 
          "sheets": [
            "Sheet1", 
            "Sheet2"
          ]
        }, 
        {
          "app": 1104, 
          "fullname": "Book4", 
          "name": "Book4", 
          "names": [], 
          "selection": "Sheet2!$A$1", 
          "sheets": [
            "Sheet1"
          ]
        }
      ]
    }

.. http:get:: /books/<book_name_or_ix>

**Example response**:

.. sourcecode:: json

    {
      "app": 1104, 
      "fullname": "C:\\Users\\felix\\DEV\\xlwings\\scripts\\Book1.xlsx", 
      "name": "Book1.xlsx", 
      "names": [
        "Sheet1!myname1", 
        "myname2"
      ], 
      "selection": "Sheet2!$A$1", 
      "sheets": [
        "Sheet1", 
        "Sheet2"
      ]
    }

.. http:get:: /books/<book_name_or_ix>/names

**Example response**:

.. sourcecode:: json

    {
      "names": [
        {
          "name": "Sheet1!myname1", 
          "refers_to": "=Sheet1!$B$2:$C$3"
        }, 
        {
          "name": "myname2", 
          "refers_to": "=Sheet1!$A$1"
        }
      ]
    }

.. http:get:: /books/<book_name_or_ix>/names/<name>

**Example response**:

.. sourcecode:: json

    {
      "name": "myname2", 
      "refers_to": "=Sheet1!$A$1"
    }

.. http:get:: /books/<book_name_or_ix>/names/<name>/range

**Example response**:

.. sourcecode:: json

    {
      "address": "$A$1", 
      "color": null, 
      "column": 1, 
      "column_width": 8.47, 
      "count": 1, 
      "current_region": "$A$1:$B$2", 
      "formula": "=1+1.1", 
      "formula_array": "=1+1,1", 
      "height": 14.25, 
      "last_cell": "$A$1", 
      "left": 0.0, 
      "name": "myname2", 
      "number_format": "General", 
      "row": 1, 
      "row_height": 14.3, 
      "shape": [
        1, 
        1
      ], 
      "size": 1, 
      "top": 0.0, 
      "value": 2.1, 
      "width": 51.0
    }

.. http:get:: /books/<book_name_or_ix>/sheets

**Example response**:

.. sourcecode:: json

    {
      "sheets": [
        {
          "charts": [
            "Chart 1"
          ], 
          "name": "Sheet1", 
          "names": [
            "Sheet1!myname1"
          ], 
          "pictures": [
            "Picture 3"
          ], 
          "shapes": [
            "Chart 1", 
            "Picture 3"
          ], 
          "used_range": "$A$1:$B$2"
        }, 
        {
          "charts": [], 
          "name": "Sheet2", 
          "names": [], 
          "pictures": [], 
          "shapes": [], 
          "used_range": "$A$1"
        }
      ]
    }

.. http:get:: /books/<book_name_or_ix>/sheets/<sheet_name_or_ix>

**Example response**:

.. sourcecode:: json

    {
      "charts": [
        "Chart 1"
      ], 
      "name": "Sheet1", 
      "names": [
        "Sheet1!myname1"
      ], 
      "pictures": [
        "Picture 3"
      ], 
      "shapes": [
        "Chart 1", 
        "Picture 3"
      ], 
      "used_range": "$A$1:$B$2"
    }

.. http:get:: /books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/charts

**Example response**:

.. sourcecode:: json

    {
      "charts": [
        {
          "chart_type": "line", 
          "height": 211.0, 
          "left": 0.0, 
          "name": "Chart 1", 
          "top": 0.0, 
          "width": 355.0
        }
      ]
    }

.. http:get:: /books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/charts/<chart_name_or_ix>

**Example response**:

.. sourcecode:: json

    {
      "chart_type": "line", 
      "height": 211.0, 
      "left": 0.0, 
      "name": "Chart 1", 
      "top": 0.0, 
      "width": 355.0
    }

.. http:get:: /books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/names

**Example response**:

.. sourcecode:: json

    {
      "names": [
        {
          "name": "Sheet1!myname1", 
          "refers_to": "=Sheet1!$B$2:$C$3"
        }
      ]
    }

.. http:get:: /books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/names/<sheet_scope_name>

**Example response**:

.. sourcecode:: json

    {
      "name": "Sheet1!myname1", 
      "refers_to": "=Sheet1!$B$2:$C$3"
    }

.. http:get:: /books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/names/<sheet_scope_name>/range

**Example response**:

.. sourcecode:: json

    {
      "address": "$B$2:$C$3", 
      "color": null, 
      "column": 2, 
      "column_width": 8.47, 
      "count": 4, 
      "current_region": "$A$1:$B$2", 
      "formula": [
        [
          "", 
          ""
        ], 
        [
          "", 
          ""
        ]
      ], 
      "formula_array": "", 
      "height": 28.5, 
      "last_cell": "$C$3", 
      "left": 51.0, 
      "name": "Sheet1!myname1", 
      "number_format": "General", 
      "row": 2, 
      "row_height": 14.3, 
      "shape": [
        2, 
        2
      ], 
      "size": 4, 
      "top": 14.25, 
      "value": [
        [
          null, 
          null
        ], 
        [
          null, 
          null
        ]
      ], 
      "width": 102.0
    }

.. http:get:: /books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/pictures

**Example response**:

.. sourcecode:: json

    {
      "pictures": [
        {
          "height": 100.0, 
          "left": 0.0, 
          "name": "Picture 3", 
          "top": 0.0, 
          "width": 100.0
        }
      ]
    }

.. http:get:: /books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/pictures/<picture_name_or_ix>

**Example response**:

.. sourcecode:: json

    {
      "height": 100.0, 
      "left": 0.0, 
      "name": "Picture 3", 
      "top": 0.0, 
      "width": 100.0
    }

.. http:get:: /books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/range

**Example response**:

.. sourcecode:: json

    {
      "address": "$A$1:$B$2", 
      "color": null, 
      "column": 1, 
      "column_width": 8.47, 
      "count": 4, 
      "current_region": "$A$1:$B$2", 
      "formula": [
        [
          "=1+1.1", 
          "a string"
        ], 
        [
          "43395.0064583333", 
          ""
        ]
      ], 
      "formula_array": null, 
      "height": 28.5, 
      "last_cell": "$B$2", 
      "left": 0.0, 
      "name": null, 
      "number_format": null, 
      "row": 1, 
      "row_height": 14.3, 
      "shape": [
        2, 
        2
      ], 
      "size": 4, 
      "top": 0.0, 
      "value": [
        [
          2.1, 
          "a string"
        ], 
        [
          "Mon, 22 Oct 2018 00:09:18 GMT", 
          null
        ]
      ], 
      "width": 102.0
    }

.. http:get:: /books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/range/<address>

**Example response**:

.. sourcecode:: json

    {
      "address": "$A$1:$B$2", 
      "color": null, 
      "column": 1, 
      "column_width": 8.47, 
      "count": 4, 
      "current_region": "$A$1:$B$2", 
      "formula": [
        [
          "=1+1.1", 
          "a string"
        ], 
        [
          "43395.0064583333", 
          ""
        ]
      ], 
      "formula_array": null, 
      "height": 28.5, 
      "last_cell": "$B$2", 
      "left": 0.0, 
      "name": null, 
      "number_format": null, 
      "row": 1, 
      "row_height": 14.3, 
      "shape": [
        2, 
        2
      ], 
      "size": 4, 
      "top": 0.0, 
      "value": [
        [
          2.1, 
          "a string"
        ], 
        [
          "Mon, 22 Oct 2018 00:09:18 GMT", 
          null
        ]
      ], 
      "width": 102.0
    }

.. http:get:: /books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/shapes

**Example response**:

.. sourcecode:: json

    {
      "shapes": [
        {
          "height": 211.0, 
          "left": 0.0, 
          "name": "Chart 1", 
          "top": 0.0, 
          "type": "chart", 
          "width": 355.0
        }, 
        {
          "height": 100.0, 
          "left": 0.0, 
          "name": "Picture 3", 
          "top": 0.0, 
          "type": "picture", 
          "width": 100.0
        }
      ]
    }

.. http:get:: /books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/shapes/<shape_name_or_ix>

**Example response**:

.. sourcecode:: json

    {
      "height": 211.0, 
      "left": 0.0, 
      "name": "Chart 1", 
      "top": 0.0, 
      "type": "chart", 
      "width": 355.0
    }

.. _apps:

/apps
*****

.. http:get:: /apps

**Example response**:

.. sourcecode:: json

    {
      "apps": [
        {
          "books": [
            "Book1", 
            "C:\\Users\\felix\\DEV\\xlwings\\scripts\\Book1.xlsx", 
            "Book4"
          ], 
          "calculation": "automatic", 
          "display_alerts": true, 
          "pid": 1104, 
          "screen_updating": true, 
          "selection": "[Book1.xlsx]Sheet2!$A$1", 
          "version": "16.0", 
          "visible": true
        }, 
        {
          "books": [
            "Book2", 
            "Book5"
          ], 
          "calculation": "automatic", 
          "display_alerts": true, 
          "pid": 7920, 
          "screen_updating": true, 
          "selection": "[Book5]Sheet2!$A$1", 
          "version": "16.0", 
          "visible": true
        }
      ]
    }

.. http:get:: /apps/<pid>

**Example response**:

.. sourcecode:: json

    {
      "books": [
        "Book1", 
        "C:\\Users\\felix\\DEV\\xlwings\\scripts\\Book1.xlsx", 
        "Book4"
      ], 
      "calculation": "automatic", 
      "display_alerts": true, 
      "pid": 1104, 
      "screen_updating": true, 
      "selection": "[Book1.xlsx]Sheet2!$A$1", 
      "version": "16.0", 
      "visible": true
    }

.. http:get:: /apps/<pid>/books

**Example response**:

.. sourcecode:: json

    {
      "books": [
        {
          "app": 1104, 
          "fullname": "Book1", 
          "name": "Book1", 
          "names": [], 
          "selection": "Sheet2!$A$1", 
          "sheets": [
            "Sheet1"
          ]
        }, 
        {
          "app": 1104, 
          "fullname": "C:\\Users\\felix\\DEV\\xlwings\\scripts\\Book1.xlsx", 
          "name": "Book1.xlsx", 
          "names": [
            "Sheet1!myname1", 
            "myname2"
          ], 
          "selection": "Sheet2!$A$1", 
          "sheets": [
            "Sheet1", 
            "Sheet2"
          ]
        }, 
        {
          "app": 1104, 
          "fullname": "Book4", 
          "name": "Book4", 
          "names": [], 
          "selection": "Sheet2!$A$1", 
          "sheets": [
            "Sheet1"
          ]
        }
      ]
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>

**Example response**:

.. sourcecode:: json

    {
      "app": 1104, 
      "fullname": "C:\\Users\\felix\\DEV\\xlwings\\scripts\\Book1.xlsx", 
      "name": "Book1.xlsx", 
      "names": [
        "Sheet1!myname1", 
        "myname2"
      ], 
      "selection": "Sheet2!$A$1", 
      "sheets": [
        "Sheet1", 
        "Sheet2"
      ]
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/names

**Example response**:

.. sourcecode:: json

    {
      "names": [
        {
          "name": "Sheet1!myname1", 
          "refers_to": "=Sheet1!$B$2:$C$3"
        }, 
        {
          "name": "myname2", 
          "refers_to": "=Sheet1!$A$1"
        }
      ]
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/names/<name>

**Example response**:

.. sourcecode:: json

    {
      "name": "myname2", 
      "refers_to": "=Sheet1!$A$1"
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/names/<name>/range

**Example response**:

.. sourcecode:: json

    {
      "address": "$A$1", 
      "color": null, 
      "column": 1, 
      "column_width": 8.47, 
      "count": 1, 
      "current_region": "$A$1:$B$2", 
      "formula": "=1+1.1", 
      "formula_array": "=1+1,1", 
      "height": 14.25, 
      "last_cell": "$A$1", 
      "left": 0.0, 
      "name": "myname2", 
      "number_format": "General", 
      "row": 1, 
      "row_height": 14.3, 
      "shape": [
        1, 
        1
      ], 
      "size": 1, 
      "top": 0.0, 
      "value": 2.1, 
      "width": 51.0
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/sheets

**Example response**:

.. sourcecode:: json

    {
      "sheets": [
        {
          "charts": [
            "Chart 1"
          ], 
          "name": "Sheet1", 
          "names": [
            "Sheet1!myname1"
          ], 
          "pictures": [
            "Picture 3"
          ], 
          "shapes": [
            "Chart 1", 
            "Picture 3"
          ], 
          "used_range": "$A$1:$B$2"
        }, 
        {
          "charts": [], 
          "name": "Sheet2", 
          "names": [], 
          "pictures": [], 
          "shapes": [], 
          "used_range": "$A$1"
        }
      ]
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>

**Example response**:

.. sourcecode:: json

    {
      "charts": [
        "Chart 1"
      ], 
      "name": "Sheet1", 
      "names": [
        "Sheet1!myname1"
      ], 
      "pictures": [
        "Picture 3"
      ], 
      "shapes": [
        "Chart 1", 
        "Picture 3"
      ], 
      "used_range": "$A$1:$B$2"
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/charts

**Example response**:

.. sourcecode:: json

    {
      "charts": [
        {
          "chart_type": "line", 
          "height": 211.0, 
          "left": 0.0, 
          "name": "Chart 1", 
          "top": 0.0, 
          "width": 355.0
        }
      ]
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/charts/<chart_name_or_ix>

**Example response**:

.. sourcecode:: json

    {
      "chart_type": "line", 
      "height": 211.0, 
      "left": 0.0, 
      "name": "Chart 1", 
      "top": 0.0, 
      "width": 355.0
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/names

**Example response**:

.. sourcecode:: json

    {
      "names": [
        {
          "name": "Sheet1!myname1", 
          "refers_to": "=Sheet1!$B$2:$C$3"
        }
      ]
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/names/<sheet_scope_name>

**Example response**:

.. sourcecode:: json

    {
      "name": "Sheet1!myname1", 
      "refers_to": "=Sheet1!$B$2:$C$3"
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/names/<sheet_scope_name>/range

**Example response**:

.. sourcecode:: json

    {
      "address": "$B$2:$C$3", 
      "color": null, 
      "column": 2, 
      "column_width": 8.47, 
      "count": 4, 
      "current_region": "$A$1:$B$2", 
      "formula": [
        [
          "", 
          ""
        ], 
        [
          "", 
          ""
        ]
      ], 
      "formula_array": "", 
      "height": 28.5, 
      "last_cell": "$C$3", 
      "left": 51.0, 
      "name": "Sheet1!myname1", 
      "number_format": "General", 
      "row": 2, 
      "row_height": 14.3, 
      "shape": [
        2, 
        2
      ], 
      "size": 4, 
      "top": 14.25, 
      "value": [
        [
          null, 
          null
        ], 
        [
          null, 
          null
        ]
      ], 
      "width": 102.0
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/pictures

**Example response**:

.. sourcecode:: json

    {
      "pictures": [
        {
          "height": 100.0, 
          "left": 0.0, 
          "name": "Picture 3", 
          "top": 0.0, 
          "width": 100.0
        }
      ]
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/pictures/<picture_name_or_ix>

**Example response**:

.. sourcecode:: json

    {
      "height": 100.0, 
      "left": 0.0, 
      "name": "Picture 3", 
      "top": 0.0, 
      "width": 100.0
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/range

**Example response**:

.. sourcecode:: json

    {
      "address": "$A$1:$B$2", 
      "color": null, 
      "column": 1, 
      "column_width": 8.47, 
      "count": 4, 
      "current_region": "$A$1:$B$2", 
      "formula": [
        [
          "=1+1.1", 
          "a string"
        ], 
        [
          "43395.0064583333", 
          ""
        ]
      ], 
      "formula_array": null, 
      "height": 28.5, 
      "last_cell": "$B$2", 
      "left": 0.0, 
      "name": null, 
      "number_format": null, 
      "row": 1, 
      "row_height": 14.3, 
      "shape": [
        2, 
        2
      ], 
      "size": 4, 
      "top": 0.0, 
      "value": [
        [
          2.1, 
          "a string"
        ], 
        [
          "Mon, 22 Oct 2018 00:09:18 GMT", 
          null
        ]
      ], 
      "width": 102.0
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/range/<address>

**Example response**:

.. sourcecode:: json

    {
      "address": "$A$1:$B$2", 
      "color": null, 
      "column": 1, 
      "column_width": 8.47, 
      "count": 4, 
      "current_region": "$A$1:$B$2", 
      "formula": [
        [
          "=1+1.1", 
          "a string"
        ], 
        [
          "43395.0064583333", 
          ""
        ]
      ], 
      "formula_array": null, 
      "height": 28.5, 
      "last_cell": "$B$2", 
      "left": 0.0, 
      "name": null, 
      "number_format": null, 
      "row": 1, 
      "row_height": 14.3, 
      "shape": [
        2, 
        2
      ], 
      "size": 4, 
      "top": 0.0, 
      "value": [
        [
          2.1, 
          "a string"
        ], 
        [
          "Mon, 22 Oct 2018 00:09:18 GMT", 
          null
        ]
      ], 
      "width": 102.0
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/shapes

**Example response**:

.. sourcecode:: json

    {
      "shapes": [
        {
          "height": 211.0, 
          "left": 0.0, 
          "name": "Chart 1", 
          "top": 0.0, 
          "type": "chart", 
          "width": 355.0
        }, 
        {
          "height": 100.0, 
          "left": 0.0, 
          "name": "Picture 3", 
          "top": 0.0, 
          "type": "picture", 
          "width": 100.0
        }
      ]
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/shapes/<shape_name_or_ix>

**Example response**:

.. sourcecode:: json

    {
      "height": 211.0, 
      "left": 0.0, 
      "name": "Chart 1", 
      "top": 0.0, 
      "type": "chart", 
      "width": 355.0
    }

