REST API
========


New in v0.13.0

xlwings offers an easy way to expose an Excel workbook via REST API both on Windows and macOS. You can run the REST API
server from a command prompt or terminal as follows::

    xlwings restapi run

This will run a default Flask development server on http://127.0.0.1:5000. You can provide ``--host`` and ``--port`` as
command line args and it also respects the Flask environment variables like ``FLASK_ENVIRONMENT``. Press ``Ctrl-C`` to terminate
the server again.

If you want to have more control, you can just run the server directly with Flask, see the
`Flask docs <http://flask.pocoo.org/docs/1.0/quickstart/>`_ for more details::

    set FLASK_APP=xlwings.rest.api
    flask run

If you are on Mac, use ``export FLASK_APP=xlwings.rest.api`` instead of ``set FLASK_APP=xlwings.rest.api``.

.. note::
    Currently, we only provide the GET methods to read the workbook. If you are also interested in the POST methods
    to edit the workbook, let us know via GitHub issues.

For production, you can use any WSGI HTTP Server like `gunicorn <https://gunicorn.org/>`_ (on Mac) or `waitress
<https://docs.pylonsproject.org/projects/waitress/en/latest/>`_ (on Mac/Windows) to serve the API. For example,
with gunicorn you would do: ``gunicorn xlwings.rest.api:api``.

The xlwings REST API is a thin wrapper around the :ref:`xlwings Object API <api>` which makes it very easy if
you have worked previously with xlwings. It also means that the REST API does require the Excel application to be up and
running which makes it a great choice if the data in your Excel workbook is constantly changing.

To try things out, run ``xlwings restapi run`` from the command line and then paste the base url together with an endpoint
from below into your web browser or something more convenient like Postman or Insomnia. As an example, going to
http://localhost:5000/apps will give you back all open Excel instances and which workbooks they contain.

Endpoint overview
-----------------

+----------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------+
| Endpoint | Description                                                                                                                                                     |
+==========+=================================================================================================================================================================+
| /book    | Finds your workbook across all open instances of Excel and will open it if it can't find it. It will not work if you have the same workbook open in 2 instances |
+----------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------+
| /books   | Goes against the active instance of Excel                                                                                                                       |
+----------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------+
| /apps    | This allows you to specify the Excel instance you want to work with                                                                                             |
+----------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------+

Endpoint details
----------------


.. http:get:: /apps

**Example response**:

.. sourcecode:: json

    {
      "apps": [
        {
          "books": [
            "/Users/Felix/DEV/xlwings/scripts/Book1.xlsx", 
            "Book2"
          ], 
          "calculation": "automatic", 
          "display_alerts": true, 
          "pid": 16731, 
          "screen_updating": true, 
          "selection": "[Book2]Sheet1!$A$1", 
          "version": "16.19", 
          "visible": true
        }, 
        {
          "books": [
            "Book3"
          ], 
          "calculation": "automatic", 
          "display_alerts": true, 
          "pid": 16734, 
          "screen_updating": true, 
          "selection": "[Book3]Sheet2!$A$1", 
          "version": "16.19", 
          "visible": true
        }
      ]
    }

.. http:get:: /apps/<pid>

**Example response**:

.. sourcecode:: json

    {
      "books": [
        "/Users/Felix/DEV/xlwings/scripts/Book1.xlsx", 
        "Book2"
      ], 
      "calculation": "automatic", 
      "display_alerts": true, 
      "pid": 16731, 
      "screen_updating": true, 
      "selection": "[Book2]Sheet1!$A$1", 
      "version": "16.19", 
      "visible": true
    }

.. http:get:: /apps/<pid>/books

**Example response**:

.. sourcecode:: json

    {
      "books": [
        {
          "app": 16731, 
          "fullname": "/Users/Felix/DEV/xlwings/scripts/Book1.xlsx", 
          "name": "Book1.xlsx", 
          "names": [
            "Sheet1!myname1", 
            "myname2"
          ], 
          "selection": "Sheet1!$A$1", 
          "sheets": [
            "Sheet1", 
            "Sheet2"
          ]
        }, 
        {
          "app": 16731, 
          "fullname": "Book2", 
          "name": "Book2", 
          "names": [], 
          "selection": "Sheet1!$A$1", 
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
      "app": 16731, 
      "fullname": "/Users/Felix/DEV/xlwings/scripts/Book1.xlsx", 
      "name": "Book1.xlsx", 
      "names": [
        "Sheet1!myname1", 
        "myname2"
      ], 
      "selection": "Sheet1!$A$1", 
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

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/names/<book_scope_name>

**Example response**:

.. sourcecode:: json

    {
      "name": "myname2", 
      "refers_to": "=Sheet1!$A$1"
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
          "index": 1, 
          "name": "Sheet1", 
          "names": [
            "Sheet1!myname1"
          ], 
          "pictures": [
            "Chart 1", 
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
          "index": 2, 
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
      "index": 1, 
      "name": "Sheet1", 
      "names": [
        "Sheet1!myname1"
      ], 
      "pictures": [
        "Chart 1", 
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

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/pictures

**Example response**:

.. sourcecode:: json

    {
      "pictures": [
        {
          "height": 211.0, 
          "left": 0.0, 
          "name": "Chart 1", 
          "top": 0.0, 
          "width": 355.0
        }, 
        {
          "height": 60.0, 
          "left": 0.0, 
          "name": "Picture 3", 
          "top": 0.0, 
          "width": 60.0
        }
      ]
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/pictures/<picture_name_or_ix>

**Example response**:

.. sourcecode:: json

    {
      "height": 211.0, 
      "left": 0.0, 
      "name": "Chart 1", 
      "top": 0.0, 
      "width": 355.0
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/range

**Example response**:

.. sourcecode:: json

    {
      "address": "$A$1:$B$2", 
      "color": null, 
      "column": 1, 
      "column_width": 10.0, 
      "count": 4, 
      "current_region": "$A$1:$B$2", 
      "formula": [
        [
          "1.1", 
          "a string"
        ], 
        [
          "43390", 
          ""
        ]
      ], 
      "formula_array": null, 
      "height": 32.0, 
      "last_cell": "$B$2", 
      "left": 0.0, 
      "name": null, 
      "number_format": null, 
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
          1.1, 
          "a string"
        ], 
        [
          "Wed, 17 Oct 2018 00:00:00 GMT", 
          null
        ]
      ], 
      "width": 130.0
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/range/<address>

**Example response**:

.. sourcecode:: json

    {
      "address": "$A$1:$B$2", 
      "color": null, 
      "column": 1, 
      "column_width": 10.0, 
      "count": 4, 
      "current_region": "$A$1:$B$2", 
      "formula": [
        [
          "1.1", 
          "a string"
        ], 
        [
          "43390", 
          ""
        ]
      ], 
      "formula_array": null, 
      "height": 32.0, 
      "last_cell": "$B$2", 
      "left": 0.0, 
      "name": null, 
      "number_format": null, 
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
          1.1, 
          "a string"
        ], 
        [
          "Wed, 17 Oct 2018 00:00:00 GMT", 
          null
        ]
      ], 
      "width": 130.0
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
          "height": 60.0, 
          "left": 0.0, 
          "name": "Picture 3", 
          "top": 0.0, 
          "type": "picture", 
          "width": 60.0
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

.. http:get:: /book/<fullname>

**Example response**:

.. sourcecode:: json

    {
      "app": 16731, 
      "fullname": "/Users/Felix/DEV/xlwings/scripts/Book1.xlsx", 
      "name": "Book1.xlsx", 
      "names": [
        "Sheet1!myname1", 
        "myname2"
      ], 
      "selection": "Sheet1!$A$1", 
      "sheets": [
        "Sheet1", 
        "Sheet2"
      ]
    }

.. http:get:: /book/<fullname>/names

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

.. http:get:: /book/<fullname>/names/<book_scope_name>

**Example response**:

.. sourcecode:: json

    {
      "name": "myname2", 
      "refers_to": "=Sheet1!$A$1"
    }

.. http:get:: /book/<fullname>/sheets

**Example response**:

.. sourcecode:: json

    {
      "sheets": [
        {
          "charts": [
            "Chart 1"
          ], 
          "index": 1, 
          "name": "Sheet1", 
          "names": [
            "Sheet1!myname1"
          ], 
          "pictures": [
            "Chart 1", 
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
          "index": 2, 
          "name": "Sheet2", 
          "names": [], 
          "pictures": [], 
          "shapes": [], 
          "used_range": "$A$1"
        }
      ]
    }

.. http:get:: /book/<fullname>/sheets

**Example response**:

.. sourcecode:: json

    {
      "sheets": [
        {
          "charts": [
            "Chart 1"
          ], 
          "index": 1, 
          "name": "Sheet1", 
          "names": [
            "Sheet1!myname1"
          ], 
          "pictures": [
            "Chart 1", 
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
          "index": 2, 
          "name": "Sheet2", 
          "names": [], 
          "pictures": [], 
          "shapes": [], 
          "used_range": "$A$1"
        }
      ]
    }

.. http:get:: /book/<fullname>/sheets/<sheet_name_or_ix>/charts

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

.. http:get:: /book/<fullname>/sheets/<sheet_name_or_ix>/charts/<chart_name_or_ix>

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

.. http:get:: /book/<fullname>/sheets/<sheet_name_or_ix>/names

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

.. http:get:: /book/<fullname>/sheets/<sheet_name_or_ix>/names/<sheet_scope_name>

**Example response**:

.. sourcecode:: json

    {
      "name": "Sheet1!myname1", 
      "refers_to": "=Sheet1!$B$2:$C$3"
    }

.. http:get:: /book/<fullname>/sheets/<sheet_name_or_ix>/pictures

**Example response**:

.. sourcecode:: json

    {
      "pictures": [
        {
          "height": 211.0, 
          "left": 0.0, 
          "name": "Chart 1", 
          "top": 0.0, 
          "width": 355.0
        }, 
        {
          "height": 60.0, 
          "left": 0.0, 
          "name": "Picture 3", 
          "top": 0.0, 
          "width": 60.0
        }
      ]
    }

.. http:get:: /book/<fullname>/sheets/<sheet_name_or_ix>/pictures/<picture_name_or_ix>

**Example response**:

.. sourcecode:: json

    {
      "height": 211.0, 
      "left": 0.0, 
      "name": "Chart 1", 
      "top": 0.0, 
      "width": 355.0
    }

.. http:get:: /book/<fullname>/sheets/<sheet_name_or_ix>/range

**Example response**:

.. sourcecode:: json

    {
      "address": "$A$1:$B$2", 
      "color": null, 
      "column": 1, 
      "column_width": 10.0, 
      "count": 4, 
      "current_region": "$A$1:$B$2", 
      "formula": [
        [
          "1.1", 
          "a string"
        ], 
        [
          "43390", 
          ""
        ]
      ], 
      "formula_array": null, 
      "height": 32.0, 
      "last_cell": "$B$2", 
      "left": 0.0, 
      "name": null, 
      "number_format": null, 
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
          1.1, 
          "a string"
        ], 
        [
          "Wed, 17 Oct 2018 00:00:00 GMT", 
          null
        ]
      ], 
      "width": 130.0
    }

.. http:get:: /book/<fullname>/sheets/<sheet_name_or_ix>/range/<address>

**Example response**:

.. sourcecode:: json

    {
      "address": "$A$1:$B$2", 
      "color": null, 
      "column": 1, 
      "column_width": 10.0, 
      "count": 4, 
      "current_region": "$A$1:$B$2", 
      "formula": [
        [
          "1.1", 
          "a string"
        ], 
        [
          "43390", 
          ""
        ]
      ], 
      "formula_array": null, 
      "height": 32.0, 
      "last_cell": "$B$2", 
      "left": 0.0, 
      "name": null, 
      "number_format": null, 
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
          1.1, 
          "a string"
        ], 
        [
          "Wed, 17 Oct 2018 00:00:00 GMT", 
          null
        ]
      ], 
      "width": 130.0
    }

.. http:get:: /book/<fullname>/sheets/<sheet_name_or_ix>/shapes

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
          "height": 60.0, 
          "left": 0.0, 
          "name": "Picture 3", 
          "top": 0.0, 
          "type": "picture", 
          "width": 60.0
        }
      ]
    }

.. http:get:: /book/<fullname>/sheets/<sheet_name_or_ix>/shapes/<shape_name_or_ix>

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

.. http:get:: /books

**Example response**:

.. sourcecode:: json

    {
      "books": [
        {
          "app": 16731, 
          "fullname": "/Users/Felix/DEV/xlwings/scripts/Book1.xlsx", 
          "name": "Book1.xlsx", 
          "names": [
            "Sheet1!myname1", 
            "myname2"
          ], 
          "selection": "Sheet1!$A$1", 
          "sheets": [
            "Sheet1", 
            "Sheet2"
          ]
        }, 
        {
          "app": 16731, 
          "fullname": "Book2", 
          "name": "Book2", 
          "names": [], 
          "selection": "Sheet1!$A$1", 
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
      "app": 16731, 
      "fullname": "/Users/Felix/DEV/xlwings/scripts/Book1.xlsx", 
      "name": "Book1.xlsx", 
      "names": [
        "Sheet1!myname1", 
        "myname2"
      ], 
      "selection": "Sheet1!$A$1", 
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

.. http:get:: /books/<book_name_or_ix>/names/<book_scope_name>

**Example response**:

.. sourcecode:: json

    {
      "name": "myname2", 
      "refers_to": "=Sheet1!$A$1"
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
          "index": 1, 
          "name": "Sheet1", 
          "names": [
            "Sheet1!myname1"
          ], 
          "pictures": [
            "Chart 1", 
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
          "index": 2, 
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
      "index": 1, 
      "name": "Sheet1", 
      "names": [
        "Sheet1!myname1"
      ], 
      "pictures": [
        "Chart 1", 
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

.. http:get:: /books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/pictures

**Example response**:

.. sourcecode:: json

    {
      "pictures": [
        {
          "height": 211.0, 
          "left": 0.0, 
          "name": "Chart 1", 
          "top": 0.0, 
          "width": 355.0
        }, 
        {
          "height": 60.0, 
          "left": 0.0, 
          "name": "Picture 3", 
          "top": 0.0, 
          "width": 60.0
        }
      ]
    }

.. http:get:: /books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/pictures/<picture_name_or_ix>

**Example response**:

.. sourcecode:: json

    {
      "height": 211.0, 
      "left": 0.0, 
      "name": "Chart 1", 
      "top": 0.0, 
      "width": 355.0
    }

.. http:get:: /books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/range

**Example response**:

.. sourcecode:: json

    {
      "address": "$A$1:$B$2", 
      "color": null, 
      "column": 1, 
      "column_width": 10.0, 
      "count": 4, 
      "current_region": "$A$1:$B$2", 
      "formula": [
        [
          "1.1", 
          "a string"
        ], 
        [
          "43390", 
          ""
        ]
      ], 
      "formula_array": null, 
      "height": 32.0, 
      "last_cell": "$B$2", 
      "left": 0.0, 
      "name": null, 
      "number_format": null, 
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
          1.1, 
          "a string"
        ], 
        [
          "Wed, 17 Oct 2018 00:00:00 GMT", 
          null
        ]
      ], 
      "width": 130.0
    }

.. http:get:: /books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/range/<address>

**Example response**:

.. sourcecode:: json

    {
      "address": "$A$1:$B$2", 
      "color": null, 
      "column": 1, 
      "column_width": 10.0, 
      "count": 4, 
      "current_region": "$A$1:$B$2", 
      "formula": [
        [
          "1.1", 
          "a string"
        ], 
        [
          "43390", 
          ""
        ]
      ], 
      "formula_array": null, 
      "height": 32.0, 
      "last_cell": "$B$2", 
      "left": 0.0, 
      "name": null, 
      "number_format": null, 
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
          1.1, 
          "a string"
        ], 
        [
          "Wed, 17 Oct 2018 00:00:00 GMT", 
          null
        ]
      ], 
      "width": 130.0
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
          "height": 60.0, 
          "left": 0.0, 
          "name": "Picture 3", 
          "top": 0.0, 
          "type": "picture", 
          "width": 60.0
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

