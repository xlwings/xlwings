REST API
========


xlwings offers an easy way to expose an Excel workbook via REST API. You can run the REST API server like this from a command prompt or terminal::

    xlwings restapi run

This will run a Flask development server with the default args on http://localhost:5000. You can provide ``--host`` and ``--port`` as
command line args and it also accepts the Flask environment variables like ``FLASK_ENVIRONMENT``.

If you want to have more control, you can just run it directly with Flask, see http://flask.pocoo.org/docs/1.0/quickstart/
for more details::

    set FLASK_APP=xlwings.rest.api
    flask run

If you are on Mac, use ``export FLASK_APP=xlwings.rest.api`` instead of ``set FLASK_APP=xlwings.rest.api``.

.. note::
    Currently, we only provide the GET methods to read the workbook. If you are also interested in the POST methods
    to edit the workbook, let us know via GitHub issues.

For production, you can use a WSGI HTTP Server like gunicorn (on Mac) or waitress (on Mac/Windows) to
serve the API. For example, with gunicorn you would do: ``gunicorn xlwings.rest.api:api``.

The xlwings REST API is a thin wrapper around the :ref:`xlwings object API <api>` which makes it easy to learn as it
translates one-to-one. It also means that the REST API still requires the Excel application to be up and running which
makes sense if the data in your Excel workbook is constantly changing.

As a little recap, if you want xlwings to find your workbook across all open instances of Excel (called ``apps``
in xlwings), then use the ``/book/...`` endpoint. ``/books/...`` goes against the active app and if you need to specify
the app (usually when you have the same book open in 2 instances), then you have to use the ``/apps/...`` endpoint.

To try things out, run ``xlwings restapi run`` from the command line and then paste the base url toghether with an endpoint
from below into your web browser or something more convenient like Postman or Insomnia. As an example, going to
http://localhost:5000/apps will give you back all open Excel instances and which workbooks they contain.


.. http:get:: /apps

**Example response**:

.. sourcecode:: json

    {
      "apps": [
        {
          "books": [
            "Book1"
          ], 
          "calculation": "automatic", 
          "display_alerts": true, 
          "pid": 97508, 
          "screen_updating": true, 
          "selection": "$R$24", 
          "version": "16.19", 
          "visible": true
        }
      ]
    }

.. http:get:: /apps/<pid>/

**Example response**:

.. sourcecode:: json

    {
      "books": [
        "Book1"
      ], 
      "calculation": "automatic", 
      "display_alerts": true, 
      "pid": 97508, 
      "screen_updating": true, 
      "selection": "$R$24", 
      "version": "16.19", 
      "visible": true
    }

.. http:get:: /apps/<pid>/books

**Example response**:

.. sourcecode:: json

    {
      "books": [
        {
          "fullname": "Book1", 
          "name": "Book1", 
          "names": [
            "myname", 
            "Sheet1!myname2"
          ], 
          "selection": "$R$24", 
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
      "fullname": "Book1", 
      "name": "Book1", 
      "names": [
        "myname", 
        "Sheet1!myname2"
      ], 
      "selection": "$R$24", 
      "sheets": [
        "Sheet1"
      ]
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/names

**Example response**:

.. sourcecode:: json

    {
      "names": [
        {
          "name": "myname", 
          "refers_to": "=Sheet1!$D$18"
        }, 
        {
          "name": "Sheet1!myname2", 
          "refers_to": "=Sheet1!$C$12"
        }
      ]
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/names/<book_scope_name>

**Example response**:

.. sourcecode:: json

    {
      "name": "myname", 
      "refers_to": "=Sheet1!$D$18"
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
            "Sheet1!myname2"
          ], 
          "pictures": [
            "Chart 1", 
            "Picture 2"
          ], 
          "shapes": [
            "Chart 1", 
            "Picture 2"
          ]
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
        "Sheet1!myname2"
      ], 
      "pictures": [
        "Chart 1", 
        "Picture 2"
      ], 
      "shapes": [
        "Chart 1", 
        "Picture 2"
      ]
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/charts

**Example response**:

.. sourcecode:: json

    {
      "charts": [
        {
          "chart_type": "column_clustered", 
          "height": 216.0, 
          "left": 502.5, 
          "name": "Chart 1", 
          "top": 199.0, 
          "width": 360.0
        }
      ]
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/charts/<chart_name_or_ix>

**Example response**:

.. sourcecode:: json

    {
      "chart_type": "column_clustered", 
      "height": 216.0, 
      "left": 502.5, 
      "name": "Chart 1", 
      "top": 199.0, 
      "width": 360.0
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/names

**Example response**:

.. sourcecode:: json

    {
      "names": [
        {
          "name": "Sheet1!myname2", 
          "refers_to": "=Sheet1!$C$12"
        }
      ]
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/names/<sheet_scope_name>

**Example response**:

.. sourcecode:: json

    {
      "name": "Sheet1!myname2", 
      "refers_to": "=Sheet1!$C$12"
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/pictures

**Example response**:

.. sourcecode:: json

    {
      "pictures": [
        {
          "height": 216.0, 
          "left": 502.5, 
          "name": "Chart 1", 
          "top": 199.0, 
          "width": 360.0
        }, 
        {
          "height": 612.0, 
          "left": 200.0, 
          "name": "Picture 2", 
          "top": 240.0, 
          "width": 625.4505004882812
        }
      ]
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/pictures/<picture_name_or_ix>

**Example response**:

.. sourcecode:: json

    {
      "height": 216.0, 
      "left": 502.5, 
      "name": "Chart 1", 
      "top": 199.0, 
      "width": 360.0
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/range

**Example response**:

.. sourcecode:: json

    {
      "address": "$C$7:$C$8", 
      "color": null, 
      "column": 3, 
      "column_width": 10.0, 
      "count": 2, 
      "current_region": "$C$7:$C$8", 
      "formula": [
        [
          "1"
        ], 
        [
          "1"
        ]
      ], 
      "formula_array": "1", 
      "height": 32.0, 
      "last_cell": "$C$8", 
      "left": 130.0, 
      "name": null, 
      "number_format": "General", 
      "row": 7, 
      "row_height": 16.0, 
      "shape": [
        2, 
        1
      ], 
      "size": 2, 
      "top": 96.0, 
      "value": [
        1.0, 
        1.0
      ], 
      "width": 65.0
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
      "current_region": "$A$1", 
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
          null, 
          null
        ], 
        [
          null, 
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
          "height": 216.0, 
          "left": 502.5, 
          "name": "Chart 1", 
          "top": 199.0, 
          "type": "chart", 
          "width": 360.0
        }, 
        {
          "height": 612.0, 
          "left": 200.0, 
          "name": "Picture 2", 
          "top": 240.0, 
          "type": "picture", 
          "width": 625.4505004882812
        }
      ]
    }

.. http:get:: /apps/<pid>/books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/shapes/<shape_name_or_ix>

**Example response**:

.. sourcecode:: json

    {
      "height": 216.0, 
      "left": 502.5, 
      "name": "Chart 1", 
      "top": 199.0, 
      "type": "chart", 
      "width": 360.0
    }

.. http:get:: /book/<fullname>

**Example response**:

.. sourcecode:: json

    {
      "fullname": "Book1", 
      "name": "Book1", 
      "names": [
        "myname", 
        "Sheet1!myname2"
      ], 
      "selection": "$R$24", 
      "sheets": [
        "Sheet1"
      ]
    }

.. http:get:: /book/<fullname>/names

**Example response**:

.. sourcecode:: json

    {
      "names": [
        {
          "name": "myname", 
          "refers_to": "=Sheet1!$D$18"
        }, 
        {
          "name": "Sheet1!myname2", 
          "refers_to": "=Sheet1!$C$12"
        }
      ]
    }

.. http:get:: /book/<fullname>/names/<book_scope_name>

**Example response**:

.. sourcecode:: json

    {
      "name": "myname", 
      "refers_to": "=Sheet1!$D$18"
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
            "Sheet1!myname2"
          ], 
          "pictures": [
            "Chart 1", 
            "Picture 2"
          ], 
          "shapes": [
            "Chart 1", 
            "Picture 2"
          ]
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
            "Sheet1!myname2"
          ], 
          "pictures": [
            "Chart 1", 
            "Picture 2"
          ], 
          "shapes": [
            "Chart 1", 
            "Picture 2"
          ]
        }
      ]
    }

.. http:get:: /book/<fullname>/sheets/<sheet_name_or_ix>/charts

**Example response**:

.. sourcecode:: json

    {
      "charts": [
        {
          "chart_type": "column_clustered", 
          "height": 216.0, 
          "left": 502.5, 
          "name": "Chart 1", 
          "top": 199.0, 
          "width": 360.0
        }
      ]
    }

.. http:get:: /book/<fullname>/sheets/<sheet_name_or_ix>/charts/<chart_name_or_ix>

**Example response**:

.. sourcecode:: json

    {
      "chart_type": "column_clustered", 
      "height": 216.0, 
      "left": 502.5, 
      "name": "Chart 1", 
      "top": 199.0, 
      "width": 360.0
    }

.. http:get:: /book/<fullname>/sheets/<sheet_name_or_ix>/names

**Example response**:

.. sourcecode:: json

    {
      "names": [
        {
          "name": "Sheet1!myname2", 
          "refers_to": "=Sheet1!$C$12"
        }
      ]
    }

.. http:get:: /book/<fullname>/sheets/<sheet_name_or_ix>/names/<sheet_scope_name>

**Example response**:

.. sourcecode:: json

    {
      "name": "Sheet1!myname2", 
      "refers_to": "=Sheet1!$C$12"
    }

.. http:get:: /book/<fullname>/sheets/<sheet_name_or_ix>/pictures

**Example response**:

.. sourcecode:: json

    {
      "pictures": [
        {
          "height": 216.0, 
          "left": 502.5, 
          "name": "Chart 1", 
          "top": 199.0, 
          "width": 360.0
        }, 
        {
          "height": 612.0, 
          "left": 200.0, 
          "name": "Picture 2", 
          "top": 240.0, 
          "width": 625.4505004882812
        }
      ]
    }

.. http:get:: /book/<fullname>/sheets/<sheet_name_or_ix>/pictures/<picture_name_or_ix>

**Example response**:

.. sourcecode:: json

    {
      "height": 216.0, 
      "left": 502.5, 
      "name": "Chart 1", 
      "top": 199.0, 
      "width": 360.0
    }

.. http:get:: /book/<fullname>/sheets/<sheet_name_or_ix>/range

**Example response**:

.. sourcecode:: json

    {
      "address": "$C$7:$C$8", 
      "color": null, 
      "column": 3, 
      "column_width": 10.0, 
      "count": 2, 
      "current_region": "$C$7:$C$8", 
      "formula": [
        [
          "1"
        ], 
        [
          "1"
        ]
      ], 
      "formula_array": "1", 
      "height": 32.0, 
      "last_cell": "$C$8", 
      "left": 130.0, 
      "name": null, 
      "number_format": "General", 
      "row": 7, 
      "row_height": 16.0, 
      "shape": [
        2, 
        1
      ], 
      "size": 2, 
      "top": 96.0, 
      "value": [
        1.0, 
        1.0
      ], 
      "width": 65.0
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
      "current_region": "$A$1", 
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
          null, 
          null
        ], 
        [
          null, 
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
          "height": 216.0, 
          "left": 502.5, 
          "name": "Chart 1", 
          "top": 199.0, 
          "type": "chart", 
          "width": 360.0
        }, 
        {
          "height": 612.0, 
          "left": 200.0, 
          "name": "Picture 2", 
          "top": 240.0, 
          "type": "picture", 
          "width": 625.4505004882812
        }
      ]
    }

.. http:get:: /book/<fullname>/sheets/<sheet_name_or_ix>/shapes/<shape_name_or_ix>

**Example response**:

.. sourcecode:: json

    {
      "height": 216.0, 
      "left": 502.5, 
      "name": "Chart 1", 
      "top": 199.0, 
      "type": "chart", 
      "width": 360.0
    }

.. http:get:: /books

**Example response**:

.. sourcecode:: json

    {
      "books": [
        {
          "fullname": "Book1", 
          "name": "Book1", 
          "names": [
            "myname", 
            "Sheet1!myname2"
          ], 
          "selection": "$R$24", 
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
      "fullname": "Book1", 
      "name": "Book1", 
      "names": [
        "myname", 
        "Sheet1!myname2"
      ], 
      "selection": "$R$24", 
      "sheets": [
        "Sheet1"
      ]
    }

.. http:get:: /books/<book_name_or_ix>/names

**Example response**:

.. sourcecode:: json

    {
      "names": [
        {
          "name": "myname", 
          "refers_to": "=Sheet1!$D$18"
        }, 
        {
          "name": "Sheet1!myname2", 
          "refers_to": "=Sheet1!$C$12"
        }
      ]
    }

.. http:get:: /books/<book_name_or_ix>/names/<book_scope_name>

**Example response**:

.. sourcecode:: json

    {
      "name": "myname", 
      "refers_to": "=Sheet1!$D$18"
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
            "Sheet1!myname2"
          ], 
          "pictures": [
            "Chart 1", 
            "Picture 2"
          ], 
          "shapes": [
            "Chart 1", 
            "Picture 2"
          ]
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
        "Sheet1!myname2"
      ], 
      "pictures": [
        "Chart 1", 
        "Picture 2"
      ], 
      "shapes": [
        "Chart 1", 
        "Picture 2"
      ]
    }

.. http:get:: /books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/charts

**Example response**:

.. sourcecode:: json

    {
      "charts": [
        {
          "chart_type": "column_clustered", 
          "height": 216.0, 
          "left": 502.5, 
          "name": "Chart 1", 
          "top": 199.0, 
          "width": 360.0
        }
      ]
    }

.. http:get:: /books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/charts/<chart_name_or_ix>

**Example response**:

.. sourcecode:: json

    {
      "chart_type": "column_clustered", 
      "height": 216.0, 
      "left": 502.5, 
      "name": "Chart 1", 
      "top": 199.0, 
      "width": 360.0
    }

.. http:get:: /books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/names

**Example response**:

.. sourcecode:: json

    {
      "names": [
        {
          "name": "Sheet1!myname2", 
          "refers_to": "=Sheet1!$C$12"
        }
      ]
    }

.. http:get:: /books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/names/<sheet_scope_name>

**Example response**:

.. sourcecode:: json

    {
      "name": "Sheet1!myname2", 
      "refers_to": "=Sheet1!$C$12"
    }

.. http:get:: /books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/pictures

**Example response**:

.. sourcecode:: json

    {
      "pictures": [
        {
          "height": 216.0, 
          "left": 502.5, 
          "name": "Chart 1", 
          "top": 199.0, 
          "width": 360.0
        }, 
        {
          "height": 612.0, 
          "left": 200.0, 
          "name": "Picture 2", 
          "top": 240.0, 
          "width": 625.4505004882812
        }
      ]
    }

.. http:get:: /books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/pictures/<picture_name_or_ix>

**Example response**:

.. sourcecode:: json

    {
      "height": 216.0, 
      "left": 502.5, 
      "name": "Chart 1", 
      "top": 199.0, 
      "width": 360.0
    }

.. http:get:: /books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/range

**Example response**:

.. sourcecode:: json

    {
      "address": "$C$7:$C$8", 
      "color": null, 
      "column": 3, 
      "column_width": 10.0, 
      "count": 2, 
      "current_region": "$C$7:$C$8", 
      "formula": [
        [
          "1"
        ], 
        [
          "1"
        ]
      ], 
      "formula_array": "1", 
      "height": 32.0, 
      "last_cell": "$C$8", 
      "left": 130.0, 
      "name": null, 
      "number_format": "General", 
      "row": 7, 
      "row_height": 16.0, 
      "shape": [
        2, 
        1
      ], 
      "size": 2, 
      "top": 96.0, 
      "value": [
        1.0, 
        1.0
      ], 
      "width": 65.0
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
      "current_region": "$A$1", 
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
          null, 
          null
        ], 
        [
          null, 
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
          "height": 216.0, 
          "left": 502.5, 
          "name": "Chart 1", 
          "top": 199.0, 
          "type": "chart", 
          "width": 360.0
        }, 
        {
          "height": 612.0, 
          "left": 200.0, 
          "name": "Picture 2", 
          "top": 240.0, 
          "type": "picture", 
          "width": 625.4505004882812
        }
      ]
    }

.. http:get:: /books/<book_name_or_ix>/sheets/<sheet_name_or_ix>/shapes/<shape_name_or_ix>

**Example response**:

.. sourcecode:: json

    {
      "height": 216.0, 
      "left": 502.5, 
      "name": "Chart 1", 
      "top": 199.0, 
      "type": "chart", 
      "width": 360.0
    }

