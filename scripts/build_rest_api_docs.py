import os
import datetime as dt
import requests
import xlwings as xw
from xlwings.rest.api import api


BASE_URL = 'http://localhost:5000'

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

sheet1['A1'].value = [[1.1, 'a string'], [dt.datetime.now(), None]]
sheet1['A1'].formula = '=1+1.1'
chart = sheet1.charts.add()
chart.set_source_data(sheet1['A1'])
chart.chart_type = 'line'


pic = os.path.abspath(os.path.join('..', 'xlwings', 'tests', 'sample_picture.png'))
sheet1.pictures.add(pic)
wb1.sheets[0].range('B2:C3').name = 'Sheet1!myname1'
wb1.sheets[0].range('A1').name = 'myname2'
wb1.save('Book1.xlsx')
wb1 = xw.Book('Book1.xlsx')  # hack as save doesn't return the wb properly
app1.activate()

# Get all routes
get_urls = []
for rule in api.url_map.iter_rules():
    if 'GET' in rule.methods and 'static' not in rule.rule:
        get_urls.append(rule.rule)

get_urls = sorted(get_urls)

text = []
text.append('REST API')
text.append('========')
text.append('')
intro = """
New in v0.13.0

xlwings offers an easy way to expose an Excel workbook via REST API both on Windows and macOS. You can run the REST API
server from a command prompt or terminal as follows::

    xlwings restapi run

This will run a default Flask development server on http://127.0.0.1:5000. You can provide ``--host`` and ``--port`` as
command line args and it also respects the Flask environment variables like ``FLASK_ENV=development``. Press ``Ctrl-C`` to terminate
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
with gunicorn you would do: ``gunicorn xlwings.rest.api:api``. Or with waitress::

    from xlwings.rest.api import api
    from waitress import serve
    serve(api, listen='127.0.0.1:5000')

The xlwings REST API is a thin wrapper around the :ref:`xlwings Object API <api>` which makes it very easy if
you have worked previously with xlwings. It also means that the REST API does require the Excel application to be up and
running which makes it a great choice if the data in your Excel workbook is constantly changing.

To try things out, run ``xlwings restapi run`` from the command line and then paste the base url together with an endpoint
from below into your web browser or something more convenient like `Postman <https://www.getpostman.com/>`_ or
`Insomnia <https://insomnia.rest/>`_ (Curl works equally fine).
As an example, going to http://localhost:5000/apps will give you back all open Excel instances with the workbooks they
contain.

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

"""
text.append(intro)

for url in get_urls:
    text.append('.. http:get:: ' + url.replace('path:', ''))
    text.append('')
    rv = requests.get(BASE_URL +
                      url.replace('<pid>', str(wb1.app.pid))
                         .replace('<book_name_or_ix>', wb1.name)
                         .replace('<chart_name_or_ix>', '0')
                         .replace('<book_scope_name>', 'myname2')
                         .replace('<sheet_scope_name>', 'myname1')
                         .replace('<sheet_name_or_ix>', 'sheet1')
                         .replace('<shape_name_or_ix>', '0')
                         .replace('<path:fullname>', wb1.name)
                         .replace('<picture_name_or_ix>', '0')
                         .replace('<address>', 'A1:B2')
                      )
    text.append('**Example response**:')
    text.append('')
    text.append('.. sourcecode:: json')
    text.append('')
    for i in rv.text.splitlines():
        text.append('    ' + i)
    text.append('')


with open('../docs/rest_api.rst', 'w') as f:
    for line in text:
        f.write(f'{line}\n')


app1.kill()
app2.kill()
