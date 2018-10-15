from xlwings.rest.api import api
import requests


BASE_URL = 'http://localhost:5000'

# Get all routes
get_urls = []
for rule in api.url_map.iter_rules():
    if 'GET' in rule.methods and 'static' not in rule.rule:
        get_urls.append(rule.rule)

get_urls = sorted(get_urls)

# params
pid = requests.get(BASE_URL + '/apps').json()['apps'][0]['pid']
wb = requests.get(BASE_URL + '/books/0').json()['name']

text = []
text.append('REST API')
text.append('========')
text.append('')
intro = """
xlwings offers an easy way to expose an Excel workbook via REST API both on macOS and Windows. You can run the REST API server from a command prompt or terminal::

    xlwings restapi run

This will run a Flask development server with the default args on http://127.0.0.1:5000. You can provide ``--host`` and ``--port`` as
command line args and it also accepts the Flask environment variables like ``FLASK_ENVIRONMENT``. Press ``Ctrl-C`` to terminate
the server again.

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

"""
text.append(intro)

for url in get_urls:
    text.append('.. http:get:: ' + url.replace('path:', ''))
    text.append('')
    rv = requests.get(BASE_URL +
                      url.replace('<pid>', str(pid))
                         .replace('<book_name_or_ix>', wb)
                         .replace('<chart_name_or_ix>', '0')
                         .replace('<book_scope_name>', 'myname')
                         .replace('<sheet_scope_name>', 'myname2')
                         .replace('<sheet_name_or_ix>', 'sheet1')
                         .replace('<shape_name_or_ix>', '0')
                         .replace('<path:fullname>', 'book1')
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
