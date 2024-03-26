Office.js Custom Functions
==========================

.. admonition:: Requirements

    * xlwings edition: PRO
    * Server OS: Windows, macOS, Linux
    * Excel platform: Windows, macOS, Web
    * Google Sheets: not supported (planned)
    * Minimum xlwings version: 0.30.0
    * Minimum Excel version: 2021 or 365

Quickstart
----------

Custom functions are based on Office.js add-ins. It's therefore a good idea to revisit the :ref:`Office.js Add-in docs <officejs_addins>`.

1. Follow the full :ref:`Office.js Add-in Quickstart <pro/server/officejs_addins:Quickstart>`. At the end of it, you should have the backend server running and the manifest sideloaded.
2. That's it! You can now use the custom functions that are defined in the quickstart project under ``app/custom_functions.py``: e.g., type ``=HELLO("xlwings")`` into a cell and hit Enter---you'll be greeted by ``Hello xlwings!``.

As long as you don't change the name or arguments of the function, you can edit the code in the ``app/custom_functions.py`` file and see the effect immediately by recalculating your formula. You can recalculate by either editing the cell and hitting Enter again, or by hitting ``Ctrl-Alt-F9`` (Windows) or ``Ctrl-Option-F9`` (macOS). If you add new functions or make changes to function names or arguments of existing functions, you'll need to sideload the add-in again.

Basic syntax
------------

As you could see in the quickstart sample, the simplest custom function only requires the ``@server.func`` decorator:

.. code-block:: python

    from xlwings import server

    @server.func
    def hello(name):
        return f"Hello {name}!"

.. note::

    The decorators for Office.js are imported from ``xlwings.server`` instead of ``xlwings`` and therefore read ``server.func`` instead of ``xw.func``. See also :ref:`pro/server/officejs_custom_functions:Custom functions vs. legacy UDFs`.

Python modules
--------------

By default, xlwings expects the functions to live in a module called ``custom_functions.py``.

* If you want to call your module differently, import it like so: ``import your_module as custom_functions``
* If you want to store your custom functions across different modules/packages, import them into ``custom_functions.py``:

  .. code-block:: python
  
      # custom_functions.py
      from mypackage.subpackage import func1, func2
      from mymodule import func3, func4

Note that ``custom_functions`` needs to be imported where you define the required endpoints in your web framework, see :ref:`pro/server/officejs_custom_functions:backend and manifest`.

pandas DataFrames
-----------------

By using the ``@server.arg`` and ``@server.ret`` decorators, you can apply converters and options to arguments and the return value, respectively.

For example, to read in the values of a range as pandas DataFrame and return the correlations without writing out the header and the index, you would write:

.. code-block:: python

    import pandas as pd
    from xlwings import server

    @server.func
    @server.arg("df", pd.DataFrame, index=False, header=False)
    @server.ret(index=False, header=False)
    def correl2(df):
        return df.corr()

For an overview of the available converters and options, have a look at :ref:`converters`.

Variable number of arguments (``*args``)
----------------------------------------

.. versionadded:: 0.30.15

Varargs are supported. You can also use a converter, which will be applied to all arguments provided by ``*args``:

.. code-block:: python

  from xlwings import server

  @server.func
  @server.arg("*args", pd.DataFrame, index=False)
  def concat(*args):
      return pd.concat(args)

Doc strings
-----------

To describe your function and its arguments, you can use a function docstring or the ``arg`` decorator, respectively:

.. code-block:: python

    from xlwings import server

    @server.func
    @server.arg("name", doc='A name such as "World"')
    def hello(name):
        """This is a classic Hello World example"""
        return f"Hello {name}!"

These doc strings will appear in Excel's function wizard/formula builder. Note that the name of the arguments will automatically be shown when typing the formula into a cell (intellisense).

Date and time
-------------

Depending on whether you're reading from Excel or writing to Excel, there are different tools available to work with date and time.

**Reading**

In the context of custom functions, xlwings will detect numbers, strings, and booleans but not cells with a date/time format. Hence, you need to use converters. For single datetime arguments do this:

.. code-block:: python

    import datetime as dt
    from xlwings import server

    @server.func
    @server.arg("date", dt.datetime)
    def isoformat(date):
        return date.isoformat()

Instead of ``dt.datetime``, you can also use ``dt.date`` to get a date object instead.

If you have multiple values that you need to convert, you can use the ``xlwings.to_datetime()`` function:

.. code-block:: python

    import datetime as dt
    import xlwings as xw
    from xlwings import server

    @server.func
    def isoformat(dates):
        dates = [xw.to_datetime(d) for d in dates]
        return [d.isoformat() for d in dates]

And if you are dealing with pandas DataFrames, you can simply use the ``parse_dates`` option. It behaves the same as with ``pandas.read_csv()``:

.. code-block:: python

    import pandas as pd
    from xlwings import server

    @server.func
    @server.arg("df", pd.DataFrame, parse_dates=[0])
    def timeseries_start(df):
        return df.index.min()

Like ``pandas.read_csv()``, you could also provide ``parse_dates`` with a list of columns names instead of indices.

**Writing**

When writing datetime object to Excel, xlwings automatically formats the cells as date if your version of Excel supports data types, so no special handling is required:

.. code-block:: python

    import datetime as dt
    import xlwings as xw
    from xlwings import server

    @server.func
    def pytoday():
        return dt.date.today()

By default, it will format the date according to the content language of your Excel instance, but you can also override this by explicitly providing the ``date_format`` option:

.. code-block:: python

    import datetime as dt
    import xlwings as xw
    from xlwings import server

    @server.func
    @server.ret(date_format="yyyy-m-d")
    def pytoday():
        return dt.date.today()

For the accepted ``date_format`` string, consult the `official Excel documentation <https://support.microsoft.com/en-us/office/format-numbers-as-dates-or-times-418bd3fe-0577-47c8-8caa-b4d30c528309>`_.

.. note::

    Some older builds of Excel don't support date formatting and will display the date as date serial instead, requiring you format it manually. See also :ref:`pro/server/officejs_custom_functions:limitations`.

Namespace
---------

A namespace groups related custom functions together by prepending the namespace to the function name, separated with a dot. For example, to have NumPy-related functions show up under the numpy namespace, you would do:

.. code-block:: python

    import numpy as np
    from xlwings import server

    @server.func(namespace="numpy")
    def standard_normal(rows, columns):
        rng = np.random.default_rng()
        return rng.standard_normal(size=(rows, columns))

This function will be shown as ``NUMPY.STANDARD_NORMAL`` in Excel.

**Sub-namespace**

You can create sub-namespaces by including a dot like so:

.. code-block:: python

    @server.func(namespace="numpy.random")

This function will be shown as ``NUMPY.RANDOM.STANDARD_NORMAL`` in Excel.

**Default namespace**

If you want all your functions to appear under a common namespace, you can include the following line under the ShortStrings sections in the manifest XML:

.. code-block:: xml

    <bt:String id="Functions.Namespace" DefaultValue="XLWINGS"/>

Have a look at ``manifest-xlwings-officejs-quickstart.xml`` where the respective line is commented out.

If you define a namespace as part of the function decorator while also having a default namespace defined, the namespace from the function decorator will define the sub-namespace.

Help URL
--------

You can include a link to an internet page with more information about your function by using the ``help_url`` option. The function wizard/formula builder will show that link under "More help on this function".

.. code-block:: python

    from xlwings import server

    @server.func(help_url="https://www.xlwings.org")
    def hello(name):
        return f"Hello {name}!"


Array Dimensions
----------------

If you want your function to accept arguments of any dimensions (as single cell or one- or two-dimensional ranges), you may need to use the ``ndim`` option to make your code work in every case. Likewise, there's an easy trick to return a simple list in a vertical orientation by using the ``transpose`` option.

**Arguments**

Depending on the dimensionality of the function parameters, xlwings either delivers a scalar, a list, or a nested list:

* Single cells (e.g., ``A1``) arrive as scalar, i.e., number, string, or boolean: ``1`` or ``"text"``, or ``True``
* A one-dimensional (vertical or horizontal!) range (e.g. ``A1:B1`` or ``A1:A2``) arrives as list: ``[1, 2]``
* A two-dimensional range (e.g., ``A1:B2``) arrives as nested list: ``[[1, 2], [3, 4]]``

This behavior is not only consistent in itself, it's also in line with how NumPy works and is often what you want: for example, you can directly loop over a vertical 1-dimensional range of cells.

However, if the argument can be anything from a single cell to a one- or two-dimensional range, you'll want to use the ``ndim`` option: this allows you to always get the inputs as a two-dimensional list, no matter what the input dimension is:

.. code-block:: python

    from xlwings import server

    @server.func
    @server.arg("x", ndim=2)
    def add_one(x):
        return [[cell + 1 for cell in row] for row in data]

The above sample would raise an error if you'd leave away the ``ndim=2`` and use a single cell as argument ``x``.

**Return value**

If you need to write out a list in vertical orientation, the ``transpose`` option comes in handy:

.. code-block:: python

    from xlwings import server

    @server.func
    @server.ret(transpose=True)
    def vertical_list():
        return [1, 2, 3, 4]

Error handling and error cells
------------------------------

Error cells in Excel such as ``#VALUE!`` are used to display an error from Python. xlwings reads error cells as ``None`` by default but also allows you to read them as strings. When writing to Excel, you can Excel have an cell formatted as error. Let's get into the details!

Error handling
**************

Whenever there's an error in Python, the cell value will show ``#VALUE!``. To understand what's going on, click on the cell with the error, then hover (don't click!) on the exclamation mark that appears: you'll see the error message.

If you see ``Internal Server Error``, you need to consult the Python server logs or you can add an exception handler for the type of Exception that you'd like to see in more detail on the frontend, see the function ``xlwings_exception_handler`` in the quickstart project under ``app/server_fastapi.py``.

Writing NaN values
******************

``np.nan`` and ``pd.NA`` will be converted to Excel's ``#NUM!`` error type.

Error cells
***********

**Reading**

By default, error cells are converted to ``None`` (scalars and lists) or ``np.nan`` (NumPy arrays and pandas DataFrames). If you'd like to get them in their string representation, use ``err_to_str`` option:

.. code-block:: python

    from xlwings import server

    @server.func
    @server.arg("x", err_to_str=True)
    def myfunc(x):
        ...

**Writing**

To format cells as proper error cells in Excel, simply use their string representation (``#DIV/0!``, ``#N/A``, ``#NAME?``, ``#NULL!``, ``#NUM!``, ``#REF!``, ``#VALUE!``):

.. code-block:: python

    from xlwings import server

    @server.func
    def myfunc(x):
        return ["#N/A", "#VALUE!"]

.. note::

    Some older builds of Excel don't support proper error types and will display the error as string instead, see also :ref:`pro/server/officejs_custom_functions:limitations`.

Dynamic arrays
--------------

If your return value is not just a single value but a one- or two-dimensional list, Excel will automatically spill the values into the surrounding cells by using the native dynamic arrays. There are no code changes required:

Returning a simple list:

.. code-block:: python

    from xlwings import server

    @server.func
    def programming_languages():
        return ["Python", "JavaScript"]

Returning a NumPy array with standard normally distributed random numbers:

.. code-block:: python

    import numpy as np
    from xlwings import server

    @server.func
    def standard_normal(rows, columns):
        rng = np.random.default_rng()
        return rng.standard_normal(size=(rows, columns))

Returning a pandas DataFrame:

.. code-block:: python

    import pandas as pd
    from xlwings import server

    @server.func
    def get_dataframe():
        df = pd.DataFrame({"Language": ["Python", "JavaScript"], "Year": [1991, 1995]})
        return df

Volatile functions
------------------

Volatile functions are recalculated whenever Excel calculates something, even if none of the function arguments have changed. To mark a function as volatile, use the ``volatile`` argument in the ``func`` decorator:

.. code-block:: python

    import datetime as dt
    from xlwings import server

    @server.func(volatile=True)
    def last_calculated():
        return f"Last calculated: {dt.datetime.now()}"

Asynchronous functions
----------------------

Custom functions are always asynchronous, meaning that the cell will show ``#BUSY!`` during calculation, allowing you to continue using Excel: custom function don't block Excel's user interface.

Streaming functions ("RTD functions")
-------------------------------------

In the traditional version of Excel, streaming functions were called "RTD functions" or "RealTimeData functions". However, unlike traditional RTD functions, streaming functions don't use a local COM server. Instead, the process runs as a background task on xlwings Server and pushes updates via WebSockets (using Socket.io) to Excel. What's great about streaming functions is that you can connect to your data source in a single place and stream the values to every Excel installation in your entire company.

To create a streaming function, you simply need to write an asynchronous generator. That is, you need to use ``async def`` and ``yield`` instead of ``return``, e.g.:

.. code-block:: python

  import asyncio
  from xlwings import server

  @server.func
  async def streaming_random(rows, cols):
      """A streaming function pushing updates of a random DataFrame every second"""
      rng = np.random.default_rng()
      while True:
          matrix = rng.standard_normal(size=(rows, cols))
          df = pd.DataFrame(matrix, columns=[f"col{i+1}" for i in range(matrix.shape[1])])
          yield df
          await asyncio.sleep(1)

As a bit of a more real-world sample, here's how you can transform a REST API into a streaming function to stream the BTC price:

.. code-block:: python

  import asyncio
  from xlwings import server

  @server.func
  @server.ret(date_format="hh:mm:ss", index=False)
  async def btc_price(base_currency="USD"):
      while True:
          async with httpx.AsyncClient() as client:
              response = await client.get(
                  f"https://cex.io/api/ticker/BTC/{base_currency}"
              )
          response_data = response.json()
          response_data["timestamp"] = pd.to_datetime(
              int(response_data["timestamp"]), unit="s"
          )
          df = pd.DataFrame(response_data, index=[0])
          df = df[["pair", "timestamp", "bid", "ask"]]
          yield df
          await asyncio.sleep(1)

Key to remember is that you're moving in the async world with streaming functions, so you shouldn't use long-running blocking operations. For example, instead of using ``requests`` to fetch the data, you should use one of the async libraries such as ``httpx`` or ``aiohttp``.

If you use the `official xlwings Server <https://github.com/xlwings/xlwings-server>`_ implementation, that's all you need because it supports streaming functions out-of-the-box. If you're using your own server implementation, you'll need to implement the Socket.io endpoints according to the official xlwings Server implementation.


Backend and Manifest
--------------------

This section highlights which part of the code in ``app/server_fastapi.py``, ``app/taskpane.html`` and ``manifest-xlwings-officejs-quickstart.xml`` are responsible for handling custom functions. They are already implemented in the quickstart project.

Backend
*******

The backend needs to implement the following three endpoints to support custom functions. You can check them out under ``app/server_fastapi.py`` or in one of the other framework implementations.

.. tab-set::
    .. tab-item:: FastAPI
      :sync: fastapi

      .. code-block::

          import xlwings as xw
          import custom_functions

          @app.get("/xlwings/custom-functions-meta")
          async def custom_functions_meta():
              return xw.server.custom_functions_meta(custom_functions)
  
  
          @app.get("/xlwings/custom-functions-code")
          async def custom_functions_code():
              return PlainTextResponse(xw.server.custom_functions_code(custom_functions))
  
  
          @app.post("/xlwings/custom-functions-call")
          async def custom_functions_call(data: dict = Body):
              rv = await xw.server.custom_functions_call(data, custom_functions)
              return {"result": rv}

    .. tab-item:: Starlette
      :sync: starlette

      .. code-block::

          import xlwings as xw
          import custom_functions

          async def custom_functions_meta(request):
              return JSONResponse(xw.server.custom_functions_meta(custom_functions))


          async def custom_functions_code(request):
              return PlainTextResponse(xw.server.custom_functions_code(custom_functions))


          async def custom_functions_call(request):
              data = await request.json()
              rv = await xw.server.custom_functions_call(data, custom_functions)
              return JSONResponse({"result": rv})

You'll also need to load the custom functions by adding the following line at the end of the ``head`` element in your HTML file, see ``app/taskpane.html`` in the quickstart project:

.. code-block:: html

    <head>
      <!-- ... -->
      <script type="text/javascript" src="/xlwings/custom-functions-code"></script>
    </head>

Manifest
********

The relevant parts in the manifest XML are:

.. code-block:: xml

    <Requirements>
        <Sets DefaultMinVersion="1.1">
        <Set Name="SharedRuntime" MinVersion="1.1"/>
        </Sets>
    </Requirements>

And:

.. code-block:: xml

    <Runtimes>
        <Runtime resid="Taskpane.Url" lifetime="long"/>
    </Runtimes>
    <AllFormFactors>
        <ExtensionPoint xsi:type="CustomFunctions">
        <Script>
            <SourceLocation resid="Functions.Script.Url"/>
        </Script>
        <Page>
            <SourceLocation resid="Taskpane.Url"/>
        </Page>
        <Metadata>
            <SourceLocation resid="Functions.Metadata.Url"/>
        </Metadata>
        <Namespace resid="Functions.Namespace"/>
        </ExtensionPoint>
    </AllFormFactors>

As mentioned under :ref:`pro/server/officejs_custom_functions:namespace`: if you want to set a default namespace for your functions, you'd do that with this line:

.. code-block:: xml

    <bt:String id="Functions.Namespace" DefaultValue="XLWINGS"/>

As usual, for the full context, have a look at ``manifest-xlwings-officejs-quickstart.xml`` in the quickstart sample.

Authentication
--------------

To authenticate (and possibly authorize) the users of your custom functions, you'll need to implement a global ``getAuth()`` function under ``app/taskpane.html``. In the quickstart project, it's set up to give back an empty string:

.. code-block:: js

    globalThis.getAuth = async function () {
      return ""
    };

The string that this function returns will be provided as Authorization header whenever a custom function executes so the backend can authenticate the user. Hence, to activate authentication, you'll need to change this function to give back the desired token/credentials.

.. note::

    The ``getAuth`` function is required for custom functions to work, even if you don't want to authenticate users, so don't delete it.

**SSO / Entra ID (previously called AzureAD) authentication**

The most convenient way to authenticate users is to use single-sign on (SSO) based on Entra ID (previously called Azure AD), which will use the identity of the signed-in Office user:

.. code-block:: js

    globalThis.getAuth = async function () {
      return await xlwings.getAccessToken();
    };

* This requires you to set up an Entra ID (previously called Azure AD) app as well as adjusting the manifest accordingly, see :ref:`pro/server/server_authentication:SSO/Entra ID (previously called Azure AD) for Office.js`.
* You'll also need to verify the AzureAD access token on the backend. This is already implemented in https://github.com/xlwings/xlwings-server

Deployment
----------

To deploy your custom functions, please refer to :ref:`pro/server/officejs_addins:production deployment` in the Office.js Add-ins docs.

Custom functions vs. legacy UDFs
--------------------------------

While Office.js-based custom functions are mostly compatible with the VBA-based UDFs, there are a few differences, which you should be aware of when switching from UDFs to custom functions or vice versa:

.. list-table::
    :header-rows: 1
  
    * -
      - Custom functions (Office.js-based)
      - User-defined functions UDFs (VBA-based)

    * - Supported platforms
      -  * Windows
         * macOS
         * Excel on the web
      - * Windows

    * - Empty cells are converted to
      - ``0`` => If you want ``None``, you have to set the following formula in Excel: ``=""``
      - ``None``

    * - Cells with integers are converted to
      - Integers
      - Floats

    * - Reading Date/Time-formatted cells
      - Requires the use of ``dt.datetime`` or ``parse_dates`` in the arg decorators
      - Automatic conversion

    * - Writing datetime objects
      - Automatic cell formatting
      - No cell formatting

    * - Can write proper Excel cell error
      - Yes
      - No

    * - Writing ``NaN`` (``np.nan`` or ``pd.NA``) arrives in Excel as
      - ``#NUM!``
      - Empty cell

    * - Functions are bound to
      - Add-in
      - Workbook

    * - Asynchronous functions
      - Always and automatically
      - Requires ``@xw.func(async_mode="threading")``
  
    * - Decorators
      - ``from xlwings import server``, then ``server.func`` etc.
      - ``import xlwings as xw``, then ``xw.func`` etc.

    * - Formula Intellisense
      - Yes
      - No

    * - Supports namespaces e.g., ``NAMESPACE.FUNCTION``
      - Yes
      - No

    * - Capitalization of function name
      - Excel formula gets automatically capitalized
      - Excel formula has same capitalization as Python function

    * - Supports (SSO) Authentication
      - Yes
      - No

    * - ``caller`` function argument
      - N/A
      - Returns Range object of calling cell

    * - ``@xw.arg(vba=...)``
      - N/A
      - Allows to access Excel VBA objects

    * - Supports pictures
      - No
      - Yes

    * - Requires a local installation of Python
      - No
      - Yes

    * - Python code must be shared with end-user
      - No
      - Yes

    * - Requires License Key
      - Yes
      - No

    * - License
      - PolyForm Noncommercial License 1.0.0 or xlwings PRO License
      - BSD 3-clause Open Source License

Limitations
-----------

* The Office.js Custom Functions API was introduced in 2018 and therefore requires at least Excel 2021 or Excel 365.
* Note that some functionality requires specific build versions, such as error cells and date formatting, but if your version of Excel doesn't support these features, xlwings will fall back to either string-formatted error messages or unformatted date serials. For more details on which builds support which function, see `Custom Functions requirement sets <https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/custom-functions-requirement-sets>`_.
* xlwings custom functions must be run with the shared runtime, i.e., the runtime that comes with a task pane add-in. The JavaScript-only runtime is not supported.

Roadmap
-------

* Streaming functions
* Object handlers
* Client-side caching
* Add support for Google Sheets
