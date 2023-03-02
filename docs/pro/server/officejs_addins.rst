.. _officejs_addins:

Office.js Add-ins
=================

This feature requires at least v0.29.0.

Office.js add-ins (officially called *Office add-ins*) are web apps that traditionally require you to use the `Excel JavaScript API <https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview>`_ by writing JavaScript or TypeScript code. Note that the Excel JavaScript API ("Office.js") is not to be confused with `Office Scripts <https://learn.microsoft.com/en-us/office/dev/scripts/overview/excel>`_, which is a layer on top of Office.js. While Office Scripts is much easier to use than Office.js, it only works for writing scripts that run via Excel's Automate tab and can't be used to create add-ins. This documentation will teach you how to build Office.js add-ins with xlwings Server, saving you from having to write Office.js code.

.. note::

  Office.js add-ins are just one option to talk to xlwings Server. The other options are VBA, Office Scripts, and Google Apps Script, see :ref:`xlwings Server documentation <remote_interpreter>`.

Why is this useful?
-------------------

Compared to using Office.js directly, using Office.js via xlwings Server has the following advantages:

* No need to learn JavaScript and the Excel JavaScript API. Instead, use the familiar xlwings syntax in Python.
* No need to install Node.js or use any JavaScript build tool such as Webpack.
* xlwings alerts saves you from having to use the `Office dialog API <https://learn.microsoft.com/en-us/office/dev/add-ins/develop/dialog-api-in-office-add-ins>`_ and from designing your own HTML for simple pop-ups.
* Error handling is built-in.
* Compatible with Excel 365 and the perpetual editions back to Excel 2016.

.. note::

  While xlwings will reduce the amount of JavaScript code to almost zero, you still have to use HTML and CSS if you want to use a task pane. However, task panes aren't mandatory as you can link your function directly to a Ribbon button, see :ref:`pro/server/officejs_addins:Commands`.

Introduction to Office.js add-ins 
---------------------------------

Office.js add-ins are web apps that can interact with Excel. In its simplest form, they consist of just two files:

* **Manifest XML file**: This is a configuration file that is loaded in Excel (either manually during development or via the add-in store for production). It defines the Ribbon buttons and includes the URL to the backend/web server.
* **HTML file**: The HTML file has to be served by a web server and defines either a visible *task pane* or *commands* that are directly linked to Ribbon buttons. There are more possibilities than just task panes and commands (see the `official documentation <https://learn.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins>`_), but we'll ignore them for the purpose of this introduction.

To get a better understanding about how the simplest possible add-in works (without Python or xlwings), have a look at the following repo: `<https://github.com/xlwings/officejs-helloworld>`_. Follow the repo's README to load the add-in in development mode, a process that is called *sideloading*.

Now that you know the basic structure of an Office.js add-in, let's see how we can replace the Excel JavaScript API with ``xlwings.runPython()`` calls!

Quickstart
----------

This quickstart shows you how you can call Python both from a button on the task pane and directly from a Ribbon button. xlwings can be used with any web framework and the quickstart repo therefore contains various implementations such as ``app/server_fastapi.py`` or ``app/server_starlette.py``. At the end of this quickstart, you'll have a working environment for local development.

1. **Download quickstart repo**: Use Git to clone the following repository: https://github.com/xlwings/xlwings-officejs-quickstart. If you don't want to use Git, you could also download the repo by clicking on the green ``Code`` button, followed by ``Download ZIP``, then unzipping it locally.
2. **Update manifest**: If you want to build your own add-in based off this quickstart repo, replace ``<Id>xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx</Id>`` in ``manifest-xlwings-officejs-quickstart.xml`` with a unique ID that you can create by visiting https://www.guidgen.com or by running the following command in Python: ``import uuid;print(uuid.uuid4())``.
3. **Create certificates**: Generate development certificates as the development server needs to be accessed via https instead of http (even on localhost). Otherwise, icons and alerts won't work and Excel on the web won't load the manifest at all. `Download mkcert <https://github.com/FiloSottile/mkcert/releases>`_ (pick the correct file according to your platform), rename the file to ``mkcert``, then run the following commands from a Terminal/Command Prompt (make sure you're in the same directory as ``mkcert``):

   .. code-block:: text

     $ ./mkcert -install
     $ ./mkcert localhost 127.0.0.1 ::1

   This will generate two files ``localhost+2.pem`` and ``localhost+2-key.pem``: move them to the root of the ``xlwings-officejs-quickstart`` quickstart repo.

4. **Install Python dependencies**: 
   
   * Local Python installation: create a virtual or Conda environment and install the Python dependencies by running: ``pip install -r requirements.txt``.
   * Docker: skip this step.
5. **xlwings license key**:

   Get a free `trial license key <https://www.xlwings.org/trial>`_ and install it as follows:

   * Local Python installation: ``xlwings license update -k your-license-key``
   * Docker: set the license key as ``XLWINGS_LICENSE_KEY`` environment variable. The easiest way to do this is to run ``cp .env.template .env`` in a Terminal/Command Prompt and fill in the license key in the ``.env`` file.
6. **Start web app**: 

   * Local Python installation: with the previously created virtual/Conda env activated, start the Python development server by running the Python file with the desired implementation. For example, to run the backend with FastAPI, run the following: ``python app/server_fastapi.py``. You could also run the file via the capabilities offered by your editor.
   * Docker: run ``docker compose up`` instead. Note that Docker by default uses the FastAPI implementation, so you'll need to edit ``docker-compose.yaml`` if you want to change that.
   
   If you see the following, the server is up and running:

   .. code-block:: text

      $ python app/server_fastapi.py 
      INFO:     Will watch for changes in these directories: ['/Users/fz/Dev/xlwings-officejs-quickstart']
      INFO:     Uvicorn running on https://127.0.0.1:8000 (Press CTRL+C to quit)
      INFO:     Started reloader process [56708] using WatchFiles
      INFO:     Started server process [56714]
      INFO:     Waiting for application startup.
      INFO:     Application startup complete.


7. **Sideload the add-in**: Manually load ``manifest-xlwings-officejs-quickstart.xml`` in Excel. This is called *sideloading* and the process differs depending on the platform you're using, see `Office.js docs <https://learn.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing>`_ for instructions. Once you've sideloaded the manifest, you'll see the ``Quickstart`` tab in the Ribbon.
8. **Time to play**: You're now ready to play around with the add-in in Excel and make changes to the source code under ``app/server_fastapi.py`` or under the respective file of your framework. Every time you edit and save the Python code, the development server will restart automatically so that you can instantly try out the code changes in Excel. If you make changes to the HTML file, you'll need to right-click on the task pane and select ``Reload``.

With a working development environment, let's see how everything works step-by-step. Let's start with looking at the Python backend server.

Backend
-------

The backend exposes your Python functions by using a Python web framework: you need to handle a POST request as shown in the following sample. Please have a look at the respective Python file in the ``app`` directory for the full context:

.. tab-set::
    .. tab-item:: FastAPI
      :sync: fastapi

      .. code-block::

          from fastapi import Body, FastAPI

          app = FastAPI()

          @app.post("/hello")
          async def hello(data: dict = Body):
              # Instantiate a Book object with the deserialized request body
              book = xw.Book(json=data)
          
              # Use xlwings as usual
              sheet = book.sheets[0]
              cell = sheet["A1"]
              if cell.value == "Hello xlwings!":
                  cell.value = "Bye xlwings!"
              else:
                  cell.value = "Hello xlwings!"
      
              # Pass the following back as the response
              return book.json()

    .. tab-item:: Flask
      :sync: flask

      .. code-block::

        from flask import Flask, jsonify, request

        app = Flask(__name__)

        @app.route("/hello", methods=["POST"])
        def hello():
            # Instantiate a Book object with the deserialized request body
            book = xw.Book(json=request.json)

            # Use xlwings as usual
            sheet = book.sheets[0]
            cell = sheet["A1"]
            if cell.value == "Hello xlwings!":
                cell.value = "Bye xlwings!"
            else:
                cell.value = "Hello xlwings!"

            # Pass the following back as the response
            return jsonify(book.json())

    .. tab-item:: Starlette
      :sync: starlette

      .. code-block::

            from starlette.applications import Starlette
            from starlette.responses import JSONResponse
            from starlette.routing import Route

            async def hello(request):
                # Instantiate a Book object with the deserialized request body
                data = await request.json()
                book = xw.Book(json=data)

                # Use xlwings as usual
                sheet = book.sheets[0]
                cell = sheet["A1"]
                if cell.value == "Hello xlwings!":
                    cell.value = "Bye xlwings!"
                else:
                    cell.value = "Hello xlwings!"

                # Pass the following back as the response
                return JSONResponse(book.json())

            routes = [
                Route("/hello", hello, methods=["POST"]),
            ]

            app = Starlette(debug=True, routes=routes)

Let's now move over to the frontend to learn how we can call these Python functions from the Office.js add-in!

Frontend
--------

In the following code snippet (an excerpt from ``app/taskpane.html``), the highlighted lines represent the relevant ones---the rest is just HTML boilerplate.

.. code-block:: html
   :emphasize-lines: 8-10, 14-15, 17-26
   :caption: app/taskpane.html (excerpt: only showing the 'Run hello' functionality)

    <!doctype html>
    <html lang="en">

    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>My Taskpane</title>
        <!-- ➊ Load office.js and xlwings.min.js -->
        <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
        <script type="text/javascript" src="https://cdn.jsdelivr.net/gh/xlwings/xlwings@0.29.0/xlwingsjs/dist/xlwings.min.js"></script>
    </head>

    <body>
        <!-- ➋ Put a button on the task pane -->
        <button id="run" type="button">Run hello</button>
        <script>
            // ➌ Initialize Office.js
            Office.onReady(function (info) { });

            // ➍ Add click event listeners to button
            document.getElementById("run").addEventListener("click", hello);

            // ❺ Use runPython with the desired endpoint of your web app
            function hello() {
                xlwings.runPython(window.location.origin + "/hello");
            }
        </script>
    </body>

    </html>

Let's see what's happening here by walking through the numbered sections!

➊ Load JavaScript libraries
~~~~~~~~~~~~~~~~~~~~~~~~~~~

Before anything else, we need to load ``office.js`` and ``xlwings.min.js`` in the ``head`` of the HTML file. While ``office.js`` is giving us access to the Excel JavaScript API, ``xlwings.min.js`` will make the ``runPython`` function available.

For ``xlwings.min.js``, make sure to adjust the version number after the ``@`` sign to match the version of the xlwings Python package you're using on the backend. In the quickstart repo, this would have to correspond to the version of xlwings defined in ``requirements.txt``.

While ``xlwings.min.js`` is not available via npm package manager at the moment, you could also download the file and its corresponding ``map`` file (by adding ``.map`` to the URL). Then refer to the file path of ``xlwings.min.js`` instead of using the URL of the CDN.

Note, however, that ``office.js`` requires you to use the CDN version in case you want to distribute the add-in publicly via the add-in store.

➋ Put a button on the task pane
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Putting a button on the task pane is a single line of HTML. Note the ``id`` that we will need under ➍ to attach a click event handler to it. To keep things as simple as possible, the button isn't styled in any way using CSS, so it will look spectacularly boring.

➌ Initialize Office.js
~~~~~~~~~~~~~~~~~~~~~~

In the body, as the first line in your ``script`` tag, you have to initialize Office.js.

Usually, this is all you need to worry about, but if you want to block your addin from running on certain versions of Excel, ``Office.onReady()`` is where you would handle this, see `the official docs <https://learn.microsoft.com/en-us/office/dev/add-ins/develop/initialize-add-in>`_.

➍ Add click event listeners
~~~~~~~~~~~~~~~~~~~~~~~~~~~

To define what should happen when you click the button, you need to attach an event listener to it. In our case, we're telling the event listener to call the ``hello`` function when the button with ``id=run`` is clicked.

❺ Use runPython
~~~~~~~~~~~~~~~

To call a function of your backend, you have to provide the ``xlwings.runPython()`` function the respective URL. Use ``window.location.origin + /myendpoint`` instead of hardcoding the full URL. This will ensure that everything still works when you change the URL e.g., when moving from development to production. Note that ``runPython`` accepts optional arguments, such as ``auth`` to send an Authorization header:

.. code-block:: js

    function hello() {
        xlwings.runPython(window.location.origin + "/hello", { auth: "mytoken" });
    }

* For more details on the optional ``runPython`` arguments, see :ref:`xlwings Server Config<xlwings_server_config>`.
* For more details on authentication, see :ref:`xlwings Server Auth<server_auth>`.

Task pane
---------

To have a Ribbon button show the task pane, you'll need to configure it properly in the manifest. The relevant blocks are the following (these lines are out of context, so search for them in ``manifest-xlwings-officejs-quickstart.xml``):

.. code-block:: xml

    <!-- ... -->

    <Control xsi:type="Button" id="TaskpaneButton">
      <!-- ... -->
      <!-- Action type must be ShowTaskpane -->
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>ButtonId1</TaskpaneId>
        <!-- resid must point to a Url Resource -->
        <SourceLocation resid="Taskpane.Url"/>
      </Action>
    </Control>

    <!-- ... -->

    <!-- This must point to the HTML document with the task pane -->
    <bt:Url id="Taskpane.Url" DefaultValue="https://127.0.0.1:8000/taskpane.html"/>

Commands
--------

.. note::

  Functions that you bind to a Ribbon button directly react a bit slower than a button on a task pane. This is because the task pane gets loaded once and stays loaded, whereas clicking a button on the Ribbon loads everything from scratch every time you click the button.

To understand how you can call ``xlwings.runPython()`` directly from a Ribbon button, have a look at ``app/commands.html`` in the quickstart repo. Its body reads as follows:

.. code-block:: html

  <body>
      <script>
          // Initialize Office.js
          Office.onReady(function (info) { });
  
          // Make sure to provide the event argument and call 
          // event.completed() at the end of functions that 
          // are directly associated with Ribbon buttons
          function hello(event) {
              xlwings.runPython(window.location.origin + "/hello");
              event.completed();
          }
          // You must associate the FunctionName from manifest.xml ("run")
          // with the JavaScript function name (hello)
          Office.actions.associate("run", hello);
      </script>
  </body>

The relevant blocks in the manifest are the following (again, these lines are out of context, so search for them in ``manifest-xlwings-officejs-quickstart.xml``). Note that compared to task panes, you need the additional reference to ``FunctionFile``:

.. code-block:: xml

    <!-- ... -->

    <!-- resid must point to a Url Resource -->
    <FunctionFile resid="Commands.Url"/>

    <!-- ... -->

    <Control xsi:type="Button" id="MyFunctionButton">
      <!-- ... -->
      <!-- Action type must be ExecuteFunction -->
      <Action xsi:type="ExecuteFunction">
        <!-- This is the name that you use in Office.actions.associate()
            to connect it to a function -->
        <FunctionName>run</FunctionName>
      </Action>
    </Control>

    <!-- ... -->

    <!-- This must point to the HTML document with the function -->
    <bt:Url id="Commands.Url" DefaultValue="https://127.0.0.1:8000/commands.html"/>

    <!-- ... -->

Having seen how you can call Python from task panes and Ribbon buttons, let's move on with alerts!

Alerts
------

Alerts require a bit of boilerplate on the Python side. Because alerts are used for unhandled exceptions, you should implement the boilerplate code even if you don't use alerts in your own code.

Alerts boilerplate
~~~~~~~~~~~~~~~~~~

The boilerplate consists of:

* Implementing the ``/xlwings/alert`` endpoint
* Giving your templating engine access to the ``xlwings-alert.html`` template, which is included in the xlwings Python package under ``xlwings.html``

Here is the relevant code. As usual, have a look at ``app/server_fastapi.py`` for the full context.

.. tab-set::
    .. tab-item:: FastAPI + Jinja2
      :sync: fastapi

      .. code-block:: python
  
          import jinja2
          import markupsafe  # This is a dependency of Jinja2
          from fastapi import Request
          from fastapi.responses import HTMLResponse
          from fastapi.templating import Jinja2Templates
      
          @app.get("/xlwings/alert", response_class=HTMLResponse)
          async def alert(
              request: Request, prompt: str, title: str, buttons: str, mode: str, callback: str
          ):
              """This endpoint is required by myapp.alert() and to show unhandled exceptions"""
              return templates.TemplateResponse(
                  "xlwings-alert.html",
                  {
                      "request": request,
                      "prompt": markupsafe.Markup(prompt.replace("\n", "<br>")),
                      "title": title,
                      "buttons": buttons,
                      "mode": mode,
                      "callback": callback,
                  },
              )

          # Add the xlwings alert template as source by making use of an additional template loader
          loader = jinja2.ChoiceLoader(
              [
                  jinja2.FileSystemLoader("mytemplates"),  # this is your default templates folder
                  jinja2.PackageLoader("xlwings", "html"),
              ]
          )
          templates = Jinja2Templates(directory="mytemplates", loader=loader)

    .. tab-item:: Starlette + Jinja2
      :sync: starlette

      .. code-block:: python

        import jinja2
        import markupsafe  # This is a dependency of Jinja2
        from starlette.templating import Jinja2Templates

        async def alert(request):
            """Boilerplate required by book.app.alert() and to show unhandled exceptions"""
            params = request.query_params
            return templates.TemplateResponse(
                "xlwings-alert.html",
                {
                    "request": request,
                    "prompt": markupsafe.Markup(params["prompt"].replace("\n", "<br>")),
                    "title": params["title"],
                    "buttons": params["buttons"],
                    "mode": params["mode"],
                    "callback": params["callback"],
                },
            )

        # Add xlwings.html as additional source for templates so the /xlwings/alert endpoint
        # will find xlwings-alert.html. "mytemplates" can be a dummy if the app doesn't use
        # own templates
        loader = jinja2.ChoiceLoader(
            [
                jinja2.FileSystemLoader("mytemplates"),
                jinja2.PackageLoader("xlwings", "html"),
            ]
        )
        templates = Jinja2Templates(directory="mytemplates", loader=loader)

        routes = [
            Route("/xlwings/alert", alert),
        ]


With the boilerplate in place, you're now ready to use alerts, as we'll see next.

Showing alerts
~~~~~~~~~~~~~~

.. note::

  Except in Excel on the web, alerts are non-modal, i.e., allow the user to continue using Excel while the alert is open. This is a limitation of Office.js.

Calling an alert with an ``OK`` button is as simple as:

.. code-block:: python

    # book is an xlwings Book object
    book.app.alert(
        "Some text",
        title="Some Title",  # optional
    )

Clicking either the "x" at the top right or the OK button will close the alert and you're done with it.

However, if you need to react differently depending on whether the user clicks on OK or Cancel, you can supply a ``callback`` argument that accepts the name of a JavaScript function. To understand how this works, consider the following example:

.. code-block:: python

    book.app.alert(
        prompt="This will capitalize all sheet names!",
        title="Are you sure?",
        buttons="ok_cancel",
        callback="capitalizeSheetNames",
    )

When the user clicks a button, it will call the JavaScript function ``capitalizeSheetNames`` with the name of the clicked button as argument in lower case. For example, if the user clicks on ``Cancel``, it would call ``capitalizeSheetNames("cancel")``. Depending on the answer, you can run another ``xlwings.runPython()`` call or do something directly in JavaScript. To make this work, we'll need to add our callback function to the script tag in the body of our HTML file. You'll also need to register that function using the ``xlwings.registerCallback`` function:


.. code-block:: js

    function capitalizeSheetNames(arg) {
        if (arg == "ok") {
            xlwings.runPython(window.location.origin + "/capitalize-sheet-names");
        } else {
            // Cancel
        }
    }
    // Make sure to register the callback function
    xlwings.registerCallback(capitalizeSheetNames);

As usual, to get a better understanding, check out ``app/taskpane.html`` and ``app/server_fastapi.py`` for the full context and play around with the respective button on the task pane.

Debugging
---------

If you need to debug errors on the client side, you'll need to open the developer tools of the browser that's being used so you can inspect the error messages in the console. Depending on the platform and version of Excel, the process is different:

* Excel on the web: open the developer tools of the browser you're using. For example, in Chrome you can type ``Ctrl+Shift+I`` (Windows) or ``Cmd-Option-I`` (macOS), then switch to the Console tab.
* Desktop Excel on Windows: right-click on the task pane and select ``Inspect``, then switch to the Console tab.
* Desktop Excel on macOS: to be able to get the Web Inspector showing up, you'll need to run the following command in a Terminal once:: 
    
    defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true
    
  Then, after restarting Excel, right-click on the task pane and select ``Inspect Element`` and switch to the Console tab. Note that after running this command, you'll also see an empty page loaded when you call a command from the Ribbon button directly. To hide it, you would need to disable debugging again by running the same command in the Terminal with ``false`` instead of ``true``.

Production deployment
---------------------

* Make sure that the ``Id`` in the manifest is your own unique UUID.
* Make sure you have authentication implemented.
* The Python backend can be deployed anywhere you like, there are some suggestions under :ref:`xlwings Server production deployment <server_production>`.
* Once you have your backend deployed, you'll need to replace ``https://127.0.0.1:8000`` with your production URL. You may want to keep multiple copies of the manifest, one for local development and one for each environment like production.
* Depending on whether you want to deploy your add-in within your company or to the whole world, there's a different process for deploying the manifest XML:
    * **Company-internal** (must be done by a Microsoft 365 admin): on office.com, click on Admin > Show all > Settings > Integrated Apps > Add-ins. There, click on the ``Deploy Add-in`` button which allows you to upload the manifest or point to it via URL.
    * **Public**: you'll need to submit your add-in for approval to Microsoft AppSource, see: https://learn.microsoft.com/en-us/azure/marketplace/submit-to-appsource-via-partner-center


Workaround for missing features
-------------------------------

In the classic version of xlwings, you can use the ``.api`` property to fall back to the underlying automation library and work around :ref:`missing features <missing_features>` in xlwings. That's not possible with xlwings Server.

Instead, call the ``book.app.macro()`` method to run functions in JavaScript. The first parameter will have to be the request context, which gives you access to the Excel JavaScript API. Note that you have to register JavaScript functions that you want to call from Python via ``xlwings.registerCallback()`` (last line):

.. code-block:: js

    async function wrapText(context, sheetName, cellAddress) {
      // The first parameter has to be the request context, the others 
      // are those parameters that you will provide via Python
      const range = context.workbook.worksheets
        .getItem(sheetName)
        .getRange(cellAddress);
      range.format.wrapText = true;
      await context.sync();
    }
    // Make sure to register the function as callback
    xlwings.registerCallback(wrapText);

Now you can call this function from Python like so:

.. code-block:: Python

    # book is an xlwings Book object
    wrap_text = book.app.macro("wrapText")
    wrap_text("Sheet1", "A1")
    wrap_text("Sheet2", "B2")

Limitations
-----------

* Currently, only a subset of the xlwings API is covered, mainly the Range and Sheet classes with a focus on reading and writing values. This, however, includes full support for type conversion including pandas DataFrames, NumPy arrays, datetime objects, etc.
* Excel 2016 and 2019 won't support automatic Date conversion when reading from Excel to Python. It works properly though on Excel 2021 and Excel 365.
* You are moving within the web's request/response cycle, meaning that values that you write to a range will only be written back to Google Sheets/Excel once the function call returns. Put differently, you'll get the state of the sheets at the moment the call was initiated, but you can't read from a cell you've just written to until the next call.
* You will need to use the same xlwings version for the Python package and the JavaScript module, otherwise, the server will raise an error.
* Currently, custom functions (a.k.a. user-defined functions or UDFs) are not supported.
