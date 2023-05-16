.. _remote_interpreter:

xlwings Server: VBA, Office Scripts, Google Apps Script
=======================================================

This feature requires xlwings PRO and at least v0.27.0.

Instead of installing Python on each end-user's machine, you can work with a server-based Python installation. It's essentially a web application, but uses your spreadsheet as the frontend instead of a web page in a browser. xlwings Server doesn't just work with the Desktop versions of Excel on Windows and macOS but additionally supports Google Sheets and Excel on the web for a full cloud experience. xlwings Server runs everywhere where Python runs, including Linux, Docker and WSL (Windows Subsystem for Linux). it can run on your local machine, as a (serverless) cloud service, or on an on-premise server.

.. important:: This feature currently only covers parts of the RunPython API. See also :ref:`pro/server/server:Limitations` and :ref:`pro/server/server:Roadmap`.

Why is this useful?
-------------------

Having to install a local installation of Python with the correct dependencies is the number one friction when using xlwings. Most excitingly though, xlwings Server adds support for the web-based spreadsheets: Google Sheets and Excel on the web.

To automate Office on the web, you have to use Office Scripts (i.e., TypeScript, a typed superset of JavaScript) and for Google Sheets, you have to use Apps Script (i.e., JavaScript). If you don't feel like learning JavaScript, xlwings allows you to write Python code instead. But even if you are comfortable with JavaScript, you are very limited in what you can do, as both Office Scripts and Apps Script are primarily designed to automate simple spreadsheet tasks such as inserting a new sheet or formatting cells rather than performing data-intensive tasks. They also make it very hard/impossible to use external JavaScript libraries and run in environments with minimal resources.

.. note:: From here on, when I refer to the **xlwings JavaScript module**, I mean either the xlwings Apps Script module if you use Google Sheets or the xlwings Office Scripts module if you use Excel on the web.

On the other hand, xlwings Server brings you these advantages:

* **Work with the whole Python ecosystem**: including pandas, machine learning libraries, database packages, web scraping, boto (for AWS S3), etc. This makes xlwings a great alternative for Power Query, which isn't currently available for Excel on the web or Google Sheets.
* **Leverage your existing development workflow**: use your favorite IDE/editor (local or cloud-based) with full Git support, allowing you to easily track changes, collaborate and perform code reviews. You can also write unit tests using pytest.
* **Remain in control of your data and code**: except for the data you expose in Excel or Google Sheets, everything stays on your server. This can include database passwords and other sensitive info such as customer data. There's also no need to give the Python code to end-users: the whole business logic with your secret sauce is protected on your own infrastructure.
* **Choose the right machine for the job**: whether that means using a GPU, a ton of CPU cores, lots of memory, or a gigantic hard disc. As long as Python runs on it, you can go from serverless functions as offered by the big cloud vendors all the way to a self-managed Kubernetes cluster under your desk (see :ref:`pro/server/server:Production Deployment`).
* **Headache-free deployment and maintenance**: there's only one location (usually a Linux server) where your Python code lives and you can automate the whole deployment process with continuous integration pipelines like GitHub actions etc.
* **Cross-platform**: xlwings Server works with Google Sheets, Excel on the web and the Desktop apps of Excel on Windows and macOS.

Prerequisites
-------------

.. tab-set::
    .. tab-item:: Excel (VBA)
      :sync: vba

      * At least xlwings 0.27.0
      * Either the xlwings add-in installed or a workbook that has been set up in standalone mode

    .. tab-item:: Excel (Office Scripts)
      :sync: officescripts

      * At least xlwings 0.27.0
      * You need the ``Automate`` tab enabled in order to access Office Scripts. Note that Office Scripts currently requires OneDrive for Business or SharePoint (it's not available on the free office.com), see also `Office Scripts Requirements <https://docs.microsoft.com/en-gb/office/dev/scripts/overview/excel#requirements>`_.
      * The ``fetch`` command in Office Scripts must **not** be disabled by your Microsoft 365 administrator.
      * Note that Office Scripts is available for Excel on the web and more recently also for Desktop Excel if you use Microsoft 365 (macOS and Windows), you may need to be on the beta channel though.

    .. tab-item:: Google Sheets
      :sync: google

      * At least xlwings 0.27.0
      * New sheets: no special requirements.
      * Older sheets: make sure that Chrome V8 runtime is enabled under ``Extensions`` > ``Apps Script`` > ``Project Settings`` > ``Enable Chrome V8 runtime``.

Introduction
------------

xlwings Server consists of two parts:

* Backend: the Python part
* Frontend: the xlwings JavaScript module (for Google Sheets/Excel via Office Scripts) or the VBA code in the form of the add-in or standalone modules (Desktop Excel via VBA)

The backend exposes your Python functions by using a Python web framework. In more detail, you need to handle a POST request along these lines (note that you can use any web framework, these are just examples of some of the most popular ones):

.. tab-set::

    .. tab-item:: FastAPI
      :sync: fastapi

      .. code-block:: python

          @app.post("/hello")
          async def hello(data: dict = Body):
              # Instantiate a Book object with the deserialized request body
              with xw.Book(json=data) as book:

                  # Use xlwings as usual
                  sheet = book.sheets[0]
                  sheet["A1"].value = "Hello xlwings!"

                  # Return a JSON response
                  return book.json()

    .. tab-item:: Flask
      :sync: flask

      .. code-block:: python

          @app.route("/hello", methods=["POST"])
          def hello():
              # Instantiate a Book object with the deserialized request body
              with xw.Book(json=request.json) as book:

                  # Use xlwings as usual
                  sheet = book.sheets[0]
                  sheet["A1"].value = "Hello xlwings!"

                  # Return a JSON response
                  return book.json()

    .. tab-item:: Django
      :sync: django

      .. code-block:: python

          def hello(request):
              # Instantiate a book object with the parsed request body
              data = json.loads(request.body.decode("utf-8"))
              with xw.Book(json=data) as book:

                  # Use xlwings as usual
                  sheet = book.sheets[0]
                  sheet["A1"].value = "Hello xlwings!"

                  # Return a JSON response
                  return JsonResponse(book.json())

    .. tab-item:: Starlette
      :sync: starlette

      .. code-block:: python

          async def hello(request):
              # Instantiate a Book object with the deserialized request body
              data = await request.json()
              with xw.Book(json=data) as book:

                  # Use xlwings as usual
                  sheet = book.sheets[0]
                  sheet["A1"].value = "Hello xlwings!"

                  # Return a JSON response
                  return JSONResponse(book.json())

.. caution:: To prevent a memory leak, it is important to close the book at the end of the request either by invoking ``book.close()`` or, as shown in the example, by using ``book`` as a context manager via the ``with`` statement.

* For Desktop Excel, you can run the web server locally and call the respective function
    * from VBA (requires the add-in installed or a workbook in standalone mode) or
    * from Office Scripts
* For the cloud-based spreadsheets, you have to run this on a web server that can be reached from Google Sheets or Excel on the web, and you have to paste the xlwings JavaScript module into the respective editor. How this all works, will be shown in detail under :ref:`pro/server/server:Cloud-based development with Gitpod`.

The next section shows you how you can play around with the xlwings Server on your local desktop before we'll dive into developing against the cloud-based spreadsheets.

Local Development with Desktop Excel
------------------------------------

The easiest way to try things out is to run the web server locally against your Desktop version of Excel. We're going to use `FastAPI <https://fastapi.tiangolo.com/>`_ as our web framework. While you can use any web framework you like, no quickstart command exists for these yet. However, for Flask, you can find the respective project on GitHub: https://github.com/xlwings/xlwings-server-helloworld-flask

Start by running the following command on a Terminal/Command Prompt. Feel free to replace ``demo`` with another project name and make sure to run this command in the desired directory::

    $ xlwings quickstart demo --fastapi

This creates a folder called ``demo`` in the current directory with the following files::

    demo.xlsm
    main.py
    requirements.txt

I would recommend you to create a virtual or Conda environment where you install the dependencies via ``pip install -r requirements.txt``. To run this server locally, run ``python main.py`` in your Terminal/Command Prompt or use your code editor/IDE's run button. You should see something along these lines:

.. code-block:: text

    $ python main.py
    INFO:     Will watch for changes in these directories: ['/Users/fz/Dev/demo']
    INFO:     Uvicorn running on http://127.0.0.1:8000 (Press CTRL+C to quit)
    INFO:     Started reloader process [36073] using WatchFiles
    INFO:     Started server process [36075]
    INFO:     Waiting for application startup.
    INFO:     Application startup complete.

Your web server is now listening, so let's open ``demo.xlsm``.

If you want to use VBA, press ``Alt+F11`` to open the VBA editor, and in ``Module1``, place your cursor somewhere inside the following function:

.. code-block:: vb.net

    Sub SampleRemoteCall()
        RunRemotePython "http://127.0.0.1:8000/hello"
    End Sub

Then hit ``F5`` to run the function---you should see ``Hello xlwings!`` in cell A1 of the first sheet.

If, however, you want to use Office Scripts, you can start from an empty file (it can be ``xlsx``, it doesn't have to be ``xlsm``), and run ``xlwings copy os`` on the Terminal/Command Prompt/Anaconda Prompt. Then add a new Office Script and paste the code from the clipboard before clicking on ``Run``.

To move this to production, you need to deploy the backend to a server, set up authentication, and point the URL to the production server, see :ref:`pro/server/server:Production Deployment`.

The next sections, however, show you how you can make this work with Google Sheets and Excel on the web.

Cloud-based development with Gitpod
-----------------------------------

Using Gitpod is the easiest solution if you'd like to develop against either Google Sheets or Excel on the web.

If you want to have a development environment up and running in less than 5 minutes (even if you're new to web development), simply click the ``Open in Gitpod`` button to open a `sample project <https://github.com/xlwings/xlwings-web-fastapi>`_ in `Gitpod <https://www.gitpod.io>`_ (Gitpod is a cloud-based development environment with a generous free tier):

.. image:: https://gitpod.io/button/open-in-gitpod.svg
   :target: https://gitpod.io/#https://github.com/xlwings/xlwings-server-helloworld-fastapi
   :alt: Open in Gitpod

Opening the project in Gitpod will require you to sign in with your GitHub account. A few moments later, you should see an online version of VS Code. In the Terminal, it will ask you to paste the xlwings license key (`get a free trial key <https://www.xlwings.org/trial>`_ if you want to try this out in a commercial context or use the ``noncommercial`` license key if your usage `qualifies as noncommercial <https://polyformproject.org/licenses/noncommercial/1.0.0>`_). Note that your browser will ask you for permission to paste. Once you confirm your license key by hitting ``Enter``, the server will automatically start with everything properly configured. You can then open the ``app`` directory and look at the ``main.py`` file, where you'll see the ``hello`` function. This is the function we're going to call from Google Sheets/Excel on the web in just a moment. Let's now look at the ``js`` folder and open the file according to your platform:

.. tab-set::
    .. tab-item:: Excel (Office Scripts)
      :sync: officescripts

      .. code-block:: text

          xlwings_excel.ts

    .. tab-item:: Google Sheets
      :sync: google

      .. code-block:: text

          xlwings_google.js

Copy all the code, then switch to Google Sheets or Excel, respectively, and continue as follows:

.. tab-set::
    .. tab-item:: Excel (Office Scripts)
      :sync: officescripts

      In the ``Automate`` tab, click on ``New Script``. This opens a code editor pane on the right-hand side with a function stub. Replace this function stub with the copied code from ``xlwings_excel.ts``. Make sure to click on ``Save script`` before clicking on ``Run``: the script will run the ``hello`` function and write ``Hello xlwings!`` into cell ``A1``.

      To run this script from a button, click on the 3 dots in the Office Scripts pane (above the script), then select ``+ Add button``.

    .. tab-item:: Google Sheets
      :sync: google

      Click on ``Extensions`` > ``Apps Script``. This will open a separate browser tab and open a file called ``Code.gs`` with a function stub. Replace this function stub with the copied code from ``xlwings_google.js`` and click on the ``Save`` icon. Then hit the ``Run`` button (the ``hello`` function should be automatically selected in the dropdown to the right of it). If you run this the very first time, Google Sheets will ask you for the permissions it needs. Once approved, the script will run the ``hello`` function and write ``Hello xlwings!`` into cell ``A1``.

      To add a button to a sheet to run this function, switch from the Apps Script editor back to Google Sheets, click on ``Insert`` > ``Drawing`` and draw a rounded rectangle. After hitting ``Save and Close``, the rectangle will appear on the sheet. Select it so that you can click on the 3 dots on the top right of the shape. Select ``Assign Script`` and write ``hello`` in the text box, then hit ``OK``.


Any changes you make to the ``hello`` function in ``app/main.py`` in Gitpod are automatically saved and reloaded by the web server and will be reflected the next time you run the script from Google Sheets or Excel on the web.

.. note:: While Excel on the web requires you to create a separate script with a function called ``main`` for each Python function, Google Sheets allows you to add multiple functions with any name.

Please note that clicking the Gitpod button gets you up and running quickly, but if you want to save your changes (i.e., commit them to Git), you should first fork the project on GitHub to your own account and open it by prepending ``https://gitpod.io/#`` to your GitHub URL instead of clicking the button (this works with GitLab and Bitbucket too). Or continue with the next section, which shows you how you can start a project from scratch on your local machine.

An alternative for Gitpod is `GitHub Codespaces <https://github.com/features/codespaces>`_, but unlike Gitpod, GitHub Codespaces only works with GitHub.

Local Development with Google Sheets or Excel via Office Scripts
----------------------------------------------------------------

This section walks you through a local development workflow as an alternative to using Gitpod/GitHub Codespaces. What's making this a little harder than using a preconfigured online IDE like Gitpod is the fact that we need to expose our local web server to the internet for easy development (even if we use the Desktop version of Excel).

As before, we're going to use `FastAPI <https://fastapi.tiangolo.com/>`_ as our web framework. While you can use any web framework you like, no quickstart command exists for these yet, so you'd have to set up the boilerplate yourself. Let's start with the server before turning our attention to the client side (i.e, Google Sheets or Excel on the web).

Part I: Backend
***************

Start a new quickstart project by running the following command on a Terminal/Command Prompt. Feel free to replace ``demo`` with another project name and make sure to run this command in the desired directory::

    $ xlwings quickstart demo --fastapi

This creates a folder called ``demo`` in the current directory with a few files::

    main.py
    demo.xlsm
    requirements.txt

I would recommend you to create a virtual or Conda environment where you install the dependencies via ``pip install -r requirements.txt``. In ``app.py``, you'll find the FastAPI boilerplate code and in ``main.py``, you'll find the ``hello`` function that is exposed under the ``/hello`` endpoint.

To run this server locally, run ``python main.py`` in your Terminal/Command Prompt or use your code editor/IDE's run button. You should see something along these lines:

.. code-block:: text

    $ python main.py
    INFO:     Will watch for changes in these directories: ['/Users/fz/Dev/demo']
    INFO:     Uvicorn running on http://127.0.0.1:8000 (Press CTRL+C to quit)
    INFO:     Started reloader process [36073] using watchgod
    INFO:     Started server process [36075]
    INFO:     Waiting for application startup.
    INFO:     Application startup complete.

Your web server is now listening, however, to enable it to communicate with Google Sheets or Excel via Office Scripts, you need to expose the port used by your local server (port 8000 in your example) securely to the internet. There are many free and paid services available to help you do this. One of the more popular ones is `ngrok <https://ngrok.com/>`_ whose free version will do the trick (for a list of ngrok alternatives, see `Awesome Tunneling <https://github.com/anderspitman/awesome-tunneling>`_):

* `ngrok Installation <https://ngrok.com/download>`_
* `ngrok Tutorial <https://ngrok.com/docs>`_

For the sake of this tutorial, let's assume you've installed ngrok, in which case you would run the following on your Terminal/Command Prompt to expose your local server to the public internet::

    $ ngrok http 8000

Note that the number of the port (8000) has to correspond to the port that is configured on your local development server as specified at the bottom of ``main.py``. ngrok will print something along these lines::

    ngrok by @inconshreveable                                                                                (Ctrl+C to quit)

    Session Status                online
    Account                       name@domain.com (Plan: Free)
    Version                       2.3.40
    Region                        United States (us)
    Web Interface                 http://127.0.0.1:4040
    Forwarding                    http://xxxx-xxxx-xx-xx-xxx-xxxx-xxxx-xxxx-xxx.ngrok.io -> http://localhost:8000
    Forwarding                    https://xxxx-xxxx-xx-xx-xxx-xxxx-xxxx-xxxx-xxx.ngrok.io -> http://localhost:8000

To configure the xlwings client in the next step, we'll need the ``https`` version of the Forwarding address that ngrok prints, i.e., ``https://xxxx-xxxx-xx-xx-xxx-xxxx-xxxx-xxxx-xxx.ngrok.io``.

.. note:: When you're not actively developing, you should stop your ngrok session by hitting ``Ctrl-C`` in the Terminal/Command Prompt.

Part II: Frontend
*****************

Now it's time to switch to Google Sheets or Excel! To paste the xlwings JavaScript module, follow these 3 steps:

1. **Copy the xlwings JavaScript module**: On a Terminal/Command Prompt on your local machine, run the following command:

   .. tab-set::
       .. tab-item:: Excel (Office Scripts)
         :sync: officescripts

         .. code-block:: text

             $ xlwings copy os

       .. tab-item:: Google Sheets
         :sync: google

         .. code-block:: text

             $ xlwings copy gs

   This will copy the correct xlwings JavaScript module to the clipboard so we can paste it in the next step.

2. **Paste the xlwings JavaScript module**

.. tab-set::
    .. tab-item:: Excel (Office Scripts)
      :sync: officescripts

      In the ``Automate`` tab, click on ``New Script``. This opens a code editor pane on the right-hand side with a function stub. Replace this function stub with the copied code from the previous step. Make sure to click on ``Save script`` before clicking on ``Run``: the script will run the ``hello`` function and write ``Hello xlwings!`` into cell ``A1``.

      To run this script from a button, click on the 3 dots in the Office Scripts pane (above the script), then select ``+ Add button``.

    .. tab-item:: Google Sheets
      :sync: google

      Click on ``Extensions`` > ``Apps Script``. This will open a separate browser tab and open a file called ``Code.gs`` with a function stub. Replace this function stub with the copied code from the previous step and click on the ``Save`` icon. Then hit the ``Run`` button (the ``hello`` function should be automatically selected in the dropdown to the right of it). If you run this the very first time, Google Sheets will ask you for the permissions it needs. Once approved, the script will run the ``hello`` function and write ``Hello xlwings!`` into cell ``A1``.

      To add a button to a sheet to run this function, switch from the Apps Script editor back to Google Sheets, click on ``Insert`` > ``Drawing`` and draw a rounded rectangle. After hitting ``Save and Close``, the rectangle will appear on the sheet. Select it so that you can click on the 3 dots on the top right of the shape. Select ``Assign Script`` and write ``hello`` in the text box, then hit ``OK``.

3. **Configuration**: The final step is to configure the xlwings JavaScript module properly, see the next section :ref:`pro/server/server:Configuration`.

.. _xlwings_server_config:

Configuration
-------------

xlwings can be configured in two ways:

* Via arguments in the ``runPython`` (via Apps Script / Office Scripts) or ``RunRemotePython`` (via VBA) function, respectively.
* Via ``xlwings.conf`` sheet (in this case, the keys are UPPER_CASE with underscore instead of camelCase, see the screenshot below).

If you provide a value via config sheet and via function argument, the function argument wins. Let's see what the available settings are:

* ``url`` (required): This is the full URL of your function. In the above example under :ref:`pro/server/server:Local Development with Google Sheets or Excel via Office Scripts`, this would be ``https://xxxx-xxxx-xx-xx-xxx-xxxx-xxxx-xxxx-xxx.ngrok.io/hello``, i.e., the ngrok URL **with the /hello endpoint appended**.
* ``auth`` (optional): This is a shortcut to set the ``Authorization`` header. See the section about :ref:`Server Auth <server_auth>` for the options.
* ``headers`` (optional): A dictionary (VBA) or object literal (JS) with name/value pairs. If you set the ``Authorization`` header, the ``auth`` argument will be ignored.
* ``exclude`` (optional): By default, xlwings sends over the complete content of the whole workbook to the server. If you have sheets with big amounts of data, this can make the calls slow or you could even hit a timeout. If your backend doesn't need the content of certain sheets, the ``exclude`` option will block the sheet's content (e.g., values, pictures, etc.) from being sent to the backend. Currently, you can only exclude entire sheets as comma-delimited string like so: ``"Sheet1, Sheet2"``.
* ``include`` (optional): It's the counterpart to ``exclude`` and allows you to submit the names of the sheets whose content (e.g., values, pictures, etc.) you want to send to the server. Like ``exclude``, ``include`` accepts a comma-delimited string, e.g., ``"Sheet1,Sheet2"``.
* ``timeout`` (optional, VBA client only): By default, the VBA client has a timeout of 30s, you can change it by providing the timeout in milliseconds, so if you want to increase it to 40s, provide the argument as ``timeout:=40000``.

Configuration Examples: Function Arguments
******************************************

.. tab-set::

    .. tab-item:: Excel (VBA)
      :sync: vba

      No arguments:

      .. code-block:: vb.net

        Sub Hello()
            RunRemotePython "http://127.0.0.1:8000/hello"
        End Sub

      Additionally providing the ``auth`` and ``exclude`` parameters as well as including a custom header:

      .. code-block:: vb.net

        Sub Hello()
            Dim headers As New Dictionary
            headers.Add "MyHeader", "my-value"
            RunRemotePython "http://127.0.0.1:8000/hello", auth:="xxxxxxxxxxxx", exclude:="xlwings.conf, Sheet1", headers:=headers
        End Sub

    .. tab-item:: Excel (Office Scripts)
      :sync: officescripts

      No arguments:

      .. code-block:: JavaScript

        async function main(workbook: ExcelScript.Workbook) {
          await runPython(
            workbook,
            "https://xxxx-xxxx-xx-xx-xxx-xxxx-xxxx-xxxx-xxx.ngrok.io/hello",
          );
        }

      Additionally providing the ``auth`` and ``exclude`` parameters as well as a custom header:

      .. code-block:: JavaScript

        async function main(workbook: ExcelScript.Workbook) {
          await runPython(
            workbook,
            "https://xxxx-xxxx-xx-xx-xxx-xxxx-xxxx-xxxx-xxx.ngrok.io/hello",
            {
              auth: "xxxxxxxxxxxx",
              exclude: "xlwings.conf, Sheet1",
              headers: { MyHeader: "my-value" },
            }
          );
        }

    .. tab-item:: Google Sheets
      :sync: google

      No arguments:

      .. code-block:: JavaScript

        function hello() {
          runPython("https://xxxx-xxxx-xx-xx-xxx-xxxx-xxxx-xxxx-xxx.ngrok.io/hello");
        }

      Additionally providing the ``auth`` and ``exclude`` parameters as well as a custom header:

      .. code-block:: JavaScript

        function hello() {
          runPython("https://xxxx-xxxx-xx-xx-xxx-xxxx-xxxx-xxxx-xxx.ngrok.io/hello", {
            auth: "xxxxxxxxxxxx",
            exclude: "xlwings.conf, Sheet1",
            headers: { MyHeader: "my-value" },
          });
        }

Configuration Examples: xlwings.conf sheet
******************************************

Create a sheet called ``xlwings.conf`` and fill in key/value pairs like so::

      | A       | B                  |
      --------------------------------
    1 | AUTH    | xxxxxxxxxxxx       |
    2 | EXCLUDE | Sheet1,xlwings.conf|


.. _server_production:

Production Deployment
---------------------

The xlwings web server can be built with any web framework and can therefore be deployed using any solution capable of running a Python backend or function. Here is a list for inspiration (non-exhaustive):

* **Fully-managed services**: `Heroku <https://www.heroku.com>`_, `Render <https://www.render.com>`_, `Fly.io <https://www.fly.io>`_, etc.
* **Interactive environments**: `PythonAnywhere <https://www.pythonanywhere.com>`_, `Anvil <https://www.anvil.works>`_, etc.
* **Serverless functions**: `AWS Lambda <https://aws.amazon.com/lambda/>`_, `Azure Functions <https://azure.microsoft.com/en-us/services/functions/>`_, `Google Cloud Functions <https://cloud.google.com/functions>`_, `Vercel <https://vercel.com>`_, etc.
* **Virtual Machines**: `DigitalOcean <https://digitalocean.com>`_, `vultr <https://www.vultr.com>`_, `Linode <https://www.linode.com/>`_, `AWS EC2 <https://aws.amazon.com/ec2/>`_, `Microsoft Azure VM <https://azure.microsoft.com/en-us/services/virtual-machines/>`_, `Google Cloud Compute Engine <https://cloud.google.com/compute>`_, etc.
* **Corporate servers**: Anything will work (including Kubernetes) as long as the respective endpoints can be accessed from your spreadsheet app.

Serverless Functions
********************

For examples how to configure the serverless function platform with xlwings see the following example repositories.

* `DigitalOcean Functions xlwings server <https://github.com/xlwings/xlwings-server-digitaloceanfunctions>`_
* `Azure Functions xlwings server <https://github.com/xlwings/xlwings-server-azurefunctions>`_
* `AWS Lambda xlwings server <https://github.com/xlwings/xlwings-server-awslambda>`_

.. important::
    For production deployments, make sure to set up authentication, see :ref:`Server Auth <server_auth>`.

Triggers
--------

.. tab-set::
    .. tab-item:: Excel (Office Scripts)
      :sync: officescripts

      Normally, you would use Power Automate to achieve similar things as with Google Sheets Triggers, but unfortunately, Power Automate can't run Office Scripts that contain a ``fetch`` command like xlwings does, so for the time being, you can only trigger xlwings calls manually on Excel on the web. Alternatively, you can open your Excel file with Google Sheets and leverage the Triggers that Google Sheets offers. This, however, requires you to store your Excel file on Google Drive.

    .. tab-item:: Google Sheets
      :sync: google

      For Google Sheets, you can take advantage of the integrated Triggers (accessible from the menu on the left-hand side of the Apps Script editor). You can trigger your xlwings functions on a schedule or by an event, such as opening or editing a sheet.


Workaround for missing features
-------------------------------

In the classic version of xlwings, you can use the ``.api`` property to fall back to the underlying automation library and work around :ref:`missing features <missing_features>` in xlwings. That's not possible with xlwings Server.

Instead, call the ``book.app.macro()`` method to run functions in JavaScript or VBA, respectively.

.. tab-set::

    .. tab-item:: Excel (VBA)
      :sync: vba

      .. code-block:: vb.net

        ' The first parameter has to be the workbook, the others 
        ' are those parameters that you will provide via Python
        ' NOTE: you're limited to 10 parameters
        Sub WrapText(wb As Workbook, sheetName As String, cellAddress As String)
            wb.Worksheets(sheetName).Range(cellAddress).WrapText = True
        End Sub

      Now you can call this function from Python like so:

      .. code-block:: Python

          # book is an xlwings Book object
          wrap_text = book.app.macro("'MyWorkbook.xlsm'!WrapText")
          wrap_text("Sheet1", "A1")
          wrap_text("Sheet2", "B2")

    .. tab-item:: Excel (Office Scripts)
      :sync: officescripts

      .. code-block:: JavaScript

          // Note that you need to register your function before calling runPython
          async function main(workbook: ExcelScript.Workbook) {
            registerCallback(wrapText);
            await runPython(workbook, "url", { auth: "DEVELOPMENT" });
          }

          // The first parameter has to be the workbook, the others 
          // are those parameters that you will provide via Python
          function wrapText(
            workbook: ExcelScript.Workbook,
            sheetName: string,
            cellAddress: string
          ) {
            const range = workbook.getWorksheet(sheetName).getRange(cellAddress);
            range.getFormat().setWrapText(true);
          }

      Now you can call this function from Python like so:

      .. code-block:: Python

          # book is an xlwings Book object
          wrap_text = book.app.macro("wrapText")
          wrap_text("Sheet1", "A1")
          wrap_text("Sheet2", "B2")

    .. tab-item:: Google Sheets
      :sync: google
  
      .. code-block:: JavaScript

        // The first parameter has to be the workbook, the others 
        // are those parameters that you will provide via Python
        function wrapText(workbook, sheetName, cellAddress) {
          workbook.getSheetByName(sheetName).getRange(cellAddress).setWrap(true);
        }

      Now you can call this function from Python like so:

      .. code-block:: Python

          # book is an xlwings Book object
          wrap_text = book.app.macro("wrapText")
          wrap_text("Sheet1", "A1")
          wrap_text("Sheet2", "B2")

Limitations
-----------

* Currently, only a subset of the xlwings API is covered, mainly the Range and Sheet classes with a focus on reading and writing values and sending pictures (including Matplotlib plots). This, however, includes full support for type conversion including pandas DataFrames, NumPy arrays, datetime objects, etc.
* You are moving within the web's request/response cycle, meaning that values that you write to a range will only be written back to Google Sheets/Excel once the function call returns. Put differently, you'll get the state of the sheets at the moment the call was initiated, but you can't read from a cell you've just written to until the next call.
* You will need to use the same xlwings version for the Python package and the JavaScript module, otherwise, the server will raise an error.
* For users with no experience in web development, this documentation may not be quite good enough just yet.

Platform-specific limitations:

.. tab-set::
    .. tab-item:: Excel on the web
      :sync: officescripts

      * xlwings relies on the ``fetch`` command in Office Scripts that cannot be used via Power Automate and that can be disabled by your Microsoft 365 administrator.
      * While Excel on the web feels generally slow, it seems to have an extreme lag depending on where in the world you open the browser with Excel on the web. For example, a hello world call takes ~4.5s if you open a browser in Amsterdam/Netherlands while it takes ~8.5s if you do it Buenos Aires/Argentina.
      * `Platform limits with Office Scripts <https://docs.microsoft.com/en-us/office/dev/scripts/testing/platform-limits>`_ apply.

    .. tab-item:: Google Sheets
      :sync: google

      * `Quotas for Google Services <https://developers.google.com/apps-script/guides/services/quotas>`_ apply.


Roadmap
-------

* Complete the RunPython API by adding features that currently aren't supported yet, e.g., charts, shapes, etc.
* Perfomance improvements
