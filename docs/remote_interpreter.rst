.. _remote_interpreter:

Remote Python Interpreter (Google Sheets & Excel on the web)
============================================================

This feature requires xlwings :guilabel:`PRO` and at least v0.26.0.

In connection with **Google Sheets** or **Excel on the web**, xlwings can be run on a **server** (cloud service or self-hosted) for a full cloud experience without local installations of neither Excel nor Python.

.. important:: This feature is currently experimental and only covers parts of the xlwings API, see also :ref:`Limitations`.

Why is this useful?
-------------------

To automate Office on the web, you have to use Office Scripts (i.e., TypeScript, a typed superset of JavaScript) and for Google Sheets, you have to use Apps Script (i.e., JavaScript). If you don't feel like learning JavaScript, xlwings will allow you to write Python code instead. But even if you are comfortable with JavaScript, you are very limited in what you can do, as both Office Scripts and Apps Script make it pretty much impossible to use external libraries.

On the other hand, xlwings with a remote Python interpreter brings you these advantages:

* **Work with the whole Python ecosystem**: including pandas, machine learning libraries, database packages, web scraping, boto (for AWS S3), etc. This makes xlwings a great alternative for Power Query, which isn't currently available for Excel on the web or Google Sheets.
* **Leverage your existing development workflow**: use your favorite (local or cloud-based) IDE/editor as well as your Git workflow, allowing you to easily collaborate and perform code reviews.
* **Remain in control of your data and code**: except for the data you expose in Excel or Google Sheets, everything stays on your server. This can include database passwords and other sensitive info such as customer data. There's also no need to give the Python code to end-users: the whole business logic with your secret sauce is protected on your own infrastructure.
* **Choose the right machine for the job**: whether that means using a GPU, a ton of CPU cores, lots of memory, or a gigantic hard disc. As long as Python runs on it, you can go from serverless functions all the way to Kubernetes (see :ref:`Production Deployment`).
* **Headache-free deployment and maintenance**: there's only one location where your Python code lives and you can automate the whole deployment process with continuous integration pipelines like GitHub actions etc.

Prerequisites
-------------

* Google Sheets
    * No special requirements.

* Excel on the web
    * You need access to Excel on the web with the ``Automate`` tab enabled, i.e., access to Office Scripts. Note that Office Scripts currently requires OneDrive for Business or SharePoint (it's not available on the free office.com), see also: https://docs.microsoft.com/en-gb/office/dev/scripts/overview/excel#requirements
    * The ``fetch`` command in Office Scripts must **not** be disabled by your Microsoft 365 administrator.

Introduction
------------

Working with a remote Python interpreter consists of two parts:

* the Python part (the "backend" or "server")
* the xlwings Office Scripts or Apps Script module (the "frontend" or "client")

This corresponds to the classic use of xlwings except that the xlwings Office Scripts/Apps Script module is used in place of the VBA add-in/module and that the Python backend runs on a server instead of on your local machine.

Working with a remote Python interpreter means that you have to expose your Python functions by using a Python web framework. In more detail, you need to handle a POST request along these lines (the sample shows an excerpt that uses `FastAPI <https://fastapi.tiangolo.com/>`_ as the web framework, but it works accordingly with any other web framework like Django or Flask):

.. code-block:: python

    @app.post("/hello")
    def hello(data: dict = Body(...)):
        # Instantiate a Book object with the deserialized request body
        book = xw.Book(json=data)

        # Use xlwings as usual
        book.sheets[0].value = 'Hello xlwings!'

        # Pass the following back as the response
        return book.json()

Once this runs on a public-facing web server, you simply have to paste the xlwings Office Scripts or Apps Script module into the editor in Excel on the web or Google Sheets, respectively, adjust the configuration, and you're all set!

Cloud-based development with Gitpod
-----------------------------------

If you want to have a development environment up and running in less than 5 minutes (even if you're new to web development), simply click the following button to open a sample project in `Gitpod <https://www.gitpod.io>`_ (Gitpod is a cloud-based development environment). If you prefer, you can also have a look first at the sample project on GitHub: https://github.com/xlwings/xlwings-web-fastapi

.. image:: https://gitpod.io/button/open-in-gitpod.svg
   :target: https://gitpod.io/#https://github.com/xlwings/xlwings-web-fastapi
   :alt: Open in Gitpod

Opening the project in Gitpod will require you to sign in with your GitHub account. A few moments later, you should see an online version of VS Code. In the Terminal, it will ask you to paste the xlwings license key (get one `here <https://www.xlwings.org/trial>`_). Note that your browser will ask you for permission to paste. Once you confirm your license key by hitting ``Enter``, the server will automatically start with everything properly configured. You can then open the file ``main.py`` in the ``app`` directory, where you'll see the ``hello`` function. Let's leave this alone for a moment and look at the ``js`` folder instead. Depending on whether you want to use Google Sheets or Excel on the web, open the following file:

* Google Sheets: ``xlwings-google.js``
* Excel on the web: ``xlwings-excel.ts``

Copy the code, then switch to Google Sheets or Excel on the web, respectively, and continue as follows:

* **Google Sheets**:
  Click on ``Extensions`` > ``Apps Script``. This will open a separate browser tab and open a file ``Code.gs`` with a function stub. Replace this with the copied code from ``xlwings-google.js``. Then hit the ``Save`` icon and after that the ``Run`` button with the ``hello`` function selected. If you run this the very first time, Google Sheets will ask you for the permissions it needs. Once approved, the script will run the ``hello`` function and write ``Hello xlwings!`` into cell ``A1``. To add a button to a sheet to run this function, switch from the Apps Script editor back to Google Sheets, click on ``Insert`` > ``Drawing`` and draw a rounded rectangle. After hitting ``Save and Close``, the rectangle will appear on the sheet. Click on it so that you can click on the 3 dots on the top right of the shape. Select ``Assign Script`` and write ``hello`` in the text box, then hit ``OK``.

* **Excel on the web**:
  In the ``Automate`` tab, click on ``New Script``. This opens a document in the right side pane where you'll paste the code from ``xlwings-excel.ts``. Make sure to click on ``Save script`` before clicking on ``Run``: the script will run the ``hello`` function and write ``Hello xlwings!`` into cell ``A1``. To run this script from a button, click on the 3 dots in the Office Scripts pane (above the script), then select ``+ Add button``.

Any changes you make to the ``hello`` function in ``app/main.py`` in Gitpod are automatically saved and reloaded by the web server and will be reflected the next time you run the script from Google Sheets or Excel on the web.

To test out the other function of the sample project (``yahoo``), simply replace ``hello`` with ``yahoo`` in the ``runPython`` function in Office Scripts or Apps Script.

.. note:: While Excel on the web requires you to create a separate script for each Python function you want to call (the function has to be called ``main``), Google Sheets allows you to add any number of functions.

Please note that clicking the Gitpod button gets you up and running quickly, but if you want to save your changes (i.e., commit them to GitHub), you should first fork the project on Github and open it via Gitpod.

An alternative to Gitpod is `GitHub CodeSpaces <https://github.com/features/codespaces>`_, but unlike Gitpod, GitHub Codespaces only works with GitHub, has no free tier, and may not be available yet on your account.

Local Development
-----------------

This tutorial walks you through a local development workflow as an alternative to Gitpod. We're going to use `FastAPI <https://fastapi.tiangolo.com/>`_ as our web framework. While you can use any web framework you like, no quickstart command exists for these yet, so you'd have to set the boilerplate up manually. Let's start building the xlwings server first before setting up the xlwings client.

Part I: xlwings Server
**********************

Start a new quickstart project by running the following command on a Terminal/Command Prompt (feel free to replace ``demo`` with another project name). Before you run this command, make sure to change into the desired directory::

    xlwings quickstart demo --fastapi

This creates a folder called ``demo`` in the current directory with the following files::

    main.py
    app.py
    requirements.txt

I would recommend you to create a virtual or Conda environment where you install these dependencies via ``pip install -r requirements.txt``. In ``app.py``, you'll find the FastAPI boilerplate code and in ``main.py``, you'll find the ``hello`` function that is exposed under the ``/hello`` endpoint.

The application expects you to set a unique ``XLWINGS_API_KEY`` as environment variable in order to protect your application from unauthorized access. If you don't set an environment variable, it will expect ``DEVELOPMENT`` as the api key (only use this for quick tests and never for production!).

To run this server locally, run ``python main.py``. Now, to make this accessible from Excel on the web, you need to expose your local server securely to the internet. There are many free and paid services available to help you do this. One of the more popular ones is `ngrok <https://ngrok.com/>`_ whose free version will do the trick:

* `ngrok Installation <https://ngrok.com/download>`_
* `ngrok Tutorial <https://ngrok.com/docs>`_

For a list of alternatives, see https://github.com/anderspitman/awesome-tunneling.

For the sake of this tutorial, let's assume you're using ngrok to expose your local web server, in which case you would run the following on your Terminal/Command Prompt to expose your local server to the public internet::

    ngrok http 8000

Note that the number of the port (8000) has to correspond to the port that is configured on your local development server as specified at the bottom of `main.py`. ngrok will print something along these lines::

    ngrok by @inconshreveable                                                                                (Ctrl+C to quit)

    Session Status                online
    Account                       name@domain.com (Plan: Free)
    Version                       2.3.40
    Region                        United States (us)
    Web Interface                 http://127.0.0.1:4040
    Forwarding                    http://xxxx-xxxx-xx-xx-xxx-xxxx-xxxx-xxxx-xxx.ngrok.io -> http://localhost:8000
    Forwarding                    https://xxxx-xxxx-xx-xx-xxx-xxxx-xxxx-xxxx-xxx.ngrok.io -> http://localhost:8000

To configure the xlwings client in the next step, we'll need the ``https`` version of the forwarding address that ngrok prints, i.e., ``https://xxxx-xxxx-xx-xx-xxx-xxxx-xxxx-xxxx-xxx.ngrok.io``.

Part II: xlwings Client
***********************

Now it's time to switch to Google Sheets or Excel on the web! To paste the xlwings Apps Script or Office Scripts module, follow these steps:

1. **Copy the Apps Script/Office Script xlwings module**: On a Terminal/Command Prompt/Anaconda Prompt on your local machine, run the following command:
    * **Excel no the web**: ``xlwings copy os``
    * **Google Sheets**: ``xlwings copy gs``

    This will copy the xlwings Office Scripts or Apps Script module to the clipboard so we can paste it in the next step.

2. **Paste the Apps Script/Office Script xlwings module**
    * **Excel no the web**: In the ``Automate`` tab, click on ``New Script``. This opens a document in the right side pane where you'll paste the code from ``xlwings-excel.ts``. Make sure to click on ``Save script`` before clicking on ``Run``: the script will run the ``hello`` function and write ``Hello xlwings!`` in cell ``A1``. To run this script from a button, click on the 3 dots in the Office Scripts pane, then select ``+ Add button``.
    * **Google Sheets**: Click on ``Extensions`` > ``Apps Script``. This will open a separate browser tab and open a file `Code.gs` with a function stub. Replace this with the copied code from ``xlwings.js``. Then hit the ``Save`` icon and hit the ``Run`` button with the ``hello`` function selected. If you run this the very first time, Google Sheets will ask you for the permissions it needs. Once approved, the script will run the ``hello`` function and write ``Hello xlwings!`` in cell ``A1``. To add a button to a sheet to run this function, switch from the Apps Script editor back to Google Sheets, then click on ``Insert`` > ``Drawing`` and draw a rounded rectangle. After hitting ``Save and Close``, the rectangle will appear on the sheet. Click on it so that you can click on the 3 dots on the top right of the shape. Select ``Assign Script`` and write ``hello`` in the text box, then hit ``OK``.

3. **Configuration**
    The final step is to configure the Apps Script/Office Scripts properly, see the next section :ref:`Configuration`.

Configuration
-------------

The Office Scripts/App Script xlwings module can be configured in two ways:

* Directly in the ``runPython`` function as arguments
* On a sheet called ``xlwings.conf``

If both ways are configured, the function arguments are used. Using the ``xlwings.conf`` sheet has the advantages that you can (a) upgrade your xlwings script without having to adjust the code and (b) you can share your configuration with multiple scripts (as Office Scripts only allows you to set up one function per script). Let's first see what the available settings are:

* ``URL`` (required): This is the full URL of your function. In the above example of :ref:`Local Development`, this would be ``https://xxxx-xxxx-xx-xx-xxx-xxxx-xxxx-xxxx-xxx.ngrok.io/hello``, i.e., the ngrok URL **with the /hello endpoint appended**.
* ``API_KEY`` (required): The API_KEY is a key that you set yourself on both the server (as ``XLWINGS_API_KEY`` environment var) and on the client (via configuation) to protect your functions from unauthorized access. You should choose a strong random key, for example by running the following on a Terminal/Command Prompt: ``python -c "import secrets; print(secrets.token_hex(32))"``. It's good practice to keep your sensitive keys such as the ``API_KEY`` out of your source code (the Office Scripts/App Scripts module), but putting in in the ``xlwings.conf`` sheet may only be marginally better. Excel on the web, however, doesn't currently provide you with a better way of handling this. Google sheets, on the other hand, allows you to work with `Properties Service <https://developers.google.com/apps-script/guides/properties>`_ to keep the ``API_KEY`` out of both the code and sheet.

  .. note:: The API_KEY is chosen by you to protect your application and has nothing to do with the xlwings license key!

* ``EXCLUDE`` (optional): By default, xlwings sends over the complete content of the whole workbook. If you have sheets with big amounts of data, this can make the calls slow. If your backend doesn't need the content of certain sheets, you can exclude the content from being sent over via the ``EXCLUDE`` setting. Currently, you can only exclude entire sheets as comma-delimited string like so: ``Sheet1, Sheet2``.

Examples for function arguments
*******************************

* **Google Sheets**:

  Only required arguments:

  .. code-block:: JavaScript

    function hello() {
      runPython(
        "https://xxxx-xxxx-xx-xx-xxx-xxxx-xxxx-xxxx-xxx.ngrok.io/hello",
        "YOUR_UNIUQE_API_KEY"
      );
    }

  Excluding the ``xlwings.conf`` and ``Sheet1``:

  .. code-block:: JavaScript

    function hello() {
      runPython(
        "https://xxxx-xxxx-xx-xx-xxx-xxxx-xxxx-xxxx-xxx.ngrok.io/hello",
        "YOUR_UNIUQE_API_KEY",
        "xlwings.conf, Sheet1"
      );
    }

* **Excel on the web**:

  Only required arguments:

  .. code-block:: JavaScript

    async function main(workbook: ExcelScript.Workbook) {
      await runPython(
        workbook,
        "https://xxxx-xxxx-xx-xx-xxx-xxxx-xxxx-xxxx-xxx.ngrok.io/hello",
        "YOUR_UNIUQE_API_KEY"
      );
    }

  Excluding the ``xlwings.conf`` and ``Sheet1``:

  .. code-block:: JavaScript

    async function main(workbook: ExcelScript.Workbook) {
      await runPython(
        workbook,
        "https://xxxx-xxxx-xx-xx-xxx-xxxx-xxxx-xxxx-xxx.ngrok.io/hello",
        "YOUR_UNIUQE_API_KEY",
        "xlwings.conf, Sheet1"
      );
    }

Examples for xlwings.conf sheet
*******************************

Create a sheet called ``xlwings.conf`` and fill in key/value pairs like so:

.. figure:: images/xlwings_conf_sheet.png

You could now use this configuration as follows:

* **Google Sheets** (both functions can be on a single xlwings module):

  .. code-block:: JavaScript

    function hello() {
      runPython(
        "URL",
        "API_KEY"
      );
    }

    function yahoo() {
      runPython(
        "URL_YAHOO",
        "API_KEY"
      );
    }

* **Excel on the web** (the calls have to be on separate xlwings modules):

  .. code-block:: JavaScript

    // Script 1
    async function main(workbook: ExcelScript.Workbook) {
      await runPython(
        workbook,
        "URL",
        "API_KEY"
      );
    }

  .. code-block:: JavaScript

    // Script 2
    async function main(workbook: ExcelScript.Workbook) {
      await runPython(
        workbook,
        "URL_YAHOO",
        "API_KEY"
      );
    }

Production Deployment
---------------------

The xlwings web server can be built with any web framework and can therefore be deployed using any solution capable of running a Python backend or function. Here is a list for inspiration (non-exhaustive):

* **Fully-managed services**: `Heroku <https://www.heroku.com>`_, `render <https://www.render.com>`_, `Fly.io <https://www.fly.io>`_, etc.
* **Interactive environments**: `PythonAnywhere <https://www.pythonanywhere.com>`_, `Anvil <https://www.anvil.works>`_, etc.
* **Serverless function**: `AWS Lambda <https://aws.amazon.com/lambda/>`_, `Azure Functions <https://azure.microsoft.com/en-us/services/functions/>`_, `Google Cloud Functions <https://cloud.google.com/functions>`_, `Vercel <https://vercel.com>`_, etc.
* **Virtual Machine**: `DigitalOcean <https://m.do.co/c/ed671b0a5a9b>`_ (referral link), `vultr <https://www.vultr.com/?ref=7155223>`_ (referral link), `Linode <https://www.linode.com/>`_, `AWS EC2 <https://aws.amazon.com/ec2/>`_, `Microsoft Azure VM <https://azure.microsoft.com/en-us/services/virtual-machines/>`_, `Google Cloud Compute Engine <https://cloud.google.com/compute>`_, etc.
* **Corporate server**: Anything will work (including Kubernetes) as long as the respective endpoints can be accessed from Excel on the web or Google Sheets.

.. important::
    For production deployment, always make sure to set a unique and random ``API_KEY``, see :ref:`Configuration`.

Triggers
--------

* **Google Sheets**:
  For Google Sheets, you can take advantage of the integrated Triggers (accessible from the menu on the left-hand side of the Apps Script editor). You can trigger your xlwings functions on a schedule or by an event, such as opening or editing a sheet.

* **Excel on the web**:
  Normally, you would use Power Automate to achieve similar things as with Google Sheets Triggers, but unfortunately, Power Automate can't run Office Scripts that contain a ``fetch`` command like xlwings does, so for the time being, you can only trigger xlwings calls manually on Excel on the web.

Limitations
-----------

* Currently, only a subset of the full xlwings API is covered, mainly the Range and Sheet classes with a focus on reading and writing values. This, however, includes full support for type conversion including pandas DataFrames, NumPy arrays, datetime objects, etc.
* You will need to use the same xlwings version for the Python package and the OfficeScript module, otherwise, the server will raise an error.
* Custom functions (a.k.a. User-defined functions or UDFs) are not currently supported.
* **Excel on the web only:** xlwings relies on the ``fetch`` command in Office Scripts that cannot be used via Power Automate and that can be disabled by your administrator.

Planned next steps
------------------

* Office Scripts integration: add support for missing functionality, e.g., charts, shapes, named ranges, tables, etc. and improve efficiency.
* Other integrations: Add support for Excel Desktop (Windows & macOS). Note that Office Scripts on Windows is in Beta (Microsoft 365 only), so if you have access to this, it should work out of the box.
