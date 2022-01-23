.. _remote_interpreter:

Remote Python Interpreter
=========================

This feature requires xlwings :guilabel:`PRO` and at least v0.26.0.

In connection with **Excel on the web**, xlwings can be run on a **Linux-based** server (powered by a web framework of your choice) for a full cloud experience without local installations of neither Excel nor Python. The server can be a cloud service or self-hosted---the only condition is that it must be reachable from Excel on the web. Note that UDFs are not supported with a remote interpreter.

.. caution:: This feature is currently experimental and only covers parts of the xlwings API. See also Limitations at the bottom of this page.

Why is this useful?
-------------------

Office Scripts is the automation language of Excel on the web (it's TypeScript, a typed superset of JavaScript). Office Scripts doesn't allow you to use external JavaScript libraries, so you're very limited in what you can do. You also can't work with multiple modules, making it obvious that Office Scripts hasn't been designed to work with a larger code basis.

On the other hand, xlwings with a remote Python interpreter brings you these advantages:

* Work with the whole Python ecosystem, including pandas, machine learning libraries, database packages, web scraping, boto (for AWS S3) etc. This makes xlwings a great replacement for Power Query that currently isn't available for Excel on the web.
* Leverage your existing development process, including your local or cloud-based IDE/editor as well as your Git workflow, allowing you to easily collaborate and perform code reviews.
* Except for the data you expose in Excel, everything stays on your server. This includes database passwords and other sensitive data such as information about customers. There's also no need to give the Python code to end-users: the whole business logic with your secret sauce is protected.
* Choose the right machine for the job, whether that means using a GPU, a ton of CPU cores, lots of memory, or a gigantic hard disc. As long as Python runs on it, it will work, from server-less functions to Kubernetes (see below).
* Headache-free deployment and maintenance: there's only one location where your Python code lives and you can automate the whole deployment process with continuous integration pipelines like GitHub actions etc.
* Compatibility with Google Sheets and Excel desktop apps (planned).

Prerequisites
-------------

* You need access to Excel on the web with the ``Automate`` tab enabled, i.e., access to Office Scripts. Note that Office Scripts currently requires OneDrive for Business or SharePoint (it's not available on the free office.com), see also: https://docs.microsoft.com/en-gb/office/dev/scripts/overview/excel#requirements
* The ``fetch`` command in Office Scripts must **not** be disabled by your Microsoft 365 administrator.

Introduction
------------

Working with a remote Python interpreter means that you have to expose your Python functions by using a Python web framework. In more detail, handle a POST request along these lines (the sample shows an excerpt that uses FastAPI as the web framework, but it works accordingly with any other web framework like Django or Flask):

.. code-block:: python

    @app.post("/hello-world")
    def hello_world(data: dict = Body(...)):
        # Instantiate a Book object with the deserialized request body
        book = xw.Book(json=data)

        # Use xlwings as usual
        book.sheets[0].value = 'Hello xlwings!'

        # Pass the following back as the response
        return book.json()

Once this runs on a public-facing web server, you simply have to paste the xlwings Office Scripts module into the editor in Excel on the web, adjust the configuration and you're all set! Sound interesting? The next two sections have all the details!

Quickstart: Demo Project
------------------------

If you're impatient and just want to see this working on your end in (literally) less than 5 minutes, head over to the demo project on:

https://github.com/xlwings/xlwings-web-fastapi

The README will provide you with instructions on how to deploy this to a free Heroku app with a single click.

Step-by-step Tutorial
---------------------

Automating Excel on the web consists of two parts: the Python part (the "backend" or "server") and the xlwings Office Script module (the "client" or "frontend"). It's really not that different from the classic use of xlwings except that the Office Scripts module is used in place of the VBA add-in and that the Python backend runs on a server instead of your local machine. In this tutorial, we're using FastAPI as our web framework. While the previous section showed you how you can eventually deploy the Python backend to production, this section shows you a setup for development.

.. note::
    While you can use any web framework other than FastAPI, no quickstart command exists for these yet, so you'd have to set them up manually.

Part I: xlwings Server
**********************

Start a new quickstart project by running the following command on a Terminal/command prompt (feel free to replace ``demo`` with another project name). Before you run this command, make sure to change into the desired directory via ``cd``::

    xlwings quickstart demo --fastapi

This creates a folder called ``demo`` in the current directory with the following files::

    main.py
    app.py
    requirements.txt

I would recommend you to create a virtual or Conda environment where you install these dependencies via ``pip install -r requirements.txt``. In ``app.py``, you'll find the FastAPI boilerplate code and in ``main.py``, you'll find the ``hello_world`` function that is exposed under the ``/hello-world`` endpoint.

To run this server locally, run ``python main.py``. Now, to make this accessible from Excel on the web, you need to either expose your local server to the internet (see below under local development) or you would need to deploy it to production (see production deployment below). For the sake of this tutorial, let's assume you're using ngrok to expose your local web server, in which case you would run the following on your Terminal/Commmand Prompt to expose your local server to the public internet::

    ngrok http 8000

Note that the port 8000 has to correspond to the port that is configured on your local development server as specified at the bottom of `main.py`.

Part II: xlwings Client
***********************

Now it's time to switch to Excel on the web! To paste the xlwings, follow the these steps:

1. On a Terminal/Command Prompt/Anaconda Prompt on your local machine, run the following command: ``xlwings copy os``. This will copy the xlwings Office Scripts module that we'll paste in the Office Script editor in the next step.
2. In Excel on the web, on the ``Automate`` tab, click on ``New Script``. In the editor that appears, paste the script from the previous step and hit ``Save script``. You can also rename it into something meaningful, e.g., ``hello_world``.

API_KEY
EXCLUDE_SHEETS

Local Development
-----------------

If Gitpod or GitHub Codespaces is not an option for you, you can also work with a local environment. The easiest way to is to expose your local development web server externally in a secure way. There are many free and paid services available to help you do this. One of the more popular one is `ngrok <https://ngrok.com/>`_ whose free version will do the trick:

* `ngrok Installation <https://ngrok.com/download>`_
* `ngrok Tutorial <https://ngrok.com/docs>`_


For a list of alternatives, see: https://github.com/anderspitman/awesome-tunneling.


Production Deployment
---------------------

.. important::
    For production deployment, always make sure to set a unique and random ``API_KEY``.

The xlwings web server can be built with any framework and can therefore be deployed with any solution capable of running a Python backend or function. Here is a list for inspiration (non exhaustive):

* **Fully managed services**: Heroku, render.com, fly.io
* **Interactive environments**: anvil, PythonAnywhere
* **Serverless function**: AWS Lambda, Azure Functions, Google Cloud Functions, Vercel, etc.
* **Virtual Machine**: AWS, Microsoft Azure, Google Cloud, DigitalOcean, Linode, vultr, etc.
* **Corporate server**: Anything will work (including Kubernetes) as long as the respective endpoints can be accessed from Excel on the web

Limitations
-----------

* xlwings relies on the ``fetch`` command in Office Scripts that cannot be used via Power Automate and that can be disabled by your administrator.
* xlwings comes with an overhead compared to native Office Scripts as it has to communicate with a remote server. With the hello world example, this is an additional ~2.5s. If you send over a lot of data back and forth, this can be more dramatic, but the overhead may also be totally neglectable, e.g., if you're downloading and analyzing large datasets on a server and only send summary statistics back to Excel. Working with a real backend server also means you can take advantage of everything it offers, e.g., you can run background jobs that pre-process large data sets all the time so that they are ready when you need them.
* Efficiency of working with big arrays can be improved.
* You will need to use the same xlwings version on both the Python package and the OfficeScript module, otherwise the server will raise an error.

Planned next steps
------------------

* Office Scripts integration: add support for missing functionality, e.g., named ranges, tables, etc. and improve efficiency.
* Other integrations: Add support for other systems like Google Sheets and Excel Desktop (Windows & macOS).