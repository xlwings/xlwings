.. _permissioning:


Permissioning of Code Execution
===============================

This feature requires xlwings :guilabel:`Enterprise`.

xlwings allows you to control which Python modules are allowed to run from Excel. In order to use this functionality, you need to run your own web server. You can choose between an HTTP POST and a GET request for the permissioning:

* **GET**: This is the simpler option as you only need to host a static JSON file that you can generate via the xlwings CLI. You can use any web server that is capable of serving static files (e.g., nginx) or use a free external service like GitHub pages. However, every permission change requires you to update the JSON file on the server.
* **POST**: This option relies on the web server to validate the incoming payload of the POST request. While this requires custom logic on your end, you are able to connect it with any internal system (such as a database or LDAP server) to dynamically decide whether a user should be able to run a specific Python module through xlwings.

Before looking at each of these two options in more detail, let's go through the prerequisites and configuration.

.. note:: This feature does not stop users from running arbitrary Python code through Python directly. Rather, think of it as a mechanism to prevent accidental execution of Python code from Excel via xlwings.

Prerequisites
-------------

* This functionality requires you to have the ``requests`` and ``cryptography`` library installed. If you don't have them yet, you can install them via pip::

    pip install requests cryptography

  or via Conda::

    conda install requests cryptography

* You need to have a ``LICENSE_KEY`` in the form of a `trial key <https://www.xlwings.org/trial>`_, a paid license key or a deploy key.

Configuration
-------------

While xlwings offers various ways to configure your workbook (see :ref:`Configuration <user_config>`), it will only respect the permissioning settings in the config file in the user's home folder (on Windows, this is ``%USERPROFILE%\.xlwings\xlwings.conf``):

* To prevent end users from overwriting ``xlwings.conf``, you'll need to make sure that the file is owned by the Administrator while giving end users read-only permissions.
* Add the following settings while replacing the ``PERMISSION_CHECK_URL`` and ``PERMISSION_CHECK_METHOD`` (``POST`` or ``GET``) with the proper value for your case::

    "LICENSE_KEY","YOUR_LICENSE_OR_DEPLOY_KEY"
    "PERMISSION_CHECK_ENABLED","True"
    "PERMISSION_CHECK_URL","https://myurl.com"
    "PERMISSION_CHECK_METHOD","POST"

GET request
-----------

You can generate the static JSON file by using the xlwings CLI:

* Print the JSON string for all Python modules in a certain folder::

    cd myfolder
    xlwings permission cwd

* Print the JSON string for all embedded modules of the active workbook::

    xlwings permission book


Both commands will print a JSON string similar to this one::

    {
      "modules": [
        {
          "file_name": "myfile.py",
          "sha256": "cea259922207049a734c88930b5c09109deb6b55f692fd0832f4e57052d85896",
          "machine_names": [
            "DESKTOP-QQ27RP3"
          ]
        },
        {
          "file_name": "myfile2.py",
          "sha256": "355200bb9ae00fcec1d7b660e7dd95fb3dbf246a9db397a6daa2471458a8e6cb",
          "machine_names": [
            "DESKTOP-QQ27RP3"
          ]
        }
      ]
    }

All you need to do at this point is:

* Add additional additional machines names e.g., ``"machine_names: [""DESKTOP-QQ27RP3", "DESKTOP-XY12AS2"]``. Alternatively, you can use the ``"*"`` wildcard if you want to allow the module to be used on all end user's computers. In case of the wildcard, it will still make sure that the file's content hasn't been changed by looking at its sha256 hash. xlwings uses ``import socket;socket.gethostname()`` as the machine name.

* Make this JSON file accessible via your web server and update the settings in the ``xlwings.conf`` file accordingly (see above).

POST request
------------

If you work with POST requests, xlwings will post the following payload::

    {
      "machine_name": "DESKTOP-QQ27RP3",
      "modules": [
        {
           "file_name": "myfile.py",
           "sha256": "cea259922207049a734c88930b5c09109deb6b55f692fd0832f4e57052d85896"
        },
        {
           "file_name": "myfile2.py",
           "sha256": "355200bb9ae00fcec1d7b660e7dd95fb3dbf246a9db397a6daa2471458a8e6cb"
        }
      ]
    }

It is now up to you to validate this request and:

* Return the HTTP status code 200 (Success) if the user is allowed to run the code of these modules
* Return the HTTP status code 403 (Forbidden) if the user is not allowed to run the code of these modules

Note that xlwings only checks for HTTP status code 200, so any other status code will fail.

Implementation Details & Limitations
------------------------------------

* Currently, ``RunPython`` and user-defined functions (UDFs) are supported. ``RunFrozenPython`` is not supported.
* Permissions checks are only done when the Python module is run via Excel/xlwings, it has no effect on Python code that is run from Python directly.
* ``RunPython`` won't allow you to run code that uses the ``from x import y`` syntax. Use ``import x;x.y`` instead.
* The answer of the permissioning server is cached for the duration of the Python session. For UDFs, this means until the functions are re-imported or the ``Restart UDF Server`` button is clicked or until Excel is restarted. The same is true if you run ``RunPython`` with the ``Use UDF Server`` option. By default, however, ``RunPython`` starts a new Python session every time, so it will contact the server whenever you call ``RunPython``.
* Only top-level modules are checked, i.e. modules that are imported as UDFs or run via ``RunPython`` call. Any modules that are imported as dependencies of these modules are not checked.
* ``RunPython`` with external Python source files depends on logic in the VBA part of xlwings. UDFs and RunPython calls that use embedded code will only rely on Python to perform the permissioning.