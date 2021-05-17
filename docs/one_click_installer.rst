.. _zero_config_installer:

One-Click Installer
===================

This feature requires xlwings :guilabel:`PRO`.

With xlwings PRO you get access to a private GitHub repository that will build your custom installer in the cloud --- no local installation required. Using a custom installer to deploy the Python runtime has the following advantages:

* Zero Python knowledge required from end users
* Zero configuration required by end users
* No admin rights required
* Works for both UDFs and RunPython
* Works for external distribution
* Easy to deploy updates

End User Instructions
---------------------

* **Installing**

  Give the end user your Excel workbook and the installer. The user only has to double-click the installer and confirm a few prompts --- no configuration is required.

* **Updating**

  If you use the embedded code feature (see: :ref:`embedded_code`), you can deploy updates by simply giving the user a new Excel file. Only when you change a dependency, you will need to create a new installer.

* **Uninstalling**

  The application can be uninstalled again via Windows Settings > Apps & Features.

Build the Installer
-------------------

Before you can build the installer, the project needs to be configured correctly, see below.

In the GitHub repo, go to ``x releases`` > ``Draft/Create a new release``. Add a version like ``1.0.0`` to ``Tag version``, then hit ``Publish release``.

Wait a few minutes and refresh the page: the installer will appear under the release from where you can download it. You can follow the progress under the ``Actions`` tab.

Configuration
-------------

**Excel file**

You can add your Excel file to the repository if you like but it's not a requirement. Configure the Excel file as follows:

* Add the standalone xlwings VBA module, e.g. via ``xlwings quickstart project --standalone``
* Make sure that in the VBA editor (``Alt-F11``) under ``Tools`` > ``References`` xlwings is unchecked
* Rename the ``_xlwings.conf`` sheet into ``xlwings.conf``
* In the ``xlwings.conf`` sheet, as ``Interpreter``, set the following value: ``%LOCALAPPDATA%\project`` while replacing ``project`` with the name of your project
* If you like, you can hide the ``xlwings.conf`` sheet

**Source code**

Source code can either be embedded in the Excel file (see :ref:`embedded_code`) or added to the ``src`` directory. The first option requires ``xlwings-pro`` in ``requirements.txt``, the second option will also work with ``xlwings``.

**Dependencies**

Add your dependencies to ``requirements.txt``. For example::

    xlwings==0.18.0
    numpy==1.18.2

**Code signing (optional)**

Using a code sign certificate will show a verified publisher in the installation prompt. Without it, it will show an unverified publisher.

* Store your code sign certificate as ``sign_cert_file`` in the root of this repository (make sure your repo is private).
* Go to ``Settings`` > ``Secrets`` and add the password as ``code_sign_password``.

**Project details**

Update the following under ``.github/main.yml``::

    PROJECT:
    APP_PUBLISHER:

**Python version**

Set your Python version under ``.github/main.yml``::

    python-version: '3.7'
    architecture: 'x64'
