.. _pro:
.. _license_key:

License Key
===========

To use xlwings PRO functionality, you need to use a license key. xlwings PRO licenses keys are verified offline (i.e., no telemetry/license server involved). There are two types of license keys:

* **Developer key**: this is the key that you will be provided with after purchase. As the name suggests, a developer key should be used by the developer, i.e., someone who writes xlwings code. Developer keys are valid for one or more developers (depending on your plan) and expire after 1 year.
* **Deploy key**: to create deploy keys, you'll need an activated developer key. As the name suggests, a deploy key should be used for the deployment of workbooks to end-users or with xlwings Server. A deploy key doesn't expire but is bound to a specific version of xlwings, which means that you need to generate a new deploy key every time you update xlwings (the ``xlwings release`` command handles this automatically as we'll see below). Note that you can't generate deploy keys with a trial license.

Let's now see how you can get and activate a license key.

How to get a license key
------------------------

**License Key for Commercial Purpose**:

* To try xlwings PRO for free in a commercial context, request a trial license key: https://www.xlwings.org/trial
* To use xlwings PRO in a commercial context beyond the trial, you need to enroll in a paid plan (they include additional services like support and the ability to create one-click installers): https://www.xlwings.org/pricing

**License Key for non-commercial Purpose**:

* To use xlwings PRO for free in a non-commercial context (as defined by the `PolyForm Noncommercial License 1.0.0 <https://polyformproject.org/licenses/noncommercial/1.0.0>`_) use the following license key: ``noncommercial`` (Note that you need at least xlwings 0.26.0).


Activate a developer key
------------------------

To use xlwings PRO locally in your development environment, it's easiest to run the following command::

    xlwings license update -k YOUR_LICENSE_KEY

Make sure to replace ``YOUR_LICENSE_KEY`` with your actual key. This will store the license key in your ``xlwings.conf`` file (see :ref:`user_config` for where this is on your system). Alternatively, you could also activate the license key by setting an environment variable, as we'll see next.

Setting developer keys or deploy keys as environment variable
-------------------------------------------------------------

For xlwings Server deployments, it is recommended to set the license key via the ``XLWINGS_LICENSE_KEY`` environment variable. How you set an environment variable depends on the operating system you use. Managed services often allow you to set environment variables as a secret via their user interface and many systems and frameworks such as Docker can be configured to read environment variables from a local ``.env`` file.

Setting an environment variable is also a convenient way for your local development environment. Just make sure to restart your code editor, IDE, or Terminal/Command Prompt after setting the environment variable.

Generate a deploy key
---------------------

With an activated developer key, you can generate deploy keys like so::

    xlwings license deploy

Make sure that you run this command with the same xlwings version as you'll be using in your deployment. For convenience, when you run this command, the xlwings version will be printed as first line, but you can also query the xlwings version by running the following command in a Terminal/Command Prompt: ``python -c "import xlwings;print(xlwings.__version__)"``

Note that if you use the ``xlwings release`` command, your workbook's ``xlwings.conf`` sheet will be automatically updated with a correct deploy key as we'll see next.

Updating the deploy key in a workbook via the "xlwings release" command
-----------------------------------------------------------------------

The ``xlwings release`` command is the recommended way to prepare a workbook for deployment. It takes care of:

* Setting the deploy key in the ``xlwings.conf`` sheet
* Embedding the code if desired
* Updating the xlwings VBA module so there's no need to use the xlwings add-in on the end-users machine

For more details, see Step 2 under :ref:`1-click installer <release>`.

Updating the deploy key in a workbook manually
----------------------------------------------

To update a workbook with a deploy key for deployment manually:

* Run ``xlwings license deploy``, see above
* Paste the deploy key in your ``xlwings.conf`` sheet as value for ``LICENSE_KEY``, see :ref:`xlwings.conf Sheet <addin_wb_settings>`

Setting the license key in code
-------------------------------

If you run use xlwings PRO by running a Python script directly (e.g., as a frozen executable), it is easiest if you set the deploy key directly in code:

.. code-block::

    import os
    os.environ["XLWINGS_LICENSE_KEY"] = "YOUR_DEPLOY_KEY"

These lines must be run before importing xlwings. It is also best practice to store the deploy key in an external config file instead of hardcoding it directly in the code.
