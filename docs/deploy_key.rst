.. _deployment_key:

Deployment Key
--------------

This feature requires xlwings :guilabel:`PRO`.

If you have an xlwings PRO developer license, you can generate a deployment key. A deployment key allows you to send an xlwings PRO tool to an end user without them requiring a paid license. A deployment key is also perpetual, i.e. doesn't expire like a developer license.

In return, a deployment key only works with the version of xlwings that was used to generate the deployment key. A developer can generate new deployment keys for new versions of xlwings as long as they have an active xlwings PRO subscription.

.. note::
    You need a paid developer license to generate a deployment key. A trial license won't work.

To create a deployment key, run the following command::

    xlwings license deploy

Then paste the generated key into the xlwings config as ``LICENSE_KEY``. For deployment purposes, usually the best place to do that is on a sheet called ``xlwings.conf``, but you can also use an ``xlwings.conf`` file in either the same folder or in the ``.xlwings`` folder within the user's home folder. To use an environment variable, use ``XLWINGS_LICENSE_KEY``. See also :ref:`settings`.