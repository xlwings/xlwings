Production Deployment
---------------------

The xlwings web server can be built with any web framework and can therefore be deployed using any solution capable of running a Python backend or function. Here is a list for inspiration (non-exhaustive):

* **Fully-managed services**: `Heroku <https://www.heroku.com>`_, `Render <https://www.render.com>`_, `Fly.io <https://www.fly.io>`_, etc.
* **Interactive environments**: `PythonAnywhere <https://www.pythonanywhere.com>`_, `Anvil <https://www.anvil.works>`_, etc.
* **Serverless functions**: `AWS Lambda <https://aws.amazon.com/lambda/>`_, `Azure Functions <https://azure.microsoft.com/en-us/services/functions/>`_, `Google Cloud Functions <https://cloud.google.com/functions>`_, `Vercel <https://vercel.com>`_, etc.
* **Virtual Machines**: `DigitalOcean <https://digitalocean.com>`_, `vultr <https://www.vultr.com>`_, `Linode <https://www.linode.com/>`_, `AWS EC2 <https://aws.amazon.com/ec2/>`_, `Microsoft Azure VM <https://azure.microsoft.com/en-us/services/virtual-machines/>`_, `Google Cloud Compute Engine <https://cloud.google.com/compute>`_, etc.
* **Corporate servers**: Anything will work (including Kubernetes) as long as the respective endpoints can be accessed from your spreadsheet app.

Serverless Functions
--------------------

For examples how to configure the serverless function platform with xlwings see the following example repositories.

* `DigitalOcean Functions https://github.com/xlwings/xlwings-server-digitaloceanfunctions`_
* `Azure Functions https://github.com/xlwings/xlwings-server-azurefunctions`_
* `AWS Lambda https://github.com/xlwings/xlwings-server-awslambda`

.. important::
    For production deployments, make sure to set up authentication, see :ref:`Server Auth <server_auth>`.