.. _server_auth:

Server Auth :bdg-secondary:`PRO`
================================

Authentication (and potentially authorization) is an important step in securing your xlwings Server app. On the server side, you can handle authentication

* within your app (via your web framework)
* outside of your app (via e.g. a reverse proxy such as nginx or oauth2-proxy that sits in front of your app)

Furthermore, you can use different authentication techniques such as HTTP Basic Auth or Bearer tokens in the form of API keys or OAuth2 access tokens.

On the client side, you set the ``Authorization`` header when you make a request from Excel or Google Sheets to your xlwings backend. To set the ``Authorization`` header, xlwings offers the ``auth`` parameter:


.. tab-set::

    .. tab-item:: Excel (via VBA)
      :sync: desktop

      .. code-block:: vb.net

        Sub Main()
            RunRemotePython "url", auth:="..."
        End Sub

    .. tab-item:: Excel (via Office Scripts)
      :sync: excel

      .. code-block:: JavaScript

        async function main(workbook: ExcelScript.Workbook) {
          await runPython(workbook, "url", {
            auth: "...",
          });
        }

    .. tab-item:: Google Sheets
      :sync: google

      .. code-block:: JavaScript

        function main() {
          runPython("url", {
            auth: "...",
          });
        }

Your backend will then have to validate the Authorization header. Let's get started with the simplest implementation of an API key before looking at HTTP Basic Auth and more advanced options like Azure AD and Google access tokens (for Google Sheets).

API Key
-------

Generate a secure random string, for example by running the following from a Terminal/Command Prompt::

    python -c "import secrets; print(secrets.token_hex(32))"

Provide this value as your ``auth`` argument in the ``RunRemotePython`` or ``runPython`` call and validate it on your backend.

| For a sample backend implementation, see:
| https://github.com/xlwings/xlwings-server-helloworld-fastapi

HTTP Basic Auth
---------------

Basic auth is a simple and popular method that sends the username and password via the Authorization header.
Reverse proxies such as nginx allow you to easily protect your app with HTTP Basic Auth but you can also handle it directly in your app.

With your username and password, run the following Python script to get the value that you need to provide for ``auth``::

    import base64
    username = "myusername"
    password = "mypassword"
    print("Basic " + base64.b64encode(f"{username}:{password}".encode()).decode())

In this case, you'd provide ``"Basic bXl1c2VybmFtZTpteXBhc3N3b3Jk"`` as your ``auth`` argument.

* To validate HTTP Basic Auth with FastAPI, see: https://fastapi.tiangolo.com/advanced/security/http-basic-auth/
* If you use ngrok, there's an easy way to protect the exposed URL via Basic auth:

  .. code-block:: Text

        ngrok http 8000 -auth='myusername:mypassword'

  .. warning::
    ngrok HTTP Basic auth will NOT work with Excel via Office Scripts as it doesn't support CORS. It's, however, an easy method for protecting your app during development if you use xlwings via VBA or Google Sheets.

Azure AD for Excel VBA
----------------------

.. versionadded:: 0.28.6

  .. note::
    Azure AD authentication is only available for Desktop Excel via VBA.

`Azure Active Directory (Azure AD) <https://azure.microsoft.com/en-us/products/active-directory>`_ is Microsoft's enterprise identity service. If you're using the xlwings add-in or VBA standalone module, xlwings allows you to comfortably log in users on their desktops, allowing you to securely validate their identity on the server and optionally implement role-base access control (RBAC).

Download ``xlwings.exe``, the standalone xlwings CLI, from the `GitHub Release page <https://github.com/xlwings/xlwings/releases>`_ and place it in a specific folder, e.g., under ``C:\Program and Files\xlwings\xlwings.exe`` or ``%LOCALAPPDATA%\xlwings\xlwings.exe``.

Now you can call the following function in VBA:

.. code-block:: vb.net

    Sub Main()
      RunRemotePython "url", _
      auth:="Bearer " & GetAzureAdAccessToken( _
        tenantId:="...", _
        clientId:="...", _
        scopes:="...", _
        port:="...", _
        username:="...", _
        cliPath:="C:\Program and Files\xlwings\xlwings.exe" _
      )
    End Sub

``port`` and ``username`` are optional:

* Use ``port`` if the randomly assigned default port causes issues
* Use ``username`` if the user is logged in with multiple Microsoft accounts

.. note::
  Instead of relying on ``xlwings.exe``, you could also use a normal Python installation with ``xlwings`` and  ``msal`` installed. In this case, simply leave away the ``cliPath`` argument.

You can also use the ``xlwings.conf`` file or ``xlwings.conf`` sheet for configuration. In this case, the settings are the following:

.. code-block::

    AZUREAD_TENANT_ID
    AZUREAD_CLIENT_ID
    AZUREAD_SCOPES
    AZUREAD_USERNAME
    AZUREAD_PORT
    CLI_PATH

Note that if you use the xlwings add-in rather than relying on the xlwings standalone VBA module, you will need to make sure that there's a reference set to xlwings in the VBA editor under ``Tools`` > ``References``.

When you now call the ``Main`` function the very first time, a browser Window will open where the user needs to login to Azure AD. The acquired OAuth2 access token is then cached for 60-90 minutes. Once an access token has expired, a new one will be requested using the refresh token, i.e., without user intervention, but it will slow that that request.

For a complete walk-through on how to set up an app on Azure AD and how to validate the access token on the backend, see: https://github.com/xlwings/xlwings-server-auth-azuread

OAuth2 Access Token for Google Sheets
-------------------------------------

Google makes it easy to verify the logged-in user via OAuth2 access token. Simply provide the following as your ``auth`` argument:

.. code-block:: JavaScript

    ScriptApp.getOAuthToken()

| To see how you can validate that token on the backend, see:
| https://github.com/xlwings/xlwings-server-auth-google
