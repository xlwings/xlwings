.. _server_auth:

Server Authentication
=====================

This feature requires xlwings PRO.

Authentication (and potentially authorization) is an important step in securing your xlwings Server app. On the server side, you can handle authentication

* within your app (via your web framework)
* outside of your app (via e.g. a reverse proxy such as nginx or oauth2-proxy that sits in front of your app)

Furthermore, you can use different authentication techniques such as HTTP Basic Auth or Bearer tokens in the form of API keys or OAuth2 access tokens. The most reliable and comfortable authentication is available for Office.js add-ins in connection with Excel 365 as this allows you to leverage the built-in SSO capabilities, see :ref:`pro/server/server_authentication:SSO/Azure AD for Excel 365`.

On the client side, you set the ``Authorization`` header when you make a request from Excel or Google Sheets to your xlwings backend. To set the ``Authorization`` header, xlwings offers the ``auth`` parameter:


.. tab-set::

    .. tab-item:: VBA
      :sync: desktop

      .. code-block:: vb.net

        Sub Main()
            RunRemotePython "url", auth:="mytoken"
        End Sub

    .. tab-item:: Office Scripts
      :sync: excel

      .. code-block:: JavaScript

        async function main(workbook: ExcelScript.Workbook) {
          await runPython(workbook, "url", { auth: "mytoken" });
        }

    .. tab-item:: Office.js
      :sync: excel

      .. code-block:: JavaScript

        async function hello {
          // This requires getAuth to be properly implemented, see below under SSO
          let token = await globalThis.getAuth();
          xlwings.runPython("your-url", { auth: token });
        }

    .. tab-item:: Google Apps Script
      :sync: google

      .. code-block:: JavaScript

        function main() {
          let accessToken = ScriptApp.getOAuthToken()
          runPython("url", { auth: "Bearer " + accessToken });
        }

Your backend will then have to validate the Authorization header. Let's get started with the simplest implementation of an API key before looking at HTTP Basic Auth and more advanced options like Azure AD/SSO and Google access tokens (for Google Sheets).

API Key
-------

Generate a secure random string, for example by running the following from a Terminal/Command Prompt::

    python -c "import secrets; print(secrets.token_hex(32))"

Provide this value as your ``auth`` argument in the ``RunRemotePython`` or ``runPython``, respectively, and validate it on your backend along the following lines (these are changes meant to be introduced to a quickstart project or https://github.com/xlwings/xlwings-server-helloworld-fastapi):

.. code-block:: python

    # Only showing additional imports
    import os
    import secrets
    from fastapi import HTTPException, Security, status
    from fastapi.security.api_key import APIKeyHeader

    async def authenticate(api_key: str = Security(APIKeyHeader(name="Authorization"))):
        """Validate the Authorization header"""

        if not secrets.compare_digest(api_key, os.environ["APP_API_KEY"]):
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail="Invalid API Key",
            )

    # If you want to require the API Key for every endpoint
    app = FastAPI(dependencies=[Security(authenticate)])

This sample assumes that you have a single ``APP_API_KEY`` key set as an environment variable on your backend: if you provide the same key as ``auth`` parameter in your ``RunRemotePython`` or ``runPython`` call, everybody with the workbook gets anonymous access. So this approach merely protects your backend from unauthorized access, but it isn't really secure, as there is no secure way to store the API key in the workbook securely, so everybody with the workbook can look up the API key.

If you use the VBA client, you could use a solution where users have to store an individual API Key in an external config file and read it from there. This way, users with the workbook alone would not be able to run the xlwings functionality and you could search for the individual API keys in a database to identify the user.

A much more secure approach is to use Azure AD authentication, see below.

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

SSO/Azure AD for Office.js
--------------------------

.. versionadded:: 0.29.0

Single Sign-on (SSO) means that users who are signed into Office 365 get access to an add-in's Azure AD-protected backend and to Microsoft Graph without needing to sign-in again. Start by reading the official Microsoft documentation:

* `Overview of authentication and authorization in Office Add-ins <https://learn.microsoft.com/en-us/office/dev/add-ins/develop/overview-authn-authz>`_
* `Enable single sign-on (SSO) in an Office Add-in <https://learn.microsoft.com/en-us/office/dev/add-ins/develop/sso-in-office-add-ins>`_

As a summary, here are the components needed to enable SSO:

1. SSO is only available for Office.js add-ins. If you want to enable multi-tenant access (i.e, access for users outside your own organization) external users need to install the add-in via their internal Office add-in store, sideloading the add-in won't work.
2. You must use a supported version of Office, see: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/common/identity-api-requirement-sets
3.  `Register your add-in as an app on the Microsoft Identity Platform <https://learn.microsoft.com/en-us/office/dev/add-ins/develop/register-sso-add-in-aad-v2>`_
4. Add the following to the end of the ``<VersionOverrides ... xsi:type="VersionOverridesV1_0">`` section of your manifest XML:

   .. code-block:: XML
 
     <WebApplicationInfo>
         <Id>Your Client ID</Id>
         <Resource>api://.../Your Client ID</Resource>
         <Scopes>
             <Scope>openid</Scope>
             <Scope>profile</Scope>
             <Scope>...</Scope>
             <Scope>...</Scope>
         </Scopes>
     </WebApplicationInfo>

5.  Acquire an access token in your client-side code and send it as Authorization header to your backend where you can verify it using e.g., Azure functions or parse/verify it manually. You could also use it to authenticate with Microsoft Graph API. The officejs quickstart repo has a dummy global function ``globalThis.getAuth()`` in the ``app/taskpane.html`` file that you can implement as follows (Note that ``Office.auth.getAccessToken`` is supposed to take care of caching automatically, but this doesn't seem to work, see: https://github.com/OfficeDev/office-js/issues/3298):

    .. code-block:: js
  
      let isRenewingToken = false;
      let tokenLock = false;
      let accessToken = null;
      let tokenTimestamp = null;

      function hasKeyExpired() {
        if (!tokenTimestamp) {
          return true;
        }
        // 55 minutes, adjust according to Azure AD token lifetime
        const expirationTime = 55 * 60 * 1000;
        const currentTime = Date.now();
        return currentTime - tokenTimestamp > expirationTime;
      }

      async function renewAccessToken() {
        console.log("Renewing access token");
        try {
          accessToken = await Office.auth.getAccessToken({
            allowSignInPrompt: true,
          });
          accessToken = "Bearer " + accessToken;
          tokenTimestamp = Date.now();
        } catch (error) {
          console.log(`Error ${error.code}: ${error.message}`);
        } finally {
          tokenLock = false;
        }
      }

      globalThis.getAuth = async function () {
        if (!accessToken || hasKeyExpired()) {
          if (!tokenLock) {
            tokenLock = true;
            isRenewingToken = true;
            await renewAccessToken();

            isRenewingToken = false;
          } else {
            while (isRenewingToken) {
              await new Promise((resolve) => setTimeout(resolve, 100));
            }
          }
        }
        return accessToken;
      };

    This then allows you to call ``runPython`` like so (note that custom functions do this automatically):
  
    .. code-block:: JavaScript
  
      async function hello {
        let token = await globalThis.getAuth();
        xlwings.runPython("your-url", { auth: token })
      }

* For a sample implementation on how to validate the token on the backend, have a look at https://github.com/xlwings/xlwings-server-auth-azuread
* A good walkthrough is also `Create a Node.js Office Add-in that uses single sign-on <https://learn.microsoft.com/en-us/office/dev/add-ins/develop/create-sso-office-add-ins-nodejs>`_, but as the title says, it uses Node.js on the backend instead of Python.
* For a reference of the error codes, see: https://learn.microsoft.com/en-us/office/dev/add-ins/develop/troubleshoot-sso-in-office-add-ins


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
