VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' CHANGES:
' * PrepareCurlRequest: allow big request body, based on https://github.com/VBA-tools/VBA-Web/pull/397
'
' WebClient v4.1.6
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' `WebClient` executes requests and handles response and is responsible for functionality shared between requests,
'  such as authentication, proxy configuration, and security.
'
' Usage:
'
' ```VB.net
' Dim Client As New WebClient
' Client.BaseUrl = "https://www.example.com/api/"
'
' Dim Auth As New HttpBasicAuthenticator
' Auth.Setup Username, Password
' Set Client.Authenticator = Auth
'
' Dim Request As New WebRequest
' Dim Response As WebResponse
' ' Setup WebRequest...
'
' Set Response = Client.Execute(Request)
' ' -> Uses Http Basic authentication and appends Request.Resource to BaseUrl
' ```
'
' Errors:
' 11010 / 80042b02 / -2147210494 - cURL error in Execute
' 11011 / 80042b03 / -2147210493 - Error in Execute
' 11012 / 80042b04 / -2147210492 - Error preparing http request
' 11013 / 80042b05 / -2147210491 - Error preparing cURL request
'
' @class WebClient
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private Const web_DefaultTimeoutMs As Long = 5000

Private Const web_HttpRequest_SetCredentials_ForServer = 0
Private Const web_HttpRequest_SetCredentials_ForProxy = 1

Private Const web_HttpRequest_ProxySetting_Default = 0
Private Const web_HttpRequest_ProxySetting_PreConfig = 0
Private Const web_HttpRequest_ProxySetting_Direct = 1
Private Const web_HttpRequest_ProxySetting_Proxy = 2

Private Enum web_WinHttpRequestOption
    web_WinHttpRequestOption_UserAgentString = 0
    web_WinHttpRequestOption_URL = 1
    web_WinHttpRequestOption_URLCodePage = 2
    web_WinHttpRequestOption_EscapePercentInURL = 3
    web_WinHttpRequestOption_SslErrorIgnoreFlags = 4
    web_WinHttpRequestOption_SelectCertificate = 5
    web_WinHttpRequestOption_EnableRedirects = 6
    web_WinHttpRequestOption_UrlEscapeDisable = 7
    web_WinHttpRequestOption_UrlEscapeDisableQuery = 8
    web_WinHttpRequestOption_SecureProtocols = 9
    web_WinHttpRequestOption_EnableTracing = 10
    web_WinHttpRequestOption_RevertImpersonationOverSsl = 11
    web_WinHttpRequestOption_EnableHttpsToHttpRedirects = 12
    web_WinHttpRequestOption_EnablePassportAuthentication = 13
    web_WinHttpRequestOption_MaxAutomaticRedirects = 14
    web_WinHttpRequestOption_MaxResponseHeaderSize = 15
    web_WinHttpRequestOption_MaxResponseDrainSize = 16
    web_WinHttpRequestOption_EnableHttp1_1 = 17
    web_WinHttpRequestOption_EnableCertificateRevocationCheck = 18
End Enum

Private web_pProxyServer As String
Private web_pAutoProxyDomain As String

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

''
' Set the base url that is shared by all requests
' and that the request `Resource` is appended to.
'
' @example
' ```VB.net
' ' Desired URLs
' ' https://api.example.com/v1/messages
' ' https://api.example.com/v1/users/id
' '                BaseUrl <- ^ -> Resource
'
' Dim Client As New WebClient
' Client.BaseUrl = "https://api.example.com/v1/"
'
' Dim Request As New WebRequest
' Request.Resource = "messages"
' Request.Resource = "users/{id}"
' ```
'
' @property BaseUrl
' @type String
''
Public BaseUrl As String

''
' Attach an authenticator to the client for authentication requests.
'
' @example
' ```VB.net
' Dim Client As New WebClient
' Dim Auth As New OAuth1Authenticator
' Auth.Setup ...
'
' Set Client.Authenticator = Auth
' ' -> All requests use Auth to add "Authorization" header
' ```
'
' @property Authenticator
' @type IWebAuthenticator
''
Public Authenticator As IWebAuthenticator

''
' Timeout (in milliseconds) to wait for timeout in each request phase
' (Resolve, Connect, Send, Receive).

'
' @property TimeoutMs
' @type Long
' @default 5000
''
Public TimeoutMs As Long

''
' Comma separated list of domains to bypass the proxy.
'
' @property ProxyBypassList
' @type String
''
Public proxyBypassList As String

''
' Username for proxy.
'
' @property ProxyUsername
' @type String
''
Public proxyUsername As String

''
' Password for proxy.
'
' @property ProxyPassword
' @type String
''
Public proxyPassword As String

''
' Load proxy server and bypass list automatically (`False` by default).
'
' @property EnableAutoProxy
' @type Boolean
' @default False
''
Public enableAutoProxy As Boolean

''
' Turn off SSL validation (`False` by default).
' Useful for self-signed certificates and should only be used with trusted servers.
'
' @property Insecure
' @type Boolean
' @default False
''
Public insecure As Boolean

''
' Follow redirects (301, 302, 307) using Location header
'
' @property FollowRedirects
' @type Boolean
' @default True
''
Public followRedirects As Boolean

''
' Proxy server to pass requests through (except for those that match `ProxyBypassList`).
'
' @property ProxyServer
' @type String
''
Public Property Get proxyServer() As String
    proxyServer = web_pProxyServer
End Property
Public Property Let proxyServer(Value As String)
    Me.enableAutoProxy = False
    web_pProxyServer = Value
End Property

' ============================================= '
' Public Methods
' ============================================= '

''
' Execute the given request
' (append the request's `FormattedResource` to the `BaseUrl`)
' and return the response.
'
' @example
' ```VB.net
' Dim Client As New WebClient
' Client.BaseUrl = "https://api.example.com/v1/"
'
' Dim Request As New WebRequest
' Request.Resource = "messages/{id}"
' Request.AddUrlSegment "id", 123
'
' ' Add querystring, body, headers, cookies, etc. for request
'
' Dim Response As WebResponse
' Set Response = Client.Execute(Request)
'
' ' -> GET https://api.example/com/v1/messages/123
' '    headers, cookies, and body...
' ```
'
' @method Execute
' @param {WebRequest} request The request to execute
' @return {WebResponse} Wrapper of server response for request
' @throws 11010 / 80042b02 / -2147210494 - cURL error in Execute
' @throws 11011 / 80042b03 / -2147210493 - Error in Execute
''
Public Function Execute(Request As WebRequest) As WebResponse
    Dim web_Http As Object
    Dim web_Response As New WebResponse

    On Error GoTo web_ErrorHandling

#If Mac Then
    Dim web_Curl As String
    Dim web_Result As ShellResult

    web_Curl = Me.PrepareCurlRequest(Request)
    web_Result = WebHelpers.ExecuteInShell(web_Curl)

    ' Handle cURL errors
    '
    ' Map to WinHttp error number, as possible
    ' https://msdn.microsoft.com/en-us/library/aa383770(VS.85).aspx
    If web_Result.ExitCode > 0 Then
        Dim web_ErrorNumber As Long
        Dim web_ErrorMessage As String
        Dim web_ErrorDetails As String
        web_ErrorNumber = web_Result.ExitCode / 256

        Select Case web_ErrorNumber
        Case 1
            ' 1 = CURLE_UNSUPPORTED_PROTOCOL
            ' 12006 = ERROR_WINHTTP_UNRECOGNIZED_SCHEME
            Err.Raise 12006 + &H30000 + vbObjectError, "The URL does not use a recognized protocol (1: CURLE_UNSUPPORTED_PROTOCOL)" & vbNewLine & _
                "URL: " & Me.GetFullUrl(Request) & vbNewLine & _
                "Protocol: " & WebHelpers.GetUrlParts(Me.GetFullUrl(Request))("Protocol")
        Case 3
            ' 3 = CURLE_URL_MALFORMAT
            ' 12005 = ERROR_WINHTTP_INVALID_URL
            Err.Raise 12005 + &H30000 + vbObjectError, "The URL is invalid (3: CURLE_URL_MALFORMAT)" & _
                "URL: " & Me.GetFullUrl(Request)
        Case 5, 6
            ' 5 = CURLE_COULDNT_RESOLVE_PROXY
            ' 6 = CURLE_COULDNT_RESOLVE_HOST
            ' 12007 = ERROR_WINHTTP_NAME_NOT_RESOLVED
            If web_ErrorNumber = 5 Then
                web_ErrorDetails = "(5: CURLE_COULDNT_RESOLVE_PROXY)"
            Else
                web_ErrorDetails = "(6: CURLE_COULDNT_RESOLVE_HOST)"
            End If

            Err.Raise 12007 + &H30000 + vbObjectError, "WebClient.Execute", "The server name or address could not be resolved " & web_ErrorDetails
        Case 7
            ' 7 = CURLE_COULDNT_CONNECT
            ' 12029 = ERROR_WINHTTP_CANNOT_CONNECT
            Err.Raise 12029 + &H30000 + vbObjectError, "WebClient.Execute", "A connection with the server could not be established (7: CURLE_COULDNT_CONNECT)"
        Case 12, 28
            ' 12 = CURLE_FTP_ACCEPT_TIMEOUT
            ' 28 = CURLE_OPERATION_TIMEDOUT
            ' 12002 = ERROR_WINHTTP_TIMEOUT
            If web_ErrorNumber = 12 Then
                web_ErrorDetails = "(12: CURLE_FTP_ACCEPT_TIMEOUT)"
            Else
                web_ErrorDetails = "(28: CURLE_OPERATION_TIMEDOUT)"
            End If

            Err.Raise 12002 + &H30000 + vbObjectError, "WebClient.Execute", "The operation timed out " & web_ErrorDetails
        Case 47
            ' 47 = CURLE_TOO_MANY_REDIRECTS
            ' 12156 = ERROR_WINHTTP_REDIRECT_FAILED
            Err.Raise 12156 + &H30000 + vbObjectError, "WebClient.Execute", "Too many redirects (47: CURLE_TOO_MANY_REDIRECTS)"
        Case Else
            Err.Raise 11010 + vbObjectError, "WebClient.Execute", "An unknown cURL error occured, #" & web_ErrorNumber & vbNewLine & _
                "Find details at http://curl.haxx.se/libcurl/c/libcurl-errors.html"
        End Select
    End If

    web_Response.CreateFromCurl Me, Request, web_Result.Output

#Else
    Set web_Http = Me.PrepareHttpRequest(Request)

    web_Http.Send Request.Body
    Do While Not web_Http.WaitForResponse(0.025)
        VBA.DoEvents
    Loop

    web_Response.CreateFromHttp Me, Request, web_Http

#End If

    WebHelpers.LogResponse Me, Request, web_Response

    If Not Me.Authenticator Is Nothing Then
        Me.Authenticator.AfterExecute Me, Request, web_Response
    End If

    Set web_Http = Nothing
    Set Execute = web_Response
    Exit Function

web_ErrorHandling:

    Set web_Http = Nothing
    Dim web_ErrorDescription As String

    ' Check lower 16 bits from error
    ' (e.g. 80072EE2 -> 2EE2 -> 12002)
    Select Case Err.Number And 65535
    Case 12002, 12007, 12029
        ' Treat timeout-related errors as 408: timeout, name not resolved, cannot connect
        web_Response.StatusCode = WebStatusCode.RequestTimeout
        web_Response.StatusDescription = "Request Timeout: " & Err.Description

        WebHelpers.LogResponse Me, Request, web_Response
        Set Execute = web_Response
        Err.Clear
    Case Else
        ' Error
        web_ErrorDescription = "An error occurred during execute" & vbNewLine & _
            Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description

        WebHelpers.LogError web_ErrorDescription, "WebClient.Execute", 11011 + vbObjectError
        Err.Raise 11011 + vbObjectError, "WebClient.Execute", web_ErrorDescription
    End Select
End Function

''
' Get JSON from the given URL
' (with options for Headers, Cookies, QuerystringParams, and UrlSegments).
'
' @example
' ```VB.net
' Dim Client As New WebClient
' Dim Url As String
' Url = "https://api.example.com/v1/messages/1"
'
' Dim Response As WebResponse
' Set Response = Client.GetJson(Url)
'
' Dim Headers As New Collection
' Headers.Add WebHelpers.CreateKeyValue("Authorization", "Bearer ...")
'
' Dim Options As New Dictionary
' Options.Add "Headers", Headers
'
' Set Response = Client.GetJson(Url, Options)
' ```
'
' @method GetJson
' @param {String} Url (appended to `BaseUrl`, if set)
' @param {Dictionary} [Options]
' @param {Collection} [Options.Headers] Collection of `KeyValue`
' @param {Collection} [Options.Cookies] Collection of `KeyValue`
' @param {Collection} [Options.QuerystringParams] Collection of `KeyValue`
' @param {Dictionary} [Options.UrlSegments]
' @return {WebResponse} Response
''
Public Function GetJson(url As String, Optional Options As Dictionary) As WebResponse
    Dim web_Request As New WebRequest
    web_Request.CreateFromOptions Options
    web_Request.Resource = url
    web_Request.Format = WebFormat.Json
    web_Request.Method = WebMethod.HttpGet

    Set GetJson = Me.Execute(web_Request)
End Function

''
' Post JSON Body (`Array`, `Collection`, `Dictionary`) to the given URL
' (with options for Headers, Cookies, QuerystringParams, and UrlSegments).
'
' @example
' ```VB.net
' Dim Client As New WebClient
' Dim Url As String
' Url = "https://api.example.com/v1/messages/1"
'
' ' Body
' ' Array, Collection, or Dictionary
' Dim Body As New Dictionary
' Body.Add "message", "Howdy!"
'
' Dim Response As WebResponse
' Set Response = Client.PostJson(Url, Body)
'
' Dim Headers As New Collection
' Headers.Add WebHelpers.CreateKeyValue("Authorization", "Bearer ...")
'
' Dim Options As New Dictionary
' Options.Add "Headers", Headers
'
' Set Response = Client.PostJson(Url, Body, Options)
' ```
'
' @method PostJson
' @param {String} Url (appended to `BaseUrl`, if set)
' @param {Dictionary} Body
' @param {Dictionary} [Options]
' @param {Collection} [Options.Headers] Collection of `KeyValue`
' @param {Collection} [Options.Cookies] Collection of `KeyValue`
' @param {Collection} [Options.QuerystringParams] Collection of `KeyValue`
' @param {Dictionary} [Options.UrlSegments]
' @return {WebResponse} Response
''
Public Function PostJson(url As String, Body As Variant, Optional Options As Dictionary) As WebResponse
    Dim web_Request As New WebRequest
    web_Request.CreateFromOptions Options
    web_Request.Resource = url
    web_Request.Format = WebFormat.Json
    web_Request.Method = WebMethod.HttpPost
    If VBA.IsObject(Body) Then
        Set web_Request.Body = Body
    Else
        web_Request.Body = Body
    End If

    Set PostJson = Me.Execute(web_Request)
End Function

''
' Set proxy for all requests
'
' @example
' ```VB.net
' Dim Client As New RestClient
'
' ' Just Server
' Client.SetProxy "proxy_server:80"
'
' ' Server + Username and Password
' Client.SetProxy "proxy_server:80", "Tim", "Password"
'
' ' Server + Username and Password + BypassList
' Client.SetProxy "proxy_server:80", "Tim", "Password", "<local>,*.bypass.com"
' ```
'
' @method SetProxy
' @param {String} ProxyServer Proxy server to pass requests through
' @param {String} [Username=""] Username for proxy
' @param {String} [Password=""] Password for proxy
' @param {String} [BypassList=""] Comma-separated list of servers that should bypass proxy
''
Public Sub SetProxy(proxyServer As String, _
    Optional Username As String = "", Optional Password As String = "", Optional BypassList As String = "")

    Me.proxyServer = proxyServer
    Me.proxyUsername = Username
    Me.proxyPassword = Password
    Me.proxyBypassList = BypassList
End Sub

''
' Get full url by joining given `WebRequest.FormattedResource` and `BaseUrl`.
'
' @method GetFullUrl
' @param {WebRequest} Request
' @return {String}
''
Public Function GetFullUrl(Request As WebRequest) As String
    GetFullUrl = WebHelpers.JoinUrl(Me.BaseUrl, Request.FormattedResource)
End Function

''
' Prepare Http request for given WebRequest
'
' @internal
' @method PrepareHttpRequest
' @param {WebRequest} Request
' @return {WinHttpRequest}
' @throws 11012 / 80042b04 / -2147210492 - Error preparing http request
''
Public Function PrepareHttpRequest(Request As WebRequest, Optional Async As Boolean = True) As Object
    Dim web_Http As Object
    Dim web_KeyValue As Dictionary

    On Error GoTo web_ErrorHandling

    Set web_Http = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' Prepare request (before open)
    web_BeforeExecute Request

    ' Open http request
    web_Http.Open WebHelpers.MethodToName(Request.Method), Me.GetFullUrl(Request), Async

    ' Set timeouts
    web_Http.SetTimeouts Me.TimeoutMs, Me.TimeoutMs, Me.TimeoutMs, Me.TimeoutMs

    ' Load auto-proxy (if needed)
    If Me.enableAutoProxy Then
        web_LoadAutoProxy Request
    End If

    ' Setup proxy
    ' See http://msdn.microsoft.com/en-us/library/windows/desktop/aa384059(v=vs.85).aspx for details
    If Me.proxyServer <> "" Then
        WebHelpers.LogDebug "SetProxy: " & Me.proxyServer, "WebClient.PrepareHttpRequest"
        web_Http.SetProxy web_HttpRequest_ProxySetting_Proxy, Me.proxyServer, Me.proxyBypassList

        If Me.proxyUsername <> "" Then
            WebHelpers.LogDebug "SetProxyCredentials: " & Me.proxyUsername & ", " & WebHelpers.Obfuscate(Me.proxyPassword), "WebClient.PrepareHttpRequest"
            web_Http.SetCredentials Me.proxyUsername, Me.proxyPassword, web_HttpRequest_SetCredentials_ForProxy
        End If
    Else
        ' Attempt to get proxy setup with Proxycfg.exe, otherwise direct
        web_Http.SetProxy web_HttpRequest_ProxySetting_PreConfig
    End If

    ' Setup security
    If Me.insecure Then
        ' - Disable certifcate revocation check
        ' - Ignore all SSL errors
        '   Unknown certification authority (CA) or untrusted root, 0x0100
        '   Wrong usage, 0x0200
        '   Invalid common name (CN), 0x1000
        '   Invalid date or certificate expired, 0x2000
        '   = 0x3300 = 13056
        ' - Enable https-to-http redirects
        web_Http.Option(web_WinHttpRequestOption.web_WinHttpRequestOption_EnableCertificateRevocationCheck) = False
        web_Http.Option(web_WinHttpRequestOption.web_WinHttpRequestOption_SslErrorIgnoreFlags) = 13056
        web_Http.Option(web_WinHttpRequestOption.web_WinHttpRequestOption_EnableHttpsToHttpRedirects) = True
    Else
        ' By default:
        ' - Enable certificate revocation check (especially useful after HeartBleed)
        ' - Ignore no SLL erros
        ' - Disable https-to-http redirects
        web_Http.Option(web_WinHttpRequestOption.web_WinHttpRequestOption_EnableCertificateRevocationCheck) = True
        web_Http.Option(web_WinHttpRequestOption.web_WinHttpRequestOption_SslErrorIgnoreFlags) = 0
        web_Http.Option(web_WinHttpRequestOption.web_WinHttpRequestOption_EnableHttpsToHttpRedirects) = False
    End If

    ' Setup redirects
    web_Http.Option(web_WinHttpRequestOption.web_WinHttpRequestOption_EnableRedirects) = Me.followRedirects

    ' Set headers on http request (after open)
    For Each web_KeyValue In Request.headers
        web_Http.SetRequestHeader web_KeyValue("Key"), web_KeyValue("Value")
    Next web_KeyValue

    For Each web_KeyValue In Request.Cookies
        web_Http.SetRequestHeader "Cookie", web_KeyValue("Key") & "=" & web_KeyValue("Value")
    Next web_KeyValue

    ' Give authenticator opportunity to update Http
    If Not Me.Authenticator Is Nothing Then
        Me.Authenticator.PrepareHttp Me, Request, web_Http
    End If

    ' Log request and return
    WebHelpers.LogRequest Me, Request
    Set PrepareHttpRequest = web_Http
    Exit Function

web_ErrorHandling:

    Set web_Http = Nothing
    Err.Raise 11012 + vbObjectError, "WebClient.PrepareHttpRequest", _
        "An error occurred while preparing http request" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description
End Function

''
' Prepare cURL request for given WebRequest
'
' @internal
' @method PrepareCurlRequest
' @param {WebRequest} Request
' @return {String}
' @throws 11013 / 80042b05 / -2147210491 - Error preparing cURL request
''
Public Function PrepareCurlRequest(Request As WebRequest) As String
    Dim web_Curl As String
    Dim web_KeyValue As Dictionary
    Dim web_CookieString As String

    On Error GoTo web_ErrorHandling

    web_Curl = "curl -i"

    ' Setup authenticator
    web_BeforeExecute Request

    ' Set timeouts
    ' (max time = resolve + sent + receive)
    web_Curl = web_Curl & " --connect-timeout " & Me.TimeoutMs / 1000
    web_Curl = web_Curl & " --max-time " & 3 * Me.TimeoutMs / 1000

    ' Setup proxy
    If Me.proxyServer <> "" Then
        web_Curl = web_Curl & " --proxy " & Me.proxyServer

        If Me.proxyBypassList <> "" Then
            web_Curl = web_Curl & " --noproxy " & Me.proxyBypassList
        End If
        If Me.proxyUsername <> "" Then
            web_Curl = web_Curl & " --proxy-user " & Me.proxyUsername & ":" & Me.proxyPassword
        End If
    End If

    ' Setup security
    If Me.insecure Then
        web_Curl = web_Curl & " --insecure"
    End If

    ' Setup redirects
    If Me.followRedirects Then
        web_Curl = web_Curl & " --location"
    End If

    ' Set headers and cookies
    For Each web_KeyValue In Request.headers
        web_Curl = web_Curl & " -H '" & web_KeyValue("Key") & ": " & web_KeyValue("Value") & "'"
    Next web_KeyValue

    For Each web_KeyValue In Request.Cookies
        web_CookieString = web_CookieString & web_KeyValue("Key") & "=" & web_KeyValue("Value") & ";"
    Next web_KeyValue
    If web_CookieString <> "" Then
        web_Curl = web_Curl & " --cookie '" & web_CookieString & "'"
    End If

    ' Add method  (CHANGED from here to end of function)
    web_Curl = web_Curl & " -X " & WebHelpers.MethodToName(Request.Method)

    ' Add data
    Dim TempFileName As String
    TempFileName = "xlwings-" & CreateGUID() & ".txt"
    Open TempFileName For Output As #1
    Print #1, Request.Body
    Close #1
    web_Curl = web_Curl & " -d @" & TempFileName

    ' Add url
    web_Curl = web_Curl & " '" & Me.GetFullUrl(Request) & "'"

    ' Give authenticator opportunity to update cURL
    If Not Me.Authenticator Is Nothing Then
        Me.Authenticator.PrepareCurl Me, Request, web_Curl
    End If

    ' Log request
    WebHelpers.LogRequest Me, Request

    ' Remove data file and return
    PrepareCurlRequest = web_Curl & "&& rm " & TempFileName
    Exit Function

web_ErrorHandling:

    Err.Raise 11013 + vbObjectError, "WebClient.PrepareCURLRequest", _
        "An error occurred while preparing cURL request" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description
End Function

''
' Clone client
'
' @internal
' @method Clone
' @return {WebClient}
''
Public Function Clone() As WebClient
    Set Clone = New WebClient
    Clone.BaseUrl = Me.BaseUrl
    Clone.proxyServer = Me.proxyServer
    Clone.proxyBypassList = Me.proxyBypassList
    Clone.proxyUsername = Me.proxyUsername
    Clone.proxyPassword = Me.proxyPassword
    Clone.enableAutoProxy = Me.enableAutoProxy
    Clone.TimeoutMs = Me.TimeoutMs
    Clone.insecure = Me.insecure
    Set Clone.Authenticator = Me.Authenticator
End Function

' ============================================= '
' Private Methods
' ============================================= '

Private Sub web_BeforeExecute(web_Request As WebRequest)
    If Not Me.Authenticator Is Nothing Then
        Me.Authenticator.BeforeExecute Me, web_Request
    End If

    ' Preparing request includes adding headers
    ' -> Needs to happen after BeforeExecute in case headers were changed
    web_Request.Prepare
End Sub

Private Sub web_LoadAutoProxy(web_Request As WebRequest)
#If Win32 Or Win64 Then
    On Error GoTo web_ErrorHandling

    Dim web_Parts As Dictionary
    Dim web_Domain As String
    Dim web_ProxyServer As String
    Dim web_ProxyBypassList As String

    Set web_Parts = WebHelpers.GetUrlParts(Me.GetFullUrl(web_Request))
    web_Domain = VBA.IIf(web_Parts("Protocol") <> "", web_Parts("Protocol") & "://", "") & _
        web_Parts("Host") & ":" & web_Parts("Port")

    ' Cache auto-proxy by domain
    If web_Domain <> web_pAutoProxyDomain Then
        ' Cache first to store error as no proxy
        web_pAutoProxyDomain = web_Domain

        WebHelpers.GetAutoProxy web_Domain, web_ProxyServer, web_ProxyBypassList

        WebHelpers.LogDebug "Loaded auto-proxy for " & web_Domain & ":" & vbNewLine & _
            "Server = " & web_ProxyServer & vbNewLine & _
            "Bypass List = " & web_ProxyBypassList

        ' Store proxy server in underlying to avoid turning off auto-proxy
        web_pProxyServer = web_ProxyServer
        Me.proxyBypassList = web_ProxyBypassList
    End If

    Exit Sub

web_ErrorHandling:

    WebHelpers.LogError "An error occurred while loading auto-proxy" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description, _
        "WebClient.LoadAutoProxy", Err.Number
#End If
End Sub

Private Sub Class_Initialize()
    Me.TimeoutMs = web_DefaultTimeoutMs
    Me.enableAutoProxy = False
    Me.insecure = False
    Me.followRedirects = True
End Sub
