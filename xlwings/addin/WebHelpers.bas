Attribute VB_Name = "WebHelpers"
''
' CHANGES:
' ParseIso and ConvertToIso have been changed to not do any timezone conversion
'
' WebHelpers v4.1.6
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Contains general-purpose helpers that are used throughout VBA-Web. Includes:
'
' - Logging
' - Converters and encoding
' - Url handling
' - Object/Dictionary/Collection/Array helpers
' - Request preparation / handling
' - Timing
' - Mac
' - Cryptography
' - Converters (JSON, XML, Url-Encoded)
'
' Errors:
' 11000 - Error during parsing
' 11001 - Error during conversion
' 11002 - No matching converter has been registered
' 11003 - Error while getting url parts
' 11099 - XML format is not currently supported
'
' @module WebHelpers
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' Contents:
' 1. Logging
' 2. Converters and encoding
' 3. Url handling
' 4. Object/Dictionary/Collection/Array helpers
' 5. Request preparation / handling
' 6. Timing
' 7. Mac
' 8. Cryptography
' 9. Converters
' VBA-JSON
' VBA-UTC
' AutoProxy
' --------------------------------------------- '

' Custom formatting uses the standard version of Application.Run,
' which is incompatible with some Office applications (e.g. Word 2011 for Mac)
'
' If you have compilation errors in ParseByFormat or ConvertToFormat,
' you can disable custom formatting by setting the following compiler flag to False
#Const EnableCustomFormatting = True

' === AutoProxy Headers
#If Mac Then
#ElseIf VBA7 Then

Private Declare PtrSafe Sub AutoProxy_CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" _
    (ByVal AutoProxy_lpDest As LongPtr, ByVal AutoProxy_lpSource As LongPtr, ByVal AutoProxy_cbCopy As Long)
Private Declare PtrSafe Function AutoProxy_SysAllocString Lib "oleaut32" Alias "SysAllocString" _
    (ByVal AutoProxy_pwsz As LongPtr) As LongPtr
Private Declare PtrSafe Function AutoProxy_GlobalFree Lib "KERNEL32" Alias "GlobalFree" _
    (ByVal AutoProxy_p As LongPtr) As LongPtr
Private Declare PtrSafe Function AutoProxy_GetIEProxy Lib "WinHTTP.dll" Alias "WinHttpGetIEProxyConfigForCurrentUser" _
    (ByRef AutoProxy_proxyConfig As AUTOPROXY_IE_PROXY_CONFIG) As Long
Private Declare PtrSafe Function AutoProxy_GetProxyForUrl Lib "WinHTTP.dll" Alias "WinHttpGetProxyForUrl" _
    (ByVal AutoProxy_hSession As LongPtr, ByVal AutoProxy_pszUrl As LongPtr, ByRef AutoProxy_pAutoProxyOptions As AUTOPROXY_OPTIONS, ByRef AutoProxy_pProxyInfo As AUTOPROXY_INFO) As Long
Private Declare PtrSafe Function AutoProxy_HttpOpen Lib "WinHTTP.dll" Alias "WinHttpOpen" _
    (ByVal AutoProxy_pszUserAgent As LongPtr, ByVal AutoProxy_dwAccessType As Long, ByVal AutoProxy_pszProxyName As LongPtr, ByVal AutoProxy_pszProxyBypass As LongPtr, ByVal AutoProxy_dwFlags As Long) As LongPtr
Private Declare PtrSafe Function AutoProxy_HttpClose Lib "WinHTTP.dll" Alias "WinHttpCloseHandle" _
    (ByVal AutoProxy_hInternet As LongPtr) As Long

Private Type AUTOPROXY_IE_PROXY_CONFIG
    AutoProxy_fAutoDetect As Long
    AutoProxy_lpszAutoConfigUrl As LongPtr
    AutoProxy_lpszProxy As LongPtr
    AutoProxy_lpszProxyBypass As LongPtr
End Type
Private Type AUTOPROXY_OPTIONS
    AutoProxy_dwFlags As Long
    AutoProxy_dwAutoDetectFlags As Long
    AutoProxy_lpszAutoConfigUrl As LongPtr
    AutoProxy_lpvReserved As LongPtr
    AutoProxy_dwReserved As Long
    AutoProxy_fAutoLogonIfChallenged As Long
End Type
Private Type AUTOPROXY_INFO
    AutoProxy_dwAccessType As Long
    AutoProxy_lpszProxy As LongPtr
    AutoProxy_lpszProxyBypass As LongPtr
End Type

#Else

Private Declare Sub AutoProxy_CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" _
    (ByVal AutoProxy_lpDest As Long, ByVal AutoProxy_lpSource As Long, ByVal AutoProxy_cbCopy As Long)
Private Declare Function AutoProxy_SysAllocString Lib "oleaut32" Alias "SysAllocString" _
    (ByVal AutoProxy_pwsz As Long) As Long
Private Declare Function AutoProxy_GlobalFree Lib "KERNEL32" Alias "GlobalFree" _
    (ByVal AutoProxy_p As Long) As Long
Private Declare Function AutoProxy_GetIEProxy Lib "WinHTTP.dll" Alias "WinHttpGetIEProxyConfigForCurrentUser" _
    (ByRef AutoProxy_proxyConfig As AUTOPROXY_IE_PROXY_CONFIG) As Long
Private Declare Function AutoProxy_GetProxyForUrl Lib "WinHTTP.dll" Alias "WinHttpGetProxyForUrl" _
    (ByVal AutoProxy_hSession As Long, ByVal AutoProxy_pszUrl As Long, ByRef AutoProxy_pAutoProxyOptions As AUTOPROXY_OPTIONS, ByRef AutoProxy_pProxyInfo As AUTOPROXY_INFO) As Long
Private Declare Function AutoProxy_HttpOpen Lib "WinHTTP.dll" Alias "WinHttpOpen" _
    (ByVal AutoProxy_pszUserAgent As Long, ByVal AutoProxy_dwAccessType As Long, ByVal AutoProxy_pszProxyName As Long, ByVal AutoProxy_pszProxyBypass As Long, ByVal AutoProxy_dwFlags As Long) As Long
Private Declare Function AutoProxy_HttpClose Lib "WinHTTP.dll" Alias "WinHttpCloseHandle" _
    (ByVal AutoProxy_hInternet As Long) As Long

Private Type AUTOPROXY_IE_PROXY_CONFIG
    AutoProxy_fAutoDetect As Long
    AutoProxy_lpszAutoConfigUrl As Long
    AutoProxy_lpszProxy As Long
    AutoProxy_lpszProxyBypass As Long
End Type
Private Type AUTOPROXY_OPTIONS
    AutoProxy_dwFlags As Long
    AutoProxy_dwAutoDetectFlags As Long
    AutoProxy_lpszAutoConfigUrl As Long
    AutoProxy_lpvReserved As Long
    AutoProxy_dwReserved As Long
    AutoProxy_fAutoLogonIfChallenged As Long
End Type
Private Type AUTOPROXY_INFO
    AutoProxy_dwAccessType As Long
    AutoProxy_lpszProxy As Long
    AutoProxy_lpszProxyBypass As Long
End Type

#End If

#If Mac Then
#Else
' Constants for dwFlags of AUTOPROXY_OPTIONS
Const AUTOPROXY_AUTO_DETECT = 1
Const AUTOPROXY_CONFIG_URL = 2

' Constants for dwAutoDetectFlags
Const AUTOPROXY_DETECT_TYPE_DHCP = 1
Const AUTOPROXY_DETECT_TYPE_DNS = 2
#End If
' === End AutoProxy

' === VBA-JSON Headers
' === VBA-UTC Headers
#If Mac Then

#If VBA7 Then

' 64-bit Mac (2016)
Private Declare PtrSafe Function utc_popen Lib "/usr/lib/libc.dylib" Alias "popen" _
    (ByVal utc_Command As String, ByVal utc_Mode As String) As LongPtr
Private Declare PtrSafe Function utc_pclose Lib "/usr/lib/libc.dylib" Alias "pclose" _
    (ByVal utc_File As LongPtr) As LongPtr
Private Declare PtrSafe Function utc_fread Lib "/usr/lib/libc.dylib" Alias "fread" _
    (ByVal utc_Buffer As String, ByVal utc_Size As LongPtr, ByVal utc_Number As LongPtr, ByVal utc_File As LongPtr) As LongPtr
Private Declare PtrSafe Function utc_feof Lib "/usr/lib/libc.dylib" Alias "feof" _
    (ByVal utc_File As LongPtr) As LongPtr

#Else

' 32-bit Mac
Private Declare Function utc_popen Lib "libc.dylib" Alias "popen" _
    (ByVal utc_Command As String, ByVal utc_Mode As String) As Long
Private Declare Function utc_pclose Lib "libc.dylib" Alias "pclose" _
    (ByVal utc_File As Long) As Long
Private Declare Function utc_fread Lib "libc.dylib" Alias "fread" _
    (ByVal utc_Buffer As String, ByVal utc_Size As Long, ByVal utc_Number As Long, ByVal utc_File As Long) As Long
Private Declare Function utc_feof Lib "libc.dylib" Alias "feof" _
    (ByVal utc_File As Long) As Long

#End If

#ElseIf VBA7 Then

' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724421.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724949.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms725485.aspx
Private Declare PtrSafe Function utc_GetTimeZoneInformation Lib "KERNEL32" Alias "GetTimeZoneInformation" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
Private Declare PtrSafe Function utc_SystemTimeToTzSpecificLocalTime Lib "KERNEL32" Alias "SystemTimeToTzSpecificLocalTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
Private Declare PtrSafe Function utc_TzSpecificLocalTimeToSystemTime Lib "KERNEL32" Alias "TzSpecificLocalTimeToSystemTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long

#Else

Private Declare Function utc_GetTimeZoneInformation Lib "KERNEL32" Alias "GetTimeZoneInformation" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
Private Declare Function utc_SystemTimeToTzSpecificLocalTime Lib "KERNEL32" Alias "SystemTimeToTzSpecificLocalTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
Private Declare Function utc_TzSpecificLocalTimeToSystemTime Lib "KERNEL32" Alias "TzSpecificLocalTimeToSystemTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long

#End If

#If Mac Then

#If VBA7 Then
Private Type utc_ShellResult
    utc_Output As String
    utc_ExitCode As LongPtr
End Type

#Else

Private Type utc_ShellResult
    utc_Output As String
    utc_ExitCode As Long
End Type

#End If

#Else

Private Type utc_SYSTEMTIME
    utc_wYear As Integer
    utc_wMonth As Integer
    utc_wDayOfWeek As Integer
    utc_wDay As Integer
    utc_wHour As Integer
    utc_wMinute As Integer
    utc_wSecond As Integer
    utc_wMilliseconds As Integer
End Type

Private Type utc_TIME_ZONE_INFORMATION
    utc_Bias As Long
    utc_StandardName(0 To 31) As Integer
    utc_StandardDate As utc_SYSTEMTIME
    utc_StandardBias As Long
    utc_DaylightName(0 To 31) As Integer
    utc_DaylightDate As utc_SYSTEMTIME
    utc_DaylightBias As Long
End Type

#End If
' === End VBA-UTC

Private Type json_Options
    ' VBA only stores 15 significant digits, so any numbers larger than that are truncated
    ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
    ' See: http://support.microsoft.com/kb/269370
    '
    ' By default, VBA-JSON will use String for numbers longer than 15 characters that contain only digits
    ' to override set `JsonConverter.JsonOptions.UseDoubleForLargeNumbers = True`
    UseDoubleForLargeNumbers As Boolean

    ' The JSON standard requires object keys to be quoted (" or '), use this option to allow unquoted keys
    AllowUnquotedKeys As Boolean

    ' The solidus (/) is not required to be escaped, use this option to escape them as \/ in ConvertToJson
    EscapeSolidus As Boolean
End Type
Public JsonOptions As json_Options
' === End VBA-JSON

#If Mac Then
#If VBA7 Then
Private Declare PtrSafe Function web_popen Lib "/usr/lib/libc.dylib" Alias "popen" (ByVal web_Command As String, ByVal web_Mode As String) As LongPtr
Private Declare PtrSafe Function web_pclose Lib "/usr/lib/libc.dylib" Alias "pclose" (ByVal web_File As LongPtr) As LongPtr
Private Declare PtrSafe Function web_fread Lib "/usr/lib/libc.dylib" Alias "fread" (ByVal web_OutStr As String, ByVal web_Size As LongPtr, ByVal web_Items As LongPtr, ByVal web_Stream As LongPtr) As LongPtr
Private Declare PtrSafe Function web_feof Lib "/usr/lib/libc.dylib" Alias "feof" (ByVal web_File As LongPtr) As LongPtr
#Else
Private Declare Function web_popen Lib "libc.dylib" Alias "popen" (ByVal web_Command As String, ByVal web_Mode As String) As Long
Private Declare Function web_pclose Lib "libc.dylib" Alias "pclose" (ByVal web_File As Long) As Long
Private Declare Function web_fread Lib "libc.dylib" Alias "fread" (ByVal web_OutStr As String, ByVal web_Size As Long, ByVal web_Items As Long, ByVal web_Stream As Long) As Long
Private Declare Function web_feof Lib "libc.dylib" Alias "feof" (ByVal web_File As Long) As Long
#End If
#End If

Public Const WebUserAgent As String = "VBA-Web v4.1.6 (https://github.com/VBA-tools/VBA-Web)"

' @internal
Public Type ShellResult
    Output As String
    ExitCode As Long
End Type

Private web_pDocumentHelper As Object
Private web_pElHelper As Object
Private web_pConverters As Dictionary

' --------------------------------------------- '
' Types and Properties
' --------------------------------------------- '

''
' Helper for common http status codes. (Use underlying status code for any codes not listed)
'
' @example
' ```VB.net
' Dim Response As WebResponse
'
' If Response.StatusCode = WebStatusCode.Ok Then
'   ' Ok
' ElseIf Response.StatusCode = 418 Then
'   ' I'm a teapot
' End If
' ```
'
' @property WebStatusCode
' @param Ok `200`
' @param Created `201`
' @param NoContent `204`
' @param NotModified `304`
' @param BadRequest `400`
' @param Unauthorized `401`
' @param Forbidden `403`
' @param NotFound `404`
' @param RequestTimeout `408`
' @param UnsupportedMediaType `415`
' @param InternalServerError `500`
' @param BadGateway `502`
' @param ServiceUnavailable `503`
' @param GatewayTimeout `504`
''
Public Enum WebStatusCode
    Ok = 200
    Created = 201
    NoContent = 204
    NotModified = 304
    BadRequest = 400
    Unauthorized = 401
    Forbidden = 403
    NotFound = 404
    RequestTimeout = 408
    UnsupportedMediaType = 415
    InternalServerError = 500
    BadGateway = 502
    ServiceUnavailable = 503
    GatewayTimeout = 504
End Enum

''
' @property WebMethod
' @param HttpGet
' @param HttpPost
' @param HttpGet
' @param HttpGet
' @param HttpGet
' @default HttpGet
''
Public Enum WebMethod
    HttpGet = 0
    HttpPost = 1
    HttpPut = 2
    HttpDelete = 3
    HttpPatch = 4
    HttpHead = 5
End Enum

''
' @property WebFormat
' @param PlainText
' @param Json
' @param FormUrlEncoded
' @param Xml
' @param Custom
' @default PlainText
''
Public Enum WebFormat
    PlainText = 0
    Json = 1
    FormUrlEncoded = 2
    Xml = 3
    Custom = 9
End Enum

''
' @property UrlEncodingMode
' @param StrictUrlEncoding RFC 3986, ALPHA / DIGIT / "-" / "." / "_" / "~"
' @param FormUrlEncoding ALPHA / DIGIT / "-" / "." / "_" / "*", (space) -> "+", &...; UTF-8 encoding
' @param QueryUrlEncoding Subset of strict and form that should be suitable for non-form-urlencoded query strings
'   ALPHA / DIGIT / "-" / "." / "_"
' @param CookieUrlEncoding strict / "!" / "#" / "$" / "&" / "'" / "(" / ")" / "*" / "+" /
'   "/" / ":" / "<" / "=" / ">" / "?" / "@" / "[" / "]" / "^" / "`" / "{" / "|" / "}"
' @param PathUrlEncoding strict / "!" / "$" / "&" / "'" / "(" / ")" / "*" / "+" / "," / ";" / "=" / ":" / "@"
''
Public Enum UrlEncodingMode
    StrictUrlEncoding
    FormUrlEncoding
    QueryUrlEncoding
    CookieUrlEncoding
    PathUrlEncoding
End Enum

''
' Enable logging of requests and responses and other internal messages from VBA-Web.
' Should be the first step in debugging VBA-Web if something isn't working as expected.
' (Logs display in Immediate Window (`View > Immediate Window` or `ctrl+g`)
'
' @example
' ```VB.net
' Dim Client As New WebClient
' Client.BaseUrl = "https://api.example.com/v1/"
'
' Dim RequestWithTypo As New WebRequest
' RequestWithTypo.Resource = "peeple/{id}"
' RequestWithType.AddUrlSegment "idd", 123
'
' ' Enable logging before the request is executed
' WebHelpers.EnableLogging = True
'
' Dim Response As WebResponse
' Set Response = Client.Execute(Request)
'
' ' Immediate window:
' ' --> Request - (Time)
' ' GET https://api.example.com/v1/peeple/{id}
' ' Headers...
' '
' ' <-- Response - (Time)
' ' 404 ...
' ```
'
' @property EnableLogging
' @type Boolean
' @default False
''
Public EnableLogging As Boolean

''
' Store currently running async requests
'
' @property AsyncRequests
' @type Dictionary
''
Public AsyncRequests As Dictionary

' ============================================= '
' 1. Logging
' ============================================= '

''
' Log message (when logging is enabled with `EnableLogging`)
' with optional location where the message is coming from.
' Useful when writing extensions to VBA-Web (like an `IWebAuthenticator`).
'
' @example
' ```VB.net
' LogDebug "Executing request..."
' ' -> VBA-Web: Executing request...
'
' LogDebug "Executing request...", "Module.Function"
' ' -> Module.Function: Executing request...
' ```
'
' @method LogDebug
' @param {String} Message
' @param {String} [From="VBA-Web"]
''
Public Sub LogDebug(Message As String, Optional From As String = "VBA-Web")
    If EnableLogging Then
        Debug.Print From & ": " & Message
    End If
End Sub

''
' Log warning (even when logging is disabled with `EnableLogging`)
' with optional location where the message is coming from.
' Useful when writing extensions to VBA-Web (like an `IWebAuthenticator`).
'
' @example
' ```VB.net
' WebHelpers.LogWarning "Something could go wrong"
' ' -> WARNING - VBA-Web: Something could go wrong
'
' WebHelpers.LogWarning "Something could go wrong", "Module.Function"
' ' -> WARNING - Module.Function: Something could go wrong
' ```
'
' @method LogWarning
' @param {String} Message
' @param {String} [From="VBA-Web"]
''
Public Sub LogWarning(Message As String, Optional From As String = "VBA-Web")
    Debug.Print "WARNING - " & From & ": " & Message
End Sub

''
' Log error (even when logging is disabled with `EnableLogging`)
' with optional location where the message is coming from and error number.
' Useful when writing extensions to VBA-Web (like an `IWebAuthenticator`).
'
' @example
' ```VB.net
' WebHelpers.LogError "Something went wrong"
' ' -> ERROR - VBA-Web: Something went wrong
'
' WebHelpers.LogError "Something went wrong", "Module.Function"
' ' -> ERROR - Module.Function: Something went wrong
'
' WebHelpers.LogError "Something went wrong", "Module.Function", 100
' ' -> ERROR - Module.Function: 100, Something went wrong
' ```
'
' @method LogError
' @param {String} Message
' @param {String} [From="VBA-Web"]
' @param {Long} [ErrNumber=0]
''
Public Sub LogError(Message As String, Optional From As String = "VBA-Web", Optional ErrNumber As Long = 0)
    Dim web_ErrorValue As String
    If ErrNumber <> 0 Then
        web_ErrorValue = ErrNumber

        If ErrNumber < 0 Then
            web_ErrorValue = web_ErrorValue & " (" & (ErrNumber - vbObjectError) & " / " & VBA.LCase$(VBA.Hex$(ErrNumber)) & ")"
        End If

        web_ErrorValue = web_ErrorValue & ", "
    End If

    Debug.Print "ERROR - " & From & ": " & web_ErrorValue & Message
End Sub

''
' Log details of the request (Url, headers, cookies, body, etc.).
'
' @method LogRequest
' @param {WebClient} Client
' @param {WebRequest} Request
''
Public Sub LogRequest(Client As WebClient, Request As WebRequest)
    If EnableLogging Then
        Debug.Print "--> Request - " & Format(Now, "Long Time")
        Debug.Print MethodToName(Request.Method) & " " & Client.GetFullUrl(Request)

        Dim web_KeyValue As Dictionary
        For Each web_KeyValue In Request.headers
            Debug.Print web_KeyValue("Key") & ": " & web_KeyValue("Value")
        Next web_KeyValue

        For Each web_KeyValue In Request.Cookies
            Debug.Print "Cookie: " & web_KeyValue("Key") & "=" & web_KeyValue("Value")
        Next web_KeyValue

        If Not IsEmpty(Request.Body) Then
            Debug.Print vbNewLine & CStr(Request.Body)
        End If

        Debug.Print
    End If
End Sub

''
' Log details of the response (Status, headers, content, etc.).
'
' @method LogResponse
' @param {WebClient} Client
' @param {WebRequest} Request
' @param {WebResponse} Response
''
Public Sub LogResponse(Client As WebClient, Request As WebRequest, Response As WebResponse)
    If EnableLogging Then
        Dim web_KeyValue As Dictionary

        Debug.Print "<-- Response - " & Format(Now, "Long Time")
        Debug.Print Response.StatusCode & " " & Response.StatusDescription

        For Each web_KeyValue In Response.headers
            Debug.Print web_KeyValue("Key") & ": " & web_KeyValue("Value")
        Next web_KeyValue

        For Each web_KeyValue In Response.Cookies
            Debug.Print "Cookie: " & web_KeyValue("Key") & "=" & web_KeyValue("Value")
        Next web_KeyValue

        Debug.Print vbNewLine & Response.Content & vbNewLine
    End If
End Sub

''
' Obfuscate any secure information before logging.
'
' @example
' ```VB.net
' Dim Password As String
' Password = "Secret"
'
' WebHelpers.LogDebug "Password = " & WebHelpers.Obfuscate(Password)
' -> Password = ******
' ```
'
' @param {String} Secure Message to obfuscate
' @param {String} [Character = *] Character to obfuscate with
' @return {String}
''
Public Function Obfuscate(Secure As String, Optional Character As String = "*") As String
    Obfuscate = VBA.String$(VBA.Len(Secure), Character)
End Function

' ============================================= '
' 2. Converters and encoding
' ============================================= '

'
' Parse JSON value to `Dictionary` if it's an object or `Collection` if it's an array.
'
' @method ParseJson
' @param {String} Json JSON value to parse
' @return {Dictionary|Collection}
'
' (Implemented in VBA-JSON embedded below)

'
' Convert `Dictionary`, `Collection`, or `Array` to JSON string.
'
' @method ConvertToJson
' @param {Dictionary|Collection|Array} Obj
' @return {String}
'
' (Implemented in VBA-JSON embedded below)

''
' Parse Url-Encoded value to `Dictionary`.
'
' @method ParseUrlEncoded
' @param {String} UrlEncoded Url-Encoded value to parse
' @return {Dictionary} Parsed
''
Public Function ParseUrlEncoded(Encoded As String) As Dictionary
    Dim web_Items As Variant
    Dim web_i As Integer
    Dim web_Parts As Variant
    Dim web_Key As String
    Dim web_Value As Variant
    Dim web_Parsed As New Dictionary

    web_Items = VBA.Split(Encoded, "&")
    For web_i = LBound(web_Items) To UBound(web_Items)
        web_Parts = VBA.Split(web_Items(web_i), "=")

        If UBound(web_Parts) - LBound(web_Parts) >= 1 Then
            ' TODO: Handle numbers, arrays, and object better here
            web_Key = UrlDecode(VBA.CStr(web_Parts(LBound(web_Parts))))
            web_Value = UrlDecode(VBA.CStr(web_Parts(LBound(web_Parts) + 1)))

            web_Parsed(web_Key) = web_Value
        End If
    Next web_i

    Set ParseUrlEncoded = web_Parsed
End Function

''
' Convert `Dictionary`/`Collection` to Url-Encoded string.
'
' @method ConvertToUrlEncoded
' @param {Dictionary|Collection|Variant} Obj Value to convert to Url-Encoded string
' @return {String} UrlEncoded string (e.g. a=123&b=456&...)
''
Public Function ConvertToUrlEncoded(Obj As Variant, Optional EncodingMode As UrlEncodingMode = UrlEncodingMode.FormUrlEncoding) As String
    Dim web_Encoded As String

    If TypeOf Obj Is Collection Then
        Dim web_KeyValue As Dictionary

        For Each web_KeyValue In Obj
            If VBA.Len(web_Encoded) > 0 Then: web_Encoded = web_Encoded & "&"
            web_Encoded = web_Encoded & web_GetUrlEncodedKeyValue(web_KeyValue("Key"), web_KeyValue("Value"), EncodingMode)
        Next web_KeyValue
    Else
        Dim web_Key As Variant

        For Each web_Key In Obj.Keys()
            If Len(web_Encoded) > 0 Then: web_Encoded = web_Encoded & "&"
            web_Encoded = web_Encoded & web_GetUrlEncodedKeyValue(web_Key, Obj(web_Key), EncodingMode)
        Next web_Key
    End If

    ConvertToUrlEncoded = web_Encoded
End Function

''
' Parse XML value to `Dictionary`.
'
' _Note_ Currently, XML is not supported in 4.0.0 due to lack of Mac support.
' An updated parser is being created that supports Mac and Windows,
' but in order to avoid future breaking changes, ParseXml and ConvertToXml are not currently implemented.
'
' See https://github.com/VBA-tools/VBA-Web/wiki/XML-Support-in-4.0 for details on how to use XML in Windows in the meantime.
'
' @param {String} Encoded XML value to parse
' @return {Dictionary|Object} Parsed
' @throws 11099 - XML format is not currently supported
''
Public Function ParseXml(Encoded As String) As Object
    Dim web_ErrorMsg As String

    web_ErrorMsg = "XML is not currently supported (An updated parser is being created that supports Mac and Windows)." & vbNewLine & _
        "To use XML parsing for Windows currently, use the instructions found here:" & vbNewLine & _
        vbNewLine & _
        "https://github.com/VBA-tools/VBA-Web/wiki/XML-Support-in-4.0"

    LogError web_ErrorMsg, "WebHelpers.ParseXml", 11099
    Err.Raise 11099, "WebHeleprs.ParseXml", web_ErrorMsg
End Function

''
' Convert `Dictionary` to XML string.
'
' _Note_ Currently, XML is not supported in 4.0.0 due to lack of Mac support.
' An updated parser is being created that supports Mac and Windows,
' but in order to avoid future breaking changes, ParseXml and ConvertToXml are not currently implemented.
'
' See https://github.com/VBA-tools/VBA-Web/wiki/XML-Support-in-4.0 for details on how to use XML in Windows in the meantime.
'
' @param {Dictionary|Variant} XML
' @return {String} XML string
' @throws 11099 / 80042b5b / -2147210405 - XML format is not currently supported
''
Public Function ConvertToXml(Obj As Variant) As String
    Dim web_ErrorMsg As String

    web_ErrorMsg = "XML is not currently supported (An updated parser is being created that supports Mac and Windows)." & vbNewLine & _
        "To use XML parsing for Windows currently, use the instructions found here:" & vbNewLine & _
        vbNewLine & _
        "https://github.com/VBA-tools/VBA-Web/wiki/XML-Support-in-4.0"

    LogError web_ErrorMsg, "WebHelpers.ParseXml", 11099 + vbObjectError
    Err.Raise 11099 + vbObjectError, "WebHeleprs.ParseXml", web_ErrorMsg
End Function

''
' Helper for parsing value to given `WebFormat` or custom format.
' Returns `Dictionary` or `Collection` based on given `Value`.
'
' @method ParseByFormat
' @param {String} Value Value to parse
' @param {WebFormat} Format
' @param {String} [CustomFormat=""] Name of registered custom converter
' @param {Variant} [Bytes] Bytes for custom convert (if `ParseType = "Binary"`)
' @return {Dictionary|Collection|Object}
' @throws 11000 - Error during parsing
''
Public Function ParseByFormat(Value As String, Format As WebFormat, _
    Optional CustomFormat As String = "", Optional Bytes As Variant) As Object

    On Error GoTo web_ErrorHandling

    ' Don't attempt to parse blank values
    If Value = "" And CustomFormat = "" Then
        Exit Function
    End If

    Select Case Format
    Case WebFormat.Json
        Set ParseByFormat = ParseJson(Value)
    Case WebFormat.FormUrlEncoded
        Set ParseByFormat = ParseUrlEncoded(Value)
    Case WebFormat.Xml
        Set ParseByFormat = ParseXml(Value)
    Case WebFormat.Custom
#If EnableCustomFormatting Then
        Dim web_Converter As Dictionary
        Dim web_Callback As String

        Set web_Converter = web_GetConverter(CustomFormat)
        web_Callback = web_Converter("ParseCallback")

        If web_Converter.Exists("Instance") Then
            Dim web_Instance As Object
            Set web_Instance = web_Converter("Instance")

            If web_Converter("ParseType") = "Binary" Then
                Set ParseByFormat = VBA.CallByName(web_Instance, web_Callback, VBA.vbMethod, Bytes)
            Else
                Set ParseByFormat = VBA.CallByName(web_Instance, web_Callback, VBA.vbMethod, Value)
            End If
        Else
            If web_Converter("ParseType") = "Binary" Then
                Set ParseByFormat = Application.Run(web_Callback, Bytes)
            Else
                Set ParseByFormat = Application.Run(web_Callback, Value)
            End If
        End If
#Else
    LogWarning "Custom formatting is disabled. To use WebFormat.Custom, enable custom formatting with the EnableCustomFormatting flag in WebHelpers"
#End If
    End Select
    Exit Function

web_ErrorHandling:

    Dim web_ErrorDescription As String
    web_ErrorDescription = "An error occurred during parsing" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description

    LogError web_ErrorDescription, "WebHelpers.ParseByFormat", 11000
    Err.Raise 11000, "WebHelpers.ParseByFormat", web_ErrorDescription
End Function

''
' Helper for converting value to given `WebFormat` or custom format.
'
' _Note_ Only some converters handle `Collection` or `Array`.
'
' @method ConvertToFormat
' @param {Dictionary|Collection|Variant} Obj
' @param {WebFormat} Format
' @param {String} [CustomFormat] Name of registered custom converter
' @return {Variant}
' @throws 11001 - Error during conversion
''
Public Function ConvertToFormat(Obj As Variant, Format As WebFormat, Optional CustomFormat As String = "") As Variant
    On Error GoTo web_ErrorHandling

    Select Case Format
    Case WebFormat.Json
        ConvertToFormat = ConvertToJson(Obj)
    Case WebFormat.FormUrlEncoded
        ConvertToFormat = ConvertToUrlEncoded(Obj)
    Case WebFormat.Xml
        ConvertToFormat = ConvertToXml(Obj)
    Case WebFormat.Custom
#If EnableCustomFormatting Then
        Dim web_Converter As Dictionary
        Dim web_Callback As String

        Set web_Converter = web_GetConverter(CustomFormat)
        web_Callback = web_Converter("ConvertCallback")

        If web_Converter.Exists("Instance") Then
            Dim web_Instance As Object
            Set web_Instance = web_Converter("Instance")
            ConvertToFormat = VBA.CallByName(web_Instance, web_Callback, VBA.vbMethod, Obj)
        Else
            ConvertToFormat = Application.Run(web_Callback, Obj)
        End If
#Else
    LogWarning "Custom formatting is disabled. To use WebFormat.Custom, enable custom formatting with the EnableCustomFormatting flag in WebHelpers"
#End If
    Case Else
        If VBA.VarType(Obj) = vbString Then
            ' Plain text
            ConvertToFormat = Obj
        End If
    End Select
    Exit Function

web_ErrorHandling:

    Dim web_ErrorDescription As String
    web_ErrorDescription = "An error occurred during conversion" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description

    LogError web_ErrorDescription, "WebHelpers.ConvertToFormat", 11001
    Err.Raise 11001, "WebHelpers.ConvertToFormat", web_ErrorDescription
End Function

''
' Encode string for URLs
'
' See https://github.com/VBA-tools/VBA-Web/wiki/Url-Encoding for details
'
' References:
' - RFC 3986, https://tools.ietf.org/html/rfc3986
' - form-urlencoded encoding algorithm,
'   https://www.w3.org/TR/html5/forms.html#application/x-www-form-urlencoded-encoding-algorithm
' - RFC 6265 (Cookies), https://tools.ietf.org/html/rfc6265
'   Note: "%" is allowed in spec, but is currently excluded due to parsing issues
'
' @method UrlEncode
' @param {Variant} Text Text to encode
' @param {Boolean} [SpaceAsPlus = False] `%20` if `False` / `+` if `True`
'   DEPRECATED Use EncodingMode:=FormUrlEncoding
' @param {Boolean} [EncodeUnsafe = True] Encode characters that could be misunderstood within URLs.
'   (``SPACE, ", <, >, #, %, {, }, |, \, ^, ~, `, [, ]``)
'   DEPRECATED This was based on an outdated URI spec and has since been removed.
'     EncodingMode:=CookieUrlEncoding is the closest approximation of this behavior
' @param {UrlEncodingMode} [EncodingMode = StrictUrlEncoding]
' @return {String} Encoded string
''
Public Function UrlEncode(Text As Variant, _
    Optional SpaceAsPlus As Boolean = False, Optional EncodeUnsafe As Boolean = True, _
    Optional EncodingMode As UrlEncodingMode = UrlEncodingMode.StrictUrlEncoding) As String

    If SpaceAsPlus = True Then
        LogWarning "SpaceAsPlus is deprecated and will be removed in VBA-Web v5. " & _
            "Use EncodingMode:=FormUrlEncoding instead", "WebHelpers.UrlEncode"
    End If
    If EncodeUnsafe = False Then
        LogWarning "EncodeUnsafe has been removed as it was based on an outdated url encoding specification. " & _
            "Use EncodingMode:=CookieUrlEncoding to approximate this behavior", "WebHelpers.UrlEncode"
    End If

    Dim web_UrlVal As String
    Dim web_StringLen As Long

    web_UrlVal = VBA.CStr(Text)
    web_StringLen = VBA.Len(web_UrlVal)

    If web_StringLen > 0 Then
        Dim web_Result() As String
        Dim web_i As Long
        Dim web_CharCode As Integer
        Dim web_Char As String
        Dim web_Space As String
        ReDim web_Result(web_StringLen)

        ' StrictUrlEncoding - ALPHA / DIGIT / "-" / "." / "_" / "~"
        ' FormUrlEncoding   - ALPHA / DIGIT / "-" / "." / "_" / "*" / (space) -> "+"
        ' QueryUrlEncoding  - ALPHA / DIGIT / "-" / "." / "_"
        ' CookieUrlEncoding - strict / "!" / "#" / "$" / "&" / "'" / "(" / ")" / "*" / "+" /
        '   "/" / ":" / "<" / "=" / ">" / "?" / "@" / "[" / "]" / "^" / "`" / "{" / "|" / "}"
        ' PathUrlEncoding   - strict / "!" / "$" / "&" / "'" / "(" / ")" / "*" / "+" / "," / ";" / "=" / ":" / "@"

        ' Set space value
        If SpaceAsPlus Or EncodingMode = UrlEncodingMode.FormUrlEncoding Then
            web_Space = "+"
        Else
            web_Space = "%20"
        End If

        ' Loop through string characters
        For web_i = 1 To web_StringLen
            ' Get character and ascii code
            web_Char = VBA.Mid$(web_UrlVal, web_i, 1)
            web_CharCode = VBA.Asc(web_Char)

            Select Case web_CharCode
                Case 65 To 90, 97 To 122
                    ' ALPHA
                    web_Result(web_i) = web_Char
                Case 48 To 57
                    ' DIGIT
                    web_Result(web_i) = web_Char
                Case 45, 46, 95
                    ' "-" / "." / "_"
                    web_Result(web_i) = web_Char

                Case 32
                    ' (space)
                    ' FormUrlEncoding -> "+"
                    ' Else -> "%20"
                    web_Result(web_i) = web_Space

                Case 33, 36, 38, 39, 40, 41, 43, 58, 61, 64
                    ' "!" / "$" / "&" / "'" / "(" / ")" / "+" / ":" / "=" / "@"
                    ' PathUrlEncoding, CookieUrlEncoding -> Unencoded
                    ' Else -> Percent-encoded
                    If EncodingMode = UrlEncodingMode.PathUrlEncoding Or EncodingMode = UrlEncodingMode.CookieUrlEncoding Then
                        web_Result(web_i) = web_Char
                    Else
                        web_Result(web_i) = "%" & VBA.Hex(web_CharCode)
                    End If

                Case 35, 45, 46, 47, 60, 62, 63, 91, 93, 94, 95, 96, 123, 124, 125
                    ' "#" / "-" / "." / "/" / "<" / ">" / "?" / "[" / "]" / "^" / "_" / "`" / "{" / "|" / "}"
                    ' CookieUrlEncoding -> Unencoded
                    ' Else -> Percent-encoded
                    If EncodingMode = UrlEncodingMode.CookieUrlEncoding Then
                        web_Result(web_i) = web_Char
                    Else
                        web_Result(web_i) = "%" & VBA.Hex(web_CharCode)
                    End If

                Case 42
                    ' "*"
                    ' FormUrlEncoding, PathUrlEncoding, CookieUrlEncoding -> "*"
                    ' Else -> "%2A"
                    If EncodingMode = UrlEncodingMode.FormUrlEncoding _
                        Or EncodingMode = UrlEncodingMode.PathUrlEncoding _
                        Or EncodingMode = UrlEncodingMode.CookieUrlEncoding Then

                        web_Result(web_i) = web_Char
                    Else
                        web_Result(web_i) = "%" & VBA.Hex(web_CharCode)
                    End If

                Case 44, 59
                    ' "," / ";"
                    ' PathUrlEncoding -> Unencoded
                    ' Else -> Percent-encoded
                    If EncodingMode = UrlEncodingMode.PathUrlEncoding Then
                        web_Result(web_i) = web_Char
                    Else
                        web_Result(web_i) = "%" & VBA.Hex(web_CharCode)
                    End If

                Case 126
                    ' "~"
                    ' FormUrlEncoding, QueryUrlEncoding -> "%7E"
                    ' Else -> "~"
                    If EncodingMode = UrlEncodingMode.FormUrlEncoding Or EncodingMode = UrlEncodingMode.QueryUrlEncoding Then
                        web_Result(web_i) = "%7E"
                    Else
                        web_Result(web_i) = web_Char
                    End If

                Case 0 To 15
                    web_Result(web_i) = "%0" & VBA.Hex(web_CharCode)
                Case Else
                    web_Result(web_i) = "%" & VBA.Hex(web_CharCode)

                ' TODO For non-ASCII characters,
                '
                ' FormUrlEncoded:
                '
                ' Replace the character by a string consisting of a U+0026 AMPERSAND character (&), a "#" (U+0023) character,
                ' one or more ASCII digits representing the Unicode code point of the character in base ten, and finally a ";" (U+003B) character.
                '
                ' Else:
                '
                ' Encode to sequence of 2 or 3 bytes in UTF-8, then percent encode
                ' Reference Implementation: https://www.w3.org/International/URLUTF8Encoder.java
            End Select
        Next web_i
        UrlEncode = VBA.Join$(web_Result, "")
    End If
End Function

''
' Decode Url-encoded string.
'
' @method UrlDecode
' @param {String} Encoded Text to decode
' @param {Boolean} [PlusAsSpace = True] Decode plus as space
'   DEPRECATED Use EncodingMode:=FormUrlEncoding Or QueryUrlEncoding
' @param {UrlEncodingMode} [EncodingMode = StrictUrlEncoding]
' @return {String} Decoded string
''
Public Function UrlDecode(Encoded As String, _
    Optional PlusAsSpace As Boolean = True, _
    Optional EncodingMode As UrlEncodingMode = UrlEncodingMode.StrictUrlEncoding) As String

    Dim web_StringLen As Long
    web_StringLen = VBA.Len(Encoded)

    If web_StringLen > 0 Then
        Dim web_i As Long
        Dim web_Result As String
        Dim web_Temp As String

        For web_i = 1 To web_StringLen
            web_Temp = VBA.Mid$(Encoded, web_i, 1)

            If web_Temp = "+" And _
                (PlusAsSpace _
                 Or EncodingMode = UrlEncodingMode.FormUrlEncoding _
                 Or EncodingMode = UrlEncodingMode.QueryUrlEncoding) Then

                web_Temp = " "
            ElseIf web_Temp = "%" And web_StringLen >= web_i + 2 Then
                web_Temp = VBA.Mid$(Encoded, web_i + 1, 2)
                web_Temp = VBA.Chr(VBA.CInt("&H" & web_Temp))

                web_i = web_i + 2
            End If

            ' TODO Handle non-ASCII characters

            web_Result = web_Result & web_Temp
        Next web_i

        UrlDecode = web_Result
    End If
End Function

''
' Base64-encode text.
'
' @param {Variant} Text Text to encode
' @return {String} Encoded string
''
Public Function Base64Encode(Text As String) As String
#If Mac Then
    Dim web_Command As String
    web_Command = "printf " & PrepareTextForPrintf(Text) & " | openssl base64"
    Base64Encode = ExecuteInShell(web_Command).Output
#Else
    Dim web_Bytes() As Byte

    web_Bytes = VBA.StrConv(Text, vbFromUnicode)
    Base64Encode = web_AnsiBytesToBase64(web_Bytes)
#End If

    Base64Encode = VBA.Replace$(Base64Encode, vbLf, "")
End Function

''
' Decode Base64-encoded text
'
' @param {Variant} Encoded Text to decode
' @return {String} Decoded string
''
Public Function Base64Decode(Encoded As Variant) As String
    ' Add trailing padding, if necessary
    If (VBA.Len(Encoded) Mod 4 > 0) Then
        Encoded = Encoded & VBA.Left("====", 4 - (VBA.Len(Encoded) Mod 4))
    End If

#If Mac Then
    Dim web_Command As String
    web_Command = "echo " & PrepareTextForShell(Encoded) & " | openssl base64 -d"
    Base64Decode = ExecuteInShell(web_Command).Output
#Else
    Dim web_XmlObj As Object
    Dim web_Node As Object

    Set web_XmlObj = CreateObject("MSXML2.DOMDocument")
    Set web_Node = web_XmlObj.createElement("b64")

    web_Node.DataType = "bin.base64"
    web_Node.Text = Encoded
    Base64Decode = VBA.StrConv(web_Node.nodeTypedValue, vbUnicode)

    Set web_Node = Nothing
    Set web_XmlObj = Nothing
#End If
End Function

''
' Register custom converter for converting request `Body` and response `Content`.
' If the `ConvertCallback` or `ParseCallback` are object methods,
' pass in an object instance.
' If the `ParseCallback` needs the raw binary response value (e.g. file download),
' set `ParseType = "Binary"`, otherwise `"String"` is used.
'
' - `ConvertCallback` signature: `Function ...(Value As Variant) As String`
' - `ParseCallback` signature: `Function ...(Value As String) As Object`
'
' @example
' ```VB.net
' ' 1. Use global module functions for Convert and Parse
' ' ---
' ' Module: CSVConverter
' Function ParseCSV(Value As String) As Object
'   ' ...
' End Function
' Function ConvertToCSV(Value As Variant) As String
'   ' ...
' End Function
'
' WebHelpers.RegisterConverter "csv", "text/csv", _
'   "CSVConverter.ConvertToCSV", "CSVConverter.ParseCSV"
'
' ' 2. Use object instance functions for Convert and Parse
' ' ---
' ' Object: CSVConverterClass
' ' same as above...
'
' Dim Converter As New CSVConverterClass
' WebHelpers.RegisterConverter "csv", "text/csv", _
'   "ConvertToCSV", "ParseCSV", Instance:=Converter
'
' ' 3. Pass raw binary value to ParseCallback
' ' ---
' ' Module: ImageConverter
' Function ParseImage(Bytes As Variant) As Object
'   ' ...
' End Function
' Function ConvertToImage(Value As Variant) As String
'   ' ...
' End Function
'
' WebHelpers.RegisterConverter "image", "image/jpeg", _
'   "ImageConverter.ConvertToImage", "ImageConverter.ParseImage", _
'   ParseType:="Binary"
' ```
'
' @method RegisterConverter
' @param {String} Name
'   Name of converter for use with `CustomRequestFormat` or `CustomResponseFormat`
' @param {String} MediaType
'   Media type to use for `Content-Type` and `Accept` headers
' @param {String} ConvertCallback Global or object function name for converting
' @param {String} ParseCallback Global or object function name for parsing
' @param {Object} [Instance]
'   Use instance methods for `ConvertCallback` and `ParseCallback`
' @param {String} [ParseType="String"]
'   "String"` (default) or `"Binary"` to pass raw binary response to `ParseCallback`
''
Public Sub RegisterConverter( _
    Name As String, MediaType As String, ConvertCallback As String, ParseCallback As String, _
    Optional Instance As Object, Optional ParseType As String = "String")

    Dim web_Converter As New Dictionary
    web_Converter("MediaType") = MediaType
    web_Converter("ConvertCallback") = ConvertCallback
    web_Converter("ParseCallback") = ParseCallback
    web_Converter("ParseType") = ParseType

    If Not Instance Is Nothing Then
        Set web_Converter("Instance") = Instance
    End If

    If web_pConverters Is Nothing Then: Set web_pConverters = New Dictionary
    Set web_pConverters(Name) = web_Converter
End Sub

' Helper for getting custom converter
' @throws 11002 - No matching converter has been registered
Private Function web_GetConverter(web_CustomFormat As String) As Dictionary
    If web_pConverters.Exists(web_CustomFormat) Then
        Set web_GetConverter = web_pConverters(web_CustomFormat)
    Else
        LogError "No matching converter has been registered for custom format: " & web_CustomFormat, _
            "WebHelpers.web_GetConverter", 11002
        Err.Raise 11002, "WebHelpers.web_GetConverter", _
            "No matching converter has been registered for custom format: " & web_CustomFormat
    End If
End Function

' ============================================= '
' 3. Url handling
' ============================================= '

''
' Join Url with /
'
' @example
' ```VB.net
' Debug.Print WebHelpers.JoinUrl("a/", "/b")
' Debug.Print WebHelpers.JoinUrl("a", "b")
' Debug.Print WebHelpers.JoinUrl("a/", "b")
' Debug.Print WebHelpers.JoinUrl("a", "/b")
' -> a/b
' ```
'
' @param {String} LeftSide
' @param {String} RightSide
' @return {String} Joined url
''
Public Function JoinUrl(LeftSide As String, RightSide As String) As String
    If Left(RightSide, 1) = "/" Then
        RightSide = Right(RightSide, Len(RightSide) - 1)
    End If
    If Right(LeftSide, 1) = "/" Then
        LeftSide = Left(LeftSide, Len(LeftSide) - 1)
    End If

    If LeftSide <> "" And RightSide <> "" Then
        JoinUrl = LeftSide & "/" & RightSide
    Else
        JoinUrl = LeftSide & RightSide
    End If
End Function

''
' Get relevant parts of the given url.
' Returns `Protocol`, `Host`, `Port`, `Path`, `Querystring`, and `Hash`
'
' @example
' ```VB.net
' WebHelpers.GetUrlParts "https://www.google.com/a/b/c.html?a=1&b=2#hash"
' ' -> Protocol = "https"
' '    Host = "www.google.com"
' '    Port = "443"
' '    Path = "/a/b/c.html"
' '    Querystring = "a=1&b=2"
' '    Hash = "hash"
'
' WebHelpers.GetUrlParts "localhost:3000/a/b/c"
' ' -> Protocol = ""
' '    Host = "localhost"
' '    Port = "3000"
' '    Path = "/a/b/c"
' '    Querystring = ""
' '    Hash = ""
' ```
'
' @method GetUrlParts
' @param {String} Url
' @return {Dictionary} Parts of url
'   Protocol, Host, Port, Path, Querystring, Hash
' @throws 11003 - Error while getting url parts
''
Public Function GetUrlParts(url As String) As Dictionary
    Dim web_Parts As New Dictionary

    On Error GoTo web_ErrorHandling

#If Mac Then
    ' Run perl script to parse url

    Dim web_AddedProtocol As Boolean
    Dim web_Command As String
    Dim web_Results As Variant
    Dim web_ResultPart As Variant
    Dim web_EqualsIndex As Long
    Dim web_Key As String
    Dim web_Value As String

    ' Add Protocol if missing
    If InStr(1, url, "://") <= 0 Then
        web_AddedProtocol = True
        If InStr(1, url, "//") = 1 Then
            url = "http" & url
        Else
            url = "http://" & url
        End If
    End If

    web_Command = "perl -e '{use URI::URL;" & vbNewLine & _
        "$url = new URI::URL """ & url & """;" & vbNewLine & _
        "print ""Protocol="" . $url->scheme;" & vbNewLine & _
        "print "" | Host="" . $url->host;" & vbNewLine & _
        "print "" | Port="" . $url->port;" & vbNewLine & _
        "print "" | FullPath="" . $url->full_path;" & vbNewLine & _
        "print "" | Hash="" . $url->frag;" & vbNewLine & _
    "}'"

    web_Results = Split(ExecuteInShell(web_Command).Output, " | ")
    For Each web_ResultPart In web_Results
        web_EqualsIndex = InStr(1, web_ResultPart, "=")
        web_Key = Trim(VBA.Mid$(web_ResultPart, 1, web_EqualsIndex - 1))
        web_Value = Trim(VBA.Mid$(web_ResultPart, web_EqualsIndex + 1))

        If web_Key = "FullPath" Then
            ' For properly escaped path and querystring, need to use full_path
            ' But, need to split FullPath into Path...?Querystring
            Dim QueryIndex As Integer

            QueryIndex = InStr(1, web_Value, "?")
            If QueryIndex > 0 Then
                web_Parts.Add "Path", Mid$(web_Value, 1, QueryIndex - 1)
                web_Parts.Add "Querystring", Mid$(web_Value, QueryIndex + 1)
            Else
                web_Parts.Add "Path", web_Value
                web_Parts.Add "Querystring", ""
            End If
        Else
            web_Parts.Add web_Key, web_Value
        End If
    Next web_ResultPart

    If web_AddedProtocol And web_Parts.Exists("Protocol") Then
        web_Parts("Protocol") = ""
    End If
#Else
    ' Create document/element is expensive, cache after creation
    If web_pDocumentHelper Is Nothing Or web_pElHelper Is Nothing Then
        Set web_pDocumentHelper = CreateObject("htmlfile")
        Set web_pElHelper = web_pDocumentHelper.createElement("a")
    End If

    web_pElHelper.href = url
    web_Parts.Add "Protocol", Replace(web_pElHelper.Protocol, ":", "", Count:=1)
    web_Parts.Add "Host", web_pElHelper.hostname
    web_Parts.Add "Port", web_pElHelper.port
    web_Parts.Add "Path", web_pElHelper.pathname
    web_Parts.Add "Querystring", Replace(web_pElHelper.Search, "?", "", Count:=1)
    web_Parts.Add "Hash", Replace(web_pElHelper.Hash, "#", "", Count:=1)
#End If

    If web_Parts("Protocol") = "localhost" Then
        ' localhost:port/... was passed in without protocol
        Dim PathParts As Variant
        PathParts = Split(web_Parts("Path"), "/")

        web_Parts("Port") = PathParts(0)
        web_Parts("Protocol") = ""
        web_Parts("Host") = "localhost"
        web_Parts("Path") = Replace(web_Parts("Path"), web_Parts("Port"), "", Count:=1)
    End If
    If Left(web_Parts("Path"), 1) <> "/" Then
        web_Parts("Path") = "/" & web_Parts("Path")
    End If

    Set GetUrlParts = web_Parts
    Exit Function

web_ErrorHandling:

    Dim web_ErrorDescription As String
    web_ErrorDescription = "An error occurred while getting url parts" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description

    LogError web_ErrorDescription, "WebHelpers.GetUrlParts", 11003
    Err.Raise 11003, "WebHelpers.GetUrlParts", web_ErrorDescription
End Function

' ============================================= '
' 4. Object/Dictionary/Collection/Array helpers
' ============================================= '

''
' Create a cloned copy of the `Dictionary`.
' This is not a deep copy, so children objects are copied by reference.
'
' @method CloneDictionary
' @param {Dictionary} Original
' @return {Dictionary} Clone
''
Public Function CloneDictionary(Original As Dictionary) As Dictionary
    Dim web_Key As Variant

    Set CloneDictionary = New Dictionary
    For Each web_Key In Original.Keys
        CloneDictionary.Add VBA.CStr(web_Key), Original(web_Key)
    Next web_Key
End Function

''
' Create a cloned copy of the `Collection`.
' This is not a deep copy, so children objects are copied by reference.
'
' _Note_ Keys are not transferred to clone
'
' @method CloneCollection
' @param {Collection} Original
' @return {Collection} Clone
''
Public Function CloneCollection(Original As Collection) As Collection
    Dim web_Item As Variant

    Set CloneCollection = New Collection
    For Each web_Item In Original
        CloneCollection.Add web_Item
    Next web_Item
End Function

''
' Helper for creating `Key-Value` pair with `Dictionary`.
' Used in `WebRequest`/`WebResponse` `Cookies`, `Headers`, and `QuerystringParams`
'
' @example
' ```VB.net
' WebHelpers.CreateKeyValue "abc", 123
' ' -> {"Key": "abc", "Value": 123}
' ```
'
' @method CreateKeyValue
' @param {String} Key
' @param {Variant} Value
' @return {Dictionary}
''
Public Function CreateKeyValue(Key As String, Value As Variant) As Dictionary
    Dim web_KeyValue As New Dictionary

    web_KeyValue("Key") = Key
    web_KeyValue("Value") = Value
    Set CreateKeyValue = web_KeyValue
End Function

''
' Search a `Collection` of `KeyValue` and retrieve the value for the given key.
'
' @example
' ```VB.net
' Dim KeyValues As New Collection
' KeyValues.Add WebHelpers.CreateKeyValue("abc", 123)
'
' WebHelpers.FindInKeyValues KeyValues, "abc"
' ' -> 123
'
' WebHelpers.FindInKeyValues KeyValues, "unknown"
' ' -> Empty
' ```
'
' @method FindInKeyValues
' @param {Collection} KeyValues
' @param {Variant} Key to find
' @return {Variant}
''
Public Function FindInKeyValues(KeyValues As Collection, Key As Variant) As Variant
    Dim web_KeyValue As Dictionary

    For Each web_KeyValue In KeyValues
        If web_KeyValue("Key") = Key Then
            FindInKeyValues = web_KeyValue("Value")
            Exit Function
        End If
    Next web_KeyValue
End Function

''
' Helper for adding/replacing `KeyValue` in `Collection` of `KeyValue`
' - Add if key not found
' - Replace if key is found
'
' @example
' ```VB.net
' Dim KeyValues As New Collection
' KeyValues.Add WebHelpers.CreateKeyValue("a", 123)
' KeyValues.Add WebHelpers.CreateKeyValue("b", 456)
' KeyValues.Add WebHelpers.CreateKeyValue("c", 789)
'
' WebHelpers.AddOrReplaceInKeyValues KeyValues, "b", "abc"
' WebHelpers.AddOrReplaceInKeyValues KeyValues, "d", "def"
'
' ' -> [
' '      {"Key":"a","Value":123},
' '      {"Key":"b","Value":"abc"},
' '      {"Key":"c","Value":789},
' '      {"Key":"d","Value":"def"}
' '    ]
' ```
'
' @method AddOrReplaceInKeyValues
' @param {Collection} KeyValues
' @param {Variant} Key
' @param {Variant} Value
' @return {Variant}
''
Public Sub AddOrReplaceInKeyValues(KeyValues As Collection, Key As Variant, Value As Variant)
    Dim web_KeyValue As Dictionary
    Dim web_Index As Long
    Dim web_NewKeyValue As Dictionary

    Set web_NewKeyValue = CreateKeyValue(CStr(Key), Value)

    web_Index = 1
    For Each web_KeyValue In KeyValues
        If web_KeyValue("Key") = Key Then
            ' Replace existing
            KeyValues.Remove web_Index

            If KeyValues.Count = 0 Then
                KeyValues.Add web_NewKeyValue
            ElseIf web_Index > KeyValues.Count Then
                KeyValues.Add web_NewKeyValue, After:=web_Index - 1
            Else
                KeyValues.Add web_NewKeyValue, Before:=web_Index
            End If
            Exit Sub
        End If

        web_Index = web_Index + 1
    Next web_KeyValue

    ' Add
    KeyValues.Add web_NewKeyValue
End Sub

' ============================================= '
' 5. Request preparation / handling
' ============================================= '

''
' Get the media-type for the given format / custom format.
'
' @method FormatToMediaType
' @param {WebFormat} Format
' @param {String} [CustomFormat] Needed if `Format = WebFormat.Custom`
' @return {String}
''
Public Function FormatToMediaType(Format As WebFormat, Optional CustomFormat As String) As String
    Select Case Format
    Case WebFormat.FormUrlEncoded
        FormatToMediaType = "application/x-www-form-urlencoded;charset=UTF-8"
    Case WebFormat.Json
        FormatToMediaType = "application/json"
    Case WebFormat.Xml
        FormatToMediaType = "application/xml"
    Case WebFormat.Custom
        FormatToMediaType = web_GetConverter(CustomFormat)("MediaType")
    Case Else
        FormatToMediaType = "text/plain"
    End Select
End Function

''
' Get the method name for the given `WebMethod`
'
' @example
' ```VB.net
' WebHelpers.MethodToName WebMethod.HttpPost
' ' -> "POST"
' ```
'
' @method MethodToName
' @param {WebMethod} Method
' @return {String}
''
Public Function MethodToName(Method As WebMethod) As String
    Select Case Method
    Case WebMethod.HttpDelete
        MethodToName = "DELETE"
    Case WebMethod.HttpPut
        MethodToName = "PUT"
    Case WebMethod.HttpPatch
        MethodToName = "PATCH"
    Case WebMethod.HttpPost
        MethodToName = "POST"
    Case WebMethod.HttpGet
        MethodToName = "GET"
    Case WebMethod.HttpHead
        MethodToName = "HEAD"
    End Select
End Function

' ============================================= '
' 6. Timing
' ============================================= '

''
' Handle timeout timers expiring
'
' @internal
' @method OnTimeoutTimerExpired
' @param {String} RequestId
''
Public Sub OnTimeoutTimerExpired(web_RequestId As String)
    If Not AsyncRequests Is Nothing Then
        If AsyncRequests.Exists(web_RequestId) Then
            Dim web_AsyncWrapper As Object
            Set web_AsyncWrapper = AsyncRequests(web_RequestId)
            web_AsyncWrapper.TimedOut
        End If
    End If
End Sub

' ============================================= '
' 7. Mac
' ============================================= '

''
' Execute the given command
'
' @internal
' @method ExecuteInShell
' @param {String} Command
' @return {ShellResult}
''
Public Function ExecuteInShell(web_Command As String) As ShellResult
#If Mac Then
#If VBA7 Then
    Dim web_File As LongPtr
#Else
    Dim web_File As Long
#End If

    Dim web_Chunk As String
    Dim web_Read As Long

    On Error GoTo web_Cleanup

    web_File = web_popen(web_Command, "r")

    If web_File = 0 Then
        ' TODO Investigate why this could happen and what should be done if it happens
        Exit Function
    End If

    Do While web_feof(web_File) = 0
        web_Chunk = VBA.Space$(50)
        web_Read = CLng(web_fread(web_Chunk, 1, Len(web_Chunk) - 1, web_File))
        If web_Read > 0 Then
            web_Chunk = VBA.Left$(web_Chunk, web_Read)
            ExecuteInShell.Output = ExecuteInShell.Output & web_Chunk
        End If
    Loop

web_Cleanup:

    ExecuteInShell.ExitCode = CLng(web_pclose(web_File))
#End If
End Function

''
' Prepare text for shell
' - Wrap in "..."
' - Replace ! with '!' (reserved in bash)
' - Escape \, `, $, %, and "
'
' @internal
' @method PrepareTextForShell
' @param {String} Text
' @return {String}
''
Public Function PrepareTextForShell(ByVal web_Text As String) As String
    ' Escape special characters (except for !)
    web_Text = VBA.Replace(web_Text, "\", "\\")
    web_Text = VBA.Replace(web_Text, "`", "\`")
    web_Text = VBA.Replace(web_Text, "$", "\$")
    web_Text = VBA.Replace(web_Text, "%", "\%")
    web_Text = VBA.Replace(web_Text, """", "\""")

    ' Wrap in quotes
    web_Text = """" & web_Text & """"

    ' Escape !
    web_Text = VBA.Replace(web_Text, "!", """'!'""")

    ' Guard for ! at beginning or end (""'!'"..." or "..."'!'"" -> '!'"..." or "..."'!')
    If VBA.Left$(web_Text, 3) = """""'" Then
        web_Text = VBA.Right$(web_Text, VBA.Len(web_Text) - 2)
    End If
    If VBA.Right$(web_Text, 3) = "'""""" Then
        web_Text = VBA.Left$(web_Text, VBA.Len(web_Text) - 2)
    End If

    PrepareTextForShell = web_Text
End Function

''
' Prepare text for using with printf command
' - Wrap in "..."
' - Replace ! with '!' (reserved in bash)
' - Escape \, `, $, and "
' - Replace % with %% (used as an argument marker in printf)
'
' @internal
' @method PrepareTextForPrintf
' @param {String} Text
' @return {String}
''
Public Function PrepareTextForPrintf(ByVal web_Text As String) As String
    ' Escape special characters (except for !)
    web_Text = VBA.Replace(web_Text, "\", "\\")
    web_Text = VBA.Replace(web_Text, "`", "\`")
    web_Text = VBA.Replace(web_Text, "$", "\$")
    web_Text = VBA.Replace(web_Text, "%", "%%")
    web_Text = VBA.Replace(web_Text, """", "\""")

    ' Wrap in quotes
    web_Text = """" & web_Text & """"

    ' Escape !
    web_Text = VBA.Replace(web_Text, "!", """'!'""")

    ' Guard for ! at beginning or end (""'!'"..." or "..."'!'"" -> '!'"..." or "..."'!')
    If VBA.Left$(web_Text, 3) = """""'" Then
        web_Text = VBA.Right$(web_Text, VBA.Len(web_Text) - 2)
    End If
    If VBA.Right$(web_Text, 3) = "'""""" Then
        web_Text = VBA.Left$(web_Text, VBA.Len(web_Text) - 2)
    End If

    PrepareTextForPrintf = web_Text
End Function

' ============================================= '
' 8. Cryptography
' ============================================= '

''
' Determine the HMAC for the given text and secret using the SHA1 hash algorithm.
'
' Reference:
' - http://stackoverflow.com/questions/8246340/does-vba-have-a-hash-hmac
'
' @example
' ```VB.net
' WebHelpers.HMACSHA1 "Howdy!", "Secret"
' ' -> c8fdf74a9d62aa41ac8136a1af471cec028fb157
' ```
'
' @method HMACSHA1
' @param {String} Text
' @param {String} Secret
' @param {String} [Format="Hex"] "Hex" or "Base64" encoding for result
' @return {String} HMAC-SHA1
''
Public Function HMACSHA1(Text As String, Secret As String, Optional Format As String = "Hex") As String
#If Mac Then
    Dim web_Command As String
    web_Command = "printf " & PrepareTextForPrintf(Text) & " | openssl dgst -sha1 -hmac " & PrepareTextForShell(Secret)

    If Format = "Base64" Then
        web_Command = web_Command & " -binary | openssl enc -base64"
    End If

    HMACSHA1 = VBA.Replace(ExecuteInShell(web_Command).Output, vbLf, "")
#Else
    Dim web_Crypto As Object
    Dim web_TextBytes() As Byte
    Dim web_SecretBytes() As Byte
    Dim web_Bytes() As Byte

    web_TextBytes = VBA.StrConv(Text, vbFromUnicode)
    web_SecretBytes = VBA.StrConv(Secret, vbFromUnicode)

    Set web_Crypto = CreateObject("System.Security.Cryptography.HMACSHA1")
    web_Crypto.Key = web_SecretBytes
    web_Bytes = web_Crypto.ComputeHash_2(web_TextBytes)

    Select Case Format
    Case "Base64"
        HMACSHA1 = web_AnsiBytesToBase64(web_Bytes)
    Case Else
        HMACSHA1 = web_AnsiBytesToHex(web_Bytes)
    End Select
#End If
End Function

''
' Determine the HMAC for the given text and secret using the SHA256 hash algorithm.
'
' @example
' ```VB.net
' WebHelpers.HMACSHA256 "Howdy!", "Secret"
' ' -> fb5d65...
' ```
'
' @method HMACSHA256
' @param {String} Text
' @param {String} Secret
' @param {String} [Format="Hex"] "Hex" or "Base64" encoding for result
' @return {String} HMAC-SHA256
''
Public Function HMACSHA256(Text As String, Secret As String, Optional Format As String = "Hex") As String
#If Mac Then
    Dim web_Command As String
    web_Command = "printf " & PrepareTextForPrintf(Text) & " | openssl dgst -sha256 -hmac " & PrepareTextForShell(Secret)

    If Format = "Base64" Then
        web_Command = web_Command & " -binary | openssl enc -base64"
    End If

    HMACSHA256 = VBA.Replace(ExecuteInShell(web_Command).Output, vbLf, "")
#Else
    Dim web_Crypto As Object
    Dim web_TextBytes() As Byte
    Dim web_SecretBytes() As Byte
    Dim web_Bytes() As Byte

    web_TextBytes = VBA.StrConv(Text, vbFromUnicode)
    web_SecretBytes = VBA.StrConv(Secret, vbFromUnicode)

    Set web_Crypto = CreateObject("System.Security.Cryptography.HMACSHA256")
    web_Crypto.Key = web_SecretBytes
    web_Bytes = web_Crypto.ComputeHash_2(web_TextBytes)

    Select Case Format
    Case "Base64"
        HMACSHA256 = web_AnsiBytesToBase64(web_Bytes)
    Case Else
        HMACSHA256 = web_AnsiBytesToHex(web_Bytes)
    End Select
#End If
End Function

''
' Determine the MD5 hash of the given text.
'
' Reference:
' - http://www.di-mgt.com.au/src/basMD5.bas.html
'
' @example
' ```VB.net
' WebHelpers.MD5 "Howdy!"
' ' -> 7105f32280940271293ee00ac97da5a7
' ```
'
' @method MD5
' @param {String} Text
' @param {String} [Format="Hex"] "Hex" or "Base64" encoding for result
' @return {String} MD5 Hash
''
Public Function MD5(Text As String, Optional Format As String = "Hex") As String
#If Mac Then
    Dim web_Command As String
    web_Command = "printf " & PrepareTextForPrintf(Text) & " | openssl dgst -md5"

    If Format = "Base64" Then
        web_Command = web_Command & " -binary | openssl enc -base64"
    End If

    MD5 = VBA.Replace(ExecuteInShell(web_Command).Output, vbLf, "")
#Else
    Dim web_Crypto As Object
    Dim web_TextBytes() As Byte
    Dim web_Bytes() As Byte

    web_TextBytes = VBA.StrConv(Text, vbFromUnicode)

    Set web_Crypto = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    web_Bytes = web_Crypto.ComputeHash_2(web_TextBytes)

    Select Case Format
    Case "Base64"
        MD5 = web_AnsiBytesToBase64(web_Bytes)
    Case Else
        MD5 = web_AnsiBytesToHex(web_Bytes)
    End Select
#End If
End Function

''
' Create random alphanumeric nonce (0-9a-zA-Z)
'
' @method CreateNonce
' @param {Integer} [NonceLength=32]
' @return {String} Randomly generated nonce
''
Public Function CreateNonce(Optional NonceLength As Integer = 32) As String
    Dim web_Str As String
    Dim web_Count As Integer
    Dim web_Result As String
    Dim web_Random As Integer

    web_Str = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUIVWXYZ"
    web_Result = ""

    VBA.Randomize
    For web_Count = 1 To NonceLength
        web_Random = VBA.Int(((VBA.Len(web_Str) - 1) * VBA.Rnd) + 1)
        web_Result = web_Result & VBA.Mid$(web_Str, web_Random, 1)
    Next
    CreateNonce = web_Result
End Function

''
' Convert string to ANSI bytes
'
' @internal
' @method StringToAnsiBytes
' @param {String} Text
' @return {Byte()}
''
Public Function StringToAnsiBytes(web_Text As String) As Byte()
    Dim web_Bytes() As Byte
    Dim web_AnsiBytes() As Byte
    Dim web_ByteIndex As Long
    Dim web_AnsiIndex As Long

    If VBA.Len(web_Text) > 0 Then
        ' Take first byte from unicode bytes
        ' VBA.Int is used for floor instead of round
        web_Bytes = web_Text
        ReDim web_AnsiBytes(VBA.Int(UBound(web_Bytes) / 2))

        web_AnsiIndex = LBound(web_Bytes)
        For web_ByteIndex = LBound(web_Bytes) To UBound(web_Bytes) Step 2
            web_AnsiBytes(web_AnsiIndex) = web_Bytes(web_ByteIndex)
            web_AnsiIndex = web_AnsiIndex + 1
        Next web_ByteIndex
    End If

    StringToAnsiBytes = web_AnsiBytes
End Function

#If Mac Then
#Else
Private Function web_AnsiBytesToBase64(web_Bytes() As Byte)
    ' Use XML to convert to Base64
    Dim web_XmlObj As Object
    Dim web_Node As Object

    Set web_XmlObj = CreateObject("MSXML2.DOMDocument")
    Set web_Node = web_XmlObj.createElement("b64")

    web_Node.DataType = "bin.base64"
    web_Node.nodeTypedValue = web_Bytes
    web_AnsiBytesToBase64 = web_Node.Text

    Set web_Node = Nothing
    Set web_XmlObj = Nothing
End Function

Private Function web_AnsiBytesToHex(web_Bytes() As Byte)
    Dim web_i As Long
    For web_i = LBound(web_Bytes) To UBound(web_Bytes)
        web_AnsiBytesToHex = web_AnsiBytesToHex & VBA.LCase$(VBA.Right$("0" & VBA.Hex$(web_Bytes(web_i)), 2))
    Next web_i
End Function
#End If

' ============================================= '
' 9. Converters
' ============================================= '

' Helper for url-encoded to create key=value pair
Private Function web_GetUrlEncodedKeyValue(Key As Variant, Value As Variant, Optional EncodingMode As UrlEncodingMode = UrlEncodingMode.FormUrlEncoding) As String
    Select Case VBA.VarType(Value)
    Case VBA.vbBoolean
        ' Convert boolean to lowercase
        If Value Then
            Value = "true"
        Else
            Value = "false"
        End If
    Case VBA.vbDate
        ' Use region invariant date (ISO-8601)
        Value = WebHelpers.ConvertToIso(CDate(Value))
    Case VBA.vbDecimal, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency
        ' Use region invariant number encoding ("." for decimal separator)
        Value = VBA.Replace(VBA.CStr(Value), ",", ".")
    End Select

    ' Url encode key and value (using + for spaces)
    web_GetUrlEncodedKeyValue = UrlEncode(Key, EncodingMode:=EncodingMode) & "=" & UrlEncode(Value, EncodingMode:=EncodingMode)
End Function

''
' VBA-JSON v2.3.1
' (c) Tim Hall - https://github.com/VBA-tools/VBA-JSON
'
' JSON Converter for VBA
'
' Errors:
' 10001 - JSON parse error
'
' @class JsonConverter
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'
' Based originally on vba-json (with extensive changes)
' BSD license included below
'
' JSONLib, http://code.google.com/p/vba-json/
'
' Copyright (c) 2013, Ryo Yokoyama
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'     * Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer.
'     * Redistributions in binary form must reproduce the above copyright
'       notice, this list of conditions and the following disclaimer in the
'       documentation and/or other materials provided with the distribution.
'     * Neither the name of the <organization> nor the
'       names of its contributors may be used to endorse or promote products
'       derived from this software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
' DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

' (Declarations moved to top)

' ============================================= '
' Public Methods
' ============================================= '

''
' Convert JSON string to object (Dictionary/Collection)
'
' @method ParseJson
' @param {String} json_String
' @return {Object} (Dictionary or Collection)
' @throws 10001 - JSON parse error
''
Public Function ParseJson(ByVal JsonString As String) As Object
    Dim json_Index As Long
    json_Index = 1

    ' Remove vbCr, vbLf, and vbTab from json_String
    JsonString = VBA.Replace(VBA.Replace(VBA.Replace(JsonString, VBA.vbCr, ""), VBA.vbLf, ""), VBA.vbTab, "")

    json_SkipSpaces JsonString, json_Index
    Select Case VBA.Mid$(JsonString, json_Index, 1)
    Case "{"
        Set ParseJson = json_ParseObject(JsonString, json_Index)
    Case "["
        Set ParseJson = json_ParseArray(JsonString, json_Index)
    Case Else
        ' Error: Invalid JSON string
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(JsonString, json_Index, "Expecting '{' or '['")
    End Select
End Function

''
' Convert object (Dictionary/Collection/Array) to JSON
'
' @method ConvertToJson
' @param {Variant} JsonValue (Dictionary, Collection, or Array)
' @param {Integer|String} Whitespace "Pretty" print json with given number of spaces per indentation (Integer) or given string
' @return {String}
''
Public Function ConvertToJson(ByVal JsonValue As Variant, Optional ByVal Whitespace As Variant, Optional ByVal json_CurrentIndentation As Long = 0) As String
    Dim json_Buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long
    Dim json_Index As Long
    Dim json_LBound As Long
    Dim json_UBound As Long
    Dim json_IsFirstItem As Boolean
    Dim json_Index2D As Long
    Dim json_LBound2D As Long
    Dim json_UBound2D As Long
    Dim json_IsFirstItem2D As Boolean
    Dim json_Key As Variant
    Dim json_Value As Variant
    Dim json_DateStr As String
    Dim json_Converted As String
    Dim json_SkipItem As Boolean
    Dim json_PrettyPrint As Boolean
    Dim json_Indentation As String
    Dim json_InnerIndentation As String

    json_LBound = -1
    json_UBound = -1
    json_IsFirstItem = True
    json_LBound2D = -1
    json_UBound2D = -1
    json_IsFirstItem2D = True
    json_PrettyPrint = Not IsMissing(Whitespace)

    Select Case VBA.VarType(JsonValue)
    Case VBA.vbNull
        ConvertToJson = "null"
    Case VBA.vbDate
        ' Date
        json_DateStr = ConvertToIso(VBA.CDate(JsonValue))

        ConvertToJson = """" & json_DateStr & """"
    Case VBA.vbString
        ' String (or large number encoded as string)
        If Not JsonOptions.UseDoubleForLargeNumbers And json_StringIsLargeNumber(JsonValue) Then
            ConvertToJson = JsonValue
        Else
            ConvertToJson = """" & json_Encode(JsonValue) & """"
        End If
    Case VBA.vbBoolean
        If JsonValue Then
            ConvertToJson = "true"
        Else
            ConvertToJson = "false"
        End If
    Case VBA.vbArray To VBA.vbArray + VBA.vbByte
        If json_PrettyPrint Then
            If VBA.VarType(Whitespace) = VBA.vbString Then
                json_Indentation = VBA.String$(json_CurrentIndentation + 1, Whitespace)
                json_InnerIndentation = VBA.String$(json_CurrentIndentation + 2, Whitespace)
            Else
                json_Indentation = VBA.Space$((json_CurrentIndentation + 1) * Whitespace)
                json_InnerIndentation = VBA.Space$((json_CurrentIndentation + 2) * Whitespace)
            End If
        End If

        ' Array
        json_BufferAppend json_Buffer, "[", json_BufferPosition, json_BufferLength

        On Error Resume Next

        json_LBound = LBound(JsonValue, 1)
        json_UBound = UBound(JsonValue, 1)
        json_LBound2D = LBound(JsonValue, 2)
        json_UBound2D = UBound(JsonValue, 2)

        If json_LBound >= 0 And json_UBound >= 0 Then
            For json_Index = json_LBound To json_UBound
                If json_IsFirstItem Then
                    json_IsFirstItem = False
                Else
                    ' Append comma to previous line
                    json_BufferAppend json_Buffer, ",", json_BufferPosition, json_BufferLength
                End If

                If json_LBound2D >= 0 And json_UBound2D >= 0 Then
                    ' 2D Array
                    If json_PrettyPrint Then
                        json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength
                    End If
                    json_BufferAppend json_Buffer, json_Indentation & "[", json_BufferPosition, json_BufferLength

                    For json_Index2D = json_LBound2D To json_UBound2D
                        If json_IsFirstItem2D Then
                            json_IsFirstItem2D = False
                        Else
                            json_BufferAppend json_Buffer, ",", json_BufferPosition, json_BufferLength
                        End If

                        json_Converted = ConvertToJson(JsonValue(json_Index, json_Index2D), Whitespace, json_CurrentIndentation + 2)

                        ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
                        If json_Converted = "" Then
                            ' (nest to only check if converted = "")
                            If json_IsUndefined(JsonValue(json_Index, json_Index2D)) Then
                                json_Converted = "null"
                            End If
                        End If

                        If json_PrettyPrint Then
                            json_Converted = vbNewLine & json_InnerIndentation & json_Converted
                        End If

                        json_BufferAppend json_Buffer, json_Converted, json_BufferPosition, json_BufferLength
                    Next json_Index2D

                    If json_PrettyPrint Then
                        json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength
                    End If

                    json_BufferAppend json_Buffer, json_Indentation & "]", json_BufferPosition, json_BufferLength
                    json_IsFirstItem2D = True
                Else
                    ' 1D Array
                    json_Converted = ConvertToJson(JsonValue(json_Index), Whitespace, json_CurrentIndentation + 1)

                    ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
                    If json_Converted = "" Then
                        ' (nest to only check if converted = "")
                        If json_IsUndefined(JsonValue(json_Index)) Then
                            json_Converted = "null"
                        End If
                    End If

                    If json_PrettyPrint Then
                        json_Converted = vbNewLine & json_Indentation & json_Converted
                    End If

                    json_BufferAppend json_Buffer, json_Converted, json_BufferPosition, json_BufferLength
                End If
            Next json_Index
        End If

        On Error GoTo 0

        If json_PrettyPrint Then
            json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength

            If VBA.VarType(Whitespace) = VBA.vbString Then
                json_Indentation = VBA.String$(json_CurrentIndentation, Whitespace)
            Else
                json_Indentation = VBA.Space$(json_CurrentIndentation * Whitespace)
            End If
        End If

        json_BufferAppend json_Buffer, json_Indentation & "]", json_BufferPosition, json_BufferLength

        ConvertToJson = json_BufferToString(json_Buffer, json_BufferPosition)

    ' Dictionary or Collection
    Case VBA.vbObject
        If json_PrettyPrint Then
            If VBA.VarType(Whitespace) = VBA.vbString Then
                json_Indentation = VBA.String$(json_CurrentIndentation + 1, Whitespace)
            Else
                json_Indentation = VBA.Space$((json_CurrentIndentation + 1) * Whitespace)
            End If
        End If

        ' Dictionary
        If VBA.TypeName(JsonValue) = "Dictionary" Then
            json_BufferAppend json_Buffer, "{", json_BufferPosition, json_BufferLength
            For Each json_Key In JsonValue.Keys
                ' For Objects, undefined (Empty/Nothing) is not added to object
                json_Converted = ConvertToJson(JsonValue(json_Key), Whitespace, json_CurrentIndentation + 1)
                If json_Converted = "" Then
                    json_SkipItem = json_IsUndefined(JsonValue(json_Key))
                Else
                    json_SkipItem = False
                End If

                If Not json_SkipItem Then
                    If json_IsFirstItem Then
                        json_IsFirstItem = False
                    Else
                        json_BufferAppend json_Buffer, ",", json_BufferPosition, json_BufferLength
                    End If

                    If json_PrettyPrint Then
                        json_Converted = vbNewLine & json_Indentation & """" & json_Key & """: " & json_Converted
                    Else
                        json_Converted = """" & json_Key & """:" & json_Converted
                    End If

                    json_BufferAppend json_Buffer, json_Converted, json_BufferPosition, json_BufferLength
                End If
            Next json_Key

            If json_PrettyPrint Then
                json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength

                If VBA.VarType(Whitespace) = VBA.vbString Then
                    json_Indentation = VBA.String$(json_CurrentIndentation, Whitespace)
                Else
                    json_Indentation = VBA.Space$(json_CurrentIndentation * Whitespace)
                End If
            End If

            json_BufferAppend json_Buffer, json_Indentation & "}", json_BufferPosition, json_BufferLength

        ' Collection
        ElseIf VBA.TypeName(JsonValue) = "Collection" Then
            json_BufferAppend json_Buffer, "[", json_BufferPosition, json_BufferLength
            For Each json_Value In JsonValue
                If json_IsFirstItem Then
                    json_IsFirstItem = False
                Else
                    json_BufferAppend json_Buffer, ",", json_BufferPosition, json_BufferLength
                End If

                json_Converted = ConvertToJson(json_Value, Whitespace, json_CurrentIndentation + 1)

                ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
                If json_Converted = "" Then
                    ' (nest to only check if converted = "")
                    If json_IsUndefined(json_Value) Then
                        json_Converted = "null"
                    End If
                End If

                If json_PrettyPrint Then
                    json_Converted = vbNewLine & json_Indentation & json_Converted
                End If

                json_BufferAppend json_Buffer, json_Converted, json_BufferPosition, json_BufferLength
            Next json_Value

            If json_PrettyPrint Then
                json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength

                If VBA.VarType(Whitespace) = VBA.vbString Then
                    json_Indentation = VBA.String$(json_CurrentIndentation, Whitespace)
                Else
                    json_Indentation = VBA.Space$(json_CurrentIndentation * Whitespace)
                End If
            End If

            json_BufferAppend json_Buffer, json_Indentation & "]", json_BufferPosition, json_BufferLength
        End If

        ConvertToJson = json_BufferToString(json_Buffer, json_BufferPosition)
    Case VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal
        ' Number (use decimals for numbers)
        ConvertToJson = VBA.Replace(JsonValue, ",", ".")
    Case Else
        ' vbEmpty, vbError, vbDataObject, vbByte, vbUserDefinedType
        ' Use VBA's built-in to-string
        On Error Resume Next
        ConvertToJson = JsonValue
        On Error GoTo 0
    End Select
End Function

' ============================================= '
' Private Functions
' ============================================= '

Private Function json_ParseObject(json_String As String, ByRef json_Index As Long) As Dictionary
    Dim json_Key As String
    Dim json_NextChar As String

    Set json_ParseObject = New Dictionary
    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> "{" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '{'")
    Else
        json_Index = json_Index + 1

        Do
            json_SkipSpaces json_String, json_Index
            If VBA.Mid$(json_String, json_Index, 1) = "}" Then
                json_Index = json_Index + 1
                Exit Function
            ElseIf VBA.Mid$(json_String, json_Index, 1) = "," Then
                json_Index = json_Index + 1
                json_SkipSpaces json_String, json_Index
            End If

            json_Key = json_ParseKey(json_String, json_Index)
            json_NextChar = json_Peek(json_String, json_Index)
            If json_NextChar = "[" Or json_NextChar = "{" Then
                Set json_ParseObject.Item(json_Key) = json_ParseValue(json_String, json_Index)
            Else
                json_ParseObject.Item(json_Key) = json_ParseValue(json_String, json_Index)
            End If
        Loop
    End If
End Function

Private Function json_ParseArray(json_String As String, ByRef json_Index As Long) As Collection
    Set json_ParseArray = New Collection

    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> "[" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '['")
    Else
        json_Index = json_Index + 1

        Do
            json_SkipSpaces json_String, json_Index
            If VBA.Mid$(json_String, json_Index, 1) = "]" Then
                json_Index = json_Index + 1
                Exit Function
            ElseIf VBA.Mid$(json_String, json_Index, 1) = "," Then
                json_Index = json_Index + 1
                json_SkipSpaces json_String, json_Index
            End If

            json_ParseArray.Add json_ParseValue(json_String, json_Index)
        Loop
    End If
End Function

Private Function json_ParseValue(json_String As String, ByRef json_Index As Long) As Variant
    json_SkipSpaces json_String, json_Index
    Select Case VBA.Mid$(json_String, json_Index, 1)
    Case "{"
        Set json_ParseValue = json_ParseObject(json_String, json_Index)
    Case "["
        Set json_ParseValue = json_ParseArray(json_String, json_Index)
    Case """", "'"
        json_ParseValue = json_ParseString(json_String, json_Index)
    Case Else
        If VBA.Mid$(json_String, json_Index, 4) = "true" Then
            json_ParseValue = True
            json_Index = json_Index + 4
        ElseIf VBA.Mid$(json_String, json_Index, 5) = "false" Then
            json_ParseValue = False
            json_Index = json_Index + 5
        ElseIf VBA.Mid$(json_String, json_Index, 4) = "null" Then
            json_ParseValue = Null
            json_Index = json_Index + 4
        ElseIf VBA.InStr("+-0123456789", VBA.Mid$(json_String, json_Index, 1)) Then
            json_ParseValue = json_ParseNumber(json_String, json_Index)
        Else
            Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['")
        End If
    End Select
End Function

Private Function json_ParseString(json_String As String, ByRef json_Index As Long) As String
    Dim json_Quote As String
    Dim json_Char As String
    Dim json_Code As String
    Dim json_Buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long

    json_SkipSpaces json_String, json_Index

    ' Store opening quote to look for matching closing quote
    json_Quote = VBA.Mid$(json_String, json_Index, 1)
    json_Index = json_Index + 1

    Do While json_Index > 0 And json_Index <= Len(json_String)
        json_Char = VBA.Mid$(json_String, json_Index, 1)

        Select Case json_Char
        Case "\"
            ' Escaped string, \\, or \/
            json_Index = json_Index + 1
            json_Char = VBA.Mid$(json_String, json_Index, 1)

            Select Case json_Char
            Case """", "\", "/", "'"
                json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "b"
                json_BufferAppend json_Buffer, vbBack, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "f"
                json_BufferAppend json_Buffer, vbFormFeed, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "n"
                json_BufferAppend json_Buffer, vbCrLf, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "r"
                json_BufferAppend json_Buffer, vbCr, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "t"
                json_BufferAppend json_Buffer, vbTab, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "u"
                ' Unicode character escape (e.g. \u00a9 = Copyright)
                json_Index = json_Index + 1
                json_Code = VBA.Mid$(json_String, json_Index, 4)
                json_BufferAppend json_Buffer, VBA.ChrW(VBA.Val("&h" + json_Code)), json_BufferPosition, json_BufferLength
                json_Index = json_Index + 4
            End Select
        Case json_Quote
            json_ParseString = json_BufferToString(json_Buffer, json_BufferPosition)
            json_Index = json_Index + 1
            Exit Function
        Case Else
            json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
            json_Index = json_Index + 1
        End Select
    Loop
End Function

Private Function json_ParseNumber(json_String As String, ByRef json_Index As Long) As Variant
    Dim json_Char As String
    Dim json_Value As String
    Dim json_IsLargeNumber As Boolean

    json_SkipSpaces json_String, json_Index

    Do While json_Index > 0 And json_Index <= Len(json_String)
        json_Char = VBA.Mid$(json_String, json_Index, 1)

        If VBA.InStr("+-0123456789.eE", json_Char) Then
            ' Unlikely to have massive number, so use simple append rather than buffer here
            json_Value = json_Value & json_Char
            json_Index = json_Index + 1
        Else
            ' Excel only stores 15 significant digits, so any numbers larger than that are truncated
            ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
            ' See: http://support.microsoft.com/kb/269370
            '
            ' Fix: Parse -> String, Convert -> String longer than 15/16 characters containing only numbers and decimal points -> Number
            ' (decimal doesn't factor into significant digit count, so if present check for 15 digits + decimal = 16)
            json_IsLargeNumber = IIf(InStr(json_Value, "."), Len(json_Value) >= 17, Len(json_Value) >= 16)
            If Not JsonOptions.UseDoubleForLargeNumbers And json_IsLargeNumber Then
                json_ParseNumber = json_Value
            Else
                ' VBA.Val does not use regional settings, so guard for comma is not needed
                json_ParseNumber = VBA.Val(json_Value)
            End If
            Exit Function
        End If
    Loop
End Function

Private Function json_ParseKey(json_String As String, ByRef json_Index As Long) As String
    ' Parse key with single or double quotes
    If VBA.Mid$(json_String, json_Index, 1) = """" Or VBA.Mid$(json_String, json_Index, 1) = "'" Then
        json_ParseKey = json_ParseString(json_String, json_Index)
    ElseIf JsonOptions.AllowUnquotedKeys Then
        Dim json_Char As String
        Do While json_Index > 0 And json_Index <= Len(json_String)
            json_Char = VBA.Mid$(json_String, json_Index, 1)
            If (json_Char <> " ") And (json_Char <> ":") Then
                json_ParseKey = json_ParseKey & json_Char
                json_Index = json_Index + 1
            Else
                Exit Do
            End If
        Loop
    Else
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '""' or '''")
    End If

    ' Check for colon and skip if present or throw if not present
    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> ":" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting ':'")
    Else
        json_Index = json_Index + 1
    End If
End Function

Private Function json_IsUndefined(ByVal json_Value As Variant) As Boolean
    ' Empty / Nothing -> undefined
    Select Case VBA.VarType(json_Value)
    Case VBA.vbEmpty
        json_IsUndefined = True
    Case VBA.vbObject
        Select Case VBA.TypeName(json_Value)
        Case "Empty", "Nothing"
            json_IsUndefined = True
        End Select
    End Select
End Function

Private Function json_Encode(ByVal json_Text As Variant) As String
    ' Reference: http://www.ietf.org/rfc/rfc4627.txt
    ' Escape: ", \, /, backspace, form feed, line feed, carriage return, tab
    Dim json_Index As Long
    Dim json_Char As String
    Dim json_AscCode As Long
    Dim json_Buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long

    For json_Index = 1 To VBA.Len(json_Text)
        json_Char = VBA.Mid$(json_Text, json_Index, 1)
        json_AscCode = VBA.AscW(json_Char)

        ' When AscW returns a negative number, it returns the twos complement form of that number.
        ' To convert the twos complement notation into normal binary notation, add 0xFFF to the return result.
        ' https://support.microsoft.com/en-us/kb/272138
        If json_AscCode < 0 Then
            json_AscCode = json_AscCode + 65536
        End If

        ' From spec, ", \, and control characters must be escaped (solidus is optional)

        Select Case json_AscCode
        Case 34
            ' " -> 34 -> \"
            json_Char = "\"""
        Case 92
            ' \ -> 92 -> \\
            json_Char = "\\"
        Case 47
            ' / -> 47 -> \/ (optional)
            If JsonOptions.EscapeSolidus Then
                json_Char = "\/"
            End If
        Case 8
            ' backspace -> 8 -> \b
            json_Char = "\b"
        Case 12
            ' form feed -> 12 -> \f
            json_Char = "\f"
        Case 10
            ' line feed -> 10 -> \n
            json_Char = "\n"
        Case 13
            ' carriage return -> 13 -> \r
            json_Char = "\r"
        Case 9
            ' tab -> 9 -> \t
            json_Char = "\t"
        Case 0 To 31, 127 To 65535
            ' Non-ascii characters -> convert to 4-digit hex
            json_Char = "\u" & VBA.Right$("0000" & VBA.Hex$(json_AscCode), 4)
        End Select

        json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
    Next json_Index

    json_Encode = json_BufferToString(json_Buffer, json_BufferPosition)
End Function

Private Function json_Peek(json_String As String, ByVal json_Index As Long, Optional json_NumberOfCharacters As Long = 1) As String
    ' "Peek" at the next number of characters without incrementing json_Index (ByVal instead of ByRef)
    json_SkipSpaces json_String, json_Index
    json_Peek = VBA.Mid$(json_String, json_Index, json_NumberOfCharacters)
End Function

Private Sub json_SkipSpaces(json_String As String, ByRef json_Index As Long)
    ' Increment index to skip over spaces
    Do While json_Index > 0 And json_Index <= VBA.Len(json_String) And VBA.Mid$(json_String, json_Index, 1) = " "
        json_Index = json_Index + 1
    Loop
End Sub

Private Function json_StringIsLargeNumber(json_String As Variant) As Boolean
    ' Check if the given string is considered a "large number"
    ' (See json_ParseNumber)

    Dim json_Length As Long
    Dim json_CharIndex As Long
    json_Length = VBA.Len(json_String)

    ' Length with be at least 16 characters and assume will be less than 100 characters
    If json_Length >= 16 And json_Length <= 100 Then
        Dim json_CharCode As String

        json_StringIsLargeNumber = True

        For json_CharIndex = 1 To json_Length
            json_CharCode = VBA.Asc(VBA.Mid$(json_String, json_CharIndex, 1))
            Select Case json_CharCode
            ' Look for .|0-9|E|e
            Case 46, 48 To 57, 69, 101
                ' Continue through characters
            Case Else
                json_StringIsLargeNumber = False
                Exit Function
            End Select
        Next json_CharIndex
    End If
End Function

Private Function json_ParseErrorMessage(json_String As String, ByRef json_Index As Long, ErrorMessage As String)
    ' Provide detailed parse error message, including details of where and what occurred
    '
    ' Example:
    ' Error parsing JSON:
    ' {"abcde":True}
    '          ^
    ' Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['

    Dim json_StartIndex As Long
    Dim json_StopIndex As Long

    ' Include 10 characters before and after error (if possible)
    json_StartIndex = json_Index - 10
    json_StopIndex = json_Index + 10
    If json_StartIndex <= 0 Then
        json_StartIndex = 1
    End If
    If json_StopIndex > VBA.Len(json_String) Then
        json_StopIndex = VBA.Len(json_String)
    End If

    json_ParseErrorMessage = "Error parsing JSON:" & VBA.vbNewLine & _
                             VBA.Mid$(json_String, json_StartIndex, json_StopIndex - json_StartIndex + 1) & VBA.vbNewLine & _
                             VBA.Space$(json_Index - json_StartIndex) & "^" & VBA.vbNewLine & _
                             ErrorMessage
End Function

Private Sub json_BufferAppend(ByRef json_Buffer As String, _
                              ByRef json_Append As Variant, _
                              ByRef json_BufferPosition As Long, _
                              ByRef json_BufferLength As Long)
    ' VBA can be slow to append strings due to allocating a new string for each append
    ' Instead of using the traditional append, allocate a large empty string and then copy string at append position
    '
    ' Example:
    ' Buffer: "abc  "
    ' Append: "def"
    ' Buffer Position: 3
    ' Buffer Length: 5
    '
    ' Buffer position + Append length > Buffer length -> Append chunk of blank space to buffer
    ' Buffer: "abc       "
    ' Buffer Length: 10
    '
    ' Put "def" into buffer at position 3 (0-based)
    ' Buffer: "abcdef    "
    '
    ' Approach based on cStringBuilder from vbAccelerator
    ' http://www.vbaccelerator.com/home/VB/Code/Techniques/RunTime_Debug_Tracing/VB6_Tracer_Utility_zip_cStringBuilder_cls.asp
    '
    ' and clsStringAppend from Philip Swannell
    ' https://github.com/VBA-tools/VBA-JSON/pull/82

    Dim json_AppendLength As Long
    Dim json_LengthPlusPosition As Long

    json_AppendLength = VBA.Len(json_Append)
    json_LengthPlusPosition = json_AppendLength + json_BufferPosition

    If json_LengthPlusPosition > json_BufferLength Then
        ' Appending would overflow buffer, add chunk
        ' (double buffer length or append length, whichever is bigger)
        Dim json_AddedLength As Long
        json_AddedLength = IIf(json_AppendLength > json_BufferLength, json_AppendLength, json_BufferLength)

        json_Buffer = json_Buffer & VBA.Space$(json_AddedLength)
        json_BufferLength = json_BufferLength + json_AddedLength
    End If

    ' Note: Namespacing with VBA.Mid$ doesn't work properly here, throwing compile error:
    ' Function call on left-hand side of assignment must return Variant or Object
    Mid$(json_Buffer, json_BufferPosition + 1, json_AppendLength) = CStr(json_Append)
    json_BufferPosition = json_BufferPosition + json_AppendLength
End Sub

Private Function json_BufferToString(ByRef json_Buffer As String, ByVal json_BufferPosition As Long) As String
    If json_BufferPosition > 0 Then
        json_BufferToString = VBA.Left$(json_Buffer, json_BufferPosition)
    End If
End Function

''
' VBA-UTC v1.0.6
' (c) Tim Hall - https://github.com/VBA-tools/VBA-UtcConverter
'
' UTC/ISO 8601 Converter for VBA
'
' Errors:
' 10011 - UTC parsing error
' 10012 - UTC conversion error
' 10013 - ISO 8601 parsing error
' 10014 - ISO 8601 conversion error
'
' @module UtcConverter
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

' (Declarations moved to top)

' ============================================= '
' Public Methods
' ============================================= '

''
' Parse UTC date to local date
'
' @method ParseUtc
' @param {Date} UtcDate
' @return {Date} Local date
' @throws 10011 - UTC parsing error
''
Public Function ParseUtc(utc_UtcDate As Date) As Date
    On Error GoTo utc_ErrorHandling

#If Mac Then
    ParseUtc = utc_ConvertDate(utc_UtcDate)
#Else
    Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
    Dim utc_LocalDate As utc_SYSTEMTIME

    utc_GetTimeZoneInformation utc_TimeZoneInfo
    utc_SystemTimeToTzSpecificLocalTime utc_TimeZoneInfo, utc_DateToSystemTime(utc_UtcDate), utc_LocalDate

    ParseUtc = utc_SystemTimeToDate(utc_LocalDate)
#End If

    Exit Function

utc_ErrorHandling:
    Err.Raise 10011, "UtcConverter.ParseUtc", "UTC parsing error: " & Err.Number & " - " & Err.Description
End Function

''
' Convert local date to UTC date
'
' @method ConvertToUrc
' @param {Date} utc_LocalDate
' @return {Date} UTC date
' @throws 10012 - UTC conversion error
''
Public Function ConvertToUtc(utc_LocalDate As Date) As Date
    On Error GoTo utc_ErrorHandling

#If Mac Then
    ConvertToUtc = utc_ConvertDate(utc_LocalDate, utc_ConvertToUtc:=True)
#Else
    Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
    Dim utc_UtcDate As utc_SYSTEMTIME

    utc_GetTimeZoneInformation utc_TimeZoneInfo
    utc_TzSpecificLocalTimeToSystemTime utc_TimeZoneInfo, utc_DateToSystemTime(utc_LocalDate), utc_UtcDate

    ConvertToUtc = utc_SystemTimeToDate(utc_UtcDate)
#End If

    Exit Function

utc_ErrorHandling:
    Err.Raise 10012, "UtcConverter.ConvertToUtc", "UTC conversion error: " & Err.Number & " - " & Err.Description
End Function

''
' CHANGED: Date/Time cells are timezone naive, so don't do any conversion!
' Parse ISO 8601 date string to local date
'
' @method ParseIso
' @param {Date} utc_IsoString
' @return {Date} Local date
' @throws 10013 - ISO 8601 parsing error
''
Public Function ParseIso(utc_IsoString As String) As Date
    On Error GoTo utc_ErrorHandling

    Dim utc_Parts() As String
    Dim utc_DateParts() As String
    Dim utc_TimeParts() As String
    Dim utc_OffsetIndex As Long
    Dim utc_HasOffset As Boolean
    Dim utc_NegativeOffset As Boolean
    Dim utc_OffsetParts() As String
    Dim utc_Offset As Date

    utc_Parts = VBA.Split(utc_IsoString, "T")
    utc_DateParts = VBA.Split(utc_Parts(0), "-")
    ParseIso = VBA.DateSerial(VBA.CInt(utc_DateParts(0)), VBA.CInt(utc_DateParts(1)), VBA.CInt(utc_DateParts(2)))

    If UBound(utc_Parts) > 0 Then
        If VBA.InStr(utc_Parts(1), "Z") Then
            utc_TimeParts = VBA.Split(VBA.Replace(utc_Parts(1), "Z", ""), ":")
        Else
            utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "+")
            If utc_OffsetIndex = 0 Then
                utc_NegativeOffset = True
                utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "-")
            End If

            If utc_OffsetIndex > 0 Then
                utc_HasOffset = True
                utc_TimeParts = VBA.Split(VBA.Left$(utc_Parts(1), utc_OffsetIndex - 1), ":")
                utc_OffsetParts = VBA.Split(VBA.Right$(utc_Parts(1), Len(utc_Parts(1)) - utc_OffsetIndex), ":")

                Select Case UBound(utc_OffsetParts)
                Case 0
                    utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), 0, 0)
                Case 1
                    utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), VBA.CInt(utc_OffsetParts(1)), 0)
                Case 2
                    ' VBA.Val does not use regional settings, use for seconds to avoid decimal/comma issues
                    utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), VBA.CInt(utc_OffsetParts(1)), Int(VBA.Val(utc_OffsetParts(2))))
                End Select

                If utc_NegativeOffset Then: utc_Offset = -utc_Offset
            Else
                utc_TimeParts = VBA.Split(utc_Parts(1), ":")
            End If
        End If

        Select Case UBound(utc_TimeParts)
        Case 0
            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), 0, 0)
        Case 1
            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), 0)
        Case 2
            ' VBA.Val does not use regional settings, use for seconds to avoid decimal/comma issues
            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), Int(VBA.Val(utc_TimeParts(2))))
        End Select
        
        'CHANGED: Don't do any timezone conversion
        'ParseIso = ParseUtc(ParseIso)

        'If utc_HasOffset Then
        '    ParseIso = ParseIso - utc_Offset
        'End If
    End If

    Exit Function

utc_ErrorHandling:
    Err.Raise 10013, "UtcConverter.ParseIso", "ISO 8601 parsing error for " & utc_IsoString & ": " & Err.Number & " - " & Err.Description
End Function

''
' CHANGED: Date/Time cells are timezone naive, so don't do any conversion!
' Convert local date to ISO 8601 string
'
' @method ConvertToIso
' @param {Date} utc_LocalDate
' @return {Date} ISO 8601 string
' @throws 10014 - ISO 8601 conversion error
''
Public Function ConvertToIso(utc_LocalDate As Date) As String
    On Error GoTo utc_ErrorHandling

    ' CHANGED: Removed ConvertToUtc
    'ConvertToIso = VBA.Format$(ConvertToUtc(utc_LocalDate), "yyyy-mm-ddTHH:mm:ss.000Z")
    ConvertToIso = VBA.Format$(utc_LocalDate, "yyyy-mm-ddTHH:mm:ss.000Z")

    Exit Function

utc_ErrorHandling:
    Err.Raise 10014, "UtcConverter.ConvertToIso", "ISO 8601 conversion error: " & Err.Number & " - " & Err.Description
End Function

' ============================================= '
' Private Functions
' ============================================= '

#If Mac Then

Private Function utc_ConvertDate(utc_Value As Date, Optional utc_ConvertToUtc As Boolean = False) As Date
    Dim utc_ShellCommand As String
    Dim utc_Result As utc_ShellResult
    Dim utc_Parts() As String
    Dim utc_DateParts() As String
    Dim utc_TimeParts() As String

    If utc_ConvertToUtc Then
        utc_ShellCommand = "date -ur `date -jf '%Y-%m-%d %H:%M:%S' " & _
            "'" & VBA.Format$(utc_Value, "yyyy-mm-dd HH:mm:ss") & "' " & _
            " +'%s'` +'%Y-%m-%d %H:%M:%S'"
    Else
        utc_ShellCommand = "date -jf '%Y-%m-%d %H:%M:%S %z' " & _
            "'" & VBA.Format$(utc_Value, "yyyy-mm-dd HH:mm:ss") & " +0000' " & _
            "+'%Y-%m-%d %H:%M:%S'"
    End If

    utc_Result = utc_ExecuteInShell(utc_ShellCommand)

    If utc_Result.utc_Output = "" Then
        Err.Raise 10015, "UtcConverter.utc_ConvertDate", "'date' command failed"
    Else
        utc_Parts = Split(utc_Result.utc_Output, " ")
        utc_DateParts = Split(utc_Parts(0), "-")
        utc_TimeParts = Split(utc_Parts(1), ":")

        utc_ConvertDate = DateSerial(utc_DateParts(0), utc_DateParts(1), utc_DateParts(2)) + _
            TimeSerial(utc_TimeParts(0), utc_TimeParts(1), utc_TimeParts(2))
    End If
End Function

Private Function utc_ExecuteInShell(utc_ShellCommand As String) As utc_ShellResult
#If VBA7 Then
    Dim utc_File As LongPtr
    Dim utc_Read As LongPtr
#Else
    Dim utc_File As Long
    Dim utc_Read As Long
#End If

    Dim utc_Chunk As String

    On Error GoTo utc_ErrorHandling
    utc_File = utc_popen(utc_ShellCommand, "r")

    If utc_File = 0 Then: Exit Function

    Do While utc_feof(utc_File) = 0
        utc_Chunk = VBA.Space$(50)
        utc_Read = CLng(utc_fread(utc_Chunk, 1, Len(utc_Chunk) - 1, utc_File))
        If utc_Read > 0 Then
            utc_Chunk = VBA.Left$(utc_Chunk, CLng(utc_Read))
            utc_ExecuteInShell.utc_Output = utc_ExecuteInShell.utc_Output & utc_Chunk
        End If
    Loop

utc_ErrorHandling:
    utc_ExecuteInShell.utc_ExitCode = CLng(utc_pclose(utc_File))
End Function

#Else

Private Function utc_DateToSystemTime(utc_Value As Date) As utc_SYSTEMTIME
    utc_DateToSystemTime.utc_wYear = VBA.Year(utc_Value)
    utc_DateToSystemTime.utc_wMonth = VBA.Month(utc_Value)
    utc_DateToSystemTime.utc_wDay = VBA.Day(utc_Value)
    utc_DateToSystemTime.utc_wHour = VBA.Hour(utc_Value)
    utc_DateToSystemTime.utc_wMinute = VBA.Minute(utc_Value)
    utc_DateToSystemTime.utc_wSecond = VBA.Second(utc_Value)
    utc_DateToSystemTime.utc_wMilliseconds = 0
End Function

Private Function utc_SystemTimeToDate(utc_Value As utc_SYSTEMTIME) As Date
    utc_SystemTimeToDate = DateSerial(utc_Value.utc_wYear, utc_Value.utc_wMonth, utc_Value.utc_wDay) + _
        TimeSerial(utc_Value.utc_wHour, utc_Value.utc_wMinute, utc_Value.utc_wSecond)
End Function

#End If

''
' AutoProxy 1.0.2
' (c) Damien Thirion
'
' Auto configure proxy server
'
' Based on code shared by Stephen Sulzer
' https://groups.google.com/d/msg/microsoft.public.winhttp/ZeWN2Xig82g/jgHIBDSfBwsJ
'
' Errors:
' 11020 - Unknown error while detecting proxy
' 11021 - WPAD detection failed
' 11022 - Unable to download proxy auto-config script
' 11023 - Error in proxy auto-config script
' 11024 - No proxy can be located for the specified URL
' 11025 - Specified URL is not valid
'
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

''
' Returns IE proxy settings
' including auto-detection and auto-config scripts results
'
' @param {String} Url
' @param[out] {String} ProxyServer
' @param[out] {String} ProxyBypass
''
Public Sub GetAutoProxy(ByVal url As String, ByRef proxyServer As String, ByRef ProxyBypass As String)
#If Mac Then
    ' (Windows only)
#ElseIf VBA7 Then
    Dim AutoProxy_ProxyStringPtr As LongPtr
    Dim AutoProxy_ptr As LongPtr
    Dim AutoProxy_hSession As LongPtr
#Else
    Dim AutoProxy_ProxyStringPtr As Long
    Dim AutoProxy_ptr As Long
    Dim AutoProxy_hSession As Long
#End If
#If Mac Then
#Else
    Dim AutoProxy_IEProxyConfig As AUTOPROXY_IE_PROXY_CONFIG
    Dim AutoProxy_AutoProxyOptions As AUTOPROXY_OPTIONS
    Dim AutoProxy_ProxyInfo As AUTOPROXY_INFO
    Dim AutoProxy_doAutoProxy As Boolean
    Dim AutoProxy_Error As Long
    Dim AutoProxy_ErrorMsg As String

    AutoProxy_AutoProxyOptions.AutoProxy_fAutoLogonIfChallenged = 1
    proxyServer = ""
    ProxyBypass = ""

    ' WinHttpGetProxyForUrl returns unexpected errors if Url is empty
    If url = "" Then url = " "

    On Error GoTo AutoProxy_Cleanup

    ' Check IE's proxy configuration
    If (AutoProxy_GetIEProxy(AutoProxy_IEProxyConfig) > 0) Then
        ' If IE is configured to auto-detect, then we will too.
        If (AutoProxy_IEProxyConfig.AutoProxy_fAutoDetect <> 0) Then
            AutoProxy_AutoProxyOptions.AutoProxy_dwFlags = AUTOPROXY_AUTO_DETECT
            AutoProxy_AutoProxyOptions.AutoProxy_dwAutoDetectFlags = _
                AUTOPROXY_DETECT_TYPE_DHCP + AUTOPROXY_DETECT_TYPE_DNS
            AutoProxy_doAutoProxy = True
        End If

        ' If IE is configured to use an auto-config script, then
        ' we will use it too
        If (AutoProxy_IEProxyConfig.AutoProxy_lpszAutoConfigUrl <> 0) Then
            AutoProxy_AutoProxyOptions.AutoProxy_dwFlags = _
                AutoProxy_AutoProxyOptions.AutoProxy_dwFlags + AUTOPROXY_CONFIG_URL
            AutoProxy_AutoProxyOptions.AutoProxy_lpszAutoConfigUrl = AutoProxy_IEProxyConfig.AutoProxy_lpszAutoConfigUrl
            AutoProxy_doAutoProxy = True
        End If
    Else
        ' If the IE proxy config is not available, then
        ' we will try auto-detection
        AutoProxy_AutoProxyOptions.AutoProxy_dwFlags = AUTOPROXY_AUTO_DETECT
        AutoProxy_AutoProxyOptions.AutoProxy_dwAutoDetectFlags = _
            AUTOPROXY_DETECT_TYPE_DHCP + AUTOPROXY_DETECT_TYPE_DNS
        AutoProxy_doAutoProxy = True
    End If

    If AutoProxy_doAutoProxy Then
        On Error GoTo AutoProxy_TryIEFallback

        ' Need to create a temporary WinHttp session handle
        ' Note: Performance of this GetProxyInfoForUrl function can be
        '       improved by saving this AutoProxy_hSession handle across calls
        '       instead of creating a new handle each time
        AutoProxy_hSession = AutoProxy_HttpOpen(0, 1, 0, 0, 0)

        If (AutoProxy_GetProxyForUrl( _
            AutoProxy_hSession, StrPtr(url), AutoProxy_AutoProxyOptions, AutoProxy_ProxyInfo) > 0) Then

            AutoProxy_ProxyStringPtr = AutoProxy_ProxyInfo.AutoProxy_lpszProxy
        Else
            AutoProxy_Error = Err.LastDllError
            Select Case AutoProxy_Error
            Case 12180
                AutoProxy_ErrorMsg = "WPAD detection failed"
                AutoProxy_Error = 10021
            Case 12167
                AutoProxy_ErrorMsg = "Unable to download proxy auto-config script"
                AutoProxy_Error = 10022
            Case 12166
                AutoProxy_ErrorMsg = "Error in proxy auto-config script"
                AutoProxy_Error = 10023
            Case 12178
                AutoProxy_ErrorMsg = "No proxy can be located for the specified URL"
                AutoProxy_Error = 10024
            Case 12005, 12006
                AutoProxy_ErrorMsg = "Specified URL is not valid"
                AutoProxy_Error = 10025
            Case Else
                AutoProxy_ErrorMsg = "Unknown error while detecting proxy"
                AutoProxy_Error = 10020
            End Select
        End If

        AutoProxy_HttpClose AutoProxy_hSession
        AutoProxy_hSession = 0
    End If

AutoProxy_TryIEFallback:
    On Error GoTo AutoProxy_Cleanup

    ' If we don't have a proxy server from WinHttpGetProxyForUrl,
    ' then pick one up from the IE proxy config (if given)
    If (AutoProxy_ProxyStringPtr = 0) Then
        AutoProxy_ProxyStringPtr = AutoProxy_IEProxyConfig.AutoProxy_lpszProxy
    End If

    ' If there's a proxy string, convert it to a Basic string
    If (AutoProxy_ProxyStringPtr <> 0) Then
        AutoProxy_ptr = AutoProxy_SysAllocString(AutoProxy_ProxyStringPtr)
        AutoProxy_CopyMemory VarPtr(proxyServer), VarPtr(AutoProxy_ptr), 4
    End If

    ' Pick up any bypass string from the IEProxyConfig
    If (AutoProxy_IEProxyConfig.AutoProxy_lpszProxyBypass <> 0) Then
        AutoProxy_ptr = AutoProxy_SysAllocString(AutoProxy_IEProxyConfig.AutoProxy_lpszProxyBypass)
        AutoProxy_CopyMemory VarPtr(ProxyBypass), VarPtr(AutoProxy_ptr), 4
    End If

    ' Ensure WinHttp session is closed, an error might have occurred
    If (AutoProxy_hSession <> 0) Then
        AutoProxy_HttpClose AutoProxy_hSession
    End If

AutoProxy_Cleanup:
    On Error GoTo 0

    ' Free any strings received from WinHttp APIs
    If (AutoProxy_IEProxyConfig.AutoProxy_lpszAutoConfigUrl <> 0) Then
        AutoProxy_GlobalFree AutoProxy_IEProxyConfig.AutoProxy_lpszAutoConfigUrl
        AutoProxy_IEProxyConfig.AutoProxy_lpszAutoConfigUrl = 0
    End If
    If (AutoProxy_IEProxyConfig.AutoProxy_lpszProxy <> 0) Then
        AutoProxy_GlobalFree AutoProxy_IEProxyConfig.AutoProxy_lpszProxy
        AutoProxy_IEProxyConfig.AutoProxy_lpszProxy = 0
    End If
    If (AutoProxy_IEProxyConfig.AutoProxy_lpszProxyBypass <> 0) Then
        AutoProxy_GlobalFree AutoProxy_IEProxyConfig.AutoProxy_lpszProxyBypass
        AutoProxy_IEProxyConfig.AutoProxy_lpszProxyBypass = 0
    End If
    If (AutoProxy_ProxyInfo.AutoProxy_lpszProxy <> 0) Then
        AutoProxy_GlobalFree AutoProxy_ProxyInfo.AutoProxy_lpszProxy
        AutoProxy_ProxyInfo.AutoProxy_lpszProxy = 0
    End If
    If (AutoProxy_ProxyInfo.AutoProxy_lpszProxyBypass <> 0) Then
        AutoProxy_GlobalFree AutoProxy_ProxyInfo.AutoProxy_lpszProxyBypass
        AutoProxy_ProxyInfo.AutoProxy_lpszProxyBypass = 0
    End If

    ' Error handling
    If Err.Number <> 0 Then
        ' Unmanaged error
        Err.Raise Err.Number, "AutoProxy:" & Err.source, Err.Description, Err.HelpFile, Err.HelpContext
    ElseIf AutoProxy_Error <> 0 Then
        Err.Raise AutoProxy_Error, "AutoProxy", AutoProxy_ErrorMsg
    End If
#End If
End Sub
