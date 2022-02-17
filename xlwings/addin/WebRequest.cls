VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' WebRequest v4.1.6
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' `WebRequest` is used to create detailed requests
' (including formatting, querystrings, headers, cookies, and much more).
'
' Usage:
' ```VB.net
' Dim Request As New WebRequest
' Request.Resource = "users/{Id}"
'
' Request.Method = WebMethod.HttpPut
' Request.RequestFormat = WebFormat.UrlEncoded
' Request.ResponseFormat = WebFormat.Json
'
' Dim Body As New Dictionary
' Body.Add "name", "Tim"
' Body.Add "project", "VBA-Web"
' Set Request.Body = Body
'
' Request.AddUrlSegment "Id", 123
' Request.AddQuerystringParam "api_key", "abcd"
' Request.AddHeader "Authorization", "Token ..."
'
' ' -> PUT (Client.BaseUrl)users/123?api_key=abcd
' '    Authorization: Token ...
' '
' '    name=Tim&project=VBA-Web
' ```
'
' Errors:
' 11020 / 80042b0c / -2147210484 - Cannot add body parameter to non-Dictionary
'
' @class WebRequest
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private web_pRequestFormat As WebFormat
Private web_pResponseFormat As WebFormat
Private web_pCustomRequestFormat As String
Private web_pCustomResponseFormat As String
Private web_pBody As Variant
Private web_pConvertedBody As Variant
Private web_pContentType As String
Private web_pAccept As String
Private web_pContentLength As Long
Private web_pId As String

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

''
' Set the request's portion of the url to be appended to the client's BaseUrl.
' Can include Url Segments for dynamic values
' and Querystring parameters are smart enough to be appended to existing querystring
' (or added to resource if there isn't an existing querystring).
'
' @example
' ```VB.net
' Dim Client As New WebClient
' Client.BaseUrl = "https://api.example.com/"
'
' Dim Request As New WebRequest
' Request.Resource = "messages"
'
' ' -> Url: https://api.example.com/messages
'
' Request.Resource = "messages/{id}?a=1"
' Request.AddUrlSegment "id", 123
' Request.AddQuerystringParam "b", 2
'
' ' -> Url: https://api.example.com/messages/123?a=1&b=2
' ```
'
' @property Resource
' @type String
''
Public Resource As String

''
' Set the HTTP method to be used for the request:
' GET, POST, PUT, PATCH, DELETE
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.Method = WebMethod.HttpGet
' Request.Method = WebMethod.HttpPost
' ' or HttpPut / HttpPatch / HttpDelete
' ```
'
' @property Method
' @type WebMethod
''
Public Method As WebMethod

''
' _Note_ To add headers, use [`AddHeader`](#/WebRequest/AddHeader).
'
' `Collection` of Headers to include with request,
' stored as `KeyValue` (`Dictionary: {Key: "...", Value: "..."}`).
'
' @property Headers
' @type Collection
''
Public headers As Collection

''
' _Note_ To add querystring parameters, use [`AddQuerystringParam`](#/WebRequest/AddQuerystringParam).
'
' `Collection` of querystring parameters to include with request,
' stored as `KeyValue` (`Dictionary: {Key: "...", Value: "..."}`).
'
' @property QuerystringParams
' @type Collection
''
Public QuerystringParams As Collection

''
' _Note_ To add Url Segments, use [`AddUrlSegment`](#/WebRequest/AddUrlSegment)
'
' Url Segments are used to easily add dynamic values to `Resource`.
' Create a Url Segement in `Resource` with curly brackets and then
' replace with dynamic value with [`AddUrlSegment`](#AddUrlSegment).
'
' @example
' ```VB.net
' Dim Request As New WebRequest
'
' Dim User As String
' Dim Id As Long
' User = "Tim"
' Id = 123
'
' ' OK: Use string concatenation for dynamic values
' Request.Resource = User & "/messages/" & Id
'
' ' BETTER: Use Url Segments for dynamic values
' Request.Resource = "{User}/messages/{Id}"
' Request.AddUrlSegment "User", User
' Request.AddUrlSegment "Id", Id
'
' Request.FormattedResource ' = "Tim/messages/123"
' ```
'
' @property UrlSegments
' @type Dictionary
''
Public UrlSegments As Dictionary

''
' _Note_ To add cookies, use [`AddCookie`](#/WebRequest/AddCookie).
'
' `Collection` of cookies to include with request,
' stored as `KeyValue` (`Dictionary: {Key: "...", Value: "..."}`).
'
' @property Cookies
' @type Collection
''
Public Cookies As Collection

''
' User agent to use with request
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.UserAgent = "Mozilla/5.0"
'
' ' -> (Header) User-Agent: Mozilla/5.0
' ```
'
' @property UserAgent
' @type String
' @default "VBA-Web v#.#.# (https://github.com/VBA-tools/VBA-Web)"
''
Public UserAgent As String

''
' Set `RequestFormat`, `ResponseFormat`, and `Content-Type` and `Accept`
' headers for the `WebRequest`
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.Format = WebFormat.Json
' ' -> Request.RequestFormat = WebFormat.Json
' '    Request.ResponseFormat = WebFormat.Json
' '    (Header) Content-Type: application/json
' '    (Header) Accept: application/json
' ```
'
' @property Format
' @type WebFormat
''
Public Property Get Format() As WebFormat
    Format = RequestFormat
End Property
Public Property Let Format(Value As WebFormat)
    Me.RequestFormat = Value
    Me.ResponseFormat = Value
End Property

''
' Set the format to use for converting the response `Body` to string and for the `Content-Type` header
'
' _Note_ If `WebFormat.Custom` is used, the [`CustomRequestFormat`](#/WebRequest/CustomRequestFormat) must be set.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.Body = Array("A", "B", "C")
'
' Request.RequestFormat = WebFormat.Json
'
' ' -> (Header) Content-Type: application/json
' ' -> Convert Body to JSON string
' Request.Body ' = "["A","B","C"]"
' ```
'
' @property RequestFormat
' @type WebFormat
' @default WebFormat.Json
''
Public Property Get RequestFormat() As WebFormat
    RequestFormat = web_pRequestFormat
End Property
Public Property Let RequestFormat(Value As WebFormat)
    If Value <> web_pRequestFormat Then
        web_pRequestFormat = Value

        ' Clear cached converted body
        web_pConvertedBody = Empty
    End If
End Property

''
' Set the format to use for converting the response `Content` to `Data` and for the `Accept` header
'
' _Note_ If `WebFormat.Custom` is used, the [`CustomResponseFormat`](#/WebRequest/CustomResponseFormat) must be set.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.ResponseFormat = WebFormat.Json
'
' ' -> (Header) Accept: application/json
'
' Dim Response As WebResponse
' ' ... from Execute
' Response.Content = "{""message"":""Howdy!""}"
'
' ' -> Parse Content to JSON Dictionary
' Debug.Print Response.Data("message") ' -> "Howdy!"
' ```
'
' @property ResponseFormat
' @type WebFormat
' @default WebFormat.Json
''
Public Property Get ResponseFormat() As WebFormat
    ResponseFormat = web_pResponseFormat
End Property
Public Property Let ResponseFormat(Value As WebFormat)
    If Value <> web_pResponseFormat Then
        web_pResponseFormat = Value

        ' Clear cached converted body
        web_pConvertedBody = Empty
    End If
End Property

''
' Use converter registered with [`WebHelpers.RegisterConverter`](#/WebHelpers/RegisterConverter)
' to convert `Body` to string and set `Content-Type` header.
'
' (Automatically sets `RequestFormat` to `WebFormat.Custom`)
'
' @example
' ```VB.net
' WebHelpers.RegisterConverter "csv", "text/csv", "Module.ConvertToCsv", "Module.ParseCsv"
'
' Dim Request As New WebRequest
' Request.CustomRequestFormat = "csv"
'
' ' -> (Header) Content-Type: text/csv
' ' -> Body converted to string with Module.ConvertToCsv
' ```
'
' @property CustomRequestFormat
' @type String
''
Public Property Get CustomRequestFormat() As String
    CustomRequestFormat = web_pCustomRequestFormat
End Property
Public Property Let CustomRequestFormat(Value As String)
    If Value <> web_pCustomRequestFormat Then
        web_pCustomRequestFormat = Value

        ' Clear cached converted body
        web_pConvertedBody = Empty

        If Value <> "" Then
            web_pRequestFormat = WebFormat.Custom
        End If
    End If
End Property

''
' Use converter registered with [`WebHelpers.RegisterConverter`](#/WebHelpers/RegisterConverter)
' to convert the response `Content` to `Data` and set `Accept` header.
'
' (Automatically sets `ResponseFormat` to `WebFormat.Custom`)
'
' @example
' ```VB.net
' WebHelpers.RegisterConverter "csv", "text/csv", "Module.ConvertToCsv", "Module.ParseCsv"
'
' Dim Request As New WebRequest
' Request.CustomResponseFormat = "csv"
'
' ' -> (Header) Accept: text/csv
' ' -> WebResponse Content converted Data with Module.ParseCsv
' ```
'
' @property CustomResponseFormat
' @type String
''
Public Property Get CustomResponseFormat() As String
    CustomResponseFormat = web_pCustomResponseFormat
End Property
Public Property Let CustomResponseFormat(Value As String)
    If Value <> web_pCustomResponseFormat Then
        web_pCustomResponseFormat = Value

        ' Clear cached converted body
        web_pConvertedBody = Empty

        If Value <> "" Then
            ResponseFormat = WebFormat.Custom
        End If
    End If
End Property

''
' Set automatically from `RequestFormat` or `CustomRequestFormat`,
' but can be overriden to set `Content-Type` header for request.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.ContentType = "text/csv"
'
' ' -> (Header) Content-Type: text/csv
' ```
'
' @property ContentType
' @type String
' @default Media-type of request format
''
Public Property Get ContentType() As String
    If web_pContentType <> "" Then
        ContentType = web_pContentType
    Else
        ContentType = WebHelpers.FormatToMediaType(Me.RequestFormat, Me.CustomRequestFormat)
    End If
End Property
Public Property Let ContentType(Value As String)
    web_pContentType = Value
End Property

''
' Set automatically from `ResponseFormat` or `CustomResponseFormat`,
' but can be overriden to set `Accept` header for request.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.Accept = "text/csv"
'
' ' -> (Header) Accept: text/csv
' ```
'
' @property Accept
' @type String
' @default Media-type of response format
''
Public Property Get Accept() As String
    If web_pAccept <> "" Then
        Accept = web_pAccept
    Else
        Accept = WebHelpers.FormatToMediaType(Me.ResponseFormat, Me.CustomResponseFormat)
    End If
End Property
Public Property Let Accept(Value As String)
    web_pAccept = Value
End Property

''
' Set automatically by length of `Body`,
' but can be overriden to set `Content-Length` header for request.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.ContentLength = 200
'
' ' -> (Header) Content-Length: 200
' ```
'
' @property ContentLength
' @type Long
' @default Length of `Body`
''
Public Property Get ContentLength() As Long
    If web_pContentLength >= 0 Then
        ContentLength = web_pContentLength
    Else
        ContentLength = Len(Me.Body)
    End If
End Property
Public Property Let ContentLength(Value As Long)
    web_pContentLength = Value
End Property

''
' - Get: Body value converted to string using `RequestFormat` or `CustomRequestFormat`
' - Let: Use `String` or `Array` for Body
' - Set: Use `Collection`, `Dictionary`, or `Object` for Body
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.RequestFormat = WebFormat.Json
'
' ' Let: String|Array
' Request.Body = "text"
' Debug.Print Request.Body ' -> "text"
'
' Request.Body = Array("A", "B", "C")
' Debug.Print Request.Body ' -> "["A","B","C"]"
'
' ' Set: Collection|Dictionary|Object
' Dim Body As Object
' Set Body = New Collection
' Body.Add "Howdy!"
' Set Request.Body = Body
' Debug.Print Request.Body ' -> "["Howdy!"]"
'
' Set Body = New Dictionary
' Body.Add "a", 123
' Body.Add "b", 456
' Set Request.Body = Body
' Debug.Print Request.Body ' -> "{"a":123,"b":456}"
' ```
'
' @property Body
' @type String|Array|Collection|Dictionary|Variant
''
Public Property Get Body() As Variant
    If Not VBA.IsEmpty(web_pBody) Then
        If VBA.VarType(web_pBody) = vbString Then
            Body = web_pBody
        ElseIf IsEmpty(web_pConvertedBody) Then
            ' Convert body and cache
            Body = WebHelpers.ConvertToFormat(web_pBody, Me.RequestFormat, Me.CustomRequestFormat)
            web_pConvertedBody = Body
        Else
            Body = web_pConvertedBody
        End If
    End If
End Property
Public Property Let Body(Value As Variant)
    web_pConvertedBody = Empty
    web_pBody = Value
End Property
Public Property Set Body(Value As Variant)
    web_pConvertedBody = Empty
    Set web_pBody = Value
End Property

''
' Get `Resource` with Url Segments replaced and Querystring added.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.Resource = "examples/{Id}"
' Request.AddUrlSegment "Id", 123
' Request.AddQuerystringParam "message", "Hello"
'
' Debug.Print Request.FormattedResource ' -> "examples/123?message=Hello"
' ```
'
' @property FormattedResource
' @type String
''
Public Property Get FormattedResource() As String
    Dim web_Segment As Variant
    Dim web_Encoding As UrlEncodingMode

    FormattedResource = Me.Resource

    ' Replace url segments
    For Each web_Segment In Me.UrlSegments.Keys
        FormattedResource = VBA.Replace(FormattedResource, "{" & web_Segment & "}", WebHelpers.UrlEncode(Me.UrlSegments(web_Segment)))
    Next web_Segment

    ' Add querystring
    If Me.QuerystringParams.Count > 0 Then
        If VBA.InStr(FormattedResource, "?") <= 0 Then
            FormattedResource = FormattedResource & "?"
        Else
            FormattedResource = FormattedResource & "&"
        End If

        ' For querystrings, W3C defines form-urlencoded as the required encoding,
        ' but the treatment of space -> "+" (rather than "%20") can cause issues
        '
        ' If the request format is explicitly form-urlencoded, use FormUrlEncoding (space -> "+")
        ' otherwise, use subset of RFC 3986 and form-urlencoded that should work for both cases (space -> "%20")
        If Me.RequestFormat = WebFormat.FormUrlEncoded Then
            web_Encoding = UrlEncodingMode.FormUrlEncoding
        Else
            web_Encoding = UrlEncodingMode.QueryUrlEncoding
        End If
        FormattedResource = FormattedResource & WebHelpers.ConvertToUrlEncoded(Me.QuerystringParams, EncodingMode:=web_Encoding)
    End If
End Property

''
' @internal
' @property Id
' @type String
''
Public Property Get Id() As String
    If web_pId = "" Then: web_pId = WebHelpers.CreateNonce
    Id = web_pId
End Property

' ============================================= '
' Public Methods
' ============================================= '

''
' Add header to be sent with request.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.AddHeader "Authentication", "Bearer ..."
'
' ' -> (Header) Authorization: Bearer ...
' ```
'
' @method AddHeader
' @param {String} Key
' @param {Variant} Value
''
Public Sub AddHeader(Key As String, Value As Variant)
    Me.headers.Add WebHelpers.CreateKeyValue(Key, Value)
End Sub

''
' Add/replace header to be sent with request.
' `SetHeader` should be used for headers that can only be included once with a request
' (e.g. Authorization, Content-Type, etc.).
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.AddHeader "Authorization", "A..."
' Request.AddHeader "Authorization", "B..."
'
' ' -> Headers:
' '    Authorization: A...
' '    Authorization: B...
'
' Request.SetHeader "Authorization", "C..."
'
' ' -> Headers:
' '    Authorization: C...
' ```
'
' @method SetHeader
' @param {String} Key
' @param {Variant} Value
''
Public Sub SetHeader(Key As String, Value As Variant)
    WebHelpers.AddOrReplaceInKeyValues Me.headers, Key, Value
End Sub

''
' Url Segments are used to easily add dynamic values to `Resource`.
' Create a Url Segement in `Resource` with curly brackets and then
' replace with dynamic value with `AddUrlSegment`.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Dim User As String
' Dim Id As Long
'
' User = "Tim"
' Id = 123
'
' ' OK: Use string concatenation for dynamic values
' Request.Resource = User & "/messages/" & Id
'
' ' BETTER: Use Url Segments for dynamic values
' Request.Resource = "{User}/messages/{Id}"
' Request.AddUrlSegment "User", User
' Request.AddUrlSegment "Id", Id
'
' Debug.Print Request.FormattedResource ' > "Tim/messages/123"
' ```
'
' @method AddUrlSegment
' @param {String} Key
' @param {String} Value
''
Public Sub AddUrlSegment(Segment As String, Value As Variant)
    Me.UrlSegments.Item(Segment) = Value
End Sub

''
' Add querysting parameter to be used in `FormattedResource` for request.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.Resource = "messages"
' Request.AddQuerystringParam "from", "Tim"
'
' Request.FormattedResource ' = "messages?from=Tim"
' ```
'
' @method AddQuerystringParam
' @param {String} Key
' @param {Variant} Value
''
Public Sub AddQuerystringParam(Key As String, Value As Variant)
    Me.QuerystringParams.Add WebHelpers.CreateKeyValue(Key, Value)
End Sub

''
' Add cookie to be sent with request.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.AddCookie "a", "abc"
' Request.AddCookie "b", 123
'
' ' -> (Header) Cookie: a=abc; b=123;
' ```
'
' @method AddCookie
' @param {String} Key
' @param {Variant} Value
''
Public Sub AddCookie(Key As String, Value As Variant)
    Me.Cookies.Add WebHelpers.CreateKeyValue( _
        web_EncodeCookieName(Key), _
        WebHelpers.UrlEncode(Value, EncodingMode:=UrlEncodingMode.CookieUrlEncoding) _
    )
End Sub

''
' Add `Key-Value` to `Body`.
' `Body` must be a `Dictionary` (if it's an `Array` or `Collection` an error is thrown)
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.Format = WebFormat.Json
'
' Request.AddBodyParameter "a", 123
' Debug.Print Request.Body ' -> "{"a":123}"
'
' ' Can add parameters to existing Dictionary
' Dim Body As New Dictionary
' Body.Add "a", 123
'
' Set Request.Body = Body
' Request.AddBodyParameter "b", 456
'
' Debug.Print Request.Body ' -> "{"a":123,"b":456}"
' ```
'
' @method AddBodyParameter
' @param {Variant} Key
' @param {Variant} Value
' @throws 11020 / 80042b0c / -2147210484 - Cannot add body parameter to non-Dictionary
''
Public Sub AddBodyParameter(Key As Variant, Value As Variant)
    If VBA.IsEmpty(web_pBody) Then
        Set web_pBody = New Dictionary
    ElseIf Not TypeOf web_pBody Is Dictionary Then
        Dim web_ErrorDescription As String
        web_ErrorDescription = "Cannot add body parameter to non-Dictionary Body (existing Body must be of type Dictionary)"

        WebHelpers.LogError web_ErrorDescription, "WebRequest.AddBodyParameter", 11020 + vbObjectError
        Err.Raise 11020 + vbObjectError, "WebRequest.AddBodyParameter", web_ErrorDescription
    End If

    If VBA.IsObject(Value) Then
        Set web_pBody(Key) = Value
    Else
        web_pBody(Key) = Value
    End If

    ' Clear cached converted body
    web_pConvertedBody = Empty
End Sub

''
' Prepare request for execution
'
' @internal
' @method Prepare
''
Public Sub Prepare()
    ' Add/replace general headers for request
    SetHeader "User-Agent", Me.UserAgent
    SetHeader "Accept", Me.Accept
    If Me.Method <> WebMethod.HttpGet Then
        SetHeader "Content-Type", Me.ContentType
        SetHeader "Content-Length", VBA.CStr(Me.ContentLength)
    End If
End Sub

''
' Clone request
'
' @internal
' @method Clone
' @return {WebRequest}
''
Public Function Clone() As WebRequest
    Set Clone = New WebRequest

    ' Note: Clone underlying for properties with default values
    Clone.Resource = Me.Resource
    Clone.Method = Me.Method
    Clone.UserAgent = Me.UserAgent
    Clone.Accept = web_pAccept
    Clone.ContentType = web_pContentType
    Clone.ContentLength = web_pContentLength
    Clone.RequestFormat = Me.RequestFormat
    Clone.ResponseFormat = Me.ResponseFormat
    Clone.CustomRequestFormat = Me.CustomRequestFormat
    Clone.CustomResponseFormat = Me.CustomResponseFormat

    Set Clone.headers = WebHelpers.CloneCollection(Me.headers)
    Set Clone.QuerystringParams = WebHelpers.CloneCollection(Me.QuerystringParams)
    Set Clone.UrlSegments = WebHelpers.CloneDictionary(Me.UrlSegments)
    Set Clone.Cookies = WebHelpers.CloneCollection(Me.Cookies)

    If VBA.IsObject(web_pBody) Then
        Set Clone.Body = web_pBody
    Else
        Clone.Body = web_pBody
    End If
End Function

''
' Create WebRequest from options
'
' @method CreateFromOptions
' @param {Dictionary} Options
' @param {Collection} [Options.Headers] Collection of `KeyValue`
' @param {Collection} [Options.Cookies] Collection of `KeyValue`
' @param {Collection} [Options.QuerystringParams] Collection of `KeyValue`
' @param {Dictionary} [Options.UrlSegments]
''
Public Sub CreateFromOptions(Options As Dictionary)
    If Not Options Is Nothing Then
        If Options.Exists("Headers") Then
            Set Me.headers = Options("Headers")
        End If
        If Options.Exists("Cookies") Then
            Set Me.Cookies = Options("Cookies")
        End If
        If Options.Exists("QuerystringParams") Then
            Set Me.QuerystringParams = Options("QuerystringParams")
        End If
        If Options.Exists("UrlSegments") Then
            Set Me.UrlSegments = Options("UrlSegments")
        End If
    End If
End Sub

' ============================================= '
' Private Functions
' ============================================= '

' Encode cookie name
'
' References:
' - RFC 6265 https://tools.ietf.org/html/rfc6265
Private Function web_EncodeCookieName(web_CookieName As Variant) As String
    Dim web_CookieVal As String
    Dim web_StringLen As Long

    web_CookieVal = VBA.CStr(web_CookieName)
    web_StringLen = VBA.Len(web_CookieVal)

    If web_StringLen > 0 Then
        Dim web_Result() As String
        Dim web_i As Long
        Dim web_CharCode As Integer
        Dim web_Char As String
        ReDim web_Result(web_StringLen)

        ' ALPHA / DIGIT / "!" / "#" / "$" / "&" / "'" / "*" / "+" / "-" / "." / "^" / "_" / "`" / "|" / "~"
        ' Note: "%" is allowed in spec, but is currently excluded due to parsing issues

        ' Loop through string characters
        For web_i = 1 To web_StringLen
            ' Get character and ascii code
            web_Char = VBA.Mid$(web_CookieVal, web_i, 1)
            web_CharCode = VBA.Asc(web_Char)

            Select Case web_CharCode
                Case 65 To 90, 97 To 122
                    ' ALPHA
                    web_Result(web_i) = web_Char
                Case 48 To 57
                    ' DIGIT
                    web_Result(web_i) = web_Char
                Case 33, 35, 36, 38, 39, 42, 43, 45, 46, 94, 95, 96, 124, 126
                    ' "!" / "#" / "$" / "&" / "'" / "*" / "+" / "-" / "." / "^" / "_" / "`" / "|" / "~"
                    web_Result(web_i) = web_Char

                Case 0 To 15
                    web_Result(web_i) = "%0" & VBA.Hex(web_CharCode)
                Case Else
                    web_Result(web_i) = "%" & VBA.Hex(web_CharCode)
            End Select
        Next web_i

        web_EncodeCookieName = VBA.Join$(web_Result, "")
    End If
End Function

Private Sub Class_Initialize()
    ' Set default values
    Me.RequestFormat = WebFormat.Json
    Me.ResponseFormat = WebFormat.Json
    Me.UserAgent = WebUserAgent

    Set Me.headers = New Collection
    Set Me.QuerystringParams = New Collection
    Set Me.UrlSegments = New Dictionary
    Set Me.Cookies = New Collection
    Me.ContentLength = -1
End Sub
