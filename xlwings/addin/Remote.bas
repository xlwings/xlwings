Attribute VB_Name = "Remote"
Option Explicit
Function RunRemotePython( _
    url As String, _
    Optional apiKey As String, _
    Optional exclude As String, _
    Optional headers As Variant, _
    Optional timeout As Integer, _
    Optional proxyServer As String, _
    Optional proxyBypassList As String, _
    Optional proxyUsername As String, _
    Optional proxyPassword As String, _
    Optional enableAutoProxy As String, _
    Optional insecure As String, _
    Optional followRedirects As String _
)

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    ' Config
    ' takes the first value it finds in this order:
    ' func arg, sheet config, directory config, user config
    If exclude = "" Then
        exclude = GetConfig("EXCLUDE")
    End If
    Dim excludeArray As Variant
    excludeArray = Split(exclude, ",")

    If timeout = 0 Then
        timeout = GetConfig("TIMEOUT", 0)
    End If
    If enableAutoProxy = "" Then
        enableAutoProxy = GetConfig("ENABLE_AUTO_PROXY", False)
    End If
    If insecure = "" Then
        insecure = GetConfig("INSECURE", False)
    End If
    If followRedirects = "" Then
        followRedirects = GetConfig("FOLLOW_REDIRECTS", False)
    End If
    If proxyPassword = "" Then
        proxyPassword = GetConfig("PROXY_PASSWORD", "")
    End If
    If proxyUsername = "" Then
        proxyUsername = GetConfig("PROXY_USERNAME", "")
    End If
    If proxyServer = "" Then
        proxyServer = GetConfig("PROXY_SERVER", "")
    End If
    If proxyBypassList = "" Then
        proxyBypassList = GetConfig("PROXY_BYPASS_LIST", "")
    End If
    If apiKey = "" Then
        apiKey = GetConfig("API_KEY", "")
    End If

    ' Request payload
    Dim payload As New Dictionary
    payload.Add "client", "VBA"
    payload.Add "version", "dev"
    
    Dim bookPayload As New Dictionary
    bookPayload.Add "name", ActiveWorkbook.Name
    bookPayload.Add "active_sheet_index", ActiveSheet.Index - 1
    bookPayload.Add "selection", Application.Selection.Address(False, False)
    payload.Add "book", bookPayload
    
    Dim sheetsPayload() As Dictionary
    ReDim sheetsPayload(wb.Worksheets.Count - 1)
    Dim i As Integer
    For i = 1 To wb.Worksheets.Count
        Dim sheetDict As Dictionary
        Set sheetDict = New Dictionary
        sheetDict.Add "name", wb.Worksheets(i).Name
        Dim values As Variant
        If IsInArray(wb.Worksheets(i).Name, excludeArray) Then
            values = Array(Array())
        ElseIf IsEmpty(wb.Worksheets(i).UsedRange.Value) Then
            values = Array(Array())
        Else
            Dim startRow As Integer, startCol As Integer
            Dim nRows As Integer, nCols As Integer
            With wb.Worksheets(i).UsedRange
                startRow = .Row
                startCol = .Column
                nRows = .Rows.Count
                nCols = .Columns.Count
            End With
            With wb.Worksheets(i)
                values = .Range( _
                    .Cells(1, 1), _
                    .Cells(startRow + nRows - 1, startCol + nCols - 1) _
                ).Value
                If nRows = 1 And nCols = 1 Then
                    values = Array(Array(values))
                End If
            End With
        End If
        sheetDict.Add "values", values
        Set sheetsPayload(i - 1) = sheetDict
    Next
    payload.Add "sheets", sheetsPayload
    
    Dim myRequest As New WebRequest
    Set myRequest.Body = payload

    ' Debug.Print myRequest.Body

    ' Headers
    ' Expected as Dictionary and currently not supported via xlwings.conf
    ' Providing the Authorization header will ignore the API_KEY
    Dim authHeader As Boolean
    authHeader = False
    If Not IsMissing(headers) Then
        Dim myKey as Variant
        For Each myKey in headers.myKeys
            myRequest.AddHeader CStr(myKey), headers(myKey)
        Next
        If headers.Exists("Authorization") Then
            authHeader = True
        End If
    End If

    If authHeader = False Then
        If apiKey <> "" Then
            myRequest.AddHeader "Authorization", apiKey
        End If
    End If

    ' API call
    myRequest.Method = WebMethod.HttpPost
    myRequest.Format = WebFormat.Json

    Dim myClient As New WebClient
    myClient.BaseUrl = url
    If timeout <> 0 Then
        myClient.TimeoutMs = timeout
    End If
    If proxyBypassList <> "" Then
        myClient.proxyBypassList = proxyBypassList
    End If
    If proxyServer <> "" Then
        myClient.proxyServer = proxyServer
    End If
    If proxyUsername <> "" Then
        myClient.proxyUsername = proxyUsername
    End If
    If proxyPassword <> "" Then
        myClient.proxyPassword = proxyPassword
    End If
    If enableAutoProxy <> False Then
        myClient.enableAutoProxy = enableAutoProxy
    End If
    If insecure <> False Then
        myClient.insecure = insecure
    End If
    If followRedirects <> False Then
        myClient.followRedirects = followRedirects
    End If

    Dim response As WebResponse
    Set response = myClient.Execute(myRequest)
    
    ' Debug.Print response.Content
    
    ' Parse JSON response and run functions
    If response.StatusCode = WebStatusCode.Ok Then
        Dim action As Dictionary
        For Each action In response.Data("actions")
            Application.Run action("func"), wb, action
        Next
    Else
        MsgBox "Server responded with error " & response.StatusCode, vbCritical, "Error"
    End If

End Function

' Helpers
Function GetRange(wb, action)
    If action("row_count") = 1 And action("column_count") = 1 Then
        Set GetRange = wb.Worksheets( _
            action("sheet_position") + 1).Cells(action("start_row") + 1, _
            action("start_column") + 1 _
        )
    Else
        With wb.Worksheets(action("sheet_position") + 1)
            Set GetRange = .Range( _
                .Cells(action("start_row") + 1, action("start_column") + 1), _
                .Cells( _
                    action("start_row") + action("row_count"), _
                    action("start_column") + action("column_count") _
                ) _
            )
        End With
    End If
End Function

Function Utf8ToUtf16(ByVal strText As String) As String
    ' macOs only: apparently, Excel uses UTF-16 to represent string literals
    ' Taken from https://stackoverflow.com/a/64624336/918626
    Dim i&, l1&, l2&, l3&, l4&, l&
    For i = 1 To Len(strText)
        l1 = Asc(Mid(strText, i, 1))
        If i + 1 <= Len(strText) Then l2 = Asc(Mid(strText, i + 1, 1))
        If i + 2 <= Len(strText) Then l3 = Asc(Mid(strText, i + 2, 1))
        If i + 3 <= Len(strText) Then l4 = Asc(Mid(strText, i + 3, 1))
        Select Case l1
        Case 1 To 127
            l = l1
        Case 194 To 223
            l = ((l1 And &H1F) * 2 ^ 6) Or (l2 And &H3F)
            i = i + 1
        Case 224 To 239
            l = ((l1 And &HF) * 2 ^ 12) Or ((l2 And &H3F) * 2 ^ 6) Or (l3 And &H3F)
            i = i + 2
        Case 240 To 255
            l = ((l1 And &H7) * 2 ^ 18) Or ((l2 And &H3F) * 2 ^ 12) Or ((l3 And &H3F) * 2 ^ 6) Or (l4 And &H3F)
            i = i + 4
        Case Else
            l = 63 ' question mark
        End Select
        Utf8ToUtf16 = Utf8ToUtf16 & IIf(l < 55296, WorksheetFunction.Unichar(l), "?")
    Next i
End Function

Function HexToRgb(ByVal HexColor As String) As Variant
    Dim red As String, green As String, blue As String
    HexColor = Replace(HexColor, "#", "")
    red = Val("&H" & Mid(HexColor, 1, 2))
    green = Val("&H" & Mid(HexColor, 3, 2))
    blue = Val("&H" & Mid(HexColor, 5, 2))
    HexToRgb = RGB(red, green, blue)
End Function

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    ' Based on https://stackoverflow.com/a/38268261/918626
    If IsEmpty(arr) Then
        IsInArray = False
        Exit Function
    End If
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        If Trim(arr(i)) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

' Functions
Sub setValues(wb, action)
    Dim arr() As Variant
    Dim i As Long, j As Long
    Dim values As Collection, valueRow As Collection

    Set values = action("values")
    ReDim arr(values.Count, values(1).Count)

    For i = 1 To values.Count
        Set valueRow = values(i)
        For j = 1 To valueRow.Count
            On Error Resume Next
                ' TODO: will be replaced when backend sends location of dates
                arr(i - 1, j - 1) = WebHelpers.ParseIso(valueRow(j))
            If Err.Number <> 0 Then
                #If Mac Then
                If Application.IsText(valueRow(j)) Then
                    arr(i - 1, j - 1) = Utf8ToUtf16(valueRow(j))
                Else
                    arr(i - 1, j - 1) = valueRow(j)
                End If
                #Else
                    arr(i - 1, j - 1) = valueRow(j)
                #End If
            End If
            On Error GoTo 0
        Next j
    Next i
    GetRange(wb, action).Value = arr
End Sub

Sub clearContents(wb, action)
    GetRange(wb, action).clearContents
End Sub

Sub addSheet(wb, action)
    Dim mysheet As Worksheet
    Set mysheet = wb.Sheets.Add
    mysheet.Move after:=Worksheets(action("args")(1) + 1)
End Sub

Sub setSheetName(wb, action)
    wb.Sheets(action("sheet_position") + 1).Name = action("args")(1)
End Sub

Sub setAutofit(wb, action)
    If action("args")(1) = "columns" Then
        GetRange(wb, action).Columns.AutoFit
    Else
        GetRange(wb, action).Rows.AutoFit
    End If
End Sub

Sub setRangeColor(wb, action)
    GetRange(wb, action).Interior.Color = HexToRgb(action("args")(1))
End Sub

Sub activateSheet(wb, action)
    wb.Sheets(action("args")(1) + 1).Activate
End Sub
