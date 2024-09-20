Attribute VB_Name = "Remote"
Option Explicit
' Keep legacy alias RunRemotePython in sync
Function RunServerPython( _
    url As String, _
    Optional auth As String, _
    Optional apiKey As String, _
    Optional include As String, _
    Optional exclude As String, _
    Optional headers As Variant, _
    Optional timeout As Long, _
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
    If include = "" Then
        include = GetConfig("INCLUDE")
    End If
    Dim includeArray() As String
    If include <> "" Then
        includeArray = Split(include, ",")
    End If

    If exclude = "" Then
        exclude = GetConfig("EXCLUDE")
    End If
    Dim excludeArray() As String
    If exclude <> "" Then
        excludeArray = Split(exclude, ",")
    End If

    If include <> "" And exclude <> "" Then
        MsgBox "Either use 'include' or 'exclude', but not both!", vbCritical
        Exit Function
    End If

    If include <> "" Then
        Dim i As Integer
        For i = 1 To wb.Worksheets.Count
            If Not IsInArray(wb.Worksheets(i).Name, includeArray) Then
                ReDim Preserve excludeArray(0 To i)
                excludeArray(i) = wb.Worksheets(i).Name
            End If
        Next
    End If

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
    If apiKey = "" Then  ' Deprecated: replaced by "auth"
        apiKey = GetConfig("API_KEY", "")
    End If
    If auth = "" Then
        auth = GetConfig("AUTH", "")
    End If

    ' Request payload
    Dim payload As New Dictionary
    payload.Add "client", "VBA"
    payload.Add "version", XLWINGS_VERSION
    
    Dim bookPayload As New Dictionary
    bookPayload.Add "name", ActiveWorkbook.Name
    bookPayload.Add "active_sheet_index", ActiveSheet.Index - 1
    If TypeOf Selection Is Range Then
        bookPayload.Add "selection", Application.Selection.Address(False, False)
    Else
        bookPayload.Add "selection", Null
    End If
    payload.Add "book", bookPayload

    ' Names
    Dim myname As Name
    Dim mynames() As Dictionary
    Dim nNames As Integer
    Dim namedRangeCount As Integer
    Dim iName As Integer

    nNames = wb.Names.Count
    namedRangeCount = 0

    If nNames > 0 Then
        For iName = 1 To nNames
            Set myname = wb.Names(iName)
            Dim nameDict As Dictionary
            Set nameDict = New Dictionary
            Dim isNamedRange As Boolean
            Dim testRange As Range
            Dim isBookScope As Boolean
            nameDict.Add "name", myname.Name
            isNamedRange = False
            On Error Resume Next
                Set testRange = myname.RefersToRange
                If Err.Number = 0 Then isNamedRange = True
            On Error GoTo 0
            If isNamedRange Then
                If TypeOf myname.Parent Is Workbook Then isBookScope = True Else isBookScope = False
                nameDict.Add "sheet_index", myname.RefersToRange.Parent.Index - 1
                nameDict.Add "address", myname.RefersToRange.Address(False, False)
                nameDict.Add "book_scope", isBookScope
                If isBookScope = True Then
                    nameDict.Add "scope_sheet_name", Null
                    nameDict.Add "scope_sheet_index", Null
                Else
                    nameDict.Add "scope_sheet_name", myname.Parent.Name
                    nameDict.Add "scope_sheet_index", myname.Parent.Index - 1
                End If
                ReDim Preserve mynames(namedRangeCount)
                Set mynames(namedRangeCount) = nameDict
                namedRangeCount = namedRangeCount + 1
            End If
        Next
        If namedRangeCount > 0 Then
            payload.Add "names", mynames
        Else
            payload.Add "names", Array()
        End If
    Else
        payload.Add "names", Array()
    End If

    Dim sheetsPayload() As Dictionary
    ReDim sheetsPayload(wb.Worksheets.Count - 1)
    For i = 1 To wb.Worksheets.Count
        Dim sheetDict As Dictionary
        Set sheetDict = New Dictionary
        sheetDict.Add "name", wb.Worksheets(i).Name

        ' Pictures
        Dim pic As Shape
        Dim pics() As Dictionary
        Dim nShapes As Integer
        Dim iShape As Integer
        Dim iPic As Integer
        nShapes = wb.Worksheets(i).Shapes.Count
        If (nShapes > 0) And Not (IsInArray(wb.Worksheets(i).Name, excludeArray)) Then
            iPic = 0
            For iShape = 1 To nShapes
                Set pic = wb.Worksheets(i).Shapes(iShape)
                If pic.Type = msoPicture Then
                    ReDim Preserve pics(iPic)
                    Dim picDict As Dictionary
                    Set picDict = New Dictionary
                    picDict.Add "name", pic.Name
                    picDict.Add "height", pic.Height
                    picDict.Add "width", pic.Width
                    Set pics(iPic) = picDict
                    iPic = iPic + 1
                End If
            Next
            sheetDict.Add "pictures", pics
        Else
            sheetDict.Add "pictures", Array()
        End If

        ' Tables
        Dim table As ListObject
        Dim tables() As Dictionary
        Dim nTables As Integer
        Dim iTable As Integer
        nTables = wb.Worksheets(i).ListObjects.Count
        If (nTables > 0) And Not (IsInArray(wb.Worksheets(i).Name, excludeArray)) Then
            For iTable = 1 To nTables
                Set table = wb.Worksheets(i).ListObjects(iTable)
                ReDim Preserve tables(iTable - 1)
                Dim tableDict As Dictionary
                Set tableDict = New Dictionary
                tableDict.Add "name", table.Name
                tableDict.Add "range_address", table.Range.Address
                If table.ShowHeaders Then
                    tableDict.Add "header_row_range_address", table.HeaderRowRange.Address
                Else
                    tableDict.Add "header_row_range_address", Null
                End If
                If table.DataBodyRange Is Nothing Then
                    tableDict.Add "data_body_range_address", Null
                Else
                    tableDict.Add "data_body_range_address", table.DataBodyRange.Address
                End If
                If table.ShowTotals Then
                    tableDict.Add "total_row_range_address", table.TotalsRowRange.Address
                Else
                    tableDict.Add "total_row_range_address", Null
                End If
                tableDict.Add "show_headers", table.ShowHeaders
                tableDict.Add "show_totals", table.ShowTotals
                tableDict.Add "table_style", table.TableStyle.Name
                tableDict.Add "show_autofilter", table.ShowAutoFilter
                Set tables(iTable - 1) = tableDict
            Next
            sheetDict.Add "tables", tables
        Else
            sheetDict.Add "tables", Array()
        End If

        ' Values
        Dim values As Variant
        If IsInArray(wb.Worksheets(i).Name, excludeArray) Then
            values = Array(Array())
        ElseIf IsEmpty(wb.Worksheets(i).UsedRange.Value) Then
            values = Array(Array())
        Else
            Dim startRow As Integer, startCol As Integer
            Dim nRows As Integer, nCols As Integer
            Dim myUsedRange As Range
            With wb.Worksheets(i).UsedRange
                startRow = .Row
                startCol = .Column
                nRows = .Rows.Count
                nCols = .Columns.Count
            End With
            With wb.Worksheets(i)
                Set myUsedRange = .Range( _
                    .Cells(1, 1), _
                    .Cells(startRow + nRows - 1, startCol + nCols - 1) _
                )
                values = myUsedRange.Value
                If myUsedRange.Count = 1 Then
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
        Dim myKey As Variant
        For Each myKey In headers.Keys
            myRequest.AddHeader CStr(myKey), headers(myKey)
        Next
        If headers.Exists("Authorization") Then
            authHeader = True
        End If
    End If

    If authHeader = False Then
        If apiKey <> "" Then  ' Deprecated: replaced by "auth"
            myRequest.AddHeader "Authorization", apiKey
        End If
        If auth <> "" Then
            myRequest.AddHeader "Authorization", auth
        End If
    End If

    ' API call
    myRequest.Method = WebMethod.HttpPost
    myRequest.Format = WebFormat.Json

    Dim myClient As New WebClient
    myClient.BaseUrl = url
    If timeout <> 0 Then
        myClient.TimeoutMs = timeout
    Else
        myClient.TimeoutMs = 30000 ' Set default to 30s
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

    Dim myResponse As WebResponse
    Set myResponse = myClient.Execute(myRequest)
    
    ' Debug.Print myResponse.Content
    
    ' Parse JSON response and run functions
    If myResponse.StatusCode = WebStatusCode.Ok Then
        Dim action As Dictionary
        For Each action In myResponse.Data("actions")
            Application.Run action("func"), wb, action
        Next
    Else
        If myResponse.Content <> "" Then
            MsgBox myResponse.Content, vbCritical, "Error"
        Else
            MsgBox myResponse.StatusDescription & " (" & myResponse.StatusCode & ")", vbCritical, "Error"
        End If
    End If

End Function

' Legacy
Function RunRemotePython( _
    url As String, _
    Optional auth As String, _
    Optional apiKey As String, _
    Optional include As String, _
    Optional exclude As String, _
    Optional headers As Variant, _
    Optional timeout As Long, _
    Optional proxyServer As String, _
    Optional proxyBypassList As String, _
    Optional proxyUsername As String, _
    Optional proxyPassword As String, _
    Optional enableAutoProxy As String, _
    Optional insecure As String, _
    Optional followRedirects As String _
)
    RunRemotePython = RunServerPython(url, auth, apiKey, include, exclude, headers, timeout, proxyServer, proxyBypassList, proxyUsername, proxyPassword, enableAutoProxy, insecure, followRedirects)
End Function

' Helpers
Function GetRange(wb As Workbook, action As Dictionary)
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

Function HexToRgb(ByVal hexColor As String) As Variant
    ' Based on https://stackoverflow.com/a/63779233/918626
    Dim red As String, green As String, blue As String
    hexColor = Replace(hexColor, "#", "")
    red = Val("&H" & Mid(hexColor, 1, 2))
    green = Val("&H" & Mid(hexColor, 3, 2))
    blue = Val("&H" & Mid(hexColor, 5, 2))
    HexToRgb = RGB(red, green, blue)
End Function

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    ' Based on https://stackoverflow.com/a/38268261/918626
    If IsEmpty(arr) Then
        IsInArray = False
        Exit Function
    End If
    Dim i As Integer
    On Error GoTo ErrHandler
    For i = LBound(arr) To UBound(arr)
        If Trim(arr(i)) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    On Error GoTo 0
    IsInArray = False
ErrHandler:
    IsInArray = False
End Function

Function base64ToPic(base64string As Variant) As String
    Dim tempPath As String
    ' TODO: handle other image formats than png
    #If Mac Then
        tempPath = GetMacDir("$HOME", False) & "xlwings-" & CreateGUID() & ".png"
        Dim rv As Variant
        rv = ExecuteInShell("echo """ & base64string & """ | base64 -d > " & tempPath).Output
    #Else
        tempPath = Environ("Temp") & "\xlwings-" & CreateGUID() & ".png"
        Open tempPath For Binary As #1
           Put #1, 1, Base64Decode(base64string)
        Close #1
    #End If
    base64ToPic = tempPath
End Function

Function GetAzureAdAccessToken( _
    Optional tenantId As String, _
    Optional clientId As String, _
    Optional port As String, _
    Optional scopes As String, _
    Optional username As String, _
    Optional cliPath As String _
)
    Dim nowTs As Long, expiresTs As Long
    Dim kwargs As String

    If tenantId = "" Then
        tenantId = GetConfig("AZUREAD_TENANT_ID")
    End If
    If clientId = "" Then
        clientId = GetConfig("AZUREAD_CLIENT_ID")
    End If
    If port = "" Then
        port = GetConfig("AZUREAD_PORT")
    End If
    If scopes = "" Then
        scopes = GetConfig("AZUREAD_SCOPES")
    End If
    If username = "" Then
        username = GetConfig("AZUREAD_USERNAME")
    End If
    If cliPath = "" Then
        cliPath = GetConfig("CLI_PATH")
    End If
    If cliPath = "" Then
        kwargs = "tenant_id='" & tenantId & "', "
        kwargs = kwargs & "client_id='" & clientId & "', "
        If port <> "" Then
            kwargs = kwargs & "port='" & port & "', "
        End If
        If scopes <> "" Then
            kwargs = kwargs & "scopes='" & scopes & "', "
        End If
        If username <> "" Then
            kwargs = kwargs & "username='" & username & "', "
        End If
    Else
        kwargs = "--tenant_id=" & tenantId & " "
        kwargs = kwargs & "--client_id=" & clientId & " "
        If port <> "" Then
            kwargs = kwargs & "--port=" & port & " "
        End If
        If scopes <> "" Then
            kwargs = kwargs & "--scopes=" & scopes & " "
        End If
        If username <> "" Then
            kwargs = kwargs & "--username=" & username & " "
        End If
    End If

    expiresTs = GetConfig("AZUREAD_ACCESS_TOKEN_EXPIRES_ON_" & clientId, 0)
    nowTs = DateDiff("s", #1/1/1970#, ConvertToUtc(Now()))

    If (expiresTs > 0) And (nowTs < (expiresTs - 30)) Then
        GetAzureAdAccessToken = GetConfig("AZUREAD_ACCESS_TOKEN_" & clientId)
        Exit Function
    Else
        If cliPath <> "" Then
            RunFrozenPython cliPath, "auth azuread " & kwargs
        Else
            RunPython "from xlwings import cli;cli._auth_aad(" & kwargs & ")"
        End If
        #If Mac Then
            ' RunPython on macOS is async: 60s should be enough if you have to login from scratch
            Dim i as Integer
            For i = 1 To 60
                expiresTs = GetConfig("AZUREAD_ACCESS_TOKEN_EXPIRES_ON_" & clientId, 0)
                If (nowTs < (expiresTs - 30)) Then
                    GetAzureAdAccessToken = GetConfig("AZUREAD_ACCESS_TOKEN_" & clientId)
                    Exit Function
                End If
                Application.Wait (Now + TimeValue("0:00:01"))
            Next i
        #Else
            GetAzureAdAccessToken = GetConfig("AZUREAD_ACCESS_TOKEN_" & clientId)
        #End If
    End If
End Function

' Functions
Sub setValues(wb As Workbook, action As Dictionary)
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

Sub rangeClearContents(wb As Workbook, action As Dictionary)
    GetRange(wb, action).ClearContents
End Sub

Sub rangeClearFormats(wb As Workbook, action As Dictionary)
    GetRange(wb, action).ClearFormats
End Sub

Sub rangeClear(wb As Workbook, action As Dictionary)
    GetRange(wb, action).Clear
End Sub

Sub addSheet(wb As Workbook, action As Dictionary)
    Dim mysheet As Worksheet
    Set mysheet = wb.Sheets.Add
    mysheet.Move After:=Worksheets(action("args")(1) + 1)
    If Not IsNull(action("args")(2)) Then
        mysheet.Name = action("args")(2)
    End If
End Sub

Sub setSheetName(wb As Workbook, action As Dictionary)
    wb.Sheets(action("sheet_position") + 1).Name = action("args")(1)
End Sub

Sub setAutofit(wb As Workbook, action As Dictionary)
    If action("args")(1) = "columns" Then
        GetRange(wb, action).Columns.AutoFit
    Else
        GetRange(wb, action).Rows.AutoFit
    End If
End Sub

Sub setRangeColor(wb As Workbook, action As Dictionary)
    GetRange(wb, action).Interior.Color = HexToRgb(action("args")(1))
End Sub

Sub activateSheet(wb As Workbook, action As Dictionary)
    wb.Sheets(action("args")(1) + 1).Activate
End Sub

Sub addHyperlink(wb As Workbook, action As Dictionary)
    GetRange(wb, action).Hyperlinks.Add _
        Anchor:=GetRange(wb, action), _
        Address:=action("args")(1), _
        TextToDisplay:=action("args")(2), _
        ScreenTip:=action("args")(3)
End Sub

Sub setNumberFormat(wb As Workbook, action As Dictionary)
    GetRange(wb, action).NumberFormat = action("args")(1)
End Sub

Sub setPictureName(wb As Workbook, action As Dictionary)
    wb.Sheets(action("sheet_position") + 1).Pictures(action("args")(1) + 1).Name = action("args")(2)
End Sub

Sub setPictureHeight(wb As Workbook, action As Dictionary)
    wb.Sheets(action("sheet_position") + 1).Pictures(action("args")(1) + 1).Height = action("args")(2)
End Sub

Sub setPictureWidth(wb As Workbook, action As Dictionary)
    wb.Sheets(action("sheet_position") + 1).Pictures(action("args")(1) + 1).Width = action("args")(2)
End Sub

Sub deletePicture(wb As Workbook, action As Dictionary)
    wb.Sheets(action("sheet_position") + 1).Pictures(action("args")(1) + 1).Delete
End Sub

Sub addPicture(wb As Workbook, action As Dictionary)
    Dim tempPath As String
    Dim anchorCell As Range
    Dim imgLeft, imgTop, imgWidth, imgHeight As Long

    tempPath = base64ToPic(action("args")(1))
    With wb.Sheets(action("sheet_position") + 1)
        Set anchorCell = .Cells(action("args")(3) + 1, action("args")(2) + 1)
    End With
    If action("args")(4) > 0 Then
        imgLeft = action("args")(4)
    Else
        imgLeft = anchorCell.Left
    End If
    If action("args")(5) > 0 Then
        imgTop = action("args")(5)
    Else
        imgTop = anchorCell.Top
    End If

    wb.Sheets(action("sheet_position") + 1).Shapes.addPicture tempPath, False, True, imgLeft, imgTop, -1, -1
    On Error Resume Next
        Kill tempPath
    On Error GoTo 0
End Sub

Sub updatePicture(wb As Workbook, action As Dictionary)
    Dim img As Picture
    Dim newImg As Shape
    Dim tempPath, imgName As String
    Dim imgLeft, imgTop, imgWidth, imgHeight As Long
    tempPath = base64ToPic(action("args")(1))
    Set img = wb.Sheets(action("sheet_position") + 1).Pictures(action("args")(2) + 1)
    imgName = img.Name
    imgLeft = img.Left
    imgTop = img.Top
    imgWidth = img.Width
    imgHeight = img.Height
    img.Delete
    Set newImg = wb.Sheets(action("sheet_position") + 1).Shapes.addPicture(tempPath, False, True, imgLeft, imgTop, imgWidth, imgHeight)
    newImg.Name = imgName
    On Error Resume Next
        Kill tempPath
    On Error GoTo 0
End Sub

Sub alert(wb As Workbook, action As Dictionary)
    Dim myPrompt As String, myTitle As String, myButtons As String, myMode As String, myCallback As String
    Dim myStyle As Integer, rv As Integer
    Dim buttonResult As String

    myPrompt = action("args")(1)
    myTitle = action("args")(2)
    myButtons = action("args")(3)
    myMode = action("args")(4)
    myCallback = action("args")(5)

    Select Case myButtons
    Case "ok"
        myStyle = VBA.vbOKOnly
    Case "ok_cancel"
        myStyle = VBA.vbOKCancel
    Case "yes_no"
        myStyle = VBA.vbYesNo
    Case "yes_no_cancel"
        myStyle = VBA.vbYesNoCancel
    End Select

    If myMode = "info" Then
        myStyle = myStyle + VBA.vbInformation
    ElseIf myMode = "critical" Then
        myStyle = myStyle + VBA.vbCritical
    End If

    rv = MsgBox(Prompt:=myPrompt, Title:=myTitle, Buttons:=myStyle)

    Select Case rv
    Case 1
        buttonResult = "ok"
    Case 2
        buttonResult = "cancel"
    Case 6
        buttonResult = "yes"
    Case 7
        buttonResult = "no"
    End Select


    If myCallback <> "" Then
        Application.Run myCallback, buttonResult
    End If

End Sub

Sub setRangeName(wb As Workbook, action As Dictionary)
    GetRange(wb, action).Name = action("args")(1)
End Sub

Sub namesAdd(wb As Workbook, action As Dictionary)
    If IsNull(action("sheet_position")) Then
        wb.Names.Add Name:=action("args")(1), RefersTo:=action("args")(2)
    Else
        wb.Worksheets(action("sheet_position") + 1).Names.Add Name:=action("args")(1), RefersTo:=action("args")(2)
    End If
End Sub

Sub nameDelete(wb As Workbook, action As Dictionary)
    Dim myname As Name
    For Each myname In wb.Names()
        If (myname.Name = action("args")(1)) And (myname.RefersTo = action("args")(2)) Then
            myname.Delete
            Exit For
        End If
    Next
End Sub

Sub runMacro(wb As Workbook, action As Dictionary)
    Dim nArgs As Integer
    nArgs = action("args").Count
    Select Case nArgs
    Case 1
        Application.Run action("args")(1), wb
    Case 2
        Application.Run action("args")(1), wb, action("args")(2)
    Case 3
        Application.Run action("args")(1), wb, action("args")(2), action("args")(3)
    Case 4
        Application.Run action("args")(1), wb, action("args")(2), action("args")(3), action("args")(4)
    Case 5
        Application.Run action("args")(1), wb, action("args")(2), action("args")(3), action("args")(4), action("args")(5)
    Case 6
        Application.Run action("args")(1), wb, action("args")(2), action("args")(3), action("args")(4), action("args")(5), action("args")(6)
    Case 7
        Application.Run action("args")(1), wb, action("args")(2), action("args")(3), action("args")(4), action("args")(5), action("args")(6), action("args")(7)
    Case 8
        Application.Run action("args")(1), wb, action("args")(2), action("args")(3), action("args")(4), action("args")(5), action("args")(6), action("args")(7), action("args")(8)
    Case 9
        Application.Run action("args")(1), wb, action("args")(2), action("args")(3), action("args")(4), action("args")(5), action("args")(6), action("args")(7), action("args")(8), action("args")(9)
    Case 10
        Application.Run action("args")(1), wb, action("args")(2), action("args")(3), action("args")(4), action("args")(5), action("args")(6), action("args")(7), action("args")(8), action("args")(9), action("args")(10)
    Case 11
        Application.Run action("args")(1), wb, action("args")(2), action("args")(3), action("args")(4), action("args")(5), action("args")(6), action("args")(7), action("args")(8), action("args")(9), action("args")(10), action("args")(11)
    Case Else
        Err.Raise vbObjectError + 513, , "macro() only supports up to 10 arguments"
    End Select
End Sub

Sub rangeDelete(wb As Workbook, action As Dictionary)
    Dim shift As String
    shift = action("args")(1)
    If shift = "up" Then
        GetRange(wb, action).Delete (XlDeleteShiftDirection.xlShiftUp)
    Else
        GetRange(wb, action).Delete (XlDeleteShiftDirection.xlShiftToLeft)
    End If
End Sub

Sub rangeInsert(wb As Workbook, action As Dictionary)
    Dim shift As String
    shift = action("args")(1)
    If shift = "down" Then
        GetRange(wb, action).Insert (XlInsertShiftDirection.xlShiftDown)
    Else
        GetRange(wb, action).Insert (XlInsertShiftDirection.xlShiftToRight)
    End If
End Sub

Sub rangeSelect(wb As Workbook, action As Dictionary)
    GetRange(wb, action).Select
End Sub

Sub addTable(wb As Workbook, action As Dictionary)
    Dim hasHeaders As Integer
    If action("args")(2) = True Then
        hasHeaders = XlYesNoGuess.xlYes
    Else
        hasHeaders = XlYesNoGuess.xlNo
    End If
    Dim table As ListObject
    Set table = wb.Worksheets(action("sheet_position") + 1).ListObjects.Add(source:=wb.Worksheets(action("sheet_position") + 1).Range(action("args")(1)), XlListObjectHasHeaders:=hasHeaders, TableStyleName:=action("args")(3))
    If Not IsNull(action("args")(4)) Then
        table.Name = action("args")(4)
    End If
End Sub

Sub setTableName(wb As Workbook, action As Dictionary)
    wb.Worksheets(action("sheet_position") + 1).ListObjects(action("args")(1) + 1).Name = action("args")(2)
End Sub

Sub resizeTable(wb As Workbook, action As Dictionary)
    wb.Worksheets(action("sheet_position") + 1).ListObjects(action("args")(1) + 1).Resize (wb.Worksheets(action("sheet_position") + 1).Range(action("args")(2)))
End Sub

Sub showAutofilterTable(wb As Workbook, action As Dictionary)
    wb.Worksheets(action("sheet_position") + 1).ListObjects(action("args")(1) + 1).ShowAutoFilter = action("args")(2)
End Sub

Sub showHeadersTable(wb As Workbook, action As Dictionary)
    wb.Worksheets(action("sheet_position") + 1).ListObjects(action("args")(1) + 1).ShowHeaders = action("args")(2)
End Sub

Sub showTotalsTable(wb As Workbook, action As Dictionary)
    wb.Worksheets(action("sheet_position") + 1).ListObjects(action("args")(1) + 1).ShowTotals = action("args")(2)
End Sub

Sub setTableStyle(wb As Workbook, action As Dictionary)
    wb.Worksheets(action("sheet_position") + 1).ListObjects(action("args")(1) + 1).TableStyle = action("args")(2)
End Sub

Sub copyRange(wb As Workbook, action As Dictionary)
    If IsNull(action("args")(1)) Then
        GetRange(wb, action).Copy
    Else
        GetRange(wb, action).Copy Destination:=wb.Worksheets(action("args")(1) + 1).Range(action("args")(2))
    End If
End Sub

Sub sheetDelete(wb As Workbook, action As Dictionary)
    Dim displayAlertsState As Boolean
    displayAlertsState = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wb.Worksheets(action("sheet_position") + 1).Delete
    Application.DisplayAlerts = displayAlertsState
End Sub

Sub sheetClear(wb As Workbook, action As Dictionary)
    wb.Worksheets(action("sheet_position") + 1).Cells.Clear
End Sub

Sub sheetClearContents(wb As Workbook, action As Dictionary)
    wb.Worksheets(action("sheet_position") + 1).Cells.ClearContents
End Sub

Sub sheetClearFormats(wb As Workbook, action As Dictionary)
    wb.Worksheets(action("sheet_position") + 1).Cells.ClearFormats
End Sub
