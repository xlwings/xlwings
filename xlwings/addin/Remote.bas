Attribute VB_Name = "Remote"
Option Explicit
Function RunRemotePython( _
    url As String, _
    Optional apiKey As String, _
    Optional include As String, _
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
    If apiKey = "" Then
        apiKey = GetConfig("API_KEY", "")
    End If

    ' Request payload
    Dim payload As New Dictionary
    payload.Add "client", "VBA"
    payload.Add "version", XLWINGS_VERSION
    
    Dim bookPayload As New Dictionary
    bookPayload.Add "name", ActiveWorkbook.Name
    bookPayload.Add "active_sheet_index", ActiveSheet.Index - 1
    bookPayload.Add "selection", Application.Selection.Address(False, False)
    payload.Add "book", bookPayload

    ' Names
    Dim myname As Name
    Dim mynames() As Dictionary
    Dim nNames As Integer
    Dim iName As Integer
    nNames = wb.Names.Count
    If nNames > 0 Then
        ReDim mynames(nNames - 1)
        For iName = 1 To nNames
            Set myname = wb.Names(iName)
            Dim nameDict As Dictionary
            Set nameDict = New Dictionary
            nameDict.Add "name", myname.Name
            If InStr(1, myname.RefersTo, "=") <> 1 Then
                ' If the reference doesn't start with an =, it's a named range
                nameDict.Add "sheet_index", myname.RefersToRange.Parent.Index - 1
                nameDict.Add "address", myname.RefersToRange.Address(False, False)
                nameDict.Add "book_scope", TypeOf myname.Parent Is Workbook
            Else
                nameDict.Add "sheet_index", Null
                nameDict.Add "address", Null
                nameDict.Add "book_scope", Null
            End If
            Set mynames(iName - 1) = nameDict
        Next
        payload.Add "names", mynames
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
        Dim pic As Picture
        Dim pics() As Dictionary
        Dim nPics As Integer
        Dim nPic As Integer
        nPics =  wb.Worksheets(i).Pictures.Count
        If nPics > 0 Then
            ReDim pics(nPics - 1)
            For nPic = 1 To nPics
                Set pic =  wb.Worksheets(i).Pictures(nPic)
                Dim picDict As Dictionary
                Set picDict = New Dictionary
                picDict.Add "name", pic.Name
                picDict.Add "height", pic.Height
                picDict.Add "width", pic.Width
                Set pics(nPic - 1) = picDict
            Next
            sheetDict.Add "pictures", pics
        Else
            sheetDict.Add "pictures", Array()
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
        For Each myKey In headers.myKeys
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
        MsgBox myResponse.Content, vbCritical, "Error"
    End If

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

Sub clearContents(wb As Workbook, action As Dictionary)
    GetRange(wb, action).clearContents
End Sub

Sub addSheet(wb As Workbook, action As Dictionary)
    Dim mysheet As Worksheet
    Set mysheet = wb.Sheets.Add
    mysheet.Move After:=Worksheets(action("args")(1) + 1)
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
    tempPath = base64ToPic(action("args")(1))
    wb.Sheets(action("sheet_position") + 1).Shapes.addPicture tempPath, False, True, action("args")(4), action("args")(5), -1, -1
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

