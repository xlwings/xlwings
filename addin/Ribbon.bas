Attribute VB_Name = "Ribbon"
' Ribbon docs: https://msdn.microsoft.com/en-us/library/dd910855(v=office.12).aspx
' Custom UI Editor: http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2006/05/26/customuieditor.aspx

Sub ImportFunctions(control As IRibbonControl)
    Set wb = ActiveWorkbook

    If LCase$(Right$(wb.Name, 5)) <> ".xlsm" And LCase$(Right$(wb.Name, 5)) <> ".xlsb" Then
        MsgBox "Please save this workbook (""" + wb.Name + """) as a macro-enabled workbook first."
        Exit Sub
    End If

    ImportPythonUDFs
    Set wb = Nothing
End Sub

Sub SetInterpreter(control As IRibbonControl, text As String)
    tf = SaveConfigToFile(GetConfigFilePath, "INTERPRETER", text)
End Sub

Sub GetInterpreter(control As IRibbonControl, ByRef returnedVal)
    Dim setting As String
    returnedVal = GetConfigFromFile(GetConfigFilePath, "INTERPRETER", setting)
    If returnedVal = False Then
        returnedVal = ""
    End If
End Sub

Sub SetPythonpath(control As IRibbonControl, text As String)
    tf = SaveConfigToFile(GetConfigFilePath, "PYTHONPATH", text)
End Sub

Sub GetPythonpath(control As IRibbonControl, ByRef returnedVal)
    Dim setting As String
    returnedVal = GetConfigFromFile(GetConfigFilePath, "PYTHONPATH", setting)
    If returnedVal = False Then
        returnedVal = ""
    End If
End Sub

Sub SetLogfile(control As IRibbonControl, text As String)
    tf = SaveConfigToFile(GetConfigFilePath, "LOG FILE", text)
End Sub

Sub GetLogfile(control As IRibbonControl, ByRef returnedVal)
    Dim setting As String
    returnedVal = GetConfigFromFile(GetConfigFilePath, "LOG FILE", setting)
    If returnedVal = False Then
        returnedVal = ""
    End If
End Sub

Sub SetUdfModules(control As IRibbonControl, text As String)
    tf = SaveConfigToFile(GetConfigFilePath, "UDF MODULES", text)
End Sub

Sub GetUdfModules(control As IRibbonControl, ByRef returnedVal)
    Dim setting As String
    returnedVal = GetConfigFromFile(GetConfigFilePath, "UDF MODULES", setting)
    If returnedVal = False Then
        returnedVal = ""
    End If
End Sub

Sub ChangeUdfDebug(control As IRibbonControl, pressed As Boolean)
    tf = SaveConfigToFile(GetConfigFilePath, "DEBUG UDFS", CStr(pressed))
End Sub

Sub GetPressedUdfDebug(control As IRibbonControl, ByRef pressed)
    Dim setting As String
    returnedVal = GetConfigFromFile(GetConfigFilePath, "DEBUG UDFS", setting)
    If returnedVal = False Then
        returnedVal = ""
    End If
End Sub

Sub ChangeUdfServer(control As IRibbonControl, pressed As Boolean)
    tf = SaveConfigToFile(GetConfigFilePath, "USE UDF SERVER", CStr(pressed))
End Sub

Sub GetPressedUdfServer(control As IRibbonControl, ByRef pressed)
    Dim setting As String
    returnedVal = GetConfigFromFile(GetConfigFilePath, "USE UDF SERVER", setting)
    If returnedVal = False Then
        returnedVal = ""
    End If
End Sub

Sub RestartPython(control As IRibbonControl)
    KillPy
End Sub
