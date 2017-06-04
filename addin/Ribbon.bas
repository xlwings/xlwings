Attribute VB_Name = "Ribbon"
' Ribbon docs: https://msdn.microsoft.com/en-us/library/dd910855(v=office.12).aspx
' Custom UI Editor: http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2006/05/26/customuieditor.aspx

Sub import_functions(control As IRibbonControl)
    Set wb = ActiveWorkbook

    If LCase$(Right$(wb.Name, 5)) <> ".xlsm" And LCase$(Right$(wb.Name, 5)) <> ".xlsb" Then
        MsgBox "Please save this workbook (""" + wb.Name + """) as a macro-enabled workbook first."
        Exit Sub
    End If

    ImportPythonUDFs
    Set wb = Nothing
End Sub

Sub set_interpreter(control As IRibbonControl, text As String)
    tf = SaveConfigToFile(GetConfigFilePath, "INTERPRETER", text)
End Sub

Sub get_interpreter(control As IRibbonControl, ByRef returnedVal)
    Dim setting As String
    returnedVal = GetConfigFromFile(GetConfigFilePath, "INTERPRETER", setting)
    If returnedVal = False Then
        returnedVal = ""
    End If
End Sub

Sub set_pythonpath(control As IRibbonControl, text As String)
    tf = SaveConfigToFile(GetConfigFilePath, "PYTHONPATH", text)
End Sub

Sub get_pythonpath(control As IRibbonControl, ByRef returnedVal)
    Dim setting As String
    returnedVal = GetConfigFromFile(GetConfigFilePath, "PYTHONPATH", setting)
    If returnedVal = False Then
        returnedVal = ""
    End If
End Sub

Sub set_logfile(control As IRibbonControl, text As String)
    tf = SaveConfigToFile(GetConfigFilePath, "LOGFILE", text)
End Sub

Sub get_logfile(control As IRibbonControl, ByRef returnedVal)
    Dim setting As String
    returnedVal = GetConfigFromFile(GetConfigFilePath, "LOGFILE", setting)
    If returnedVal = False Then
        returnedVal = ""
    End If
End Sub

Sub set_udfmodules(control As IRibbonControl, text As String)
    tf = SaveConfigToFile(GetConfigFilePath, "UDF_MODULES", text)
End Sub

Sub get_udfmodules(control As IRibbonControl, ByRef returnedVal)
    Dim setting As String
    returnedVal = GetConfigFromFile(GetConfigFilePath, "UDF_MODULES", setting)
    If returnedVal = False Then
        returnedVal = ""
    End If
End Sub

Sub change_udfdebug(control As IRibbonControl, pressed As Boolean)
    tf = SaveConfigToFile(GetConfigFilePath, "UDF_DEBUG", CStr(pressed))
End Sub

Sub getpressed_udfdebug(control As IRibbonControl, ByRef pressed)
    Dim setting As String
    returnedVal = GetConfigFromFile(GetConfigFilePath, "UDF_DEBUG", setting)
    If returnedVal = False Then
        returnedVal = ""
    End If
End Sub

Sub change_comserver(control As IRibbonControl, pressed As Boolean)
    tf = SaveConfigToFile(GetConfigFilePath, "UDF_SERVER", CStr(pressed))
End Sub

Sub getpressed_comserver(control As IRibbonControl, ByRef pressed)
    Dim setting As String
    returnedVal = GetConfigFromFile(GetConfigFilePath, "UDF_SERVER", setting)
    If returnedVal = False Then
        returnedVal = ""
    End If
End Sub

Sub restart_python(control As IRibbonControl)
    KillPy
End Sub
