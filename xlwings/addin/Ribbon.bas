Attribute VB_Name = "Ribbon"
Option Explicit
' Ribbon docs: https://msdn.microsoft.com/en-us/library/dd910855(v=office.12).aspx
' Custom UI Editor: http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2006/05/26/customuieditor.aspx

Sub ImportFunctions(control As IRibbonControl)
    #If Mac Then
    #Else
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    If LCase$(Right$(wb.Name, 5)) <> ".xlsm" And LCase$(Right$(wb.Name, 5)) <> ".xlsb" Then
        MsgBox "Please save this workbook (""" + wb.Name + """) as a macro-enabled workbook first."
        Exit Sub
    End If

    ImportPythonUDFs
    Set wb = Nothing
    #End If
End Sub

Sub GetVisible(control As IRibbonControl, ByRef returnedVal)
    #If Mac Then
        returnedVal = False
    #Else
        returnedVal = True
    #End If
End Sub

Sub GetVersion(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "Version: 0.11.4"
End Sub

Sub SetInterpreter(control As IRibbonControl, text As String)
    Dim tf As Boolean
    tf = SaveConfigToFile(GetConfigFilePath, "INTERPRETER", text)
End Sub

Sub GetInterpreter(control As IRibbonControl, ByRef returnedVal)
    Dim setting As String
    If GetConfigFromFile(GetConfigFilePath(), "INTERPRETER", setting) Then
        returnedVal = setting
    Else
        returnedVal = ""
    End If
End Sub

Sub SetPythonpath(control As IRibbonControl, text As String)
    Dim tf As Boolean
    tf = SaveConfigToFile(GetConfigFilePath, "PYTHONPATH", text)
End Sub

Sub GetPythonpath(control As IRibbonControl, ByRef returnedVal)
    Dim setting As String
    If GetConfigFromFile(GetConfigFilePath, "PYTHONPATH", setting) Then
        returnedVal = setting
    Else
        returnedVal = ""
    End If
End Sub

Sub SetLogfile(control As IRibbonControl, text As String)
    Dim tf As Boolean
    tf = SaveConfigToFile(GetConfigFilePath, "LOG FILE", text)
End Sub

Sub GetLogfile(control As IRibbonControl, ByRef returnedVal)
    Dim setting As String
    If GetConfigFromFile(GetConfigFilePath, "LOG FILE", setting) Then
        returnedVal = setting
    Else
        returnedVal = ""
    End If
End Sub

Sub SetUdfModules(control As IRibbonControl, text As String)
    #If Mac Then
    #Else
        Dim tf As Boolean
        tf = SaveConfigToFile(GetConfigFilePath, "UDF MODULES", text)
    #End If
End Sub

Sub GetUdfModules(control As IRibbonControl, ByRef returnedVal)
    #If Mac Then
    #Else
        Dim setting As String
        If GetConfigFromFile(GetConfigFilePath, "UDF MODULES", setting) Then
            returnedVal = setting
        Else
            returnedVal = ""
        End If
    #End If
End Sub

Sub ChangeUdfDebug(control As IRibbonControl, pressed As Boolean)
    #If Mac Then
    #Else
    Dim tf As Boolean
    tf = SaveConfigToFile(GetConfigFilePath, "DEBUG UDFS", CStr(pressed))
    #End If
End Sub

Sub GetPressedUdfDebug(control As IRibbonControl, ByRef pressed)
    #If Mac Then
    #Else
    Dim setting As String
    If GetConfigFromFile(GetConfigFilePath, "DEBUG UDFS", setting) Then
        pressed = setting
    Else
        pressed = False
    End If
    #End If
End Sub

Sub ChangeUdfServer(control As IRibbonControl, pressed As Boolean)
    #If Mac Then
    #Else
    Dim tf As Boolean
    tf = SaveConfigToFile(GetConfigFilePath, "USE UDF SERVER", CStr(pressed))
    #End If
End Sub

Sub GetPressedUdfServer(control As IRibbonControl, ByRef pressed)
    #If Mac Then
    #Else
    Dim setting As String
    If GetConfigFromFile(GetConfigFilePath, "USE UDF SERVER", setting) Then
        pressed = setting
    Else
        pressed = ""
    End If
    #End If
End Sub

Sub RestartPython(control As IRibbonControl)
    #If Mac Then
    #Else
    KillPy
    #End If
End Sub
