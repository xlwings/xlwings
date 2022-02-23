Attribute VB_Name = "RibbonXlwings"
Option Explicit
' Ribbon docs: https://msdn.microsoft.com/en-us/library/dd910855(v=office.12).aspx
' Custom UI Editor: https://github.com/fernandreu/office-ribbonx-editor
Sub RunMain(control As IRibbonControl)
    Dim wb As Workbook
    Dim mymodule As String
    Set wb = ActiveWorkbook
    
    If ActiveWorkbook.Path = vbNullString Then
        MsgBox "Please save this workbook (""" + wb.name + """) first."
        Exit Sub
    Else
        mymodule = Left(wb.name, (InStrRev(wb.name, ".", -1, vbTextCompare) - 1))
    End If
    
    Application.StatusBar = "Running..."
    RunPython "import " & mymodule & ";" & mymodule & ".main()"
    Application.StatusBar = False
End Sub


Sub ImportFunctions(control As IRibbonControl)
    #If Mac Then
    #Else
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    If LCase$(Right$(wb.name, 5)) <> ".xlsm" And LCase$(Right$(wb.name, 5)) <> ".xlsb" And LCase$(Right$(wb.name, 5)) <> ".xlam" Then
        MsgBox "Please save this workbook (""" + wb.name + """) as a macro-enabled workbook first."
        Exit Sub
    End If
    KillPy
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
    returnedVal = "Version: " & XLWINGS_VERSION
End Sub

Sub SetInterpreter(control As IRibbonControl, text As String)
    Dim tf As Boolean
    Dim interpreter As String
    #If Mac Then
        interpreter = "INTERPRETER_MAC"
    #Else
        interpreter = "INTERPRETER_WIN"
    #End If
    tf = SaveConfigToFile(GetConfigFilePath, interpreter, text)
End Sub

Sub GetInterpreter(control As IRibbonControl, ByRef returnedVal)
    Dim setting As String, interpreter As String
    #If Mac Then
        interpreter = "INTERPRETER_MAC"
    #Else
        interpreter = "INTERPRETER_WIN"
    #End If

    If GetConfigFromFile(GetConfigFilePath(), interpreter, setting) Then
        returnedVal = setting
    Else
        If GetConfigFromFile(GetConfigFilePath(), "INTERPRETER", setting) Then
            ' Legacy
            returnedVal = setting
        Else
            returnedVal = ""
        End If
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

Sub SetCondaPath(control As IRibbonControl, text As String)
    Dim tf As Boolean
    tf = SaveConfigToFile(GetConfigFilePath, "CONDA PATH", text)
End Sub

Sub GetCondaPath(control As IRibbonControl, ByRef returnedVal)
    Dim setting As String
    If GetConfigFromFile(GetConfigFilePath, "CONDA PATH", setting) Then
        returnedVal = setting
    Else
        returnedVal = ""
    End If
End Sub

Sub SetCondaEnv(control As IRibbonControl, text As String)
    Dim tf As Boolean
    tf = SaveConfigToFile(GetConfigFilePath, "CONDA ENV", text)
End Sub

Sub GetCondaEnv(control As IRibbonControl, ByRef returnedVal)
    Dim setting As String
    If GetConfigFromFile(GetConfigFilePath, "CONDA ENV", setting) Then
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
        If setting = "True" Then
            pressed = True
        Else
            pressed = False
        End If
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
        If setting = "True" Then
            pressed = True
        Else
            pressed = False
        End If
    Else
        pressed = False
    End If
    #End If
End Sub

Sub ChangeShowConsole(control As IRibbonControl, pressed As Boolean)
    #If Mac Then
    #Else
    Dim tf As Boolean
    tf = SaveConfigToFile(GetConfigFilePath, "SHOW CONSOLE", CStr(pressed))
    #End If
End Sub

Sub GetPressedShowConsole(control As IRibbonControl, ByRef pressed)
    #If Mac Then
    #Else
    Dim setting As String
    If GetConfigFromFile(GetConfigFilePath, "SHOW CONSOLE", setting) Then
        If setting = "True" Then
            pressed = True
        Else
            pressed = False
        End If
    Else
        pressed = False
    End If
    #End If
End Sub

Sub RestartPython(control As IRibbonControl)
    #If Mac Then
    #Else
    KillPy
    Py.Exec ""
    #End If
End Sub

