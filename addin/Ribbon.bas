Attribute VB_Name = "ribbon"
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
    settings.SetInterpreter (text)
End Sub

Sub get_interpreter(control As IRibbonControl, ByRef returnedVal)
    returnedVal = settings.GetInterpreter
End Sub

Sub set_pythonpath(control As IRibbonControl, text As String)
    settings.SetPythonpath (text)
End Sub

Sub get_pythonpath(control As IRibbonControl, ByRef returnedVal)
    returnedVal = settings.GetPythonpath
End Sub

Sub set_logfile(control As IRibbonControl, text As String)
    settings.SetLogfile (text)
End Sub

Sub get_logfile(control As IRibbonControl, ByRef returnedVal)
    returnedVal = settings.GetLogfile
End Sub

Sub set_udfmodules(control As IRibbonControl, text As String)
    settings.SetUdfmodules (text)
End Sub

Sub get_udfmodules(control As IRibbonControl, ByRef returnedVal)
    returnedVal = settings.GetUdfmodules
End Sub

Sub change_udfdebug(control As IRibbonControl, pressed As Boolean)
    settings.SetUdfDebug (pressed)
End Sub

Sub getpressed_udfdebug(control As IRibbonControl, ByRef pressed)
    pressed = settings.GetUdfDebug
End Sub

Sub change_comserver(control As IRibbonControl, pressed As Boolean)
    settings.SetComServer (pressed)
End Sub

Sub getpressed_comserver(control As IRibbonControl, ByRef pressed)
    pressed = settings.GetComServer
End Sub

Sub restart_python(control As IRibbonControl)
    KillPy
End Sub
