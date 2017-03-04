Attribute VB_Name = "menu"
' Password: xlwings

' Ribbon docs: https://msdn.microsoft.com/en-us/library/dd910855(v=office.12).aspx
' Custom UI Editor: http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2006/05/26/customuieditor.aspx

Function ModuleIsPresent(ByVal wb As Workbook, moduleName As String) As Boolean
    On Error GoTo not_present
    Set x = wb.VBProject.VBComponents.Item(moduleName)
    ModuleIsPresent = True
    Exit Function
not_present:
    ModuleIsPresent = False
End Function

Sub import_functions(control As IRibbonControl)
    Set wb = ActiveWorkbook
    If Not ModuleIsPresent(wb, "xlwings") Then
        MsgText = "Make sure that this workbook contains the xlwings module "
        MsgText = MsgText & "and you are trusting access to the VBA project object module (Options)."
        MsgBox MsgText, vbCritical, "Error"
        Exit Sub
    End If

    If LCase$(Right$(wb.Name, 5)) <> ".xlsm" And LCase$(Right$(wb.Name, 5)) <> ".xlsb" Then
        MsgBox "Please save this workbook (""" + wb.Name + """) as a macro-enabled workbook first."
        Exit Sub
    End If

    Application.Run "'" + ActiveWorkbook.Name + "'!ImportPythonUDFs"
    Set wb = Nothing
End Sub

'Callback for interpreter onChange
Sub set_interpreter(control As IRibbonControl, text As String)
    Debug.Print text
    'tf = SaveSetting(SETTINGSFILE, "interpreter", text)
End Sub

'Callback for interpreter getText
Sub get_interpreter(control As IRibbonControl, ByRef returnedVal)
End Sub

'Callback for pythonpath onChange
Sub set_pythonpath(control As IRibbonControl, text As String)
End Sub

'Callback for pythonpath getText
Sub get_pythonpath(control As IRibbonControl, ByRef returnedVal)
End Sub

'Callback for logfile onChange
Sub set_logpath(control As IRibbonControl, text As String)
End Sub

'Callback for logfile getText
Sub get_logpath(control As IRibbonControl, ByRef returnedVal)
End Sub

'Callback for comserver onAction
Sub change_comserver(control As IRibbonControl, pressed As Boolean)
End Sub

'Callback for udfmodules onChange
Sub set_udfmodules(control As IRibbonControl, text As String)
End Sub

'Callback for udfmodules getText
Sub get_udfmodules(control As IRibbonControl, ByRef returnedVal)
End Sub

'Callback for udfdebug onAction

Sub change_udfdebug(control As IRibbonControl, pressed As Boolean)
End Sub
