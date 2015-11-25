Attribute VB_Name = "Module1"
' Password: xlwings

Function ModuleIsPresent(ByVal wb As Workbook, moduleName As String) As Boolean
    On Error GoTo not_present
    Set x = wb.VBProject.VBComponents.Item(moduleName)
    ModuleIsPresent = True
    Exit Function
not_present:
    ModuleIsPresent = False
End Function

Sub ImportPythonUDFsAddIn(control As IRibbonControl)
    Set wb = ActiveWorkbook
    If Not ModuleIsPresent(wb, "xlwings") Then
        MsgBox "This workbook must contain the xlwings VBA module."
        Exit Sub
    End If

    If LCase$(Right$(wb.Name, 5)) <> ".xlsm" And LCase$(Right$(wb.Name, 5)) <> ".xlsb" Then
        MsgBox "Please save this workbook (""" + wb.Name + """) as a macro-enabled workbook first."
        Exit Sub
    End If

    Application.Run ActiveWorkbook.Name + "!ImportPythonUDFs"
End Sub
