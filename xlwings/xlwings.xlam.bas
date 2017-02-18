Attribute VB_Name = "Module1"
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


'Adopted from http://peltiertech.com/save-retrieve-information-text-files/
Function SaveSetting(sFileName As String, sName As String, Optional sValue As String) As Boolean

  Dim iFileNumA As Long
  Dim iFileNumB As Long
  Dim sFile As String
  Dim sXFile As String
  Dim sVarName As String
  Dim sVarValue As String
  Dim lErrLast As Long

  ' assume false unless variable is successfully saved
  SaveSetting = False

  ' temporary file
  sFile = sFileName
  sXFile = sFileName & "_temp"

  ' open text file to read settings
  If FileExists(sFile) Then
    'replace existing settings file
    iFileNumA = FreeFile
    Open sFile For Input As iFileNumA
    iFileNumB = FreeFile
    Open sXFile For Output As iFileNumB
      Do While Not EOF(iFileNumA)
        Input #iFileNumA, sVarName, sVarValue
        If sVarName <> sName Then
          Write #iFileNumB, sVarName, sVarValue
        End If
      Loop
      Write #iFileNumB, sName, sValue
      SaveSetting = True
    Close #iFileNumA
    Close #iFileNumB
    FileCopy sXFile, sFile
    Kill sXFile
  Else
    ' make new file
    iFileNumB = FreeFile
    Open sFile For Output As iFileNumB
      Write #iFileNumB, sName, sValue
      SaveSetting = True
    Close #iFileNumB
  End If

End Function

Function GetSetting(sFile As String, sName As String, Optional sValue As String) As Boolean

  Dim iFileNum As Long
  Dim sVarName As String
  Dim sVarValue As String
  Dim lErrLast As Long

  ' assume false unless variable is found
  GetSetting = False

  ' open text file to read settings
  If FileExists(sFile) Then
    iFileNum = FreeFile
    Open sFile For Input As iFileNum
      Do While Not EOF(iFileNum)
        Input #iFileNum, sVarName, sVarValue
        If sVarName = sName Then
          sValue = sVarValue
          GetSetting = True
          Exit Do
        End If
      Loop
    Close #iFileNum
  End If

End Function

Function IsFullName(sFile As String) As Boolean
  ' if sFile includes path, it contains path separator "\" or "/"
  IsFullName = InStr(sFile, "\") + InStr(sFile, "/") > 0
End Function

Function FileExists(ByVal FileSpec As String) As Boolean
   ' by Karl Peterson MS MVP VB
   Dim Attr As Long
   ' Guard against bad FileSpec by ignoring errors
   ' retrieving its attributes.
   On Error Resume Next
   Attr = GetAttr(FileSpec)
   If Err.Number = 0 Then
      ' No error, so something was found.
      ' If Directory attribute set, then not a file.
      FileExists = Not ((Attr And vbDirectory) = vbDirectory)
   End If
End Function
