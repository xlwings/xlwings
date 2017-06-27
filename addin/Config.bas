Attribute VB_Name = "Config"
Public Const CONFIG_SHEET = "xlwings.conf"

Function ConfigSheetExists() As Boolean
    Dim sht As Worksheet
    On Error Resume Next
        Set sht = ActiveWorkbook.Sheets(CONFIG_SHEET)
    On Error GoTo 0
    ConfigSheetExists = Not sht Is Nothing
End Function

Function ConfigFileExists() As Boolean
    On Error GoTo Err 'Fails on Mac if it doesn't exist
    If Dir(GetConfigFilePath) <> "" Then
    On Error GoTo 0
        ConfigFileExists = True
    Else
        ConfigFileExists = False
    End If
    Exit Function
Err:
    ConfigFileExists = False
End Function

Function GetConfigFilePath() As String
    #If Mac Then
        ' /Users/<User>/Library/Containers/com.microsoft.Excel/Data/xlwings.conf
        GetConfigFilePath = GetMacDir("Home") & "/xlwings.conf"
    #Else
        GetConfigFilePath = Environ("USERPROFILE") & "\.xlwings\xlwings.conf"
    #End If
End Function

Function GetSheetConfig() As Dictionary
    Dim d As Dictionary
    Dim lastCell As Range
    Set d = New Dictionary
    Set sht = ActiveWorkbook.Sheets(CONFIG_SHEET)

    If sht.Range("A2") = "" Then
        Set lastCell = sht.Range("A1")
    Else
        Set lastCell = sht.Range("A1").End(xlDown)
    End If

    For Each cell In Range(sht.Range("A1"), lastCell)
        d.Add UCase(cell.Value), cell.Offset(0, 1).Value
    Next cell
    Set GetSheetConfig = d
End Function

Function GetConfig(configKey As String, default As String) As Variant
    Dim configValue As String
    ' A entry in xlwings.conf sheet overrides the config file/ribbon

    If ConfigSheetExists = True Then
        GetConfig = GetSheetConfig.Item(configKey)
    End If

    If GetConfig = "" And ConfigFileExists = True Then
        If GetConfigFromFile(GetConfigFilePath(), configKey, configValue) Then
            GetConfig = configValue
        End If
    End If

    If GetConfig = "" Then
        GetConfig = default
    End If
End Function

Function SaveConfigToFile(sFileName As String, sName As String, Optional sValue As String) As Boolean
'Adopted from http://peltiertech.com/save-retrieve-information-text-files/

  Dim iFileNumA As Long
  Dim iFileNumB As Long
  Dim sFile As String
  Dim sXFile As String
  Dim sVarName As String
  Dim sVarValue As String
  Dim lErrLast As Long

  ' assume false unless variable is successfully saved
  SaveConfigToFile = False

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
      SaveConfigToFile = True
    Close #iFileNumA
    Close #iFileNumB
    FileCopy sXFile, sFile
    Kill sXFile
  Else
    ' make new file
    iFileNumB = FreeFile
    Open sFile For Output As iFileNumB
      Write #iFileNumB, sName, sValue
      SaveConfigToFile = True
    Close #iFileNumB
  End If

End Function

Function GetConfigFromFile(sFile As String, sName As String, Optional sValue As String) As Boolean
'Adopted from http://peltiertech.com/save-retrieve-information-text-files/

  Dim iFileNum As Long
  Dim sVarName As String
  Dim sVarValue As String
  Dim lErrLast As Long

  ' assume false unless variable is found
  GetConfigFromFile = False

  ' open text file to read settings
  If FileExists(sFile) Then
    iFileNum = FreeFile
    Open sFile For Input As iFileNum
      Do While Not EOF(iFileNum)
        Input #iFileNum, sVarName, sVarValue
        If sVarName = sName Then
          sValue = sVarValue
          GetConfigFromFile = True
          Exit Do
        End If
      Loop
    Close #iFileNum
  End If

End Function
