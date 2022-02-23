Attribute VB_Name = "Config"
Option Explicit
#Const App = "Microsoft Excel" 'Adjust when using outside of Excel

#If App = "Microsoft Excel" Then
Function GetDirectoryPath(Optional wb As Workbook) As String
#Else
Function GetDirectoryPath(Optional wb As Variant) As String
#End If
    ' Leaving this here for now because we currently don't have #Const App in Utils
    Dim Path As String
    #If App = "Microsoft Excel" Then
        On Error Resume Next 'On Mac, this is called when exiting the Python interpreter
            Path = GetDirectory(GetFullName(wb))
        On Error GoTo 0
    #ElseIf App = "Microsoft Word" Then
        Path = ActiveDocument.Path
    #ElseIf App = "Microsoft Access" Then
        Path = CurrentProject.Path ' Won't be transformed for standalone module as ThisProject doesn't exit
    #ElseIf App = "Microsoft PowerPoint" Then
        Path = ActivePresentation.Path ' Won't be transformed for standalone module ThisPresentation doesn't exist
    #Else
        Exit Function
    #End If
    GetDirectoryPath = Path
End Function

Function GetConfigFilePath() As String
    #If Mac Then
        ' ~/Library/Containers/com.microsoft.Excel/Data/xlwings.conf
        GetConfigFilePath = GetMacDir("$HOME", False) & "/" & PROJECT_NAME & ".conf"
    #Else
        GetConfigFilePath = Environ("USERPROFILE") & "\." & PROJECT_NAME & "\" & PROJECT_NAME & ".conf"
    #End If
End Function

Function GetDirectoryConfigFilePath() As String
    Dim pathSeparator As String
    
    #If Mac Then ' Application.PathSeparator doesn't seem to exist in Access...
        pathSeparator = "/"
    #Else
        pathSeparator = "\"
    #End If
    
    GetDirectoryConfigFilePath = GetDirectoryPath(ActiveWorkbook) & pathSeparator & PROJECT_NAME & ".conf"
End Function

#If App = "Microsoft Excel" Then
Function GetConfigFromSheet(wb As Workbook)
    Dim lastCell As Range, cell As Range
    #If Mac Then
    Dim d As Dictionary
    Set d = New Dictionary
    #Else
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    #End If
    Dim sht As Worksheet

    Set sht = wb.Sheets(PROJECT_NAME & ".conf")

    If sht.Range("A2") = "" Then
        Set lastCell = sht.Range("A1")
    Else
        Set lastCell = sht.Range("A1").End(xlDown)
    End If

    For Each cell In Range(sht.Range("A1"), lastCell)
        d.Add UCase(cell.Value), cell.Offset(0, 1).Value
    Next cell
    Set GetConfigFromSheet = d
End Function
#End If

Function GetConfig(configKey As String, Optional default As String = "", Optional source As String = "") As Variant
    ' If source is provided, returns the value from this source only, otherwise it goes through all layers until
    ' it finds a value (sheet -> directory -> user -> default)
    ' An entry in xlwings.conf sheet overrides the config file/ribbon
    Dim configValue As String
    
    ' Sheet
    #If App = "Microsoft Excel" Then
    If source = "" Or source = "sheet" Then
        If Application.name = "Microsoft Excel" Then
            'Workbook Sheet Config
            If SheetExists(ActiveWorkbook, PROJECT_NAME & ".conf") = True Then
                If GetConfigFromSheet(ActiveWorkbook).Exists(configKey) = True Then
                    GetConfig = GetConfigFromSheet(ActiveWorkbook).Item(configKey)
                    GetConfig = ExpandEnvironmentStrings(GetConfig)
                    Exit Function
                End If
            End If
    
            'Add-in Sheet Config (only for custom add-ins, unused by xlwings add-in)
            If SheetExists(ThisWorkbook, PROJECT_NAME & ".conf") = True Then
                If GetConfigFromSheet(ThisWorkbook).Exists(configKey) = True Then
                    GetConfig = GetConfigFromSheet(ThisWorkbook).Item(configKey)
                    GetConfig = ExpandEnvironmentStrings(GetConfig)
                    Exit Function
                End If
            End If
        End If
    End If
    #End If

    ' Directory Config
    If source = "" Or source = "directory" Then
        #If App = "Microsoft Excel" Then
            If GetFullName(ActiveWorkbook) <> "" Then ' Empty if local dir can't be figured out (e.g. SharePoint)
        #Else
            If InStr(GetDirectoryPath(), "://") = 0 Then ' Other Office apps: skip for synced SharePoint/OneDrive files
        #End If
            If FileExists(GetDirectoryConfigFilePath()) = True Then
                If GetConfigFromFile(GetDirectoryConfigFilePath(), configKey, configValue) Then
                    GetConfig = configValue
                    GetConfig = ExpandEnvironmentStrings(GetConfig)
                    Exit Function
                End If
            End If
        End If
    End If

    ' User Config
    If source = "" Or source = "user" Then
        If FileExists(GetConfigFilePath()) = True Then
            If GetConfigFromFile(GetConfigFilePath(), configKey, configValue) Then
                GetConfig = configValue
                GetConfig = ExpandEnvironmentStrings(GetConfig)
                Exit Function
            End If
        End If
    End If

    ' Defaults
    GetConfig = default
    GetConfig = ExpandEnvironmentStrings(GetConfig)

End Function

Function SaveConfigToFile(sFileName As String, sName As String, Optional sValue As String) As Boolean
'Adopted from http://peltiertech.com/save-retrieve-information-text-files/

  Dim iFileNumA As Long, iFileNumB As Long, lErrLast As Long
  Dim sFile As String, sXFile As String, sVarName As String, sVarValue As String
      
    
  #If Mac Then
    If Not FileOrFolderExistsOnMac(ParentFolder(sFileName)) Then
  #Else
    If Len(Dir(ParentFolder(sFileName), vbDirectory)) = 0 Then
  #End If
     MkDir ParentFolder(sFileName)
  End If

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
'Based on http://peltiertech.com/save-retrieve-information-text-files/

  Dim iFileNum As Long, lErrLast As Long
  Dim sVarName As String, sVarValue As String


  ' assume false unless variable is found
  GetConfigFromFile = False

  ' open text file to read settings
  If FileExists(sFile) Then
    iFileNum = FreeFile
    Open sFile For Input As iFileNum
      Do While Not EOF(iFileNum)
        Input #iFileNum, sVarName, sVarValue
        If LCase(sVarName) = LCase(sName) Then
          sValue = sVarValue
          GetConfigFromFile = True
          Exit Do
        End If
      Loop
    Close #iFileNum
  End If

End Function
