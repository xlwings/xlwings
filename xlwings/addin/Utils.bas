Attribute VB_Name = "Utils"
Option Explicit
#Const App = "Microsoft Excel" 'Adjust when using outside of Excel

Function WScript(Optional CreateNew As Boolean) As Object
  Static Value As Object
  If CreateNew Or Value Is Nothing Then Set Value = CreateObject("WScript.Shell")
  Set WScript = Value
End Function

Function IsFullName(sFile As String) As Boolean
  ' if sFile includes path, it contains path separator "\" or "/"
  IsFullName = InStr(sFile, "\") + InStr(sFile, "/") > 0
End Function

Function FileExists(ByVal FileSpec As String) As Boolean
    #If Mac Then
        FileExists = FileOrFolderExistsOnMac(FileSpec)
    #Else
        FileExists = FileExistsOnWindows(FileSpec)
    #End If
End Function

Function FileExistsOnWindows(ByVal FileSpec As String) As Boolean
   ' by Karl Peterson MS MVP VB
   Dim Attr As Long
   ' Guard against bad FileSpec by ignoring errors
   ' retrieving its attributes.
   On Error Resume Next
   Attr = GetAttr(FileSpec)
   If Err.Number = 0 Then
      ' No error, so something was found.
      ' If Directory attribute set, then not a file.
      FileExistsOnWindows = Not ((Attr And vbDirectory) = vbDirectory)
   End If
End Function


Function FileOrFolderExistsOnMac(FileOrFolderstr As String) As Boolean
'Ron de Bruin : 26-June-2015
'Function to test whether a file or folder exist on a Mac in office 2011 and up
'Uses AppleScript to avoid the problem with long names in Office 2011,
'limit is max 32 characters including the extension in 2011.
    Dim ScriptToCheckFileFolder As String
    Dim TestStr As String
    
    #If Mac Then
    If Val(Application.VERSION) < 15 Then
        ScriptToCheckFileFolder = "tell application " & Chr(34) & "System Events" & Chr(34) & _
         "to return exists disk item (" & Chr(34) & FileOrFolderstr & Chr(34) & " as string)"
        FileOrFolderExistsOnMac = MacScript(ScriptToCheckFileFolder)
    Else
        On Error Resume Next
        TestStr = Dir(FileOrFolderstr, vbDirectory)
        On Error GoTo 0
        If Not TestStr = vbNullString Then FileOrFolderExistsOnMac = True
    End If
    #End If
End Function

Function ParentFolder(ByVal Folder)
  #If Mac Then
      ParentFolder = Left$(Folder, InStrRev(Folder, "/") - 1)
  #Else
      ParentFolder = Left$(Folder, InStrRev(Folder, "\") - 1)
  #End If
End Function

Function GetDirectory(Path)
    #If Mac Then
    GetDirectory = Left(Path, InStrRev(Path, "/"))
    #Else
    GetDirectory = Left(Path, InStrRev(Path, "\"))
    #End If
End Function

Function KillFileOnMac(Filestr As String)
    'Ron de Bruin
    '30-July-2012
    'Delete files from a Mac.
    'Uses AppleScript to avoid the problem with long file names (on 2011 only)

    Dim ScriptToKillFile As String
    
    #If Mac Then
    ScriptToKillFile = "tell application " & Chr(34) & "Finder" & Chr(34) & Chr(13)
    ScriptToKillFile = ScriptToKillFile & "do shell script ""rm "" & quoted form of posix path of " & Chr(34) & Filestr & Chr(34) & Chr(13)
    ScriptToKillFile = ScriptToKillFile & "end tell"

    On Error Resume Next
        MacScript (ScriptToKillFile)
    On Error GoTo 0
    #End If
End Function

Function ToMacPath(PosixPath As String) As String
    ' This function transforms a Posix Path into a MacOS Path
    ' E.g. "/Users/<User>" --> "MacintoshHD:Users:<User>"
    #If Mac Then
    ToMacPath = MacScript("set mac_path to POSIX file " & Chr(34) & PosixPath & Chr(34) & " as string")
    #End If
End Function

Function GetMacDir(Name As String, Normalize As Boolean) As String
    #If Mac Then
        Select Case Name
            Case "$HOME"
                Name = "home folder"
            Case "$APPLICATIONS"
                Name = "applications folder"
            Case "$DOCUMENTS"
                Name = "documents folder"
            Case "$DOWNLOADS"
                Name = "downloads folder"
            Case "$DESKTOP"
                Name = "desktop folder"
            Case "$TMPDIR"
                Name = "temporary items"
        End Select
        GetMacDir = MacScript("return POSIX path of (path to " & Name & ") as string")
        If Normalize = True Then
            'Normalize Excel sandbox location
            GetMacDir = Replace(GetMacDir, "/Library/Containers/com.microsoft.Excel/Data", "")
        End If
    #Else
    #End If
End Function


Function ToPosixPath(ByVal MacPath As String) As String
    'This function accepts relative paths with backward and forward slashes: ActiveWorkbook & "\test"
    ' E.g. "MacintoshHD:Users:<User>" --> "/Users/<User>"

    Dim s As String
    Dim LeadingSlash As Boolean
    
    #If Mac Then
    If MacPath = "" Then
        ToPosixPath = ""
    Else
        ToPosixPath = Replace(MacPath, "\", "/")
        ToPosixPath = MacScript("return POSIX path of (" & Chr(34) & MacPath & Chr(34) & ") as string")
    End If
    #End If
End Function

Sub ShowError(FileName As String, Optional Message As String = "")
    ' Shows a MsgBox with the content of a text file

    Dim Content As String
    Dim ErrorSheet As Worksheet

    Const OK_BUTTON_ERROR = 16
    Const AUTO_DISMISS = 0
    
    If Message = "" Then
        Content = ReadFile(FileName)
    Else
        Content = Message
    End If
    

    If GetConfig("SHOW_ERROR_POPUPS", "True") = "False" Then
        If SheetExists(ActiveWorkbook, "Error") = False Then
            Set ErrorSheet = ActiveWorkbook.Sheets.Add()
            ErrorSheet.Name = "Error"
        Else
            Set ErrorSheet = ActiveWorkbook.Sheets("Error")
        End If
        ErrorSheet.Range("A1").Value = Content
    Else
        #If Mac Then
            MsgBox Content, vbCritical, "Error"
        #Else
            Content = Content & vbCrLf
            Content = Content & "Press Ctrl+C to copy this message to the clipboard."
    
            WScript.Popup Content, AUTO_DISMISS, "Error", OK_BUTTON_ERROR
        #End If
    End If
End Sub

Function ExpandEnvironmentStrings(ByVal s As String)
    ' Expand environment variables
    Dim EnvString As String
    Dim PathParts As Variant
    Dim i As Integer
    #If Mac Then
        If Left(s, 1) = "$" Then
            PathParts = Split(s, "/")
            EnvString = PathParts(0)
            ExpandEnvironmentStrings = GetMacDir(EnvString, True)
            For i = 1 To UBound(PathParts)
                If Right$(ExpandEnvironmentStrings, 1) = "/" Then
                    ExpandEnvironmentStrings = ExpandEnvironmentStrings & PathParts(i)
                Else
                    ExpandEnvironmentStrings = ExpandEnvironmentStrings & "/" & PathParts(i)
                End If
            Next i
        Else
            ExpandEnvironmentStrings = s
        End If
    #Else
        ExpandEnvironmentStrings = WScript.ExpandEnvironmentStrings(s)
    #End If
End Function

Function ReadFile(ByVal FileName As String)
    ' Read a text file

    Dim Content As String
    Dim Token As String
    Dim FileNum As Integer
    Dim objShell As Object
    Dim LineBreak As Variant

    #If Mac Then
        FileName = ToMacPath(FileName)
        LineBreak = vbLf
    #Else
        FileName = ExpandEnvironmentStrings(FileName)
        LineBreak = vbCrLf
    #End If

    FileNum = FreeFile
    Content = ""

    ' Read Text File
    Open FileName For Input As #FileNum
        Do While Not EOF(FileNum)
            Line Input #FileNum, Token
            Content = Content & Token & LineBreak
        Loop
    Close #FileNum

    ReadFile = Content
End Function

#If App = "Microsoft Excel" Then
Function SheetExists(wb As Workbook, sheetName As String) As Boolean
    Dim sht As Worksheet
    On Error Resume Next
        Set sht = wb.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not sht Is Nothing
End Function
#End If

Function GetBaseName(wb As String) As String
    Dim extension As String
    extension = LCase$(Right$(wb, 4))
    If extension = ".xls" Or extension = ".xla" Or extension = ".xlt" Then
        GetBaseName = Left$(wb, Len(wb) - 4)
    Else
        GetBaseName = Left$(wb, Len(wb) - 5)
    End If
End Function

Function has_dynamic_array() As Boolean
    has_dynamic_array = False
    On Error GoTo ErrHandler
        Application.WorksheetFunction.Unique ("dummy")
        has_dynamic_array = True
    Exit Function
ErrHandler:
    has_dynamic_array = False
End Function

Public Function CreateGUID() As String
    Randomize Timer() + Application.Hwnd
    ' https://stackoverflow.com/a/46474125/918626
    Do While Len(CreateGUID) < 32
        If Len(CreateGUID) = 16 Then
            '17th character holds version information
            CreateGUID = CreateGUID & Hex$(8 + CInt(Rnd * 3))
        End If
        CreateGUID = CreateGUID & Hex$(CInt(Rnd * 15))
    Loop
    CreateGUID = Mid(CreateGUID, 1, 8) & "-" & Mid(CreateGUID, 9, 4) & "-" & Mid(CreateGUID, 13, 4) & "-" & Mid(CreateGUID, 17, 4) & "-" & Mid(CreateGUID, 21, 12)
End Function

Function CheckConda(CondaPath As String) As Boolean
    ' Check if the conda executable exists.
    ' If it doesn't, conda is too old and the Interpreter setting has to be used instead of Conda settings
    Dim condaExecutable As String
    Dim condaExists As Boolean
    #If Mac Then
        condaExecutable = CondaPath & "\condabin\conda"
    #Else
        condaExecutable = CondaPath & "\condabin\conda.bat"
    #End If
    ' Replace space escape character ^ to check if path exists
    condaExists = FileExists(Replace(condaExecutable, "^", ""))
    If condaExists = False And CondaPath <> "" Then
        MsgBox "Your Conda version seems to be too old for the Conda settings. Use the Interpreter setting instead."
    End If
    CheckConda = condaExists
End Function

#If App = "Microsoft Excel" Then
Function GetFullName(wb As Workbook) As String
    ' The only case where this is still used is for directory-based config files, otherwise this is now handled in Python
    ' Unlike the Python version, this doesn't work for SharePoint and will just ignore a directory-based config file silently

    Dim total_found, i_parsing, i_env_var, slash_number As Integer
    Dim found_path, one_drive_path, full_path_name, this_found_path As String

    ' In the majority of cases, ThisWorkbook.FullName will provide the path of the
    ' Excel workbook correctly. Unfortunately, when the user is using OneDrive
    ' this doesn't work. This function will attempt to find the LOCAL path.
    ' This uses code from Daniel Guetta and
    ' https://stackoverflow.com/questions/33734706/excels-fullname-property-with-onedrive
    
    If InStr(wb.FullName, "://") = 0 Or wb.Path = "" Then
        GetFullName = wb.FullName
        Exit Function
    End If
        
    ' According to the link above, there are three possible environment variables
    ' the user's OneDrive folder could be located in
    '      "OneDriveCommercial", "OneDriveConsumer", "OneDrive"
    '
    ' Furthermore, there are two possible formats for OneDrive URLs
    '    1. "https://companyName-my.sharepoint.com/personal/userName_domain_com/Documents" & file.FullName
    '    2. "https://d.docs.live.net/d7bbaa#######1/" & file.FullName
    ' In the first case, we can find the true path by just looking for everything after /Documents. In the
    ' second, we need to look for the fourth slash in the URL
    '
    ' The code below will try every combination of the three environment variables above, and
    ' each of the two methods of parsing the URL. The file is found in *exactly* one of those
    ' locations, then we're good to go.
    '
    ' Note that this still leaves a gap - if this file (file A) is in a location that is NOT covered by the
    ' eventualities above AND a file of the exact same name (file B) exists in one of the locations that is
    ' covered above, then this function will identify File B's location as the location of this workbook,
    ' which would be wrong
    total_found = 0
    
    For i_parsing = 1 To 2
        If i_parsing = 1 Then
            ' Parse using method 1 above; find /Documents and take everything after, INCLUDING the
            ' leading slash
            If InStr(1, wb.FullName, "/Documents") Then
                full_path_name = Mid(wb.FullName, InStr(1, wb.FullName, "/Documents") + Len("/Documents"))
            Else
                full_path_name = ""
            End If
        Else
            ' Parse using method 2; find everything after the fourth slash, including that fourth
            ' slash
            Dim i_pos As Integer
            
            ' Start at the last slash in https://
            i_pos = 8

            For slash_number = 1 To 2
                i_pos = InStr(i_pos + 1, wb.FullName, "/")
            Next slash_number
            
            full_path_name = Mid(wb.FullName, i_pos)
        End If
        
        ' Replace forward slahes with backslashes on Windows
        full_path_name = Replace(full_path_name, "/", Application.pathSeparator)
        
        
        If full_path_name <> "" Then
            #If Not Mac Then
            For i_env_var = 1 To 3
                    one_drive_path = Environ(Choose(i_env_var, "OneDriveCommercial", "OneDriveConsumer", "OneDrive"))
                
                    If (one_drive_path <> "") And FileExists(one_drive_path & full_path_name) Then
                        this_found_path = one_drive_path & full_path_name
                        
                        If this_found_path <> found_path Then
                            total_found = total_found + 1
                            found_path = this_found_path
                        End If
                    End If
            Next i_env_var
            #End If
        End If
    Next i_parsing
        
    If total_found = 1 Then
        GetFullName = found_path
        Exit Function
    End If

End Function
#End If

Function GetAzureAdAccessToken()
    Dim nowTs As Long, expiresTs As Long

    expiresTs = GetConfig("AZUREAD_ACCESS_TOKEN_EXPIRES_ON", 0)
    nowTs = DateDiff("s", #1/1/1970#, ConvertToUtc(Now()))

    If (expiresTs > 0) And (nowTs < (expiresTs - 30)) Then
        GetAzureAdAccessToken = GetConfig("AZUREAD_ACCESS_TOKEN")
        Exit Function
    Else
        RunPython "from xlwings import cli;cli._auth_aad()"
        GetAzureAdAccessToken = GetConfig("AZUREAD_ACCESS_TOKEN")
    End If
End Function