Attribute VB_Name = "Utils"
Option Explicit

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

Function ParentFolder(ByVal Folder)
  #If Mac Then
      ParentFolder = Left$(Folder, InStrRev(Folder, "/") - 1)
  #Else
      ParentFolder = Left$(Folder, InStrRev(Folder, "\") - 1)
  #End If
End Function

Function KillFileOnMac(Filestr As String)
    'Ron de Bruin
    '30-July-2012
    'Delete files from a Mac.
    'Uses AppleScript to avoid the problem with long file names (on 2011 only)

    Dim ScriptToKillFile As String

    ScriptToKillFile = "tell application " & Chr(34) & "Finder" & Chr(34) & Chr(13)
    ScriptToKillFile = ScriptToKillFile & "do shell script ""rm "" & quoted form of posix path of " & Chr(34) & Filestr & Chr(34) & Chr(13)
    ScriptToKillFile = ScriptToKillFile & "end tell"

    On Error Resume Next
        MacScript (ScriptToKillFile)
    On Error GoTo 0
End Function

Function ToMacPath(PosixPath As String) As String
    ' This function transforms a Posix Path into a MacOS Path
    ' E.g. "/Users/<User>" --> "MacintoshHD:Users:<User>"

    ToMacPath = MacScript("set mac_path to POSIX file " & Chr(34) & PosixPath & Chr(34) & " as string")
End Function

Function GetMacDir(dirName As String) As String
    ' Get Mac special folders. Protetcted so they don't exectue on Windows.

    Dim Path As String

    #If Mac Then
        Select Case dirName
            Case "Home"
                Path = MacScript("return POSIX path of (path to home folder) as string")
             Case "Desktop"
                Path = MacScript("return POSIX path of (path to desktop folder) as string")
            Case "Applications"
                Path = MacScript("return POSIX path of (path to applications folder) as string")
            Case "Documents"
                Path = MacScript("return POSIX path of (path to documents folder) as string")
        End Select
            GetMacDir = Left$(Path, Len(Path) - 1) ' get rid of trailing "/"
    #Else
        GetMacDir = ""
    #End If
End Function

Function ToPosixPath(ByVal MacPath As String) As String
    'This function accepts relative paths with backward and forward slashes: ActiveWorkbook & "\test"
    ' E.g. "MacintoshHD:Users:<User>" --> "/Users/<User>"

    Dim s As String
    Dim LeadingSlash As Boolean

    If MacPath = "" Then
        ToPosixPath = ""
    Else
        #If MAC_OFFICE_VERSION < 15 Then
            If Left$(MacPath, 1) = "/" Then
                LeadingSlash = True
            End If
            MacPath = Replace(MacPath, "\", ":")
            MacPath = Replace(MacPath, "/", ":")
            s = "tell application " & Chr(34) & "Finder" & Chr(34) & Chr(13)
            s = s & "POSIX path of " & Chr(34) & MacPath & Chr(34) & Chr(13)
            s = s & "end tell" & Chr(13)
            If LeadingSlash = True Then
                ToPosixPath = "/" + MacScript(s)
            Else
                ToPosixPath = MacScript(s)
            End If
            If Left$(ToPosixPath, 2) = "/$" Then
                ' If it starts with an env variables, it's otherwise not correctly returned
                ToPosixPath = Right$(ToPosixPath, Len(ToPosixPath) - 1)
            End If

        #Else
            ToPosixPath = Replace(MacPath, "\", "/")
        #End If
    End If
End Function

Sub ShowError(FileName As String)
    ' Shows a MsgBox with the content of a text file

    Dim Content As String
    Dim objShell

    Const OK_BUTTON_ERROR = 16
    Const AUTO_DISMISS = 0

    Content = ReadFile(FileName)
    #If Mac Then
        MsgBox Content, vbCritical, "Error"
    #Else
        Content = Content & vbCrLf
        Content = Content & "Press Ctrl+C to copy this message to the clipboard."

        Set objShell = CreateObject("Wscript.Shell")
        objShell.Popup Content, AUTO_DISMISS, "Error", OK_BUTTON_ERROR
    #End If

End Sub

Function ReadFile(ByVal FileName As String)
    ' Read a text file

    Dim Content As String
    Dim Token As String
    Dim FileNum As Integer
    Dim objShell As Object

    #If Mac Then
        FileName = ToMacPath(FileName)
    #Else
        Set objShell = CreateObject("WScript.Shell")
        FileName = objShell.ExpandEnvironmentStrings(FileName)
    #End If

    FileNum = FreeFile
    Content = ""

    ' Read Text File
    Open FileName For Input As #FileNum
        Do While Not EOF(FileNum)
            Line Input #FileNum, Token
            Content = Content & Token & vbCrLf
        Loop
    Close #FileNum

    ReadFile = Content
End Function

Function SheetExists(sheetName As String) As Boolean
    Dim sht As Worksheet
    On Error Resume Next
        Set sht = ActiveWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not sht Is Nothing
End Function
