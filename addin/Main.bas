Attribute VB_Name = "Main"
'Option Explicit
#If VBA7 Then
    #If Mac Then
        Private Declare PtrSafe Function system Lib "libc.dylib" (ByVal Command As String) As Long
    #End If
    #If Win64 Then
        Const XLPyDLLName As String = "xlwings64.dll"
        Declare PtrSafe Function XLPyDLLActivateAuto Lib "xlwings64.dll" (ByRef result As Variant, Optional ByVal config As String = "", Optional ByVal mode As Long = 1) As Long
        Declare PtrSafe Function XLPyDLLNDims Lib "xlwings64.dll" (ByRef src As Variant, ByRef dims As Long, ByRef transpose As Boolean, ByRef dest As Variant) As Long
        Declare PtrSafe Function XLPyDLLVersion Lib "xlwings64.dll" (tag As String, version As Double, arch As String) As Long
    #Else
        Private Const XLPyDLLName As String = "xlwings32.dll"
        Declare PtrSafe Function XLPyDLLActivateAuto Lib "xlwings32.dll" (ByRef result As Variant, Optional ByVal config As String = "", Optional ByVal mode As Long = 1) As Long
        Private Declare PtrSafe Function XLPyDLLNDims Lib "xlwings32.dll" (ByRef src As Variant, ByRef dims As Long, ByRef transpose As Boolean, ByRef dest As Variant) As Long
        Private Declare PtrSafe Function XLPyDLLVersion Lib "xlwings32.dll" (tag As String, version As Double, arch As String) As Long
    #End If
    Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
#Else
    #If Mac Then
        Private Declare Function system Lib "libc.dylib" (ByVal Command As String) As Long
    #End If
    Private Const XLPyDLLName As String = "xlwings32.dll"
    Private Declare Function XLPyDLLActivateAuto Lib "xlwings32.dll" (ByRef result As Variant, Optional ByVal config As String = "", Optional ByVal mode As Long = 1) As Long
    Private Declare Function XLPyDLLNDims Lib "xlwings32.dll" (ByRef src As Variant, ByRef dims As Long, ByRef transpose As Boolean, ByRef dest As Variant) As Long
    Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
    Declare Function XLPyDLLVersion Lib "xlwings32.dll" (tag As String, version As Double, arch As String) As Long
#End If

Public Function RunPython(PythonCommand As String)
    ' Public API: Runs the Python command, e.g.: to run the function foo() in module bar, call the function like this:
    ' RunPython ("import bar; bar.foo()")

    Dim INTERPRETER As String, PYTHONPATH As String, LOG_FILE As String, OPTIMIZED_CONNECTION As String

    INTERPRETER = GetConfig("INTERPRETER", "")
    PYTHONPATH = ActiveWorkbook.Path & ";" & GetConfig("PYTHONPATH", "")
    LOG_FILE = GetConfig("LOG_FILE", "")
    OPTIMIZED_CONNECTION = GetConfig("UDF_SERVER", "False")

    ' Call Python platform-dependent
    #If Mac Then
            Application.StatusBar = "Running..."  ' Non-blocking way of giving feedback that something is happening
        #If MAC_OFFICE_VERSION >= 15 Then
            ExecuteMac PythonCommand, INTERPRETER, LOG_FILE, PYTHONPATH
        #Else
            ExcecuteMac2011 PythonCommand, INTERPRETER, LOG_FILE, PYTHONPATH
        #End If
    #Else
        If OPTIMIZED_CONNECTION = "True" Then
            Py.SetAttr Py.Module("xlwings._xlwindows"), "BOOK_CALLER", ActiveWorkbook
            Py.Exec "" & PythonCommand & ""
        Else
            ExecuteWindows False, PythonCommand, INTERPRETER, LOG_FILE, PYTHONPATH
        End If
    #End If
End Function

Sub ExcecuteMac2011(PythonCommand As String, PYTHON_MAC As String, LOG_FILE As String, Optional PYTHONPATH As String)
    #If Mac Then
    ' Run Python with the "-c" command line switch: add the path of the python file and run the
    ' Command as first argument, then provide the WORKBOOK_FULLNAME and "from_xl" as 2nd and 3rd arguments.
    ' Finally, redirect stderr to the LOG_FILE and run as background process.

    Dim PythonInterpreter As String, RunCommand As String, WORKBOOK_FULLNAME As String, Log As String

    If LOG_FILE = "" Then
        LOG_FILE = "/tmp/xlwings_log.txt"
    Else
        LOG_FILE = ToPosixPath(LOG_FILE)
    End If

    ' Delete Log file just to make sure we don't show an old error
    On Error Resume Next
        KillFileOnMac ToMacPath(LOG_FILE)
    On Error GoTo 0

    ' Transform from MacOS Classic path style (":") and Windows style ("\") to Bash friendly style ("/")
    PYTHONPATH = ToPosixPath(PYTHONPATH)
    If PYTHON_MAC <> "" Then
        PythonInterpreter = ToPosixPath(PYTHON_MAC)
    Else
        PythonInterpreter = "python"
    End If
    WORKBOOK_FULLNAME = ToPosixPath(ActiveWorkbook.Path & ":" & ActiveWorkbook.Name) 'ActiveWorkbook.FullName doesn't handle unicode on Excel 2011

    ' Build the command (ignore warnings to be in line with Windows where we only show the popup if the ExitCode <> 0
    ' -u is needed because on PY3 stderr is buffered by default and so wouldn't be available on time for the pop-up to show
    RunCommand = PythonInterpreter & " -u -B -W ignore -c ""import sys, os; sys.path.extend(os.path.normcase(os.path.expandvars(r'" & PYTHONPATH & "')).split(';')); " & PythonCommand & """ "

    ' Send the command to the shell. Courtesy of Robert Knight (http://stackoverflow.com/a/12320294/918626)
    ' Since Excel blocks AppleScript as long as a VBA macro is running, we have to excecute the call as background call
    ' so it can do its magic after this Function has terminated. Python calls ClearUp via the atexit handler.

    'Check if .bash_profile is existing and source it
    Res = system("source ~/.bash_profile")
    If Res = 0 Then
        Res = system("source ~/.bash_profile;" & RunCommand & """" & WORKBOOK_FULLNAME & """ ""from_xl""" & " " & Chr(34) & ToPosixPath(Application.Path) & "/" & Application.Name & Chr(34) & ">" & Chr(34) & LOG_FILE & Chr(34) & " 2>&1 &")
    Else
        Res = system(RunCommand & """" & WORKBOOK_FULLNAME & """ ""from_xl""" & " " & Chr(34) & ToPosixPath(Application.Path) & "/" & Application.Name & Chr(34) & ">" & Chr(34) & LOG_FILE & Chr(34) & " 2>&1 &")
    End If

    ' If there's a log at this point (normally that will be from the Shell only, not Python) show it and reset the StatusBar
    On Error Resume Next
        Log = ReadFile(LOG_FILE)
        If Log = "" Then
            Exit Sub
        Else
            ShowError (LOG_FILE)
            Application.StatusBar = False
        End If
    On Error GoTo 0
    #End If
End Sub

Sub ExecuteMac(PythonCommand As String, PYTHON_MAC As String, LOG_FILE As String, Optional PYTHONPATH As String)
    #If Mac Then
    Dim PythonInterpreter As String, RunCommand As String, WORKBOOK_FULLNAME As String, Log As String, ParameterString As String, ExitCode As String
    Dim Res As Integer

    ' Transform paths
    PYTHONPATH = ToPosixPath(PYTHONPATH)

    If PYTHON_MAC <> "" Then
        PythonInterpreter = ToPosixPath(PYTHON_MAC)
    Else
        PythonInterpreter = "python"
    End If

    WORKBOOK_FULLNAME = ToPosixPath(ActiveWorkbook.FullName)
    If LOG_FILE = "" Then
        ' Sandbox location that requires no file access confirmation
        LOG_FILE = Environ("HOME") + "/xlwings_log.txt" '/Users/<User>/Library/Containers/com.microsoft.Excel/Data/xlwings_log.txt
    Else
        LOG_FILE = ToPosixPath(LOG_FILE)
    End If

    ' Delete Log file just to make sure we don't show an old error
    On Error Resume Next
        Kill LOG_FILE
    On Error GoTo 0

    ' ParameterSting with all paramters (AppleScriptTask only accepts a single parameter)
    ParameterString = PYTHONPATH + ";"
    ParameterString = ParameterString + "," + PythonInterpreter
    ParameterString = ParameterString + "," + PythonCommand
    ParameterString = ParameterString + "," + ActiveWorkbook.FullName
    ParameterString = ParameterString + "," + Left(Application.Path, Len(Application.Path) - 4)
    ParameterString = ParameterString + "," + LOG_FILE

    On Error GoTo AppleScriptErrorHandler
        ExitCode = AppleScriptTask("xlwings.applescript", "VbaHandler", ParameterString)
    On Error GoTo 0

    ' If there's a log at this point (normally that will be from the Shell only, not Python) show it and reset the StatusBar
    On Error Resume Next
        Log = ReadFile(LOG_FILE)
        If Log = "" Then
            Exit Sub
        Else
            ShowError (LOG_FILE)
            Application.StatusBar = False
        End If
        Exit Sub
    On Error GoTo 0

AppleScriptErrorHandler:
    MsgBox "To enable RunPython, please run 'xlwings runpython install' in a terminal once and try again.", vbCritical
    #End If
End Sub

Sub ExecuteWindows(IsFrozen As Boolean, PythonCommand As String, PYTHON_WIN As String, LOG_FILE As String, Optional PYTHONPATH As String)
    ' Call a command window and change to the directory of the Python installation or frozen executable
    ' Note: If Python is called from a different directory with the fully qualified path, pywintypesXX.dll won't be found.
    ' This seems to be a general issue with pywin32, see http://stackoverflow.com/q/7238403/918626

    Dim Wsh As Object
    Dim WaitOnReturn As Boolean: WaitOnReturn = True
    Dim WindowStyle As Integer: WindowStyle = 0
    Set Wsh = CreateObject("WScript.Shell")
    Dim DriveCommand As String, RunCommand As String, WORKBOOK_FULLNAME As String, PythonInterpreter As String, PythonDir As String
    Dim ExitCode As Integer

    If LOG_FILE = "" Then
        LOG_FILE = Environ("APPDATA") + "\xlwings_log.txt"
    End If

    If Not IsFrozen And (PYTHON_WIN <> "python" And PYTHON_WIN <> "pythonw") Then
        PythonDir = ParentFolder(PYTHON_WIN)
    Else
        PythonDir = ""  ' TODO: hack
    End If

    If Left$(PYTHON_WIN, 2) Like "[A-Za-z]:" Then
        ' If Python is installed on a mapped or local drive, change to drive, then cd to path
        DriveCommand = Left$(PYTHON_WIN, 2) & " & cd """ & PythonDir & """ & "
    ElseIf Left$(PYTHON_WIN, 2) = "\\" Then
        ' If Python is installed on a UNC path, temporarily mount and activate a drive letter with pushd
        DriveCommand = "pushd """ & PythonDir & """ & "
    End If

    ' Run Python with the "-c" command line switch: add the path of the python file and run the
    ' Command as first argument, then provide the WORKBOOK_FULLNAME and "from_xl" as 2nd and 3rd arguments.
    ' Then redirect stderr to the LOG_FILE and wait for the call to return.
    WORKBOOK_FULLNAME = ActiveWorkbook.FullName

    If PYTHON_WIN <> "python" And PYTHON_WIN <> "pythonw" Then
        PythonInterpreter = Chr(34) & PYTHON_WIN & Chr(34)
    Else
        PythonInterpreter = "python"
    End If

    If IsFrozen = False Then
        RunCommand = PythonInterpreter & " -B -c ""import sys, os; sys.path.extend(os.path.normcase(os.path.expandvars(r'" & PYTHONPATH & "')).split(';')); " & PythonCommand & """ "
    ElseIf IsFrozen = True Then
        RunCommand = PythonCommand & " "
    End If

    ExitCode = Wsh.Run("cmd.exe /C " & DriveCommand & _
                   RunCommand & _
                   """" & WORKBOOK_FULLNAME & """ ""from_xl""" & " " & Chr(34) & _
                   Application.Path & "\" & Application.Name & Chr(34) & " " & Chr(34) & Application.Hwnd & Chr(34) & _
                   " 2> """ & LOG_FILE & """ ", _
                   WindowStyle, WaitOnReturn)

    'If ExitCode <> 0 then there's something wrong
    If ExitCode <> 0 Then
        Call ShowError(LOG_FILE)
    End If

    ' Delete file after the error message has been shown
    On Error Resume Next
        Kill LOG_FILE
    On Error GoTo 0

    ' Clean up
    Set Wsh = Nothing
End Sub

Public Function RunFrozenPython(Executable As String)
    ' Runs a Python executable that has been frozen by cx_Freeze or py2exe. Call the function like this:
    ' RunFrozenPython("frozen_executable.exe"). Currently not implemented for Mac.

    Dim PYTHON_FROZEN As String, LOG_FILE As String

    PYTHON_FROZEN = GetConfig("PYTHON_FROZEN", ThisWorkbook.Path & "build\exe.win32-2.7")
    LOG_FILE = GetConfig("LOG_FILE", "")

    ' Call Python
    #If Mac Then
        MsgBox "This functionality is not yet supported on Mac." & vbNewLine & _
               "Please run your scripts directly in Python!", vbCritical + vbOKOnly, "Unsupported Feature"
    #Else
        ExecuteWindows True, Executable, PYTHON_FROZEN, LOG_FILE
    #End If
End Function

Function GetUdfModules() As String
    Dim UDF_MODULES As String

    UDF_MODULES = GetConfig("UDF_MODULES", "")

    If UDF_MODULES = "" Then
        GetUdfModules = Left$(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) ' assume that it ends in .xlsm
    Else
        GetUdfModules = UDF_MODULES
    End If
End Function

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

Sub ShowError(FileName As String)
    ' Shows a MsgBox with the content of a text file

    Dim Content As String
    Dim objShell

    Const OK_BUTTON_ERROR = 16
    Const AUTO_DISMISS = 0

    Content = ReadFile(FileName)
    #If Win32 Or Win64 Then
        Content = Content & vbCrLf
        Content = Content & "Press Ctrl+C to copy this message to the clipboard."

        Set objShell = CreateObject("Wscript.Shell")
        objShell.Popup Content, AUTO_DISMISS, "Error", OK_BUTTON_ERROR
    #Else
        MsgBox Content, vbCritical, "Error"
    #End If

End Sub

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

Function ToMacPath(PosixPath As String) As String
    ' This function transforms a Posix Path into a MacOS Path
    ' E.g. "/Users/<User>" --> "MacintoshHD:Users:<User>"

    ToMacPath = MacScript("set mac_path to POSIX file " & Chr(34) & PosixPath & Chr(34) & " as string")
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

Private Sub CleanUp()
    'On Mac only, this function is being called after Python is done (using Python's atexit handler)

    Dim PYTHON_MAC As String, PYTHON_FROZEN As String, PYTHONPATH As String
    Dim WORKBOOK_FULLNAME As String, LOG_FILE As String
    Dim Res As Integer
    Dim UDF_DEBUG_SERVER As Boolean

    If LOG_FILE = "" Then
        #If MAC_OFFICE_VERSION >= 15 Then
            LOG_FILE = Environ("HOME") + "/xlwings_log.txt" '/Users/<User>/Library/Containers/com.microsoft.Excel/Data/xlwings_log.txt
        #Else
            LOG_FILE = "/tmp/xlwings_log.txt"
        #End If
    Else
        LOG_FILE = ToPosixPath(LOG_FILE)
    End If

    'Show the LOG_FILE as MsgBox if not empty
    On Error Resume Next
    If ReadFile(LOG_FILE) <> "" Then
        Call ShowError(LOG_FILE)
    End If
    On Error GoTo 0

    'Clean up
    Application.StatusBar = False
    Application.ScreenUpdating = True
    On Error Resume Next
        #If MAC_OFFICE_VERSION >= 15 Then
            Kill LOG_FILE
        #Else
            KillFileOnMac ToMacPath(ToPosixPath(LOG_FILE))
        #End If
    On Error GoTo 0
End Sub

Function ParentFolder(ByVal Folder)
  ParentFolder = Left$(Folder, InStrRev(Folder, "\") - 1)
End Function

Function XLPyCommand()
    Dim PYTHON_WIN As String, PYTHONPATH As String, LOG_FILE As String, tail As String

    PYTHONPATH = ActiveWorkbook.Path & ";" & GetConfig("PYTHONPATH", "")
    PYTHON_WIN = GetConfig("INTERPRETER", "pythonw")
    UDF_DEBUG = GetConfig("UDF_DEBUG", "False")

    If UDF_DEBUG = "True" Then
        XLPyCommand = "{506e67c3-55b5-48c3-a035-eed5deea7d6d}"
    Else
        tail = " -B -c ""import sys, os;sys.path.extend(os.path.normcase(os.path.expandvars(r'" & PYTHONPATH & "')).split(';'));import xlwings.server; xlwings.server.serve('$(CLSID)')"""
            XLPyCommand = PYTHON_WIN + tail
    End If
End Function

Private Sub XLPyLoadDLL()
    Dim PYTHON_WIN As String

    PYTHON_WIN = GetConfig("INTERPRETER", "pythonw")

    If PYTHON_WIN <> "python" And PYTHON_WIN <> "pythonw" Then
        If LoadLibrary(ParentFolder(PYTHON_WIN) + "\" + XLPyDLLName) = 0 Then  ' Standard installation
            If LoadLibrary(ParentFolder(ParentFolder(PYTHON_WIN)) + "\" + XLPyDLLName) = 0 Then  ' Virtualenv
                Err.Raise 1, Description:= _
                    "Could not load " + XLPyDLLName + " from either of the following folders: " _
                    + vbCrLf + ParentFolder(PYTHON_WIN) _
                    + vbCrLf + ", " + ParentFolder(ParentFolder(PYTHON_WIN))
            End If
        End If
    End If
End Sub

Function NDims(ByRef src As Variant, dims As Long, Optional transpose As Boolean = False)
    XLPyLoadDLL
    If 0 <> XLPyDLLNDims(src, dims, transpose, NDims) Then Err.Raise 1001, Description:=NDims
End Function

Function Py()
    XLPyLoadDLL
    If 0 <> XLPyDLLActivateAuto(Py, XLPyCommand, 1) Then Err.Raise 1000, Description:=Py
End Function

Sub KillPy()
    XLPyLoadDLL
    Dim unused
    If 0 <> XLPyDLLActivateAuto(unused, XLPyCommand, -1) Then Err.Raise 1000, Description:=unused
End Sub

Private Sub GetDLLVersion()
    ' Currently only for testing
    Dim tag As String, arch As String
    Dim ver As Double
    XLPyDLLVersion tag, ver, arch
    Debug.Print tag
    Debug.Print ver
    Debug.Print arch
End Sub

Sub ImportPythonUDFs()
    Dim tempPath As String
    tempPath = Py.Str(Py.Call(Py.Module("xlwings"), "import_udfs", Py.Tuple(GetUdfModules, ActiveWorkbook)))
End Sub

