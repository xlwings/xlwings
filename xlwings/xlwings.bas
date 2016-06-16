Attribute VB_Name = "xlwings"
' xlwings.org, version: 0.7.2
'
' Copyright (C) 2014-2016, Zoomer Analytics LLC (www.zoomeranalytics.com)
' License: BSD 3-clause (see LICENSE.txt for details)
Option Explicit
#If Mac Then
    Private Declare Function system Lib "libc.dylib" (ByVal Command As String) As Long
#End If
#If VBA7 Then
    #If Win64 Then
        Const XLPyDLLName As String = "xlwings64.dll"
        Declare PtrSafe Function XLPyDLLActivateAuto Lib "xlwings64.dll" (ByRef result As Variant, Optional ByVal config As String = "") As Long
        Declare PtrSafe Function XLPyDLLNDims Lib "xlwings64.dll" (ByRef src As Variant, ByRef dims As Long, ByRef transpose As Boolean, ByRef dest As Variant) As Long
        Declare PtrSafe Function XLPyDLLVersion Lib "xlwings64.dll" (tag As String, version As Double, arch As String) As Long
    #Else
        Private Const XLPyDLLName As String = "xlwings32.dll"
        Private Declare PtrSafe Function XLPyDLLActivateAuto Lib "xlwings32.dll" (ByRef result As Variant, Optional ByVal config As String = "") As Long
        Private Declare PtrSafe Function XLPyDLLNDims Lib "xlwings32.dll" (ByRef src As Variant, ByRef dims As Long, ByRef transpose As Boolean, ByRef dest As Variant) As Long
        Private Declare PtrSafe Function XLPyDLLVersion Lib "xlwings32.dll" (tag As String, version As Double, arch As String) As Long
    #End If
    Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
#Else
    Private Const XLPyDLLName As String = "xlwings32.dll"
    Private Declare Function XLPyDLLActivateAuto Lib "xlwings32.dll" (ByRef result As Variant, Optional ByVal config As String = "") As Long
    Private Declare Function XLPyDLLNDims Lib "xlwings32.dll" (ByRef src As Variant, ByRef dims As Long, ByRef transpose As Boolean, ByRef dest As Variant) As Long
    Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
    Declare Function XLPyDLLVersion Lib "xlwings32.dll" (tag As String, version As Double, arch As String) As Long
#End If

Function Settings(ByRef PYTHON_WIN As String, ByRef PYTHON_MAC As String, ByRef PYTHON_FROZEN As String, ByRef PYTHONPATH As String, ByRef UDF_MODULES As String, ByRef UDF_DEBUG_SERVER As Boolean, ByRef LOG_FILE As String, ByRef SHOW_LOG As Boolean, ByRef OPTIMIZED_CONNECTION As Boolean)
    ' PYTHON_WIN: Full path of Python Interpreter on Windows, e.g. "C:\Python35\pythonw.exe". "" resolves to default on PATH
    ' PYTHON_MAC: Full path of Python Interpreter on Mac OSX, e.g. "/usr/local/bin/python3.5". "" resolves to default path in ~/.bash_profile
    ' PYTHON_FROZEN [Optional]: Currently only on Windows, indicate directory of exe file
    ' PYTHONPATH [Optional]: If the source file of your code is not found, add the path here.
    '                        Separate multiple directories by ";". Otherwise set to "".
    ' UDF_MODULES [Optional, Windows only]: Names of Python modules (without .py extension) from which the UDFs are being imported.
    '                                       Separate multiple modules by ";".
    '                                       Example: UDF_MODULES = "common_udfs;myproject"
    '                                       Default: UDF_MODULES = "" defaults to a module in the same directory of the Excel spreadsheet with
    '                                       the same name but ending in ".py".
    ' UDF_DEBUG_SERVER: Set this to True if you want to run the xlwings COM server manually for debugging
    ' LOG_FILE [Optional]: Leave empty for default location (see docs) or provide directory including file name.
    ' SHOW_LOG: If False, no pop-up with the Log messages (usually errors) will be shown
    ' OPTIMIZED_CONNECTION (EXPERIMENTAL!): Currently only on Windows, use a COM Server for an efficient connection
    '
    ' For cross-platform compatibility, use backslashes in relative directories
    ' For details, see http://docs.xlwings.org

    PYTHON_WIN = ""
    PYTHON_MAC = ""
    PYTHON_FROZEN = ThisWorkbook.Path & "\build\exe.win32-2.7"
    PYTHONPATH = ThisWorkbook.Path
    UDF_MODULES = ""
    UDF_DEBUG_SERVER = False
    LOG_FILE = ""
    SHOW_LOG = True
    OPTIMIZED_CONNECTION = False

End Function
' DO NOT EDIT BELOW THIS LINE

Public Function RunPython(PythonCommand As String)
    ' Public API: Runs the Python command, e.g.: to run the function foo() in module bar, call the function like this:
    ' RunPython ("import bar; bar.foo()")

    Dim PYTHON_WIN As String, PYTHON_MAC As String, PYTHON_FROZEN As String, PYTHONPATH As String, UDF_MODULES As String
    Dim WORKBOOK_FULLNAME As String, LOG_FILE As String, DriveCommand As String, RunCommand As String
    Dim ExitCode As Integer, Res As Integer
    Dim SHOW_LOG As Boolean, OPTIMIZED_CONNECTION As Boolean, UDF_DEBUG_SERVER As Boolean

    ' Get the settings by using the ByRef trick
    Res = Settings(PYTHON_WIN, PYTHON_MAC, PYTHON_FROZEN, PYTHONPATH, UDF_MODULES, UDF_DEBUG_SERVER, LOG_FILE, SHOW_LOG, OPTIMIZED_CONNECTION)

    ' Call Python platform-dependent
    #If Mac Then
            Application.StatusBar = "Running..."  ' Non-blocking way of giving feedback that something is happening
        #If MAC_OFFICE_VERSION >= 15 Then
            ExecuteMac PythonCommand, PYTHON_MAC, LOG_FILE, SHOW_LOG, PYTHONPATH
        #Else
            ExcecuteMac2011 PythonCommand, PYTHON_MAC, LOG_FILE, SHOW_LOG, PYTHONPATH
        #End If
    #Else
        If OPTIMIZED_CONNECTION = True Then
            Py.SetAttr Py.Module("xlwings._xlwindows"), "xl_workbook_current", ThisWorkbook
            Py.Exec "" & PythonCommand & ""
        Else
            ExecuteWindows False, PythonCommand, PYTHON_WIN, LOG_FILE, SHOW_LOG, PYTHONPATH
        End If
    #End If
End Function

Sub ExcecuteMac2011(PythonCommand As String, PYTHON_MAC As String, LOG_FILE As String, SHOW_LOG As Boolean, Optional PYTHONPATH As String)
    ' Run Python with the "-c" command line switch: add the path of the python file and run the
    ' Command as first argument, then provide the WORKBOOK_FULLNAME and "from_xl" as 2nd and 3rd arguments.
    ' Finally, redirect stderr to the LOG_FILE and run as background process.

    Dim PythonInterpreter As String, RunCommand As String, WORKBOOK_FULLNAME As String, Log As String
    Dim Res As Integer

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
    WORKBOOK_FULLNAME = ToPosixPath(ThisWorkbook.Path & ":" & ThisWorkbook.Name) 'ThisWorkbook.FullName doesn't handle unicode on Excel 2011

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
        ElseIf SHOW_LOG = True Then
            ShowError (LOG_FILE)
            Application.StatusBar = False
        End If
    On Error GoTo 0
End Sub

Sub ExecuteMac(PythonCommand As String, PYTHON_MAC As String, LOG_FILE As String, SHOW_LOG As Boolean, Optional PYTHONPATH As String)

    Dim PythonInterpreter As String, RunCommand As String, WORKBOOK_FULLNAME As String, Log As String, ParameterString As String, ExitCode As String
    Dim Res As Integer

    ' Transform paths
    PYTHONPATH = ToPosixPath(PYTHONPATH)

    If PYTHON_MAC <> "" Then
        PythonInterpreter = ToPosixPath(PYTHON_MAC)
    Else
        PythonInterpreter = "python"
    End If

    WORKBOOK_FULLNAME = ToPosixPath(ThisWorkbook.FullName)
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
    ParameterString = ParameterString + "," + ThisWorkbook.FullName
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
        ElseIf SHOW_LOG = True Then
            ShowError (LOG_FILE)
            Application.StatusBar = False
        End If
        Exit Sub
    On Error GoTo 0

AppleScriptErrorHandler:
    MsgBox "To enable RunPython, please run 'xlwings runpython install' in a terminal once and try again.", vbCritical

End Sub

Sub ExecuteWindows(IsFrozen As Boolean, PythonCommand As String, PYTHON_WIN As String, LOG_FILE As String, SHOW_LOG As Boolean, Optional PYTHONPATH As String)
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

    If Not IsFrozen And PYTHON_WIN <> "" Then
        PythonDir = ParentFolder(PYTHON_WIN)
    Else
        PythonDir = PYTHON_WIN
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
    WORKBOOK_FULLNAME = ThisWorkbook.FullName

    If PYTHON_WIN <> "" Then
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
    If ExitCode <> 0 And SHOW_LOG = True Then
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

    Dim PYTHON_WIN As String, PYTHON_MAC As String, PYTHON_FROZEN As String, PYTHONPATH As String, LOG_FILE As String, UDF_MODULES As String
    Dim SHOW_LOG As Boolean, OPTIMIZED_CONNECTION As Boolean, UDF_DEBUG_SERVER As Boolean
    Dim Res As Integer

    ' Get the settings by using the ByRef trick
    Res = Settings(PYTHON_WIN, PYTHON_MAC, PYTHON_FROZEN, PYTHONPATH, UDF_MODULES, UDF_DEBUG_SERVER, LOG_FILE, SHOW_LOG, OPTIMIZED_CONNECTION)

    ' Call Python
    #If Mac Then
        MsgBox "This functionality is not yet supported on Mac." & vbNewLine & _
               "Please run your scripts directly in Python!", vbCritical + vbOKOnly, "Unsupported Feature"
    #Else
        ExecuteWindows True, Executable, PYTHON_FROZEN, LOG_FILE, SHOW_LOG
    #End If
End Function

Function GetUdfModules() As String
    Dim PYTHON_WIN As String, PYTHON_MAC As String, PYTHON_FROZEN As String, PYTHONPATH As String
    Dim LOG_FILE As String, UDF_MODULES As String
    Dim Res As Integer
    Dim SHOW_LOG As Boolean, OPTIMIZED_CONNECTION As Boolean, UDF_DEBUG_SERVER As Boolean

    ' Get the settings
    Res = Settings(PYTHON_WIN, PYTHON_MAC, PYTHON_FROZEN, PYTHONPATH, UDF_MODULES, UDF_DEBUG_SERVER, LOG_FILE, SHOW_LOG, OPTIMIZED_CONNECTION)

    If UDF_MODULES = "" Then
        GetUdfModules = Left$(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5) ' assume that it ends in .xlsm
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
    'This function accepts relative paths with backward and forward slashes: ThisWorkbook & "\test"
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

Function GetMacDir(Name As String) As String
    ' Get Mac special folders. Protetcted so they don't exectue on Windows.

    Dim Path As String

    #If Mac Then
        Select Case Name
            Case "Home"
                Path = MacScript("return (path to home folder) as string")
             Case "Desktop"
                Path = MacScript("return (path to desktop folder) as string")
            Case "Applications"
                Path = MacScript("return (path to applications folder) as string")
            Case "Documents"
                Path = MacScript("return (path to documents folder) as string")
        End Select
            GetMacDir = Left$(Path, Len(Path) - 1) ' get rid of trailing ":"
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

    Dim PYTHON_WIN As String, PYTHON_MAC As String, PYTHON_FROZEN As String, PYTHONPATH As String, UDF_MODULES As String
    Dim WORKBOOK_FULLNAME As String, LOG_FILE As String
    Dim Res As Integer
    Dim SHOW_LOG As Boolean, OPTIMIZED_CONNECTION As Boolean, UDF_DEBUG_SERVER As Boolean

    'Get LOG_FILE
    Res = Settings(PYTHON_WIN, PYTHON_MAC, PYTHON_FROZEN, PYTHONPATH, UDF_MODULES, UDF_DEBUG_SERVER, LOG_FILE, SHOW_LOG, OPTIMIZED_CONNECTION)

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
    If SHOW_LOG = True Then
        On Error Resume Next
        If ReadFile(LOG_FILE) <> "" Then
            Call ShowError(LOG_FILE)
        End If
        On Error GoTo 0
    End If

    'Clean up
    Application.StatusBar = False
    Application.ScreenUpdating = True
    On Error Resume Next
        KillFileOnMac ToMacPath(ToPosixPath(LOG_FILE))
    On Error GoTo 0
End Sub

Function ParentFolder(ByVal Folder)
  ParentFolder = Left$(Folder, InStrRev(Folder, "\") - 1)
End Function

Function XLPyCommand()
    Dim PYTHON_WIN As String, PYTHON_MAC As String, PYTHON_FROZEN As String, PYTHONPATH As String
    Dim LOG_FILE As String, UDF_MODULES As String, Tail As String
    Dim Res As Integer
    Dim SHOW_LOG As Boolean, OPTIMIZED_CONNECTION As Boolean, UDF_DEBUG_SERVER As Boolean

    Res = Settings(PYTHON_WIN, PYTHON_MAC, PYTHON_FROZEN, PYTHONPATH, UDF_MODULES, UDF_DEBUG_SERVER, LOG_FILE, SHOW_LOG, OPTIMIZED_CONNECTION)

    If UDF_DEBUG_SERVER = True Then
        XLPyCommand = "{506e67c3-55b5-48c3-a035-eed5deea7d6d}"
    Else
        Tail = " -B -c ""import sys, os;sys.path.extend(os.path.normcase(os.path.expandvars(r'" & PYTHONPATH & "')).split(';'));import xlwings.server; xlwings.server.serve('$(CLSID)')"""
        If PYTHON_WIN = "" Then
            XLPyCommand = "pythonw.exe" + Tail
        Else
            XLPyCommand = PYTHON_WIN + Tail
        End If
    End If
End Function

Private Sub XLPyLoadDLL()
    Dim PYTHON_WIN As String, PYTHON_MAC As String, PYTHON_FROZEN As String, PYTHONPATH As String
    Dim LOG_FILE As String, UDF_MODULES As String, Tail As String
    Dim Res As Integer
    Dim SHOW_LOG As Boolean, OPTIMIZED_CONNECTION As Boolean, UDF_DEBUG_SERVER As Boolean

    Res = Settings(PYTHON_WIN, PYTHON_MAC, PYTHON_FROZEN, PYTHONPATH, UDF_MODULES, UDF_DEBUG_SERVER, LOG_FILE, SHOW_LOG, OPTIMIZED_CONNECTION)

    If PYTHON_WIN <> "" Then
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
    If 0 <> XLPyDLLActivateAuto(Py, XLPyCommand) Then Err.Raise 1000, Description:=Py
End Function

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
    tempPath = Py.Str(Py.Call(Py.Module("xlwings"), "import_udfs", Py.Tuple(GetUdfModules, ThisWorkbook)))
End Sub
