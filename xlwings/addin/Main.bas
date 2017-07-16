Attribute VB_Name = "Main"
Option Explicit

#If VBA7 Then
    #If Mac Then
        Private Declare PtrSafe Function system Lib "libc.dylib" (ByVal Command As String) As Long
    #End If
    #If Win64 Then
        Const XLPyDLLName As String = "xlwings64.dll"
        Declare PtrSafe Function XLPyDLLActivateAuto Lib "xlwings64.dll" (ByRef result As Variant, Optional ByVal Config As String = "", Optional ByVal mode As Long = 1) As Long
        Declare PtrSafe Function XLPyDLLNDims Lib "xlwings64.dll" (ByRef src As Variant, ByRef dims As Long, ByRef transpose As Boolean, ByRef dest As Variant) As Long
        Declare PtrSafe Function XLPyDLLVersion Lib "xlwings64.dll" (tag As String, version As Double, arch As String) As Long
    #Else
        Private Const XLPyDLLName As String = "xlwings32.dll"
        Declare PtrSafe Function XLPyDLLActivateAuto Lib "xlwings32.dll" (ByRef result As Variant, Optional ByVal Config As String = "", Optional ByVal mode As Long = 1) As Long
        Private Declare PtrSafe Function XLPyDLLNDims Lib "xlwings32.dll" (ByRef src As Variant, ByRef dims As Long, ByRef transpose As Boolean, ByRef dest As Variant) As Long
        Private Declare PtrSafe Function XLPyDLLVersion Lib "xlwings32.dll" (tag As String, version As Double, arch As String) As Long
    #End If
    Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
#Else
    #If Mac Then
        Private Declare Function system Lib "libc.dylib" (ByVal Command As String) As Long
    #End If
    Private Const XLPyDLLName As String = "xlwings32.dll"
    Private Declare Function XLPyDLLActivateAuto Lib "xlwings32.dll" (ByRef result As Variant, Optional ByVal Config As String = "", Optional ByVal mode As Long = 1) As Long
    Private Declare Function XLPyDLLNDims Lib "xlwings32.dll" (ByRef src As Variant, ByRef dims As Long, ByRef transpose As Boolean, ByRef dest As Variant) As Long
    Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
    Declare Function XLPyDLLVersion Lib "xlwings32.dll" (tag As String, version As Double, arch As String) As Long
#End If

Public Function RunPython(PythonCommand As String)
    ' Public API: Runs the Python command, e.g.: to run the function foo() in module bar, call the function like this:
    ' RunPython ("import bar; bar.foo()")

    Dim INTERPRETER As String, PYTHONPATH As String, LOG_FILE As String
    Dim OPTIMIZED_CONNECTION As Boolean

    INTERPRETER = GetConfig("INTERPRETER", "python")
    PYTHONPATH = ActiveWorkbook.Path & ";" & GetConfig("PYTHONPATH")
    LOG_FILE = GetConfig("LOG FILE")
    OPTIMIZED_CONNECTION = GetConfig("USE UDF SERVER", False)

    ' Call Python platform-dependent
    #If Mac Then
            Application.StatusBar = "Running..."  ' Non-blocking way of giving feedback that something is happening
        #If MAC_OFFICE_VERSION >= 15 Then
            ExecuteMac PythonCommand, INTERPRETER, LOG_FILE, PYTHONPATH
        #Else
            ExcecuteMac2011 PythonCommand, INTERPRETER, LOG_FILE, PYTHONPATH
        #End If
    #Else
        If OPTIMIZED_CONNECTION = True Then
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
    Dim Res As Integer

    If LOG_FILE = "" Then
        LOG_FILE = "/tmp/xlwings.log"
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
        If PYTHON_MAC <> "python" And PYTHON_MAC <> "pythonw" Then
            PythonInterpreter = ToPosixPath(PYTHON_MAC)
        Else
            PythonInterpreter = PYTHON_MAC
        End If
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
        Res = system("source ~/.bash_profile;" & RunCommand & """" & WORKBOOK_FULLNAME & """ ""from_xl""" & " " & Chr(34) & ToPosixPath(Application.Path) & "/" & Application.Name & Chr(34) & "> /dev/null 2>" & Chr(34) & LOG_FILE & Chr(34) & " &")
    Else
        Res = system(RunCommand & """" & WORKBOOK_FULLNAME & """ ""from_xl""" & " " & Chr(34) & ToPosixPath(Application.Path) & "/" & Application.Name & Chr(34) & "> /dev/null 2>" & Chr(34) & LOG_FILE & Chr(34) & " &")
    End If

    ' If there's a log at this point (normally that will be from the shell only, not Python) show it and reset the StatusBar
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

    ' Transform paths
    PYTHONPATH = ToPosixPath(PYTHONPATH)

    If PYTHON_MAC <> "" Then
        If PYTHON_MAC <> "python" And PYTHON_MAC <> "pythonw" Then
            PythonInterpreter = ToPosixPath(PYTHON_MAC)
        Else
            PythonInterpreter = PYTHON_MAC
        End If
    Else
        PythonInterpreter = "python"
    End If

    WORKBOOK_FULLNAME = ToPosixPath(ActiveWorkbook.FullName)
    If LOG_FILE = "" Then
        ' Sandbox location that requires no file access confirmation
        LOG_FILE = Environ("HOME") + "/xlwings.log" '/Users/<User>/Library/Containers/com.microsoft.Excel/Data/xlwings_log.txt
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

    ' If there's a log at this point (normally that will be from the shell only, not Python) show it and reset the StatusBar
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
        LOG_FILE = Environ("APPDATA") + "\xlwings.log"
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
        'Kill LOG_FILE
    On Error GoTo 0

    ' Clean up
    Set Wsh = Nothing
End Sub

Public Function RunFrozenPython(Executable As String)
    ' Runs a Python executable that has been frozen by cx_Freeze or py2exe. Call the function like this:
    ' RunFrozenPython("C:\path\to\frozen_executable.exe"). Currently not implemented for Mac.

    Dim LOG_FILE As String

    LOG_FILE = GetConfig("LOG FILE")

    ' Call Python
    #If Mac Then
        MsgBox "This functionality is not yet supported on Mac." & vbNewLine & _
               "Please run your scripts directly in Python!", vbCritical + vbOKOnly, "Unsupported Feature"
    #Else
        ExecuteWindows True, Executable, ParentFolder(Executable), LOG_FILE
    #End If
End Function

Function GetUdfModules() As String
    Dim UDF_MODULES As String

    UDF_MODULES = GetConfig("UDF MODULES")

    If UDF_MODULES = "" Then
        GetUdfModules = Left$(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) ' assume that it ends in .xlsm
    Else
        GetUdfModules = UDF_MODULES
    End If
End Function

Private Sub CleanUp()
    'On Mac only, this function is being called after Python is done (using Python's atexit handler)
    Dim LOG_FILE As String

    LOG_FILE = GetConfig("LOG FILE")

    If LOG_FILE = "" Then
        #If MAC_OFFICE_VERSION >= 15 Then
            LOG_FILE = Environ("HOME") + "/xlwings.log" '~/Library/Containers/com.microsoft.Excel/Data/xlwings_log.txt
        #Else
            LOG_FILE = "/tmp/xlwings.log"
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

Function XLPyCommand()
    Dim PYTHON_WIN As String, PYTHONPATH As String, LOG_FILE As String, tail As String
    Dim DEBUG_UDFS As Boolean

    PYTHONPATH = ActiveWorkbook.Path & ";" & GetConfig("PYTHONPATH")
    PYTHON_WIN = GetConfig("INTERPRETER", "pythonw")
    DEBUG_UDFS = GetConfig("DEBUG UDFS", False)

    If DEBUG_UDFS = True Then
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

Sub ImportPythonUDFs()
    Dim tempPath As String, errorMsg As String
    On Error GoTo ImportError
        tempPath = Py.Str(Py.Call(Py.Module("xlwings"), "import_udfs", Py.Tuple(GetUdfModules, ActiveWorkbook)))
    Exit Sub
ImportError:
    errorMsg = Err.Description & " " & Err.Number
    errorMsg = errorMsg & ". Make sure you are trusting access to the VBA project object module (Excel Options)."
    MsgBox errorMsg, vbCritical, "Error"
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


