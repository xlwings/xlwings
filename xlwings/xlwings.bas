Attribute VB_Name = "xlwings"
' Make Excel fly with Python!
'
' Homepage and documentation: http://xlwings.org
' See also: http://zoomeranalytics.com
'
' Copyright (C) 2014-2015, Zoomer Analytics LLC.
' Version: 0.6.0dev
'
' License: BSD 3-clause (see LICENSE.txt for details)
#If Mac Then
    Private Declare Function system Lib "libc.dylib" (ByVal Command As String) As Long
#End If
#If VBA7 Then
    Private Declare PtrSafe Function GetTempPath32 Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As LongPtr, ByVal lpBuffer As String) As Long
    Private Declare PtrSafe Function GetTempFileName32 Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

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
    Private Declare Function GetTempPath32 Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
    Private Declare Function GetTempFileName32 Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
    Private Const XLPyDLLName As String = "xlwings32.dll"
    Private Declare Function XLPyDLLActivateAuto Lib "xlwings32.dll" (ByRef result As Variant, Optional ByVal config As String = "") As Long
    Private Declare Function XLPyDLLNDims Lib "xlwings32.dll" (ByRef src As Variant, ByRef dims As Long, ByRef transpose As Boolean, ByRef dest As Variant) As Long
    Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
    Declare Function XLPyDLLVersion Lib "xlwings32.dll" (tag As String, version As Double, arch As String) As Long
#End If

Function Settings(ByRef PYTHON_WIN As String, ByRef PYTHON_MAC As String, ByRef PYTHON_FROZEN As String, ByRef PYTHONPATH As String, ByRef UDF_PATH As String, ByRef LOG_FILE As String, ByRef SHOW_LOG As Boolean, ByRef OPTIMIZED_CONNECTION As Boolean)
    ' PYTHON_WIN: Directory of Python Interpreter on Windows, "" resolves to default on PATH
    ' PYTHON_MAC: Directory of Python Interpreter on Mac OSX, "" resolves to default path in ~/.bash_profile
    ' PYTHON_FROZEN [Optional]: Currently only on Windows, indicate directory of exe file
    ' PYTHONPATH [Optional]: If the source file of your code is not found, add the path here.
    '                        Separate multiple directories by ";". Otherwise set to "".
    ' UDF_PATH [Optional, Windows only]: Full path to a Python file from wich the User Defined Functions are being imported.
    '                                    Example: UDF_PATH = ThisWorkbook.Path & "\functions.py"
    '                                    Default: UDF_PATH = "" defaults to a file in the same directory of the Excel spreadsheet with
    '                                    the same name but ending in ".py".
    ' LOG_FILE: Directory including file name, necessary for error handling.
    ' SHOW_LOG: If False, no pop-up with the Log messages (usually errors) will be shown
    ' OPTIMIZED_CONNECTION (EXPERIMENTAL!): Currently only on Windows, use a COM Server for an efficient connection
    '
    ' For cross-platform compatibility, use backslashes in relative directories
    ' For details, see http://docs.xlwings.org

    PYTHON_WIN = ""
    PYTHON_MAC = ""
    PYTHON_FROZEN = ThisWorkbook.Path & "\build\exe.win32-2.7"
    PYTHONPATH = ThisWorkbook.Path
    UDF_PATH = ThisWorkbook.Path & "\functions.py"
    'UDF_PATH = ""
    LOG_FILE = ThisWorkbook.Path & "\xlwings_log.txt"
    SHOW_LOG = True
    OPTIMIZED_CONNECTION = False

End Function
' DO NOT EDIT BELOW THIS LINE

Function PyScriptPath() As String
    Dim PYTHON_WIN As String, PYTHON_MAC As String, PYTHON_FROZEN As String, PYTHONPATH As String
    Dim LOG_FILE As String, UDF_PATH As String
    Dim Res As Integer
    Dim SHOW_LOG As Boolean, OPTIMIZED_CONNECTION As Boolean

    ' Get the settings
    Res = Settings(PYTHON_WIN, PYTHON_MAC, PYTHON_FROZEN, PYTHONPATH, UDF_PATH, LOG_FILE, SHOW_LOG, OPTIMIZED_CONNECTION)

    If UDF_PATH = "" Then
        PyScriptPath = Left$(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5) ' assume that it ends in .xlsm
        PyScriptPath = ThisWorkbook.Path + Application.PathSeparator + PyScriptPath + ".py"
    Else
        PyScriptPath = UDF_PATH
    End If
End Function

Public Function RunPython(PythonCommand As String)
    ' Public API: Runs the Python command, e.g.: to run the function foo() in module bar, call the function like this:
    ' RunPython ("import bar; bar.foo()")

    Dim PYTHON_WIN As String, PYTHON_MAC As String, PYTHON_FROZEN As String, PYTHONPATH As String
    Dim WORKBOOK_FULLNAME As String, LOG_FILE As String, DriveCommand As String, RunCommand As String
    Dim ExitCode As Integer, Res As Integer
    Dim SHOW_LOG As Boolean, OPTIMIZED_CONNECTION As Boolean

    ' Get the settings by using the ByRef trick
    Res = Settings(PYTHON_WIN, PYTHON_MAC, PYTHON_FROZEN, PYTHONPATH, UDF_PATH, LOG_FILE, SHOW_LOG, OPTIMIZED_CONNECTION)

    ' Call Python platform-dependent
    #If Mac Then
        #If MAC_OFFICE_VERSION >= 15 Then
            MsgBox "This functionality is not yet supported on Excel 2016 for Mac." & vbNewLine & _
               "Please run your scripts directly in Python or call them from Excel 2011!", vbCritical + vbOKOnly, "Unsupported Feature"
        #Else
            Application.StatusBar = "Running..."  ' Non-blocking way of giving feedback that something is happening
            ExcecuteMac PythonCommand, PYTHON_MAC, LOG_FILE, SHOW_LOG, PYTHONPATH
        #End If
    #Else
        ' Make sure that the calling Workbook is the active Workbook
        ' This is necessary because under certain circumstances, only the GetActiveObject
        ' call will work (e.g. when Excel opens with a Security Warning, the Workbook
        ' will not be registered in the RunningObjectTable and thus not accessible via GetObject)
        ThisWorkbook.Activate

        If OPTIMIZED_CONNECTION = True Then
            Py.SetAttr Py.Module("xlwings._xlwindows"), "xl_workbook_current", ThisWorkbook
            Py.Exec "" & PythonCommand & ""
        Else
            ExecuteWindows False, PythonCommand, PYTHON_WIN, LOG_FILE, SHOW_LOG, PYTHONPATH
        End If
    #End If
End Function

Sub ExcecuteMac(Command As String, PYTHON_MAC As String, LOG_FILE As String, SHOW_LOG As Boolean, Optional PYTHONPATH As String)
    ' Run Python with the "-c" command line switch: add the path of the python file and run the
    ' Command as first argument, then provide the WORKBOOK_FULLNAME and "from_xl" as 2nd and 3rd arguments.
    ' Finally, redirect stderr to the LOG_FILE and run as background process.

    Dim PythonInterpreter As String, RunCommand As String, WORKBOOK_FULLNAME As String, Log As String
    Dim Res As Integer

    ' Delete Log file just to make sure we don't show an old error
    On Error Resume Next
        KillFileOnMac ToMacPath(ToPosixPath(LOG_FILE))
    On Error GoTo 0

    ' Transform from MacOS Classic path style (":") and Windows style ("\") to Bash friendly style ("/")
    PYTHONPATH = ToPosixPath(PYTHONPATH)
    LOG_FILE = ToPosixPath(LOG_FILE)
    PythonInterpreter = ToPosixPath(PYTHON_MAC & "/python")
    WORKBOOK_FULLNAME = ToPosixPath(ThisWorkbook.Path & ":" & ThisWorkbook.Name) 'ThisWorkbook.FullName doesn't handle unicode on Excel 2011

    ' Build the command (ignore warnings to be in line with Windows where we only show the popup if the ExitCode <> 0
    ' -u is needed because on PY3 stderr is buffered by default and so wouldn't be available on time for the pop-up to show
    RunCommand = PythonInterpreter & " -u -W ignore -c ""import sys; sys.path.extend(r'" & PYTHONPATH & "'.split(';')); " & Command & """ "

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

Sub ExecuteWindows(IsFrozen As Boolean, Command As String, PYTHON_WIN As String, LOG_FILE As String, SHOW_LOG As Boolean, Optional PYTHONPATH As String)
    ' Call a command window and change to the directory of the Python installation or frozen executable
    ' Note: If Python is called from a different directory with the fully qualified path, pywintypesXX.dll won't be found.
    ' This seems to be a general issue with pywin32, see http://stackoverflow.com/q/7238403/918626

    Dim Wsh As Object
    Dim WaitOnReturn As Boolean: WaitOnReturn = True
    Dim WindowStyle As Integer: WindowStyle = 0
    Set Wsh = CreateObject("WScript.Shell")
    Dim DriveCommand As String, RunCommand As String, WORKBOOK_FULLNAME As String
    Dim ExitCode As Integer

    If Left$(PYTHON_WIN, 2) Like "[A-Za-z]:" Then
        ' If Python is installed on a mapped or local drive, change to drive, then cd to path
        DriveCommand = Left$(PYTHON_WIN, 2) & " & cd """ & PYTHON_WIN & """ & "
    ElseIf Left$(PYTHON_WIN, 2) = "\\" Then
        ' If Python is installed on a UNC path, temporarily mount and activate a drive letter with pushd
        DriveCommand = "pushd " & PYTHON_WIN & " & "
    End If

    ' Run Python with the "-c" command line switch: add the path of the python file and run the
    ' Command as first argument, then provide the WORKBOOK_FULLNAME and "from_xl" as 2nd and 3rd arguments.
    ' Then redirect stderr to the LOG_FILE and wait for the call to return.
    WORKBOOK_FULLNAME = ThisWorkbook.FullName

    If IsFrozen = False Then
        RunCommand = "python -c ""import sys; sys.path.extend(r'" & PYTHONPATH & "'.split(';')); " & Command & """ "
    ElseIf IsFrozen = True Then
        RunCommand = Command & " "
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

    Dim PYTHON_WIN As String, PYTHON_MAC As String, PYTHON_FROZEN As String, PYTHONPATH As String, LOG_FILE As String
    Dim SHOW_LOG As Boolean, OPTIMIZED_CONNECTION As Boolean
    Dim Res As Integer

    ' Get the settings by using the ByRef trick
    Res = Settings(PYTHON_WIN, PYTHON_MAC, PYTHON_FROZEN, PYTHONPATH, UDF_PATH, LOG_FILE, SHOW_LOG, OPTIMIZED_CONNECTION)

    ' Call Python
    #If Mac Then
        MsgBox "This functionality is not yet supported on Mac." & vbNewLine & _
               "Please run your scripts directly in Python!", vbCritical + vbOKOnly, "Unsupported Feature"
    #Else
        ExecuteWindows True, Executable, PYTHON_FROZEN, LOG_FILE, SHOW_LOG
    #End If
End Function

Function ReadFile(ByVal FileName As String)
    ' Read a text file

    Dim Content As String
    Dim Token As String
    Dim FileNum As Integer

    #If Mac Then
        FileName = ToMacPath(FileName)
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

    MacPath = Replace(MacPath, "\", ":")
    MacPath = Replace(MacPath, "/", ":")
    s = "tell application " & Chr(34) & "Finder" & Chr(34) & Chr(13)
    s = s & "POSIX path of " & Chr(34) & MacPath & Chr(34) & Chr(13)
    s = s & "end tell" & Chr(13)
    ToPosixPath = MacScript(s)
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
    'Uses AppleScript to avoid the problem with long file names

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

    Dim PYTHON_WIN As String, PYTHON_MAC As String, PYTHON_FROZEN As String, PYTHONPATH As String
    Dim WORKBOOK_FULLNAME As String, LOG_FILE As String
    Dim Res As Integer
    Dim SHOW_LOG As Boolean, OPTIMIZED_CONNECTION As Boolean

    'Get LOG_FILE
    Res = Settings(PYTHON_WIN, PYTHON_MAC, PYTHON_FROZEN, PYTHONPATH, UDF_PATH, LOG_FILE, SHOW_LOG, OPTIMIZED_CONNECTION)
    LOG_FILE = ToPosixPath(LOG_FILE)

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

'ExcelPython
Private Function GetTempFileName()
    Dim sTmpPath As String * 512
    Dim sTmpName As String * 576
    Dim nRet As Long
    nRet = GetTempPath32(512, sTmpPath)
    If nRet = 0 Then Err.Raise 1234, Description:="GetTempPath failed."
    nRet = GetTempFileName32(sTmpPath, "vba", 0, sTmpName)
    If nRet = 0 Then Err.Raise 1234, Description:="GetTempFileName failed."
    GetTempFileName = Left$(sTmpName, InStr(sTmpName, vbNullChar) - 1)
End Function

Function ModuleIsPresent(ByVal wb As Workbook, moduleName As String) As Boolean
    On Error GoTo not_present
    Set x = wb.VBProject.VBComponents.Item(moduleName)
    ModuleIsPresent = True
    Exit Function
not_present:
    ModuleIsPresent = False
End Function

Sub XLPMacroOptions2010(macroName As String, desc, argdescs() As String)
    Application.MacroOptions macroName, Description:=desc, ArgumentDescriptions:=argdescs
End Sub

Function XLPyCommand()
    Dim PYTHON_WIN As String, PYTHON_MAC As String, PYTHON_FROZEN As String, PYTHONPATH As String
    Dim LOG_FILE As String, UDF_PATH As String, Tail As String
    Dim Res As Integer
    Dim SHOW_LOG As Boolean, OPTIMIZED_CONNECTION As Boolean

    Res = Settings(PYTHON_WIN, PYTHON_MAC, PYTHON_FROZEN, PYTHONPATH, UDF_PATH, LOG_FILE, SHOW_LOG, OPTIMIZED_CONNECTION)
    Tail = " -c ""import sys;sys.path.extend(r'" & PYTHONPATH & "'.split(';'));import xlwings.server; xlwings.server.serve('$(CLSID)')"""
    If PYTHON_WIN = "" Then
        XLPyCommand = "pythonw.exe" + Tail
    Else
        XLPyCommand = PYTHON_WIN + "\pythonw.exe" + Tail
    End If
End Function

Private Sub XLPyLoadDLL()
    Dim PYTHON_WIN As String, PYTHON_MAC As String, PYTHON_FROZEN As String, PYTHONPATH As String
    Dim LOG_FILE As String, UDF_PATH As String, Tail As String
    Dim Res As Integer
    Dim SHOW_LOG As Boolean, OPTIMIZED_CONNECTION As Boolean

    Res = Settings(PYTHON_WIN, PYTHON_MAC, PYTHON_FROZEN, PYTHONPATH, UDF_PATH, LOG_FILE, SHOW_LOG, OPTIMIZED_CONNECTION)

    If PYTHON_WIN <> "" Then
        On Error Resume Next
            LoadLibrary PYTHON_WIN + "\" + XLPyDLLName 'Standard installation
            LoadLibrary ParentFolder(PYTHON_WIN) + "\" + XLPyDLLName 'Virtualenv
        On Error GoTo 0
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

'Sub ImportPythonUDFs(control As IRibbonControl)
Sub ImportPythonUDFs()
    sTab = "    "

    Set wb = ActiveWorkbook
    If Not ModuleIsPresent(wb, "xlwings") Then
        MsgBox "This workbook must contain the xlwings VBA module."
        Exit Sub
    End If
    
    ' Needed when run as add-in
    'Set Py = Application.Run("'" + wb.Name + "'!Py")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileName = GetTempFileName()
    Set f = fso.CreateTextFile(FileName, True)
    f.WriteLine "Attribute VB_Name = ""xlwings_udfs"""
    
    scriptPath = PyScriptPath()
    Set scriptVars = Py.Call(Py.Module("xlwings"), "udf_script", Py.Tuple(scriptPath))
    
    For Each svar In Py.Call(scriptVars, "values")
        If Py.HasAttr(svar, "__xlfunc__") Then
            Set xlfunc = Py.GetAttr(svar, "__xlfunc__")
            Set xlret = Py.GetItem(xlfunc, "ret")
            fname = Py.Str(Py.GetItem(xlfunc, "name"))
            
            Dim ftype As String
            If Py.Var(Py.GetItem(xlfunc, "sub")) Then ftype = "Sub" Else ftype = "Function"
            
            f.Write ftype + " " + fname + "("
            first = True
            vararg = ""
            nArgs = Py.Len(Py.GetItem(xlfunc, "args"))
            For Each arg In Py.GetItem(xlfunc, "args")
                If Not Py.Bool(Py.GetItem(arg, "vba")) Then
                    argname = Py.Str(Py.GetItem(arg, "name"))
                    If Not first Then f.Write ", "
                    If Py.Bool(Py.GetItem(arg, "vararg")) Then
                        f.Write "ParamArray "
                        vararg = argname
                    End If
                    f.Write argname
                    If Py.Bool(Py.GetItem(arg, "vararg")) Then
                        f.Write "()"
                    End If
                    first = False
                End If
            Next arg
            f.WriteLine ")"
            If ftype = "Function" Then
                f.WriteLine sTab + "If TypeOf Application.Caller Is Range Then On Error GoTo failed"
            End If
            
            If vararg <> "" Then
                f.WriteLine sTab + "ReDim argsArray(1 to UBound(" + vararg + ") - LBound(" + vararg + ") + " + CStr(nArgs) + ")"
            End If
            j = 1
            For Each arg In Py.GetItem(xlfunc, "args")
                If Not Py.Bool(Py.GetItem(arg, "vba")) Then
                    argname = Py.Str(Py.GetItem(arg, "name"))
                    If Py.Bool(Py.GetItem(arg, "vararg")) Then
                        f.WriteLine sTab + "For k = lbound(" + vararg + ") to ubound(" + vararg + ")"
                        argname = vararg + "(k)"
                    End If
                    If Not Py.Var(Py.GetItem(arg, "range")) Then
                        f.WriteLine sTab + "If TypeOf " + argname + " Is Range Then " + argname + " = " + argname + ".Value2"
                    End If
                    dims = Py.Var(Py.GetItem(arg, "dims"))
                    marshal = Py.Str(Py.GetItem(arg, "marshal"))
                    If dims <> -2 Or marshal = "nparray" Or marshal = "list" Then
                        f.WriteLine sTab + "If Not TypeOf " + argname + " Is Object Then"
                        If dims <> -2 Then
                            f.WriteLine sTab + sTab + argname + " = NDims(" + argname + ", " + CStr(dims) + ")"
                        End If
                        If marshal = "nparray" Then
                            dtype = Py.Var(Py.GetItem(arg, "dtype"))
                            If IsNull(dtype) Then
                                f.WriteLine sTab + sTab + "Set " + argname + " = Py.Call(Py.Module(""numpy""), ""array"", Py.Tuple(" + argname + "))"
                            Else
                                f.WriteLine sTab + sTab + "Set " + argname + " = Py.Call(Py.Module(""numpy""), ""array"", Py.Tuple(" + argname + ", """ + dtype + """))"
                            End If
                        ElseIf marshal = "list" Then
                            f.WriteLine sTab + sTab + "Set " + argname + " = Py.Call(Py.Eval(""lambda t: [ list(x) if isinstance(x, tuple) else x for x in t ] if isinstance(t, tuple) else t""), Py.Tuple(" + argname + "))"
                        End If
                        f.WriteLine sTab + "End If"
                    End If
                    If Py.Bool(Py.GetItem(arg, "vararg")) Then
                        f.WriteLine sTab + "argsArray(" + CStr(j) + " + k - LBound(" + vararg + ")) = " + argname
                        f.WriteLine sTab + "Next k"
                    Else
                        If vararg <> "" Then
                            f.WriteLine sTab + "argsArray(" + CStr(j) + ") = " + argname
                            j = j + 1
                        End If
                    End If
                End If
            Next arg
            
            If vararg <> "" Then
                f.WriteLine sTab + "Set args = Py.TupleFromArray(argsArray)"
            Else
                f.Write sTab + "Set args = Py.Tuple("
                first = True
                For Each arg In Py.GetItem(xlfunc, "args")
                    If Not first Then f.Write ", "
                    If Not Py.Bool(Py.GetItem(arg, "vba")) Then
                        f.Write Py.Str(Py.GetItem(arg, "name"))
                    Else
                        f.Write Py.Str(Py.GetItem(arg, "vba"))
                    End If
                    first = False
                Next arg
                f.WriteLine ")"
            End If
            
            f.WriteLine sTab + "Set xlpy = Py.Module(""xlwings"")"
            f.WriteLine sTab + "Set script = Py.Call(xlpy, ""udf_script"", Py.Tuple(PyScriptPath))"
            f.WriteLine sTab + "Set func = Py.GetItem(script, """ + fname + """)"
            If ftype = "Sub" Then
                f.WriteLine sTab + "Py.SetAttr Py.Module(""xlwings._xlwindows""), ""xl_workbook_current"", ThisWorkbook"
                f.WriteLine sTab + "Py.Call func, args"
            Else
                f.WriteLine sTab + "Set " + fname + " = Py.Call(func, args)"
                marshal = Py.Str(Py.GetItem(xlret, "marshal"))
                Select Case marshal
                Case "auto"
                    f.WriteLine sTab + "If TypeOf Application.Caller Is Range Then " + fname + " = Py.Var(" + fname + ", " + Py.Str(Py.GetItem(xlret, "lax")) + ")"
                Case "var"
                    f.WriteLine sTab + fname + " = Py.Var(" + fname + ", " + Py.Str(Py.GetItem(xlret, "lax")) + ")"
                Case "str"
                    f.WriteLine sTab + fname + " = Py.Str(" + fname + ")"
                End Select
            End If
            
            If ftype = "Function" Then
                f.WriteLine sTab + "Exit " + ftype
                f.WriteLine "failed:"
                f.WriteLine sTab + fname + " = Err.Description"
            End If
            f.WriteLine "End " + ftype
            f.WriteLine
        End If
    Next svar
    f.Close
    
    On Error GoTo not_present
    wb.VBProject.VBComponents.Remove wb.VBProject.VBComponents("xlwings_udfs")
not_present:
    On Error GoTo 0
    wb.VBProject.VBComponents.Import FileName
    
    For Each svar In Py.Call(scriptVars, "values")
        If Py.HasAttr(svar, "__xlfunc__") Then
            Set xlfunc = Py.GetAttr(svar, "__xlfunc__")
            Set xlret = Py.GetItem(xlfunc, "ret")
            Set xlargs = Py.GetItem(xlfunc, "args")
            fname = Py.Str(Py.GetItem(xlfunc, "name"))
            fdoc = Py.Str(Py.GetItem(xlret, "doc"))
            nArgs = 0
            For Each arg In xlargs
                If Not Py.Bool(Py.GetItem(arg, "vba")) Then nArgs = nArgs + 1
            Next arg
            If nArgs > 0 And Val(Application.version) >= 14 Then
                ReDim argdocs(1 To WorksheetFunction.Max(1, nArgs)) As String
                nArgs = 0
                For Each arg In xlargs
                    If Not Py.Bool(Py.GetItem(arg, "vba")) Then
                        nArgs = nArgs + 1
                        argdocs(nArgs) = Left$(Py.Str(Py.GetItem(arg, "doc")), 255)
                    End If
                Next arg
                XLPMacroOptions2010 "'" + wb.Name + "'!" + fname, Left$(fdoc, 255), argdocs
            Else
                Application.MacroOptions "'" + wb.Name + "'!" + fname, Description:=Left$(fdoc, 255)
            End If
        End If
    Next svar
    
End Sub
