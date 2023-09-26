Attribute VB_Name = "Main"
Option Explicit
#Const App = "Microsoft Excel" 'Adjust when using outside of Excel

#If VBA7 Then
    #If Mac Then
        Private Declare PtrSafe Function system Lib "libc.dylib" (ByVal Command As String) As Long
    #End If
    #If Win64 Then
        Const XLPyDLLName As String = "xlwings64-dev.dll"
        Declare PtrSafe Function XLPyDLLActivateAuto Lib "xlwings64-dev.dll" (ByRef Result As Variant, Optional ByVal Config As String = "", Optional ByVal mode As Long = 1) As Long
        Declare PtrSafe Function XLPyDLLNDims Lib "xlwings64-dev.dll" (ByRef src As Variant, ByRef dims As Long, ByRef transpose As Boolean, ByRef dest As Variant) As Long
        Declare PtrSafe Function XLPyDLLVersion Lib "xlwings64-dev.dll" (tag As String, VERSION As Double, arch As String) As Long
    #Else
        Private Const XLPyDLLName As String = "xlwings32-dev.dll"
        Declare PtrSafe Function XLPyDLLActivateAuto Lib "xlwings32-dev.dll" (ByRef Result As Variant, Optional ByVal Config As String = "", Optional ByVal mode As Long = 1) As Long
        Private Declare PtrSafe Function XLPyDLLNDims Lib "xlwings32-dev.dll" (ByRef src As Variant, ByRef dims As Long, ByRef transpose As Boolean, ByRef dest As Variant) As Long
        Private Declare PtrSafe Function XLPyDLLVersion Lib "xlwings32-dev.dll" (tag As String, VERSION As Double, arch As String) As Long
    #End If
    Private Declare PtrSafe Function LoadLibrary Lib "KERNEL32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
#Else
    #If Mac Then
        Private Declare Function system Lib "libc.dylib" (ByVal Command As String) As Long
    #End If
    Private Const XLPyDLLName As String = "xlwings32-dev.dll"
    Private Declare Function XLPyDLLActivateAuto Lib "xlwings32-dev.dll" (ByRef Result As Variant, Optional ByVal Config As String = "", Optional ByVal mode As Long = 1) As Long
    Private Declare Function XLPyDLLNDims Lib "xlwings32-dev.dll" (ByRef src As Variant, ByRef dims As Long, ByRef transpose As Boolean, ByRef dest As Variant) As Long
    Private Declare Function LoadLibrary Lib "KERNEL32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
    Declare Function XLPyDLLVersion Lib "xlwings32-dev.dll" (tag As String, VERSION As Double, arch As String) As Long
#End If

Public Const XLWINGS_VERSION As String = "dev"
Public Const PROJECT_NAME As String = "xlwings"

Public Function RunPython(PythonCommand As String)
    ' Public API: Runs the Python command, e.g.: to run the function foo() in module bar, call the function like this:
    ' RunPython "import bar; bar.foo()"
    
    Dim i As Integer
    Dim SourcePythonCommand As String, interpreter As String, PYTHONPATH As String, licenseKey, ActiveFullName As String, ThisFullName As String, AddExcelDir As String
    Dim OPTIMIZED_CONNECTION As Boolean, uses_embedded_code As Boolean
    Dim wb As Workbook
    Dim sht As Worksheet
    
    SourcePythonCommand = PythonCommand
    
    #If Mac Then
        interpreter = GetConfig("INTERPRETER_MAC", "")
    #Else
        interpreter = GetConfig("INTERPRETER_WIN", "")
    #End If
    If interpreter = "" Then
        ' Legacy
        interpreter = GetConfig("INTERPRETER", "python")
    End If

    ' Check for embedded Python code
    uses_embedded_code = False
    For i = 1 To 2
        If i = 1 Then
            Set wb = ActiveWorkbook
        Else
            Set wb = ThisWorkbook
        End If
        For Each sht In wb.Worksheets
            If Right$(sht.Name, 3) = ".py" Then
                uses_embedded_code = True
                Exit For
            End If
        Next
    Next i

    If uses_embedded_code = True Then
        AddExcelDir = "false"
    Else
        AddExcelDir = GetConfig("ADD_WORKBOOK_TO_PYTHONPATH", "true")
    End If

    ' The first 5 args are not technically part of the PYTHONPATH, but it's just easier to add it here (used by xlwings.utils.prepare_sys_path)
    #If Mac Then
        If InStr(ActiveWorkbook.FullName, "://") = 0 Then
            ActiveFullName = ToPosixPath(ActiveWorkbook.FullName)
            ThisFullName = ToPosixPath(ThisWorkbook.FullName)
        Else
            ActiveFullName = ActiveWorkbook.FullName
            ThisFullName = ThisWorkbook.FullName
        End If
    #Else
        ActiveFullName = ActiveWorkbook.FullName
        ThisFullName = ThisWorkbook.FullName
    #End If
    
    #If Mac Then
        PYTHONPATH = AddExcelDir & ";" & ActiveFullName & ";" & ThisFullName & ";" & GetConfig("ONEDRIVE_CONSUMER_MAC") & ";" & GetConfig("ONEDRIVE_COMMERCIAL_MAC") & ";" & GetConfig("SHAREPOINT_MAC") & ";" & GetConfig("PYTHONPATH")
    #Else
        PYTHONPATH = AddExcelDir & ";" & ActiveFullName & ";" & ThisFullName & ";" & GetConfig("ONEDRIVE_CONSUMER_WIN") & ";" & GetConfig("ONEDRIVE_COMMERCIAL_WIN") & ";" & GetConfig("SHAREPOINT_WIN") & ";" & GetConfig("PYTHONPATH")
    #End If

    OPTIMIZED_CONNECTION = GetConfig("USE UDF SERVER", False)

    ' PythonCommand with embedded code
    If uses_embedded_code = True Then
        licenseKey = GetConfig("LICENSE_KEY")
        If licenseKey = "" Then
            MsgBox "Embedded code requires a valid LICENSE_KEY."
            Exit Function
        Else
            PythonCommand = "import xlwings.pro;xlwings.pro.runpython_embedded_code('" & SourcePythonCommand & "')"
        End If
    End If

    ' Call Python platform-dependent
    #If Mac Then
        Application.StatusBar = "Running..."  ' Non-blocking way of giving feedback that something is happening
        ExecuteMac PythonCommand, interpreter, PYTHONPATH
    #Else
        If OPTIMIZED_CONNECTION = True Then
            XLPy.SetAttr XLPy.Module("xlwings._xlwindows"), "BOOK_CALLER", ActiveWorkbook
            
            On Error GoTo err_handling
            
            XLPy.Exec "" & PythonCommand & ""
            GoTo end_err_handling
err_handling:
            ShowError "", Err.Description
            RunPython = -1
            On Error GoTo 0
end_err_handling:
        Else
            RunPython = ExecuteWindows(False, PythonCommand, interpreter, PYTHONPATH)
        End If
    #End If
End Function


Sub ExecuteMac(PythonCommand As String, PYTHON_MAC As String, Optional PYTHONPATH As String)
    #If Mac Then
    Dim PythonInterpreter As String, RunCommand As String, Log As String
    Dim ParameterString As String, ExitCode As String, CondaCmd As String, CondaPath As String, CondaEnv As String, LOG_FILE As String

    ' Transform paths
    PYTHONPATH = Replace(PYTHONPATH, "'", "\'") ' Escaping quotes

    If PYTHON_MAC <> "" Then
        If PYTHON_MAC <> "python" And PYTHON_MAC <> "pythonw" Then
            PythonInterpreter = ToPosixPath(PYTHON_MAC)
        Else
            PythonInterpreter = PYTHON_MAC
        End If
    Else
        PythonInterpreter = "python"
    End If

    ' Sandbox location that requires no file access confirmation
    ' TODO: Use same logic with GUID like for Windows. Only here the GUID will need to be passed back to CleanUp()
    LOG_FILE = Environ("HOME") + "/xlwings.log" '/Users/<User>/Library/Containers/com.microsoft.Excel/Data/xlwings.log

    ' Delete Log file just to make sure we don't show an old error
    On Error Resume Next
        Kill LOG_FILE
    On Error GoTo 0

    ' ParameterSting with all paramters (AppleScriptTask only accepts a single parameter)
    ParameterString = PYTHONPATH + ";"
    ParameterString = ParameterString + "|" + PythonInterpreter
    ParameterString = ParameterString + "|" + PythonCommand
    ParameterString = ParameterString + "|" + ActiveWorkbook.Name
    ParameterString = ParameterString + "|" + Left(Application.Path, Len(Application.Path) - 4)
    ParameterString = ParameterString + "|" + LOG_FILE

    On Error GoTo AppleScriptErrorHandler
        ExitCode = AppleScriptTask("xlwings-" & XLWINGS_VERSION & ".applescript", "VbaHandler", ParameterString)
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

Function ExecuteWindows(IsFrozen As Boolean, PythonCommand As String, PYTHON_WIN As String, _
                        Optional PYTHONPATH As String, Optional FrozenArgs As String) As Integer
    ' Call a command window and change to the directory of the Python installation or frozen executable
    ' Note: If Python is called from a different directory with the fully qualified path, pywintypesXX.dll won't be found.
    ' This seems to be a general issue with pywin32, see http://stackoverflow.com/q/7238403/918626
    Dim ShowConsole As Integer
    Dim TempDir As String
    If GetConfig("SHOW CONSOLE", False) = True Then
        ShowConsole = 1
    Else
        ShowConsole = 0
    End If

    Dim WaitOnReturn As Boolean: WaitOnReturn = True
    Dim WindowStyle As Integer: WindowStyle = ShowConsole
    Dim DriveCommand As String, RunCommand, condaExcecutable As String
    Dim PythonInterpreter As String, PythonDir As String, CondaCmd As String, CondaPath As String, CondaEnv As String
    Dim ExitCode As Long
    Dim LOG_FILE As String
    
    TempDir = GetConfig("TEMP DIR", Environ("Temp")) 'undocumented setting
    
    LOG_FILE = TempDir & "\xlwings-" & CreateGUID() & ".log"

    If Not IsFrozen And (PYTHON_WIN <> "python" And PYTHON_WIN <> "pythonw") Then
        If FileExists(PYTHON_WIN) Then
            PythonDir = ParentFolder(PYTHON_WIN)
        Else
            MsgBox "Could not find Interpreter!", vbCritical
            Exit Function
        End If
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
    ' Command as first argument, then provide the Name and "from_xl" as 2nd and 3rd arguments.
    ' Then redirect stderr to the LOG_FILE and wait for the call to return.

    If PYTHON_WIN <> "python" And PYTHON_WIN <> "pythonw" Then
        PythonInterpreter = Chr(34) & PYTHON_WIN & Chr(34)
    Else
        PythonInterpreter = "python"
    End If

    CondaPath = GetConfig("CONDA PATH")
    CondaEnv = GetConfig("CONDA ENV")
    
    ' Handle spaces in path (for UDFs, this is handled via nested quotes instead, see XLPyCommand)
    CondaPath = Replace(CondaPath, " ", "^ ")
    
    ' Handle ampersands and backslashes in file paths
    PYTHONPATH = Replace(PYTHONPATH, "&", "^&")
    PYTHONPATH = Replace(PYTHONPATH, "\", "\\")
    
    If CondaPath <> "" And CondaEnv <> "" Then
        If CheckConda(CondaPath) = False Then
            Exit Function
        End If
        CondaCmd = CondaPath & "\condabin\conda activate " & CondaEnv & " && "
    Else
        CondaCmd = ""
    End If

    If IsFrozen = False Then
        RunCommand = CondaCmd & PythonInterpreter & " -B -c ""import xlwings.utils;xlwings.utils.prepare_sys_path(\""" & PYTHONPATH & "\""); " & PythonCommand & """ "
    ElseIf IsFrozen = True Then
        RunCommand = Chr(34) & PythonCommand & Chr(34) & " " & FrozenArgs & " "
    End If
    
    ExitCode = WScript.Run("cmd.exe /C " & DriveCommand & _
                       RunCommand & _
                       " --wb=" & """" & ActiveWorkbook.Name & """ --from_xl=1" & " --app=" & Chr(34) & _
                       Application.Path & "\" & Application.Name & Chr(34) & " --hwnd=" & Chr(34) & Application.Hwnd & Chr(34) & _
                       " 2> """ & LOG_FILE & """ ", _
                       WindowStyle, WaitOnReturn)

    'If ExitCode <> 0 then there's something wrong
    If ExitCode <> 0 Then
        Call ShowError(LOG_FILE)
        ExecuteWindows = -1
    End If

    ' Delete file after the error message has been shown
    On Error Resume Next
        Kill LOG_FILE
    On Error GoTo 0
End Function

Public Function RunFrozenPython(Executable As String, Optional Args As String)
    ' Runs a Python executable that has been frozen by PyInstaller and the like. Call the function like this:
    ' RunFrozenPython "C:\path\to\frozen_executable.exe", "arg1 arg2". Currently not implemented for Mac.

    ' Call Python
    #If Mac Then
        MsgBox "This functionality is not yet supported on Mac." & vbNewLine & _
               "Please run your scripts directly in Python!", vbCritical + vbOKOnly, "Unsupported Feature"
    #Else
        ExecuteWindows True, Executable, ParentFolder(Executable), , Args
    #End If
End Function

#If App = "Microsoft Excel" Then
Function GetUdfModules(Optional wb As Workbook) As String
#Else
Function GetUdfModules(Optional wb As Variant) As String
#End If
    Dim i As Integer
    Dim UDF_MODULES As String
    Dim sht As Worksheet

    GetUdfModules = GetConfig("UDF MODULES")
    ' Remove trailing ";"
    If Right$(GetUdfModules, 1) = ";" Then
        GetUdfModules = Left$(GetUdfModules, Len(GetUdfModules) - 1)
    End If
    
    ' Automatically add embedded code sheets
    For Each sht In wb.Worksheets
        If Right$(sht.Name, 3) = ".py" Then
            If GetUdfModules = "" Then
                GetUdfModules = Left$(sht.Name, Len(sht.Name) - 3)
            Else
                GetUdfModules = GetUdfModules & ";" & Left$(sht.Name, Len(sht.Name) - 3)
            End If
        End If
    Next

    ' Default
    If GetUdfModules = "" Then
        GetUdfModules = Left$(wb.Name, Len(wb.Name) - 5) ' assume that it ends in .xls*
    End If
    
End Function

Private Sub CleanUp()
    'On Mac only, this function is being called after Python is done (using Python's atexit handler)
    Dim LOG_FILE As String

    #If MAC_OFFICE_VERSION >= 15 Then
        LOG_FILE = Environ("HOME") + "/xlwings.log" '~/Library/Containers/com.microsoft.Excel/Data/xlwings.log
    #Else
        LOG_FILE = "/tmp/xlwings.log"
    #End If

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
    'TODO: the whole python vs. pythonw should be obsolete now that the console is shown/hidden by the dll
    Dim PYTHON_WIN As String, PYTHONPATH As String, LOG_FILE As String, tail As String, licenseKey As String, LicenseKeyEnvString As String, AddExcelDir As String
    Dim CondaCmd As String, CondaPath As String, CondaEnv As String, ConsoleSwitch As String, FName As String

    Dim DEBUG_UDFS As Boolean
    #If App = "Microsoft Excel" Then
    Dim wb As Workbook
    #End If

    ' TODO: Doesn't automatically check if code is embedded
    AddExcelDir = GetConfig("ADD_WORKBOOK_TO_PYTHONPATH", "true")

    ' The first 6 args are not technically part of the PYTHONPATH, but it's just easier to add it here (used by xlwings.utils.prepare_sys_path)
    #If App = "Microsoft Excel" Then
        PYTHONPATH = AddExcelDir & ";" & ActiveWorkbook.FullName & ";" & ThisWorkbook.FullName & ";" & GetConfig("ONEDRIVE_CONSUMER_WIN") & ";" & GetConfig("ONEDRIVE_COMMERCIAL_WIN") & ";" & GetConfig("SHAREPOINT_WIN") & ";" & GetConfig("PYTHONPATH")
    #Else
        ' Other office apps
        #If App = "Microsoft Word" Then
            FName = ThisDocument.FullName
        #ElseIf App = "Microsoft Access" Then
            FName = CurrentProject.FullName
        #ElseIf App = "Microsoft PowerPoint" Then
            FName = ActivePresentation.FullName
        #End If
        PYTHONPATH = FName & ";" & ";" & GetConfig("ONEDRIVE_CONSUMER_WIN") & ";" & GetConfig("ONEDRIVE_COMMERCIAL_WIN") & ";" & GetConfig("SHAREPOINT_WIN") & ";" & GetConfig("PYTHONPATH")
    #End If

    ' Escaping backslashes and quotes
    PYTHONPATH = Replace(PYTHONPATH, "\", "\\")
    PYTHONPATH = Replace(PYTHONPATH, "'", "\'")
    PYTHONPATH = Replace(PYTHONPATH, "&", "^&")
    
    PYTHON_WIN = GetConfig("INTERPRETER_WIN", "")
    If PYTHON_WIN = "" Then
        ' Legacy
        PYTHON_WIN = GetConfig("INTERPRETER", "pythonw")
    End If
    DEBUG_UDFS = GetConfig("DEBUG UDFS", False)

    ' /showconsole is a fictitious command line switch that's ignored by cmd.exe but used by CreateProcessA in the dll
    ' It's the only setting that's sent over like this at the moment
    If GetConfig("SHOW CONSOLE", False) = True Then
        ConsoleSwitch = "/showconsole"
    Else
        ConsoleSwitch = ""
    End If

    CondaPath = GetConfig("CONDA PATH")
    CondaEnv = GetConfig("CONDA ENV")

    If (PYTHON_WIN = "python" Or PYTHON_WIN = "pythonw") And (CondaPath <> "" And CondaEnv <> "") Then
        CondaCmd = Chr(34) & Chr(34) & CondaPath & "\condabin\conda" & Chr(34) & " activate " & CondaEnv & " && "
        PYTHON_WIN = "cmd.exe " & ConsoleSwitch & " /K " & CondaCmd & "python"
    Else
        PYTHON_WIN = "cmd.exe " & ConsoleSwitch & " /K " & Chr(34) & Chr(34) & PYTHON_WIN & Chr(34)
    End If

    licenseKey = GetConfig("LICENSE_KEY", "")
    If licenseKey <> "" Then
        LicenseKeyEnvString = "os.environ['XLWINGS_LICENSE_KEY']='" & licenseKey & "';"
    Else
        LicenseKeyEnvString = ""
    End If

    If DEBUG_UDFS = True Then
        XLPyCommand = "{506e67c3-55b5-48c3-a035-eed5deea7d6d}"
    Else
        ' Spaces in path of python.exe require quote around path AND quotes around whole command, see:
        ' https://stackoverflow.com/questions/6376113/how-do-i-use-spaces-in-the-command-prompt
        tail = " -B -c ""import sys, os;" & LicenseKeyEnvString & "import xlwings.utils;xlwings.utils.prepare_sys_path(\""" & PYTHONPATH & "\"");import xlwings; xlwings.serve('$(CLSID)')"""
        XLPyCommand = PYTHON_WIN & tail & Chr(34)
    End If
End Function

Private Sub XLPyLoadDLL()
    Dim PYTHON_WIN As String, CondaCmd As String, CondaPath As String, CondaEnv As String

    PYTHON_WIN = GetConfig("INTERPRETER_WIN", "")
    If PYTHON_WIN = "" Then
        ' Legacy
        PYTHON_WIN = GetConfig("INTERPRETER", "pythonw")
    End If
    CondaPath = GetConfig("CONDA PATH")
    CondaEnv = GetConfig("CONDA ENV")

    If (PYTHON_WIN = "python" Or PYTHON_WIN = "pythonw") And (CondaPath <> "" And CondaEnv <> "") Then
        ' This only works if the envs are in their default location
        ' Otherwise you'll have to add the full path for the interpreter in addition to the conda infos
        If CondaEnv = "base" Then
            PYTHON_WIN = CondaPath & "\" & PYTHON_WIN
        Else
            PYTHON_WIN = CondaPath & "\envs\" & CondaEnv & "\" & PYTHON_WIN
        End If
    End If

    If (PYTHON_WIN <> "python" And PYTHON_WIN <> "pythonw") Or (CondaPath <> "" And CondaEnv <> "") Then
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

Function XLPy()
    XLPyLoadDLL
    If 0 <> XLPyDLLActivateAuto(XLPy, XLPyCommand, 1) Then Err.Raise 1000, Description:=XLPy
End Function

Sub KillPy()
    XLPyLoadDLL
    Dim unused
    If 0 <> XLPyDLLActivateAuto(unused, XLPyCommand, -1) Then Err.Raise 1000, Description:=unused
End Sub

Sub ImportPythonUDFsBase(Optional addin As Boolean = False)
    ' This is called from the Ribbon button
    Dim tempPath As String, errorMsg As String
    Dim wb As Workbook

    If GetConfig("CONDA PATH") <> "" And CheckConda(GetConfig("CONDA PATH")) = False Then
        Exit Sub
    End If

    If addin = True Then
        Set wb = ThisWorkbook
    Else
        Set wb = ActiveWorkbook
    End If

    On Error GoTo ImportError
        tempPath = XLPy.Str(XLPy.Call(XLPy.Module("xlwings"), "import_udfs", XLPy.Tuple(GetUdfModules(wb), wb)))
    Exit Sub
ImportError:
    errorMsg = Err.Description & " " & Err.Number
    ShowError "", errorMsg
End Sub

Sub ImportPythonUDFs()
    ImportPythonUDFsBase
End Sub

Sub ImportPythonUDFsToAddin()
    ImportPythonUDFsBase addin:=True
End Sub

Sub ImportXlwingsUdfsModule(tf As String)
    ' Fallback: This is called from Python as direct pywin32 calls were sometimes failing, see comments in the Python code
    On Error Resume Next
    ActiveWorkbook.VBProject.VBComponents.Remove ActiveWorkbook.VBProject.VBComponents("xlwings_udfs")
    On Error GoTo 0
    ActiveWorkbook.VBProject.VBComponents.Import tf
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
