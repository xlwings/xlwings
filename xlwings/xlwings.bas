Attribute VB_Name = "xlwings"
' Make Excel fly!
'
' Homepage and documentation: http://xlwings.org
' See also: http://zoomeranalytics.com
'
' Copyright (C) 2014, Zoomer Analytics LLC.
' Version: 0.1.0
'
' License: BSD 3-clause (see LICENSE.txt for details)

Option Explicit

Dim PYTHON_DIR As String, SOURCECODE_DIR As String, WORKBOOK_FULLNAME As String
Dim LOG_FILE As String, DriveCommand As String, RunCommand As String
Dim ExitCode As Integer
Dim Wsh As Object
Dim IsFrozen As Boolean

Public Function RunFrozenPython(Executable As String)
    ' Runs a Python executable that has been frozen by cx_Freeze or py2exe. Call the function like this:
    ' RunFrozenPython("frozen_executable.exe")
    
    ' Adjust according to where the frozen executable is
    ' Preferrably use a path relative to this workbook: ThisWorkbook.Path & "\..."
    PYTHON_DIR = ThisWorkbook.Path & "\build\exe.win-amd64-2.7"
    
    ' Fully qualified name of temporary error log file
    LOG_FILE = ThisWorkbook.Path & "\log.txt"
    
    ' Call Python
    ExecuteProgram True, Executable, PYTHON_DIR
End Function

Public Function RunPython(PythonCommand As String)
    ' Runs the Python command, e.g.: to run the function foo() in module bar, call the function like this:
    ' RunPython ("import bar; bar.foo()")
    
    ' Adjust according to where python.exe is on your system, e.g.: "C:\Python27"
    ' Use an empty string if you want to call the default installation from your PATH environment variable,
    ' i.e. you want to use the installation you can start by typing "python" at the command prompt
    PYTHON_DIR = ""
    
    ' Adjust according to the directory of the Python files
    SOURCECODE_DIR = ThisWorkbook.Path
    
    ' Fully qualified name of temporary error log file
    LOG_FILE = ThisWorkbook.Path & "\" & "log.txt"
    
    ' Call Python
    ExecuteProgram False, PythonCommand, PYTHON_DIR, SOURCECODE_DIR
End Function

Sub ExecuteProgram(IsFrozen As Boolean, Command As String, PYTHON_DIR As String, Optional SOURCECODE_DIR As String)
    ' Call a command window and change to the directory of the Python installation or frozen executable
    ' Note: If Python is called from a different directory with the fully qualified path, pywintypesXX.dll won't be found.
    ' This is likely a pywin32 bug, see http://stackoverflow.com/q/7238403/918626

    ' Log the errors in the LOG_FILE and wait with proceeding until the call returns.
    Dim WaitOnReturn As Boolean: WaitOnReturn = True
    Dim WindowStyle As Integer: WindowStyle = 0
    Set Wsh = CreateObject("WScript.Shell")
    
    If Left$(PYTHON_DIR, 2) Like "[A-Z]:" Then
        ' If Python is installed on a mapped or local drive, change to drive, then cd to path
        DriveCommand = Left$(PYTHON_DIR, 2) & " & cd " & PYTHON_DIR & " & "
    ElseIf Left$(PYTHON_DIR, 2) = "\\" Then
        ' If Python is installed on a UNC path, temporarily mount and activate a drive letter with pushd
        DriveCommand = "pushd " & PYTHON_DIR & " & "
    End If
    
    ' Run Python with the "-c" command line switch: add the path of the python file and run the
    ' PythonCommand as first argument, then provide the WORKBOOK_FULLNAME as second argument.
    If IsFrozen = False Then
        ' Run Python with the "-c" command line switch: add the path of the python file
        RunCommand = "python -c ""import sys; sys.path.append(r'" & SOURCECODE_DIR & "'); " & Command & """ "
    ElseIf IsFrozen = True Then
        RunCommand = Command & " "
    End If
    
    ' Get the fully qualified name of Workbook
    WORKBOOK_FULLNAME = ThisWorkbook.FullName
    
    ExitCode = Wsh.Run("cmd.exe /C " & DriveCommand & _
                   RunCommand & _
                   """" & WORKBOOK_FULLNAME & """ ""from_xl"" 2> """ & LOG_FILE & """ ", _
                   WindowStyle, WaitOnReturn)

    'If ExitCode <> 0 then there's something wrong
    If ExitCode <> 0 Then
        Call ShowError(LOG_FILE)
    End If
    
    ' Delete file after the error message has been shown
    On Error Resume Next
        Kill LOG_FILE
    On Error GoTo 0
    
    ' Make sure Wsh is cleared as otherwise moving the file between directories could create troubles
    Set Wsh = Nothing
End Sub

Sub ShowError(FileName As String)
    Dim ReadData As String
    Dim Token As String
    Dim FileNum As Integer
    
    FileNum = FreeFile
    ReadData = ""
    
    ' Read Text File
    Open FileName For Input As #FileNum
        Do While Not EOF(FileNum)
            Line Input #FileNum, Token
            ReadData = ReadData & Token & vbCr
        Loop
    Close #FileNum
    
    ReadData = ReadData & vbCr
    ReadData = ReadData & "Press Ctrl+C to copy this message to the clipboard."
    
    ' MsgBox
    MsgBox ReadData, vbCritical, "Error"
End Sub
