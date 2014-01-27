Attribute VB_Name = "xlwings"
Option Explicit

Public Function RunPython(PythonCommand As String)
    ' Runs the Python command, e.g.: to run the function foo() in module bar, call the function like this:
    ' RunPython ("import bar; bar.foo(*args, **kwargs)")
    '
    ' Python installation and Python source file directories can be changed, the defaults are:
    ' Python interpreter: Default interpreter from your PATH environment variable,i.e. the one you get by typing "python" at the command prompt
    ' Python file location: Same as the Excel file
    '
    ' xlwings makes it easy to deploy your Python powered Excel tools on Windows.
    ' Homepage and documentation: http://xlwings.org/
    '
    ' Copyright (c) 2014, Zoomer Analytics.
    ' Version: 0.1.0-dev
    ' License: MIT (see LICENSE.txt for details)
    
    Dim PYTHON_DIR As String, SOURCECODE_DIR As String, LOG_FILE As String, WORKBOOK_FULLNAME As String
    Dim DriveCommand As String
    Dim ExitCode As Integer
    Dim Wsh As Object
    Dim WaitOnReturn As Boolean: WaitOnReturn = True
    Dim WindowStyle As Integer: WindowStyle = 0
    
    ' Adjust according to where python.exe is on your system, e.g.: "C:\Python27"
    ' Leave empty if you want to use the default installation from your PATH environment variable,
    ' i.e. you want to use the installation you can start by just typing "python" at the command prompt
    PYTHON_DIR = ""
    
    ' Adjust according to the directory of the Python files
    SOURCECODE_DIR = ThisWorkbook.Path
    
    ' Fully qualified name of temporary error log file
    LOG_FILE = ThisWorkbook.Path & "\" & "log.txt"
    
    ' Get fully qualified name of Workbook
    WORKBOOK_FULLNAME = ThisWorkbook.FullName
    
    ' Call a command window and change to the directory of the Python installation
    ' Note: If Python is called from a different directory with the fully qualified path, pywintypesXX.dll won't be found.
    ' This is likely a pywin32 bug, see http://stackoverflow.com/q/7238403/918626
    ' Run Python with the "-c" command line switch: add the path of the python file and run the
    ' PythonCommand as first argument, then provide the WORKBOOK_FULLNAME as second argument.
    ' Log the errors in the LOG_FILE and wait with proceeding until the call returns.
    Set Wsh = CreateObject("WScript.Shell")
    If Left$(PYTHON_DIR, 2) Like "[A-Z]:" Then
        ' If Python is installed on a mapped or local drive, change to drive, then cd to path
        DriveCommand = Left$(PYTHON_DIR, 2) & " & cd " & PYTHON_DIR & " & "
    ElseIf Left$(PYTHON_DIR, 2) = "\\" Then
        ' In the unlikely event that Python is installed on a UNC path, temporarily mount and activate a drive letter with pushd
        DriveCommand = "pushd " & PYTHON_DIR & " & "
    End If
    
    ExitCode = Wsh.Run("cmd.exe /C " & DriveCommand & _
                   "python -c " & """import sys; sys.path.append(r'" & SOURCECODE_DIR & "'); " & PythonCommand & _
                    """ """ & WORKBOOK_FULLNAME & """ ""from_xl"" 2> """ & LOG_FILE & """  ", _
                   WindowStyle, WaitOnReturn)

    'If ExitCode <> 0 then there's something wrong
    If ExitCode <> 0 Then
        Call ShowError(LOG_FILE)
    End If
    
    ' Delete file after the error message has been shown
    On Error Resume Next
        Kill LOG_FILE
    On Error GoTo 0
    
    ' Make sure Wsh is cleared as otherwise moving the file between directoreis could create troubles
    Set Wsh = Nothing
    
End Function

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
