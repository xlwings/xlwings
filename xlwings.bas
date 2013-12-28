Attribute VB_Name = "xlwings"
Option Explicit

Public Function RunPython(PythonCommand As String)
    ' Runs the Python command, e.g.: to run the function foo() in module bar, call the function like this:
    ' RunPython ("import bar; bar.foo(*args, **kwargs)")
    '
    ' Python interpreter and Python source file directories can be changed, the defaults are:
    ' Python interpreter: Default interpreter from PATH, i.e. the one you get by typing "python" at the command prompt
    ' Python file location: Same as the Excel file
    '
    ' xlwings makes it easy to deploy your Python powered Excel tools (Windows only).
    ' Homepage and documentation: http://xlwings.org/
    '
    ' Copyright (c) 2013, Felix Zumstein.
    ' Version: 0.1-dev
    ' License: MIT (see LICENSE.txt for details)
    
    Dim PYTHON_DIR As String
    Dim PYTHONFILE_PATH As String
    Dim wbName As String
    Dim driveCommand As String
    Dim exitCode As Integer
    Dim wsh As Object
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 0
    
    ' Adjust according to the desired Python installation directory, e.g.: "C:\Python27"
    ' Leave as empty string if you want to use the default installation from your PATH environment variable,
    ' i.e. the installation you can start by just typing "python" at the command prompt
    
    PYTHON_DIR = ""
    
    ' Adjust according to the directory of the Python file
    PYTHONFILE_PATH = ThisWorkbook.Path
    
    ' Get Workbook name
    wbName = ThisWorkbook.Name
    
    ' Call a command window and change to the directory of the Python installation
    ' Note: If Python is called from a different directory with the full qualified path, pywintypesXX.dll won't be found.
    ' This is likely a pywin32 bug, see http://stackoverflow.com/q/7238403/918626
    ' Run Python with the Command Interface Option (-c): add the path of the python file and run the
    ' PythonCommand as first argument, then provide the wbName as second argument.
    ' Wait with proceeding until the call returns.
    Set wsh = CreateObject("WScript.Shell")
    If Left$(PYTHON_DIR, 2) Like "[A-Z]:" Then
        ' If Python is installed on a mapped or local drive, change to drive, then cd to path
        driveCommand = Left$(PYTHON_DIR, 2) & " & cd " & PYTHON_DIR & " & "
    ElseIf Left$(PYTHON_DIR, 2) = "\\" Then
        ' In the unlikely event that Python is installed on a UNC path, temporarily mount and activate a drive letter with pushd
        driveCommand = "pushd " & PYTHON_DIR & " & "
    End If
    
    exitCode = wsh.Run("cmd.exe /C " & driveCommand & _
                       "python -c """ & "import sys; sys.path.append(r'" & PYTHONFILE_PATH & "'); " & PythonCommand & _
                       """ """ & wbName & """", _
                       windowStyle, waitOnReturn)

    'If exitCode <> 0 then there's something wrong
    If exitCode <> 0 Then
        MsgBox "Oops - Something went wrong."
    End If
    
    ' Make sure wsh is cleared as moving the file could create troubles otherwise
    Set wsh = Nothing
    
End Function
