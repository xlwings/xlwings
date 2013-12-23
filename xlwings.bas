Attribute VB_Name = "xlwings"
Option Explicit

Public Function RunPython(PythonCommand As String)
    ' Runs the Python command, e.g.: to run the function foo() in module bar, call the function like this:
    ' RunPython ("import bar; bar.foo(*args, **kwargs)")
    '
    ' Python interpreter and Python file location can be adjusted, the defaults are:
    ' Python interpreter: "python"
    ' Python file location: Same as the calling Excel file
    '
	' xlwings is an easy way to connect your Excel tools with Python (Windows only).
    ' The aim is to make it as easy as possible to distribute the Excel files.
    '
    ' Homepage and documentation: http://xlwings.org/
    '
    ' Copyright (c) 2013, Felix Zumstein.
    ' Version: 0.1-dev
    ' License: MIT (see LICENSE.txt for details)
    
    Dim wsh As Object
    Dim pyFilePath As String
    Dim wbName As String
    Dim pythonExe As String
    Dim drive As String
    Dim returnValue As Integer
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 0
    
    ' Adjust according to the desired Python executable
    ' e.g.: "python", "python3", "C:\Python27\python"
    pythonExe = "python"
    
    ' Adjust according to the directory of the Python file
    pyFilePath = ThisWorkbook.Path
    
    ' Get Workbook name
    wbName = ThisWorkbook.Name
    
    ' Call a command window, change to filePath, run the Python file with the Command Interface Option (-c) and provide the
    ' PythonCommand as first argument and the wbName as second argument. Wait with proceeding until the call returns.
    Set wsh = VBA.CreateObject("WScript.Shell")
    If Left$(pyFilePath, 2) = "\\" Then
        ' If UNC path, temporarily mount and activate a drive letter with pushd
        returnValue = wsh.Run("cmd.exe /C pushd " & pyFilePath & " & " & pythonExe & " -c """ & PythonCommand & """ """ & wbName & """", windowStyle, waitOnReturn)
    ElseIf Left$(pyFilePath, 2) Like "[A-Z]:" Then
        ' If mapped or local drive, change to drive, then cd to path
        drive = Left$(pyFilePath, 2)
        returnValue = wsh.Run("cmd.exe /C " & drive & " & cd " & pyFilePath & " & " & pythonExe & " -c """ & PythonCommand & """ """ & wbName & """", windowStyle, waitOnReturn)
    End If
    
    'If returnValue <> 0 then there's something wrong
    If returnValue <> 0 Then
        MsgBox "Oops - Something went wrong."
    End If
    
    ' Make sure wsh is cleared as moving the file could create troubles otherwise
    Set wsh = Nothing
    
End Function
