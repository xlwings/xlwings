Attribute VB_Name = "settings"
Public Const SETTINGSFILE = "C:\Users\Felix\Desktop\xlwings.ini"

Sub SetInterpreter(value As String)
    tf = SaveSetting(SETTINGSFILE, "INTERPRETER", value)
End Sub

Function GetInterpreter()
    Dim setting As String
    If GetSetting(SETTINGSFILE, "INTERPRETER", setting) Then
        GetInterpreter = setting
    Else
        tf = SaveSetting(SETTINGSFILE, "INTERPRETER", "python")
        GetInterpreter = setting
    End If
End Function

Function SetPythonpath(value As String)
    tf = SaveSetting(SETTINGSFILE, "PYTHONPATH", value)
End Function

Function GetPythonpath()
    Dim setting As String
    If GetSetting(SETTINGSFILE, "PYTHONPATH", setting) Then
        GetPythonpath = ActiveWorkbook.Path & ";" & setting
    Else
        GetPythonpath = ActiveWorkbook.Path
    End If
End Function

Function SetLogfile(value As String)
    tf = SaveSetting(SETTINGSFILE, "LOGFILE", value)
End Function

Function GetLogfile()
    Dim setting As String
    If GetSetting(SETTINGSFILE, "LOGFILE", setting) Then
        GetLogfile = setting
    End If
End Function

Function SetUdfmodules(value As String)
    tf = SaveSetting(SETTINGSFILE, "UDF_MODULES", value)
End Function

Function GetUdfmodules()
    Dim setting As String
    If GetSetting(SETTINGSFILE, "UDF_MODULES", setting) Then
        GetUdfmodules = setting
    End If
End Function

Function SetUdfDebug(value As String)
    tf = SaveSetting(SETTINGSFILE, "UDF_DEBUG", value)
End Function

Function GetUdfDebug()
    Dim setting As String
    If GetSetting(SETTINGSFILE, "UDF_DEBUG", setting) Then
        GetUdfDebug = setting
    End If
End Function

Function SetComServer(value As String)
    tf = SaveSetting(SETTINGSFILE, "COM_SERVER", value)
End Function

Function GetComServer()
    Dim setting As String
    If GetSetting(SETTINGSFILE, "COM_SERVER", setting) Then
        GetComServer = setting
    End If
End Function
