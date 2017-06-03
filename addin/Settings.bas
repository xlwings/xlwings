Attribute VB_Name = "settings"
Function GetSettingsFile() As String
    #If Mac Then

    #Else
        GetSettingsFile = Environ("USERPROFILE") & "\.xlwings\xlwings.conf"
    #End If
End Function

Function SettingsSheetExists()
    Dim sht As Worksheet
    On Error Resume Next
        Set sht = ActiveWorkbook.Sheets("xlwings.conf")
    On Error GoTo 0
    SettingsSheetExists = Not sht Is Nothing
End Function

Function GetSheetSettings() As Object
    Dim lastCell As Range
    Set GetSheetSettings = CreateObject("Scripting.Dictionary")

    If ActiveWorkbook.Sheets("xlwings.conf").Range("A2") = "" Then
        Set lastCell = ActiveWorkbook.Sheets("xlwings.conf").Range("A1")
    Else
        Set lastCell = ActiveWorkbook.Sheets("xlwings.conf").Range("A1").End(xlDown)
    End If

    For Each cell In Range(ActiveWorkbook.Sheets("xlwings.conf").Range("A1"), lastCell)
        GetSheetSettings.Add cell.value, cell.Offset(0, 1).value
    Next cell
End Function

Sub SetInterpreter(value As String)
    tf = SaveSetting(GetSettingsFile, "INTERPRETER", value)
End Sub

Function GetInterpreter()
    Dim setting As String
    If GetSetting(GetSettingsFile(), "INTERPRETER", setting) Then
        GetInterpreter = setting
    Else
        tf = SaveSetting(GetSettingsFile(), "INTERPRETER", "python")
        GetInterpreter = setting
    End If
End Function

Function SetPythonpath(value As String)
    tf = SaveSetting(GetSettingsFile(), "PYTHONPATH", value)
End Function

Function GetPythonpath()
    Dim setting As String
    If GetSetting(GetSettingsFile(), "PYTHONPATH", setting) Then
        GetPythonpath = ActiveWorkbook.Path & ";" & setting
    Else
        GetPythonpath = ActiveWorkbook.Path
    End If
End Function

Function SetLogfile(value As String)
    tf = SaveSetting(GetSettingsFile(), "LOGFILE", value)
End Function

Function GetLogfile()
    Dim setting As String
    If GetSetting(GetSettingsFile(), "LOGFILE", setting) Then
        GetLogfile = setting
    End If
End Function

Function SetUdfmodules(value As String)
    tf = SaveSetting(GetSettingsFile(), "UDF_MODULES", value)
End Function

Function GetUdfmodules()
    Dim setting As String
    If GetSetting(GetSettingsFile(), "UDF_MODULES", setting) Then
        GetUdfmodules = setting
    End If
End Function

Function SetUdfDebug(value As String)
    tf = SaveSetting(GetSettingsFile(), "UDF_DEBUG", value)
End Function

Function GetUdfDebug()
    Dim setting As String
    If GetSetting(GetSettingsFile(), "UDF_DEBUG", setting) Then
        GetUdfDebug = setting
    End If
End Function

Function SetComServer(value As String)
    tf = SaveSetting(GetSettingsFile(), "COM_SERVER", value)
End Function

Function GetComServer()
    Dim setting As String
    If GetSetting(GetSettingsFile(), "COM_SERVER", setting) Then
        GetComServer = setting
    End If
End Function
