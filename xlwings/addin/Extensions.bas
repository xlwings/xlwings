Attribute VB_Name = "Extensions"
Function sql(query, ParamArray tables())
        If TypeOf Application.Caller Is Range Then On Error GoTo failed
        ReDim argsArray(1 To UBound(tables) - LBound(tables) + 2)
        argsArray(1) = query
        For K = LBound(tables) To UBound(tables)
        argsArray(2 + K - LBound(tables)) = tables(K)
        Next K
        If has_dynamic_array() Then
            sql = XLPy.CallUDF("xlwings.ext", "sql_dynamic", argsArray, ActiveWorkbook, Application.Caller)
        Else
            sql = XLPy.CallUDF("xlwings.ext", "sql", argsArray, ActiveWorkbook, Application.Caller)
        End If
        Exit Function
failed:
        sql = Err.Description
End Function
