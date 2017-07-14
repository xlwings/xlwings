Attribute VB_Name = "Extensions"
Function sql(query, ParamArray tables())
        If TypeOf Application.Caller Is Range Then On Error GoTo failed
        ReDim argsArray(1 To UBound(tables) - LBound(tables) + 2)
        argsArray(1) = query
        For k = LBound(tables) To UBound(tables)
        argsArray(2 + k - LBound(tables)) = tables(k)
        Next k
        sql = Py.CallUDF("xlwings.ext", "sql", argsArray, ActiveWorkbook, Application.Caller)
        Exit Function
failed:
        sql = Err.Description
End Function
