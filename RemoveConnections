Sub EliminarConexiones()
    Dim conexion As Object
    
    ' Eliminar todas las conexiones de libro externo
    For Each conexion In ThisWorkbook.Connections
        conexion.Delete
    Next conexion
    
    ' Eliminar todas las conexiones de datos en las hojas
    For Each hoja In ThisWorkbook.Sheets
        For Each conexion In hoja.QueryTables
            conexion.Delete
        Next conexion
    Next hoja
End Sub

Sub RemoveConnections()
    
    Dim wb As Workbook

    For Each wb In Workbooks
        wb.AcceptAllChanges
        Call RemoveConnections(wb)
    Next wb
    
    Dim conn As Long
    With ActiveWorkbook
        For conn = .Connections.Count To 1 Step -1
            .Connections(conn).Delete
        Next conn
    End With
End Sub
Sub EliminarConexiones()
    Dim conexion As Object
    
    ' Eliminar todas las conexiones de libro externo
    For Each conexion In ThisWorkbook.Connections
        conexion.Delete
    Next conexion
    
    ' Eliminar todas las conexiones de datos en las hojas
    For Each hoja In ThisWorkbook.Sheets
        For Each conexion In hoja.QueryTables
            conexion.Delete
        Next conexion
    Next hoja
End Sub
