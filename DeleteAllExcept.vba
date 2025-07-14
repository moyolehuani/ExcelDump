Sub AnainvBorraHojas()
    Dim xWs As Worksheet
    Application.DisplayAlerts = False
    For Each xWs In Application.ActiveWorkbook.Worksheets
        If xWs.Name <> "[NAME]" Then
            xWs.Delete
        End If
    Next
    'Application.DisplayAlerts = True
End Sub
