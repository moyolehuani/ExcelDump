Sub ImageFromURL()

fila = ActiveCell.Row
columna = ActiveCell.Column
Do While ActiveCell.Value <> Empty

    ActiveCell.Offset(0, 2).Select
    ActiveCell.FormulaR1C1 = "=IMAGE(""[URL]" & ActiveCell.Offset(0, -[X]) & ".[FORMAT]"","""",0)"
    
    fila = fila + 1
    ActiveSheet.Cells(fila, columna).Select

Loop

End Sub
