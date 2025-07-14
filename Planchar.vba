Sub Planchar()
'
' Planchar Macro
'

'
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A1").Select
End Sub
