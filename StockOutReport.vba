Sub InicializarHojaAgotados()
    Dim RutaFuente As String

    RutaFuente = "[REDACTED]"
    
    Dim ListaAgotados As Variant
    ListaAgotados = Array()
    j = 0
    
    With Workbooks(RutaFuente).Worksheets("Hoja1")
        For Each clave In .Range("C")
            If Not IsInArray(clave, ListaAgotados) Then
                ReDim Preserve ListaAgotados(j)
                ListaAgotados(j) = clave
                j = j + 1
            End If
        Next clave
    End With
    
    'Workbooks(RutaFuente).Worksheets("Hoja1").Range("C").AdvancedFilter _
        'Action:=xlFilterCopy, CopyToRange:=.Range()
    
    'Range("C").AdvancedFilter _
    '    Action:=xlFilterCopy, CopyToRange:=Range("E1"), Unique:=True
    
End Sub

Sub LlenarHojaAgotados()
    'define una lista con todos los productos agotados
    Dim ListaAgotados As Variant
    ListaAgotados = Array()
    
    'lee la columna de productos agotados y va agregando esas claves en la lista de agotados
    i = 13
    j = 0
    Do While Not IsEmpty(Worksheets("AGOTADOS").Cells(i, 4).Value)
        ReDim Preserve ListaAgotados(j)
        ListaAgotados(j) = Worksheets("AGOTADOS").Cells(i, 4).Value
        i = i + 1
        j = j + 1
    Loop
    'reinicializa valores de renglón
    i = 13
    'Imprime Categoría
    'For Each clave In ListaAgotados
        'Worksheets("AGOTADOS").Cells(i, 3).Value =
        'Worksheets("AGOTADOS").Cells(i, 3).Value = WorksheetFunction.VLookup(clave, Range("Y:\1. Poyectos\A. Catalogos\Tablas\[Base de productos.xlsm]CATALOGO"!A:F), 4, 0)
     '   i = i + 1
    'Next clave
End Sub
Sub FormatoAgotados()
    Dim ListaAgotados As Variant
    ListaAgotados = Array(1)
    
    'Ruta del archivo Fuente
    RutaFuente = "Fuente Hipotetica.xlsx"
    'Ruta del archivo de bases de datos
    RutaBD = "ClavesPaquete C10.xlsx"
    
    'inicializa índice del array de agotados
    Dim indice As Integer
    indice = 0
    
    'lee la columna de productos agotados y va agregando esas claves en la lista de agotados
    'las claves están en la columna C
    With Workbooks(RutaFuente).Worksheets("Hoja1")
        Rows("1").Delete
        For Each clave In Range("C:C")
            If UBound(Filter(ListaAgotados, clave)) = -1 Then
                ReDim Preserve ListaAgotados(indice)
                ListaAgotados(indice) = clave
                indice = indice + 1
            End If
        Next clave
    End With
    Sheets.Add.Name = "Productos"
    Workbooks(RutaFuente).Worksheets("Hoja1").Name = "Reporte de Agotados"
    
    'Bandera para confirmar que el arreglo está bien definido:
    'For i = LBound(ListaAgotados) To UBound(ListaAgotados)
    '    msg = msg & ListaAgotados(i) & vbNewLine
    'Next i
    'MsgBox "Los productos agotados son: " & vbNewLine & msg
    
    'Imprime en Nueva Hoja los productos agotados
    CamposXProducto = Array("Clave", "Imagen", "Descripción", "Categoría", "Canal", "Status", "Origen")
    'Configura los campos
    For i = LBound(CamposXProducto) To UBound(CamposXProducto)
        Workbooks(RutaFuente).Worksheets("Productos").Cells(1, i + 1).Value = CamposXProducto(i)
    Next i
    'Pone los productos
    For i = LBound(ListaAgotados) To UBound(ListaAgotados)
        Workbooks(RutaFuente).Worksheets("Productos").Cells(i + 2, 1).Value = ListaAgotados(i)
    Next i
    
    'Imágenes
    
    'Dirección en el servidor de Amazon
    Dim ImgPath As String
    ImgPath = "[REDACTED]"
    
    Dim RutasImagenes As Variant
    RutasImagenes = Array(i)
    
    For i = LBound(ListaAgotados) To UBound(ListaAgotados)
        ReDim Preserve RutasImagenes(i)
        RutasImagenes(i) = ImgPath & ListaAgotados(i) & ".png"
    Next i
    
    'Pone las imágenes en las celdas
    'For i = LBound(RutasImagenes) To UBound(RutasImagenes)
    '    Workbooks(RutaFuente).Worksheets("Productos").Cells(i + 2, 2).Value = "=IMAGE('" & RutaFuente(i) & "',"""",0)"
        '"=IMAGE(""" & RutaFuente(i) & ""","""",0)"
    'Next i
    
    'Pone la descripción del producto
    Dim ProductDescriptions As Variant
    ProductDescriptions = Array()
    
    With Workbooks(RutaBD).Worksheets("Hoja1")
        For i = LBound(ListaAgotados) To UBound(ListaAgotados)
            Set RangoBD = .Range("G:M")
            ReDim Preserve ProductDescriptions(i)
            ProductDescriptions(i) = Application.WorksheetFunction.VLookup(ListaAgotados(i), RangoBD, 3, False)
        Next i
    End With
    
    'Imprime Descripción de Productos
    For i = LBound(ProductDescriptions) To UBound(ProductDescriptions)
        Workbooks(RutaFuente).Worksheets("Productos").Cells(i + 2, 3).Value = ProductDescriptions(i)
    Next i
    
    'Categoría de Productos
    Dim ProductCategory As Variant
    ProductCategory = Array()
    
    For i = LBound(ListaAgotados) To UBound(ListaAgotados)
            ReDim Preserve ProductCategory(i)
            ProductCategory(i) = Application.WorksheetFunction.VLookup(ListaAgotados(i), RangoBD, 4, False)
    Next i
    
    For i = LBound(ProductCategory) To UBound(ProductCategory)
        Workbooks(RutaFuente).Worksheets("Productos").Cells(i + 2, 4).Value = ProductCategory(i)
    Next i
    
    'Canal de Productos
    Dim ProductChannel As Variant
    ProductChannel = Array()
    
    For i = LBound(ListaAgotados) To UBound(ListaAgotados)
            ReDim Preserve ProductChannel(i)
            ProductChannel(i) = Application.WorksheetFunction.VLookup(ListaAgotados(i), RangoBD, 5, False)
    Next i
    
    For i = LBound(ProductChannel) To UBound(ProductChannel)
        Workbooks(RutaFuente).Worksheets("Productos").Cells(i + 2, 5).Value = ProductChannel(i)
    Next i
    
    'Status del Producto
    Dim ProductStatus As Variant
    ProductStatus = Array()
    
    For i = LBound(ListaAgotados) To UBound(ListaAgotados)
            ReDim Preserve ProductStatus(i)
            ProductStatus(i) = Application.WorksheetFunction.VLookup(ListaAgotados(i), RangoBD, 6, False)
    Next i
    
    For i = LBound(ProductStatus) To UBound(ProductStatus)
        Workbooks(RutaFuente).Worksheets("Productos").Cells(i + 2, 6).Value = ProductStatus(i)
    Next i
    
    'Origen del Producto
    Dim ProductOrigin As Variant
    ProductOrigin = Array()
    
    For i = LBound(ListaAgotados) To UBound(ListaAgotados)
            ReDim Preserve ProductOrigin(i)
            ProductOrigin(i) = Application.WorksheetFunction.VLookup(ListaAgotados(i), RangoBD, 7, False)
    Next i
    
    For i = LBound(ProductOrigin) To UBound(ProductOrigin)
        Workbooks(RutaFuente).Worksheets("Productos").Cells(i + 2, 7).Value = ProductOrigin(i)
    Next i
    
    
End Sub
Sub VTANETA()
'
' VTANETA Macro
'

'
    Range("I1").Select
    Selection.EntireColumn.Insert
    ActiveCell.FormulaR1C1 = "CY"
    Range("M1").Select
    Selection.EntireColumn.Insert
    ActiveCell.FormulaR1C1 = "VTA_NETA"
    Columns("M:M").Select
    Selection.NumberFormat = "$#,##0.00"
'
    Dim lastRow As Long
    Dim row_no As Long
    
    'Dim sht As Worksheet
    
    'sht = ThisWorkbook.Worksheets("Sheet1")
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For row_no = 2 To lastRow
        ActiveSheet.Cells(row_no, 13).Value = (ActiveSheet.Cells(row_no, 10).Value - ActiveSheet.Cells(row_no, 12).Value)
    Next
    
End Sub

Sub BaseDatos_Agotados_DarFormato()
'
' BaseDatos_Agotados_DarFormato Macro
'

'
    ActiveCell.FormulaR1C1 = "Descripción"
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "Categoría"
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "Canal"
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "Status"
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Origen"
    Range("M2").Select
End Sub
Sub DarFormato()
'
' DarFormato Macro
'

'
    ActiveCell.FormulaR1C1 = ""
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "Descripción"
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "Categoría"
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "Canal"
    Range("P1").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.FormulaR1C1 = "Status"
    Range("Q1").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.FormulaR1C1 = "Origen"
    Range("M2").Select
    ActiveSheet.PasteSpecial Format:="Texto Unicode", Link:=False, _
        DisplayAsIcon:=False, NoHTMLFormatting:=True
    ActiveCell.FormulaR1C1 = _
        "BUSCARV(G2,'CATALOGO'!A:F,2,0)"
    Range("M2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-6],'CATALOGO'!C[-12]:C[-7],2,0)"
    Range("N2").Select
    ActiveSheet.Paste
    Range("O2").Select
    ActiveSheet.Paste
    Range("O2").Select
    Windows("Fuente Agotados.xlsx").Activate
    Range("P2").Select
    ActiveCell.FormulaR1C1 = _
        "=INDEX('ANAINVXANT'!C3,MATCH(RC[-9],'ANAINVXANT'!C4,0))"
    Range("Q2").Select
    ActiveSheet.Paste
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-8],[TOs1708.xlsx]TOs!C1:C5,3,0)"
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-8],[TOs1708.xlsx]TOs!C1:C5,3,0)"
    Range("O9").Select
End Sub
Sub DarFormatoBDAgotados()
'
' DarFormatoBDAgotados Macro
'

'
    Columns("I:M").Select
    Selection.ClearContents
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Descripción"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Categoría"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Canal"
    Range("L1").Select
    Selection.NumberFormat = "@"
    ActiveCell.FormulaR1C1 = "Status"
    Range("M1").Select
    Selection.NumberFormat = "@"
    ActiveCell.FormulaR1C1 = "Origen"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = _
        "=BUSCARV(G2,'CATALOGO'!A:F,2,0)"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = _
        "=BUSCARV(G2,'CATALOGO'!A:F,2,0)"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = _
        "=BUSCARV(G2,'CATALOGO'!A:F,2,0)"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = _
        "=BUSCARV(G2,'CATALOGO'!A:F,2,0)"
    Range("I3").Select
    ActiveCell.FormulaR1C1 = _
        "BUSCARV(G3,'CATALOGO'!A:F,2,0)"
    Range("I3").Select
    ActiveCell.FormulaR1C1 = _
        "=BUSCARV(G3,'CATALOGO'!A:F,2,0)"
    Range("I4").Select
    Columns("I:I").ColumnWidth = 53.91
    Range("I2").Select
    ActiveCell.FormulaR1C1 = _
        "=BUSCARV(G2,'CATALOGO'!A:F,2,0)"
    Range("I2").Select
    Selection.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-2],'CATALOGO'!C[-8]:C[-3],2,0)"
    Range("I2").Select
    Selection.AutoFill Destination:=Range("I2:I32"), Type:=xlFillDefault
    Range("I2:I32").Select
    Range("J2").Select
    Columns("I:I").ColumnWidth = 31.27
    Range("J2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-3],'CATALOGO'!C[-9]:C[-4],4,0)"
    Range("J6").Select
    Columns("J:J").ColumnWidth = 21.45
    Range("J2").Select
    Selection.AutoFill Destination:=Range("J2:J32"), Type:=xlFillDefault
    Range("J2:J32").Select
    Range("K2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-3],'TOs'!C1:C5,3,0)"
    Range("K3").Select
    Columns("K:K").ColumnWidth = 21.45
    Range("K2").Select
    Selection.AutoFill Destination:=Range("K2:K36"), Type:=xlFillDefault
    Range("K2:K36").Select
    Range("L2").Select
    ActiveCell.FormulaR1C1 = _
        "=INDEX('ANAINVXANT'!C3,MATCH(RC[-5],'ANAINVXANT'!C4,0))"
    Range("L2").Select
    Selection.AutoFill Destination:=Range("L2:L34"), Type:=xlFillDefault
    Range("L2:L34").Select
    Range("M2").Select
    ActiveCell.FormulaR1C1 = _
        "=INDEX('ANAINVXANT'!C34,MATCH(RC[-6],'ANAINVXANT'!C4,0))"
    Range("M2").Select
    Selection.AutoFill Destination:=Range("M2:M35"), Type:=xlFillDefault
    Range("M2:M35").Select
    Range("M22").Select
End Sub

Sub grabarfecha()
'
' grabarfecha Macro
'

'
    Range("L4:P5").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = "=+TODAY()"
    Range("L4:P5").Select
    Selection.NumberFormat = "[$-x-sysdate]dddd, mmmm dd, yyyy"
    With Selection.Font
        .Name = "Calibri"
        .Size = 16
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Range("L4:P5").Select
End Sub

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

Sub HojaChidaAgotados()
'
' HojaChidaAgotados Macro
'

'
    Range("A6:H6").Select
    Application.CutCopyMode = False
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection.Font
        .Name = "Calibri"
        .Size = 18
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    ActiveCell.FormulaR1C1 = "PRODUCTOS DE LÍNEA"
    Range("B7").Select
    ActiveCell.FormulaR1C1 = "ORIGEN"
    Range("C7").Select
    ActiveCell.FormulaR1C1 = "CLAVE"
    Range("C7").Select
    ActiveCell.FormulaR1C1 = "CATEGORÍA"
    Range("D7").Select
    ActiveCell.FormulaR1C1 = "CLAVE"
    Range("E7").Select
    ActiveCell.FormulaR1C1 = "DESCRIPCIÓN"
    Range("F7").Select
    ActiveCell.FormulaR1C1 = "IMAGEN"
    Range("G7").Select
    ActiveCell.FormulaR1C1 = "CANAL"
    Range("H7").Select
    ActiveCell.FormulaR1C1 = "RAZÓN"
    Range("I7").Select
    ActiveCell.FormulaR1C1 = "JUEVES"
    Range("J7").Select
    ActiveCell.FormulaR1C1 = "VIERNES"
    Range("K7").Select
    ActiveCell.FormulaR1C1 = "LUNES"
    Range("L7").Select
    ActiveCell.FormulaR1C1 = "MARTES"
    Range("M7").Select
    ActiveCell.FormulaR1C1 = "MIÉRCOLES"
    Range("N7").Select
    ActiveCell.FormulaR1C1 = "JUEVES"
    Range("O7").Select
    ActiveCell.FormulaR1C1 = "VIERNES"
    Range("P7").Select
    ActiveCell.FormulaR1C1 = "LUNES"
    Range("Q7").Select
    ActiveCell.FormulaR1C1 = "MARTES"
    Range("R7").Select
    ActiveCell.FormulaR1C1 = "MIÉRCOLES"
    Range("I6:M6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = "SEMANA 1"
    Range("N6:R6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = "SEMANA 2"
    Range("I6:R6").Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 18
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Range("I7").Select
    Windows("Formato agotados.xls").Activate
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    Windows("REPORTES_AGOTADOS.xlsx").Activate
    Range("S7").Select
    ActiveCell.FormulaR1C1 = "S1"
    Range("T7").Select
    ActiveCell.FormulaR1C1 = "S2"
    Range("U7").Select
    ActiveCell.FormulaR1C1 = "ACUMULADOS"
    Range("V7").Select
    ActiveCell.FormulaR1C1 = "%"
    Range("W7").Select
End Sub
Sub CrearHojaAgotados(AAXX)
    'Crea la hoja de agotados por primera vez (idealmente, sería en el día 1 de la campaña)
    Call BorraHojas
    'Crea hojas
    Sheets.Add After:=ActiveSheet
    Sheets.Add After:=ActiveSheet
    Sheets.Add After:=ActiveSheet
    Sheets(2).Select
    Sheets(2).Name = "AGOTADOS POR RUTA"
    Sheets(3).Select
    Sheets(3).Name = "AGOTADOS POR DIVISION"
    Sheets(4).Select
    Sheets(4).Name = "AGOTADOS POR DIVISION DIA"
    Sheets.Add After:=ActiveSheet
    Sheets(5).Select
    Sheets(5).Name = "AGOTADOS EN CAMPAÑA"
    Sheets.Add After:=ActiveSheet
    Sheets(6).Select
    Sheets(6).Name = "SUSTITUCIONES"
    
    'Da formato
    For Each hoja In Worksheets
        If hoja.Name <> "PROCESOS" Then
            hoja.Activate
            Range("A1:Z1").Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.Merge
            With Selection.Font
                .Name = "Century Gothic"
                .Size = 28
                .Bold = True
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontMinor
            End With
            With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(239, 51, 64)
            .TintAndShade = 0
            .PatternTintAndShade = 0
            End With
            ActiveCell.Value = "FULLER COSMETICS, S.A. DE C.V."
            Range("A2:Z2").Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.Merge
            With Selection.Font
                .Name = "Century Gothic"
                .Bold = True
                .Size = 28
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontMinor
            End With
            With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(239, 51, 64)
            .TintAndShade = 0
            .PatternTintAndShade = 0
            End With
            Selection.Value = "GERENCIA DE LOGÍSTICA"
            Range("A3:Z3").Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.Merge
            With Selection.Font
                .Name = "Century Gothic"
                .Size = 22
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontMinor
            End With
            With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(239, 51, 64)
            .TintAndShade = 0
            .PatternTintAndShade = 0
            End With
            Selection.Value = "REPORTE DE AGOTADOS CORRESPONDIENTES A CAMPAÑA " & Mid(AAXX, 3, 2) & " FY 20" & Mid(AAXX, 1, 2)
            Range("U4:Z5").Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.Merge
            ActiveCell.FormulaR1C1 = "=+TODAY()"
            Range("W4:Z5").Select
            Selection.NumberFormat = "[$-x-sysdate]dddd, mmmm dd, yyyy"
            With Selection.Font
                .Name = "Calibri"
                .Size = 18
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontMinor
            End With
        End If
    Next hoja
    
    'Hoja del reporte secreto
    Sheets.Add After:=ActiveSheet
    Sheets(7).Select
    Sheets(7).Name = "Reporte Secreto"
    Sheets("Reporte Secreto").Visible = False
    Sheets("Reporte Secreto").Activate
    Cells.Select
    
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    'Encabezados de agotados por ruta
    weekdays = Array("Jueves", "Viernes", "Lunes", "Martes", "Miércoles")
    
    Dim j As Integer
    j = 1
    
    Worksheets("AGOTADOS POR RUTA").Activate
    For i = 1 To 10
        Range(Cells(7, j), Cells(7, j + 5)).Select
        With Selection
            With .Font
                .Name = "Calibri"
                .Size = 26
            End With
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .MergeCells = True
        End With
        ActiveCell.Value = weekdays((i - 1) Mod 5)
        j = j + 6
    Next i
    
    'encabezados de agotados por división
    j = 1
    
    Worksheets("AGOTADOS POR DIVISION").Activate
    For i = 1 To 10
        Range(Cells(7, j), Cells(7, j + 5)).Select
        With Selection
            With .Font
                .Name = "Calibri"
                .Size = 26
            End With
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .MergeCells = True
        End With
        ActiveCell.Value = weekdays((i - 1) Mod 5)
        j = j + 6
    Next i
    
    'encabezados de agotados por división día
    j = 1
    Worksheets("AGOTADOS POR DIVISION DIA").Activate
    For i = 1 To 10
        Range(Cells(7, j), Cells(7, j + 6)).Select
        With Selection
            With .Font
                .Name = "Calibri"
                .Size = 26
            End With
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .MergeCells = True
        End With
        ActiveCell.Value = weekdays((i - 1) Mod 5)
        j = j + 7
    Next i
    
    'ActiveWorkbook.SaveAs Filename:=Environ("USERPROFILE") & "\Documentos\Cargas de Agotados\FY" & Left(AAXX, 2) & "_C" & Right(AAXX, 2) & "\REPORTES_AGOTADOS"
    ActiveWorkbook.Save
End Sub

Sub AgregarReporteAHojaSecreta(archnom)
    'Application.ScreenUpdating = False
    'este sub recibe como argumentos la ruta del archivo reporte individual por día del que se extraerá la información, la semana será S1,S2 y el día será el día de la campaña del 1 al 10
    'Es imperativo que para este sub ya exista REPORTE_AGOTADOS, para que lo abra
    Set reporte = Workbooks.Open(archnom)
    Set destino = Workbooks("REPORTES_AGOTADOS")
    
    Dim xLastRowReporte As Integer
    Dim i As Integer
    
    'borra lo que haya en esa hoja, por si acaso
    destino.Worksheets("Reporte Secreto").Cells.ClearContents
    
    'saca el lastrow del archivo reporte, sólo de la primera hoja,
    'que en teoría debería ser la única, pero se manda llamar explícitamente
    'por si acaso
    xLastRowReporte = reporte.Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row
    
    'Coloca el reporte en la hoja oculta dentro del archivo del informe de agotados
    reporte.Worksheets(1).Range(Cells(1, 1), Cells(xLastRowReporte, 5)).Copy destino.Worksheets("Reporte Secreto").Range("A1")
    
    'cierra el archivo reporte, sin guardar cambios
    reporte.Close savechanges:=False
    
    'llena en la hoja secreta los datos que faltan con el archivo BD
    Dim basedatos As Workbook
    Set basedatos = Workbooks.Open(destino.Path & "\BD.xlsx")
        'pone los nombres de las categorías
        'a partir de la columna F
        categories = Array("DESCRIPCIÓN", "CATEGORÍA", "CANAL", "STATUS", "ORIGEN")
        columnas = Array("4", "5", "6", "7", "8")
        Workbooks("REPORTES_AGOTADOS").Worksheets("Reporte Secreto").Activate
        i = 0
        For Each elemento In categories
            Cells(1, 6 + i).Value = elemento
            i = i + 1
        Next elemento
        'jala los datos faltantes
        If xLastRowReporte > 1 Then
            For i = 1 To 5
                For j = 2 To xLastRowReporte
                    destino.Worksheets("Reporte Secreto").Cells(j, 5 + i).Formula = "=VLOOKUP(C" & j & ",'[BD.xlsx]Hoja1'!F:M," & columnas(i - 1) & ",0)"
                Next j
            Next i
        End If
    basedatos.Close savechanges:=False
    
    'Plancha los datos
    destino.Worksheets("Reporte Secreto").Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub ActivarHojas(filenom)
'esta función sirve para abrir archivos
    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.Name = filenom Then
            wb.Activate
            MsgBox "¡El archivo se abrió exitosamente!"
            Exit Sub
        End If
    Next wb
        MsgBox "Por favor abre el archivo " & filenom & ", que se puede hallar en la misma carpeta"
End Sub

Sub MAIN_Agotados()
'esta macro será la que interactúe con el usuario
    Dim xLastRowReporte As Integer
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    weekdays = Array("Jueves", "Viernes", "Lunes", "Martes", "Miercoles")
    weeks = Array("S1", "S2")
    
    'pregunta qué día de qué semana quiere cargar el reporte
    thisweek = InputBox("Escriba 1 si estamos en S1, 2 si estamos en S2", "¿En qué semana de la campaña estamos?")
    thisday = InputBox("Introduzca 1 si estamos en el primer jueves de la campaña, 2 si es el primer viernes, y así sucesivamente hasta llegar al último miércoles, que es el día 10.", "¿En qué día de la campaña estamos?")
    
    thisday = CInt(thisday)
    thisweek = CInt(thisweek)
    
    'si es el día 1 de la semana 1 o si el archivo no existe, que cree el archivo REPORTES_AGOTADOS usando el sub correspondiente
    If thisday = 1 And thisweek = 1 Then
        Call BorraHojas
        AAXX = InputBox("Por favor digite la campaña en curso en formato FY##", "¿En qué campaña estamos?")
        Call CrearHojaAgotados(AAXX)
    'en caso contrario, el archivo debe existir y debe abrirlo
    Else
        'Workbooks.Open "REPORTES_AGOTADOS"
        'Call ActivarHojas("REPORTES_AGOTADOS")
    End If
    'el nombre del archivo en cuestión será:
    filenom = Workbooks("REPORTES_AGOTADOS").Path & "\S" & thisweek & weekdays((thisday - 1) Mod 5) & ".xlsx"
    
    'a través de otra macro manda abrir dicho archivo, y si no existe, que marque error por inexistencia
    'que cargue a la hoja "Reporte Secreto" el reporte junto con la información pertinente
    Call AgregarReporteAHojaSecreta(filenom)
    'en cada hoja poner la tabla correspondiente, toda sacada de reporte secreto
        'aquí se designan los rangos correspondientes
            'Rango para agotados por división
        rangoDiv = Array("A1:A", "C1:C", "D1:D", "E1:E")
            'Rango para agotados por ruta
        rangoRuta = Array("A1:A", "C1:C", "F1:F", "D1:D")
        xLastRowReporte = Worksheets("Reporte Secreto").Cells(Rows.Count, 1).End(xlUp).Row
        
        If xLastRowReporte > 1 Then
        For indice = 0 To 3
            'hoja de división
            Worksheets("Reporte Secreto").Range(rangoDiv(indice) & xLastRowReporte).Copy
            Worksheets("AGOTADOS POR DIVISION").Activate
            Range(Cells(8, (thisday * 6) - 4 + indice), Cells(8 + xLastRowReporte, (thisday * 6) - 4 + indice)).Select
            Selection.PasteSpecial
            
            'hoja de división x día
            Worksheets("AGOTADOS POR DIVISION DIA").Activate
            Range(Cells(8, (thisday * 7) - 4 + indice), Cells(8 + xLastRowReporte, (thisday * 7) - 4 + indice)).Select
            Selection.PasteSpecial
            
            'hoja de ruta
            Worksheets("Reporte Secreto").Range(rangoRuta(indice) & xLastRowReporte).Copy
            Worksheets("AGOTADOS POR RUTA").Activate
            Range(Cells(8, (thisday * 6) - 4 + indice), Cells(8 + xLastRowReporte, (thisday * 6) - 4 + indice)).Select
            Selection.PasteSpecial
            Worksheets("AGOTADOS POR RUTA").Cells.EntireColumn.AutoFit
        Next indice
        
        'inserta los totales
            'agotados por ruta
                Dim totaldia As Long
                Worksheets("AGOTADOS POR RUTA").Activate
                Range(Cells(9, (thisday * 6) - 1), Cells(7 + xLastRowReporte, (thisday * 6) - 1)).Select
                totaldia = Application.WorksheetFunction.Sum(Selection)
                Cells(8 + xLastRowReporte, (thisday * 6) - 1).Activate
                ActiveCell.Value = totaldia
                ActiveCell.HorizontalAlignment = xlCenter
                ActiveCell.Offset(, -1).Value = "TOTAL"
                ActiveCell.Offset(, -1).Activate
                ActiveCell.HorizontalAlignment = xlRight
                ActiveCell.Font.Bold = True
                
            'agotados por división
                Worksheets("AGOTADOS POR DIVISION").Activate
                Range(Cells(9, (thisday * 6) - 2), Cells(7 + xLastRowReporte, (thisday * 6) - 2)).Select
                totaldia = Application.WorksheetFunction.Sum(Selection)
                Cells(8 + xLastRowReporte, (thisday * 6) - 2).Activate
                ActiveCell.Value = totaldia
                ActiveCell.HorizontalAlignment = xlCenter
                ActiveCell.Offset(, -1).Value = "TOTAL"
                ActiveCell.Offset(, -1).Activate
                ActiveCell.HorizontalAlignment = xlRight
                ActiveCell.Font.Bold = True
            
            'agotados por división día
                Worksheets("AGOTADOS POR DIVISION DIA").Activate
                Range(Cells(9, (thisday * 7) - 2), Cells(7 + xLastRowReporte, (thisday * 7) - 2)).Select
                totaldia = Application.WorksheetFunction.Sum(Selection)
                Cells(8 + xLastRowReporte, (thisday * 7) - 2).Activate
                ActiveCell.Value = totaldia
                ActiveCell.HorizontalAlignment = xlCenter
                ActiveCell.Offset(, -1).Value = "TOTAL"
                ActiveCell.Offset(, -1).Activate
                ActiveCell.HorizontalAlignment = xlRight
                ActiveCell.Font.Bold = True
                            
    'rellena la columna que indica S1 o S2 que hace falta en agotados división día:
    Worksheets("AGOTADOS POR DIVISION DIA").Activate
    Cells(8, (thisday * 7) - 5).Value = "SEMANA"
    
    Dim fila As Integer
    fila = 9
    
    Do While Not IsEmpty(Cells(fila, (thisday * 7) - 4))
        Cells(fila, (thisday * 7) - 5).Value = "S" & thisweek
        fila = fila + 1
    Loop
    
    Call LlenarReporteFinal_Agotados
    Call llenarSustituciones
    'Application.ScreenUpdating = True
    End If
End Sub
Sub BorraHojas()
    Dim xWs As Worksheet
    'Application.DisplayAlerts = False
    For Each xWs In Application.ActiveWorkbook.Worksheets
        If xWs.Name <> "PROCESOS" Then
            xWs.Delete
        End If
    Next
    'Application.DisplayAlerts = True
End Sub

Sub LlenarReporteFinal_Agotados()
    'da formato a la hoja
    
    Worksheets("AGOTADOS EN CAMPAÑA").Activate
    
    Range("8:1000").Select
    Selection.UnMerge
    Selection.ClearFormats
    Selection.ClearContents
    
    campos_producto = Array("ORIGEN (PRODUCTO)", "CATEGORÍA", "CLAVE", "DESCRIPCIÓN", "IMAGEN", "CANAL", "RAZÓN (CLAVE)", "RAZÓN", "ORIGEN")
    campaigndays = Array("JUEVES", "VIERNES", "LUNES", "MARTES", "MIÉRCOLES")
    campos_totales = Array("S1", "S2", "ACUMULADOS", "%")
    
    Range("B9:J9").Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    With Selection.Font
        .Name = "Calibri"
        .Bold = True
        .Size = 18
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(111, 75, 158)
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    ActiveCell.Value = "PRODUCTOS DE LÍNEA"
    
    Range("K9:O9").Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    With Selection.Font
        .Name = "Calibri"
        .Bold = True
        .Size = 18
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(111, 75, 158)
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    ActiveCell.Value = "SEMANA 1"
    
    Range("P9:T9").Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    With Selection.Font
        .Name = "Calibri"
        .Bold = True
        .Size = 18
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(111, 75, 158)
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    ActiveCell.Value = "SEMANA 2"
    
    Range("U9:X9").Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    With Selection.Font
        .Name = "Calibri"
        .Bold = True
        .Size = 18
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(111, 75, 158)
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    ActiveCell.Value = "TOTALES"
    
    For i = 0 To 8
        Cells(10, 2 + i).Activate
        ActiveCell.Value = campos_producto(i)
        ActiveCell.Interior.Color = RGB(195, 195, 225)
        ActiveCell.Font.Bold = True
        ActiveCell.Font.Size = 14
    Next i
    
    For i = 0 To 4
        Cells(10, 11 + i).Activate
        ActiveCell.Value = campaigndays(i)
        ActiveCell.Interior.Color = RGB(195, 195, 225)
        ActiveCell.Font.Bold = True
        ActiveCell.Font.Size = 14
    Next i
    
    For i = 0 To 4
        Cells(10, 16 + i).Activate
        ActiveCell.Value = campaigndays(i)
        ActiveCell.Interior.Color = RGB(195, 195, 225)
        ActiveCell.Font.Bold = True
        ActiveCell.Font.Size = 14
    Next i
    
    For i = 0 To 3
        Cells(10, 21 + i).Activate
        ActiveCell.Value = campos_totales(i)
        ActiveCell.Interior.Color = RGB(195, 195, 225)
        ActiveCell.Font.Bold = True
        ActiveCell.Font.Size = 14
    Next i
    
    'crea el arreglo de los agotados en total
    Dim ListaAgotados As Variant
    ListaAgotados = Array(1)
    
    'objeto tipo libro de excel para la base de datos
    Dim basedatos As Workbook
    Dim RutaBD As String
    RutaBD = Workbooks("REPORTES_AGOTADOS").Path & "\BD.xlsx"
    Set basedatos = Workbooks.Open(RutaBD)
    
    'inicializa índice del array de agotados
    Dim indice As Integer
    indice = 0
    Dim k As Integer
    
    'lee las columnas de las claves de los productos agotados y va agregando esas claves a ListaAgotados
    Workbooks("REPORTES_AGOTADOS").Worksheets("AGOTADOS POR RUTA").Activate
    
    With ActiveSheet
        'jala cada clave al arreglo de claves ListaAgotados
        For j = 1 To 10
            k = 0
            Do While Not IsEmpty(.Cells(9 + k, (6 * j) - 3))
                If UBound(Filter(ListaAgotados, .Cells(9 + k, (6 * j) - 3).Value)) = -1 Then
                ReDim Preserve ListaAgotados(indice)
                ListaAgotados(indice) = .Cells(9 + k, (6 * j) - 3).Value
                indice = indice + 1
            End If
                k = k + 1
            Loop
        Next j
    End With
    
    If IsEmpty(ListaAgotados) Then
        MsgBox "No hay productos agotados registrados."
    End If
    
    'Pone el apartado de HASTA AGOTAR EXISTENCIAS
    Worksheets("AGOTADOS EN CAMPAÑA").Activate
    Range(Cells(11 + UBound(ListaAgotados), 2), Cells(11 + UBound(ListaAgotados), 10)).Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    With Selection.Font
        .Name = "Calibri"
        .Bold = True
        .Size = 16
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(111, 75, 158)
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    ActiveCell.Value = "PRODUCTOS DESCONTINUADOS Y/O HASTA AGOTAR EXISTENCIAS"
    
    Range(Cells(11 + UBound(ListaAgotados), 11), Cells(11 + UBound(ListaAgotados), 15)).Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    With Selection.Font
        .Name = "Calibri"
        .Size = 18
        .Bold = True
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(111, 75, 158)
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    ActiveCell.Value = "SEMANA 1"
    
    Range(Cells(11 + UBound(ListaAgotados), 16), Cells(11 + UBound(ListaAgotados), 20)).Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    With Selection.Font
        .Name = "Calibri"
        .Bold = True
        .Size = 18
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(111, 75, 158)
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    ActiveCell.Value = "SEMANA 2"
    
    Range(Cells(11 + UBound(ListaAgotados), 21), Cells(11 + UBound(ListaAgotados), 24)).Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    With Selection.Font
        .Name = "Calibri"
        .Size = 18
        .Bold = True
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(111, 75, 158)
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    ActiveCell.Value = "TOTALES"

    For i = 0 To 8
        Cells(12 + UBound(ListaAgotados), 2 + i).Activate
        ActiveCell.Value = campos_producto(i)
        ActiveCell.Interior.Color = RGB(233, 199, 222)
        ActiveCell.Font.Bold = True
        ActiveCell.Font.Size = 14
    Next i
    
    For i = 0 To 4
        Cells(12 + UBound(ListaAgotados), 11 + i).Activate
        ActiveCell.Value = campaigndays(i)
        ActiveCell.Interior.Color = RGB(233, 199, 222)
        ActiveCell.Font.Bold = True
        ActiveCell.Font.Size = 14
    Next i
    
    For i = 0 To 4
        Cells(12 + UBound(ListaAgotados), 16 + i).Activate
        ActiveCell.Value = campaigndays(i)
        ActiveCell.Interior.Color = RGB(233, 199, 222)
        ActiveCell.Font.Bold = True
        ActiveCell.Font.Size = 14
    Next i
    
    For i = 0 To 3
        Cells(12 + UBound(ListaAgotados), 21 + i).Activate
        ActiveCell.Value = campos_totales(i)
        ActiveCell.Interior.Color = RGB(233, 199, 222)
        ActiveCell.Font.Bold = True
        ActiveCell.Font.Size = 14
    Next i
    
    
    'Pone las descripciones de los productos
    Dim ProductDescriptions As Variant
    ProductDescriptions = Array()
    If UBound(ListaAgotados) > 0 Then
        With basedatos.Worksheets(1)
            For i = LBound(ListaAgotados) To UBound(ListaAgotados)
                Set RangoBD = .Range("F:M")
                ReDim Preserve ProductDescriptions(i)
                ProductDescriptions(i) = Application.WorksheetFunction.VLookup(ListaAgotados(i), RangoBD, 4, False)
            Next i
                'Categoría de Productos
            Dim ProductCategory As Variant
            ProductCategory = Array()
            
            For i = LBound(ListaAgotados) To UBound(ListaAgotados)
                    ReDim Preserve ProductCategory(i)
                    ProductCategory(i) = Application.WorksheetFunction.VLookup(ListaAgotados(i), RangoBD, 5, False)
            Next i
            
        
            'Canal de Productos
            Dim ProductChannel As Variant
            ProductChannel = Array()
            
            For i = LBound(ListaAgotados) To UBound(ListaAgotados)
                    ReDim Preserve ProductChannel(i)
                    ProductChannel(i) = Application.WorksheetFunction.VLookup(ListaAgotados(i), RangoBD, 6, False)
            Next i
            
        
            'Status del Producto
            Dim ProductStatus As Variant
            ProductStatus = Array()
            
            For i = LBound(ListaAgotados) To UBound(ListaAgotados)
                    ReDim Preserve ProductStatus(i)
                    ProductStatus(i) = Application.WorksheetFunction.VLookup(ListaAgotados(i), RangoBD, 7, False)
            Next i
            
        
            'Origen del Producto
            Dim ProductOrigin As Variant
            ProductOrigin = Array()
            
            For i = LBound(ListaAgotados) To UBound(ListaAgotados)
                    ReDim Preserve ProductOrigin(i)
                    ProductOrigin(i) = Application.WorksheetFunction.VLookup(ListaAgotados(i), RangoBD, 8, False)
            Next i
            
            'SKU del Producto
            Dim productSKU As Variant
            productSKU = Array()
            
            For i = LBound(ListaAgotados) To UBound(ListaAgotados)
                    ReDim Preserve productSKU(i)
                    productSKU(i) = Application.WorksheetFunction.VLookup(ListaAgotados(i), RangoBD, 2, False)
            Next i
            
        End With
        
        
        'Imágenes
        
        'Dirección en el servidor de Amazon
        Dim ImgPath As String
        ImgPath = "REDACTED"
    
        Dim RutasImagenes As Variant
        RutasImagenes = Array()
        
        For i = LBound(ListaAgotados) To UBound(ListaAgotados)
            ReDim Preserve RutasImagenes(i)
            RutasImagenes(i) = ImgPath & ListaAgotados(i) & ".png"
        Next i
        
        'Bandera para confirmar que un arreglo dado está bien definido:
    '    For i = LBound(RutasImagenes) To UBound(RutasImagenes)
    '        msg = msg & RutasImagenes(i) & vbNewLine
    '    Next i
    '    MsgBox "Los productos agotados son: " & vbNewLine & msg
        
        basedatos.Close savechanges:=False
        
        Dim xLastRowRutas As Integer
        
        Worksheets("AGOTADOS EN CAMPAÑA").Activate
        
        'variable que recorre las celdas
        Dim retraso As Integer
        Dim atraso As Integer
        
        retraso = 0
        atraso = 0
        
        For i = LBound(ListaAgotados) To UBound(ListaAgotados)
            If (ProductStatus(i) = "XD" Or Not ProductChannel(i) = "Folleto") Then
                Cells(13 + UBound(ListaAgotados) + atraso, 2).Value = ProductOrigin(i)
                Cells(13 + UBound(ListaAgotados) + atraso, 3).Value = ProductCategory(i)
                Cells(13 + UBound(ListaAgotados) + atraso, 4).Value = ListaAgotados(i)
                Cells(13 + UBound(ListaAgotados) + atraso, 5).Value = ProductDescriptions(i)
                Cells(13 + UBound(ListaAgotados) + atraso, 6).Select
                ActiveCell.FormulaR1C1 = "=IMAGE(""https://s3./" & productSKU(i) & ".png"")"
                Cells(13 + UBound(ListaAgotados) + atraso, 7).Value = ProductChannel(i)
                Rows(ActiveCell.Row).RowHeight = 50
                atraso = atraso + 1
            Else
                Cells(11 + retraso, 2).Value = ProductOrigin(i)
                Cells(11 + retraso, 3).Value = ProductCategory(i)
                Cells(11 + retraso, 4).Value = ListaAgotados(i)
                Cells(11 + retraso, 5).Value = ProductDescriptions(i)
                Cells(11 + retraso, 6).Select
                ActiveCell.FormulaR1C1 = "=IMAGE(""https://s/" & productSKU(i) & ".png"")"
                Cells(11 + retraso, 7).Value = ProductChannel(i)
                Rows(ActiveCell.Row).RowHeight = 50
                retraso = retraso + 1
            End If
        Next i
        
        'retira las filas vacías
        i = 1
        Do While (IsEmpty(Cells(11 + retraso + i, 2)))
           i = i + 1
        Loop
        
        Dim marcador As Integer
        
    
        If (UBound(ListaAgotados) - retraso) > 2 Then
            Range(Rows(11 + retraso), Rows(8 + retraso + i)).Select
            Selection.Delete Shift:=xlUp
            marcador = 1
        End If
        
        'Sumas de totales productos de línea
        ColumnasClaves = Array("C", "I", "O", "U", "AA", "AG", "AM", "AS", "AY", "BE")
        ColumnasCantidades = Array("E", "K", "Q", "W", "AC", "AI", "AO", "AU", "BA", "BG")
        
        Dim sumatotal
        sumatotal = 0
        
        indice = 0
        Do While Not IsEmpty(Cells(11 + indice, 4))
            For i = 0 To 9
                xLastRowRutas = Worksheets("AGOTADOS POR RUTA").Cells(Rows.Count, (6 * i) + 2).End(xlUp).Row
                Worksheets("AGOTADOS EN CAMPAÑA").Cells(11 + indice, 4).Select
                Worksheets("AGOTADOS EN CAMPAÑA").Cells(11 + indice, 11 + i).Value = Application.WorksheetFunction.SumIfs(Worksheets("AGOTADOS POR RUTA").Range(ColumnasCantidades(i) & "9:" & ColumnasCantidades(i) & xLastRowRutas), Worksheets("AGOTADOS POR RUTA").Range(ColumnasClaves(i) & "9:" & ColumnasClaves(i) & xLastRowRutas), Selection.Value)
            Next i
            
                Worksheets("AGOTADOS EN CAMPAÑA").Cells(11 + indice, 21).Formula = Application.WorksheetFunction.Sum(Range(Cells(11 + indice, 11), Cells(11 + indice, 15)))
                Worksheets("AGOTADOS EN CAMPAÑA").Cells(11 + indice, 22).Formula = Application.WorksheetFunction.Sum(Range(Cells(11 + indice, 16), Cells(11 + indice, 20)))
                Worksheets("AGOTADOS EN CAMPAÑA").Cells(11 + indice, 23).Formula = Application.WorksheetFunction.Sum(Range(Cells(11 + indice, 21), Cells(11 + indice, 22)))
                sumatotal = sumatotal + Worksheets("AGOTADOS EN CAMPAÑA").Cells(11 + indice, 23).Value
            
            indice = indice + 1
        Loop
        
        Worksheets("AGOTADOS EN CAMPAÑA").Cells(11 + indice, 10).Activate
        ActiveCell.Value = "TOTAL:"
        With ActiveCell.Font
            .Name = "Calibri"
            .Bold = True
            .Size = 14
        End With
        
        For j = 1 To 13
            ActiveCell.Offset(, j).Formula = Application.WorksheetFunction.Sum(Range(Cells(11, 10 + j), Cells(10 + indice, 10 + j)))
        Next j
        
        'suma de totales productos de otros canales de inventarios
        
        If marcador = 1 Then
            indice = indice + 4
        Else
            indice = indice + 3
        End If
        
        Do While Not IsEmpty(Cells(11 + indice, 4))
            For i = 0 To 9
                xLastRowRutas = Worksheets("AGOTADOS POR RUTA").Cells(Rows.Count, (6 * i) + 2).End(xlUp).Row
                Worksheets("AGOTADOS EN CAMPAÑA").Cells(11 + indice, 4).Select
                Worksheets("AGOTADOS EN CAMPAÑA").Cells(11 + indice, 11 + i).Value = Application.WorksheetFunction.SumIfs(Worksheets("AGOTADOS POR RUTA").Range(ColumnasCantidades(i) & "9:" & ColumnasCantidades(i) & xLastRowRutas), Worksheets("AGOTADOS POR RUTA").Range(ColumnasClaves(i) & "9:" & ColumnasClaves(i) & xLastRowRutas), Selection.Value)
            Next i
            
                Worksheets("AGOTADOS EN CAMPAÑA").Cells(11 + indice, 21).Formula = Application.WorksheetFunction.Sum(Range(Cells(11 + indice, 11), Cells(11 + indice, 15)))
                Worksheets("AGOTADOS EN CAMPAÑA").Cells(11 + indice, 22).Formula = Application.WorksheetFunction.Sum(Range(Cells(11 + indice, 16), Cells(11 + indice, 20)))
                Worksheets("AGOTADOS EN CAMPAÑA").Cells(11 + indice, 23).Formula = Application.WorksheetFunction.Sum(Range(Cells(11 + indice, 21), Cells(11 + indice, 22)))
                sumatotal = sumatotal + Worksheets("AGOTADOS EN CAMPAÑA").Cells(11 + indice, 23).Value
            
            indice = indice + 1
        Loop
        
        Worksheets("AGOTADOS EN CAMPAÑA").Cells(11 + indice, 10).Activate
        ActiveCell.Value = "TOTAL:"
        With ActiveCell.Font
            .Name = "Calibri"
            .Bold = True
            .Size = 14
        End With
        
        For j = 1 To 13
            If marcador = 1 Then
                ActiveCell.Offset(, j).Formula = Application.WorksheetFunction.Sum(Range(Cells(11 + retraso + 4, 10 + j), Cells(10 + indice, 10 + j)))
            Else
                ActiveCell.Offset(, j).Formula = Application.WorksheetFunction.Sum(Range(Cells(11 + retraso + 3, 10 + j), Cells(10 + indice, 10 + j)))
            End If
        Next j
        
        Range(ActiveCell.Offset(3, -1), ActiveCell.Offset(3, 11)).Select
        Selection.Merge
        With Selection
            .HorizontalAlignment = xlCenter
        End With
            With Selection.Font
            .Name = "Calibri"
            .Size = 18
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = RGB(241, 88, 93)
                .TintAndShade = 0
                .PatternTintAndShade = 0
        End With
        ActiveCell.Value = "PEDIDOS"
        
        ActiveCell.Offset(1).Activate
        
        For i = 0 To 1
            For j = 0 To 4
                ActiveCell.Offset(, 1).Activate
                ActiveCell.Value = campaigndays(j)
                ActiveCell.Interior.Color = RGB(249, 192, 187)
                ActiveCell.Offset(9).Value = campaigndays(j)
                ActiveCell.Offset(9).Interior.Color = RGB(249, 192, 187)
            Next j
        Next i
        
        ActiveCell.Offset(, 1).Activate
        ActiveCell.Value = "ACUMULADOS"
        ActiveCell.Interior.Color = RGB(249, 192, 187)
        ActiveCell.Offset(9).Value = "ACUMULADOS"
        ActiveCell.Offset(9).Interior.Color = RGB(249, 192, 187)
        ActiveCell.Offset(, 1).Activate
        ActiveCell.Value = "ACUMULADOS FY"
        ActiveCell.Interior.Color = RGB(249, 192, 187)
        
        ActiveCell.Offset(1, -12).Activate
        
        pedidos1 = Array("PEDIDOS SURTIDOS", "UNIDADES FACTURADAS", "UNIDADES/PEDIDO FACTURADAS", "UNIDADES SURTIDAS", "UNIDADES/PEDIDOS SURTIDAS", "ESTIMADO CAMPAÑA NACIONAL REAL", "TOTAL")
        pedidos2 = Array("PEDIDOS AFECTADOS AL DÍA", "PEDIDOS AFECTADOS CON HE", "PEDIDOS AFECTADOS CON SV", "PEDIDOS AFECTADOS (SIN CONSIDERAR SV Y XD)", "NIVEL DE SERVICIO", "NIVEL DE SERVICIO (SIN CONSIDERAR SV Y XD)", "NIVEL DE SERVICIO (SIN DESCONTINUADOS)", "RENUEVE/VENTA ($)")
        
        For Each aspecto In pedidos1
            ActiveCell.HorizontalAlignment = xlRight
            ActiveCell.Value = aspecto
            ActiveCell.Interior.Color = RGB(241, 88, 93)
            ActiveCell.Font.ThemeColor = xlThemeColorDark1
            ActiveCell.Font.TintAndShade = 0
            ActiveCell.Offset(1).Activate
        Next aspecto
        
        ActiveCell.Offset(2).Activate
        
        For Each aspecto In pedidos2
            ActiveCell.HorizontalAlignment = xlRight
            ActiveCell.Value = aspecto
            ActiveCell.Interior.Color = RGB(241, 88, 93)
            ActiveCell.Font.ThemeColor = xlThemeColorDark1
            ActiveCell.Font.TintAndShade = 0
            ActiveCell.Offset(1).Activate
        Next aspecto
        
        ActiveCell.Select
        
        Cells(1, 1).Activate
            With ActiveCell.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = RGB(239, 51, 64)
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With ActiveCell.Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        
        Cells(2, 1).Activate
            With ActiveCell.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = RGB(239, 51, 64)
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With ActiveCell.Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        
        Cells(3, 1).Activate
            With ActiveCell.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = RGB(239, 51, 64)
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With ActiveCell.Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
            
        If retraso > 0 Then
            retraso = UBound(ListaAgotados) - atraso
        Else
            retraso = 0
        End If
        
        Range("X10").Activate
        
        If retraso > 0 Then
            For i = 0 To (retraso + 1)
                ActiveCell.Offset(1).Activate
                ActiveCell.Value = (ActiveCell.Offset(, -1).Value / sumatotal) * 100
                ActiveCell.NumberFormat = "00.0"
            Next i
        Else
            ActiveCell.Offset(1).Activate
            ActiveCell.Value = (ActiveCell.Offset(, -1).Value / sumatotal) * 100
            ActiveCell.NumberFormat = "00.0"
        End If
        
        If marcador = 1 Then
            ActiveCell.Offset(3).Activate
        Else
            ActiveCell.Offset(2).Activate
        End If
        
        For i = 0 To atraso
            ActiveCell.Offset(1).Activate
            ActiveCell.Value = (ActiveCell.Offset(, -1).Value / sumatotal) * 100
            ActiveCell.NumberFormat = "00.0"
        Next i
        
        'Ajusta el Tamaño de las celdas de forma adecuada
        
        Worksheets("AGOTADOS EN CAMPAÑA").Cells.Select
        
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        Worksheets("AGOTADOS EN CAMPAÑA").Cells.EntireColumn.AutoFit
        'Worksheets("AGOTADOS EN CAMPAÑA").Cells.EntireRow.AutoFit
        
    '    Dim archivoagotados As Variant
    '
    '
    '    Worksheets("AGOTADOS EN CAMPAÑA").Copy
    '
    '    archivoagotados = Application.GetSaveAsFilename(Title:="Por favor seleccione el directorio de almacenamiento", InitialFileName:="Reporte Final de Agotados " & Format(Date, "yyyy-mm-dd"))
    '
    '    ActiveWorkbook.Close savechanges = True
    End If
End Sub

Sub llenarSustituciones()
    Worksheets("SUSTITUCIONES").Activate
    
    'Range("8:1000").Select
    'Selection.UnMerge
    'Selection.ClearFormats
    'Selection.ClearContents
    
    campos_producto = Array("CATEGORÍA", "SUSTITUCIÓN", "ORIGINAL", "DESCRIPCIÓN", "ORIGEN", "RAZÓN (CLAVE)")
    campaigndays = Array("JUEVES", "VIERNES", "LUNES", "MARTES", "MIÉRCOLES")
    campos_totales = Array("S1", "S2", "ACUMULADOS", "%")
    
    Range("B9:G9").Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    With Selection.Font
        .Name = "Calibri"
        .Bold = True
        .Size = 18
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(111, 75, 158)
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    ActiveCell.Value = "SUSTITUCIONES"
    
    Range("H9:L9").Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    With Selection.Font
        .Name = "Calibri"
        .Bold = True
        .Size = 18
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(111, 75, 158)
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    ActiveCell.Value = "SEMANA 1"
    
    Range("M9:Q9").Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    With Selection.Font
        .Name = "Calibri"
        .Bold = True
        .Size = 18
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(111, 75, 158)
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    ActiveCell.Value = "SEMANA 2"
    
    Range("R9:U9").Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    With Selection.Font
        .Name = "Calibri"
        .Bold = True
        .Size = 18
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(111, 75, 158)
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    ActiveCell.Value = "TOTALES"
    
    For i = 0 To 5
        Cells(10, 2 + i).Activate
        ActiveCell.Value = campos_producto(i)
        ActiveCell.Interior.Color = RGB(195, 195, 225)
        ActiveCell.Font.Bold = True
        ActiveCell.Font.Size = 14
    Next i
    
    For i = 0 To 4
        Cells(10, 8 + i).Activate
        ActiveCell.Value = campaigndays(i)
        ActiveCell.Interior.Color = RGB(195, 195, 225)
        ActiveCell.Font.Bold = True
        ActiveCell.Font.Size = 14
    Next i
    
    For i = 0 To 4
        Cells(10, 13 + i).Activate
        ActiveCell.Value = campaigndays(i)
        ActiveCell.Interior.Color = RGB(195, 195, 225)
        ActiveCell.Font.Bold = True
        ActiveCell.Font.Size = 14
    Next i
    
    For i = 0 To 3
        Cells(10, 18 + i).Activate
        ActiveCell.Value = campos_totales(i)
        ActiveCell.Interior.Color = RGB(195, 195, 225)
        ActiveCell.Font.Bold = True
        ActiveCell.Font.Size = 14
    Next i
    
    'SUSTITUCIONES DE TE LLEGARÁ CUALQUIERA
    Range("B25:G25").Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    With Selection.Font
        .Name = "Calibri"
        .Bold = True
        .Size = 18
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(111, 75, 158)
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    ActiveCell.Value = "SUSTITUCIONES DE TE LLEGARÁ CUALQUIERA"
    
    Range("H25:L25").Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    With Selection.Font
        .Name = "Calibri"
        .Bold = True
        .Size = 18
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(111, 75, 158)
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    ActiveCell.Value = "SEMANA 1"
    
    Range("M25:Q25").Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    With Selection.Font
        .Name = "Calibri"
        .Bold = True
        .Size = 18
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(111, 75, 158)
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    ActiveCell.Value = "SEMANA 2"
    
    Range("R25:U25").Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    With Selection.Font
        .Name = "Calibri"
        .Bold = True
        .Size = 18
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(111, 75, 158)
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    ActiveCell.Value = "TOTALES"
    
    For i = 0 To 5
        Cells(26, 2 + i).Activate
        ActiveCell.Value = campos_producto(i)
        ActiveCell.Interior.Color = RGB(195, 195, 225)
        ActiveCell.Font.Bold = True
        ActiveCell.Font.Size = 14
    Next i
    
    For i = 0 To 4
        Cells(26, 8 + i).Activate
        ActiveCell.Value = campaigndays(i)
        ActiveCell.Interior.Color = RGB(195, 195, 225)
        ActiveCell.Font.Bold = True
        ActiveCell.Font.Size = 14
    Next i
    
    For i = 0 To 4
        Cells(26, 13 + i).Activate
        ActiveCell.Value = campaigndays(i)
        ActiveCell.Interior.Color = RGB(195, 195, 225)
        ActiveCell.Font.Bold = True
        ActiveCell.Font.Size = 14
    Next i
    
    For i = 0 To 3
        Cells(26, 18 + i).Activate
        ActiveCell.Value = campos_totales(i)
        ActiveCell.Interior.Color = RGB(195, 195, 225)
        ActiveCell.Font.Bold = True
        ActiveCell.Font.Size = 14
    Next i
    
    Worksheets("SUSTITUCIONES").Cells.Select
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    'SUSTITUCIONES DE MERCADOTECNIA
    Range("B38:G38").Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    With Selection.Font
        .Name = "Calibri"
        .Bold = True
        .Size = 18
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(111, 75, 158)
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    ActiveCell.Value = "SUSTITUCIONES DE TE LLEGARÁ CUALQUIERA"
    
    Range("H38:L38").Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    With Selection.Font
        .Name = "Calibri"
        .Bold = True
        .Size = 18
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(111, 75, 158)
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    ActiveCell.Value = "SEMANA 1"
    
    Range("M38:Q38").Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    With Selection.Font
        .Name = "Calibri"
        .Bold = True
        .Size = 18
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(111, 75, 158)
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    ActiveCell.Value = "SEMANA 2"
    
    Range("R38:U38").Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    With Selection.Font
        .Name = "Calibri"
        .Bold = True
        .Size = 18
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(111, 75, 158)
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    ActiveCell.Value = "TOTALES"
    
    For i = 0 To 5
        Cells(39, 2 + i).Activate
        ActiveCell.Value = campos_producto(i)
        ActiveCell.Interior.Color = RGB(195, 195, 225)
        ActiveCell.Font.Bold = True
        ActiveCell.Font.Size = 14
    Next i
    
    For i = 0 To 4
        Cells(39, 8 + i).Activate
        ActiveCell.Value = campaigndays(i)
        ActiveCell.Interior.Color = RGB(195, 195, 225)
        ActiveCell.Font.Bold = True
        ActiveCell.Font.Size = 14
    Next i
    
    For i = 0 To 4
        Cells(39, 13 + i).Activate
        ActiveCell.Value = campaigndays(i)
        ActiveCell.Interior.Color = RGB(195, 195, 225)
        ActiveCell.Font.Bold = True
        ActiveCell.Font.Size = 14
    Next i
    
    For i = 0 To 3
        Cells(39, 18 + i).Activate
        ActiveCell.Value = campos_totales(i)
        ActiveCell.Interior.Color = RGB(195, 195, 225)
        ActiveCell.Font.Bold = True
        ActiveCell.Font.Size = 14
    Next i
    
    Worksheets("SUSTITUCIONES").Cells.Select
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    
    Range("G20").Activate
    ActiveCell.Value = "TOTALES"
    ActiveCell.Font.Bold = True
    ActiveCell.Font.Size = 14
    ActiveCell.HorizontalAlignment = xlRight
    
    Range("G35").Activate
    ActiveCell.Value = "TOTALES"
    ActiveCell.Font.Bold = True
    ActiveCell.Font.Size = 14
    ActiveCell.HorizontalAlignment = xlRight
    
    Range("G46").Activate
    ActiveCell.Value = "TOTALES"
    ActiveCell.Font.Bold = True
    ActiveCell.Font.Size = 14
    ActiveCell.HorizontalAlignment = xlRight
    
    Worksheets("SUSTITUCIONES").Cells.EntireColumn.AutoFit
    Worksheets("SUSTITUCIONES").Cells.EntireRow.AutoFit
End Sub
