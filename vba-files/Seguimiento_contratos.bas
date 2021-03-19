Attribute VB_Name = "Seguimiento_contratos"

Dim valor As Date, celda As Range
Dim fecha As String, fecha_año As String
Dim fecha_texto As String, mes As Integer, año As Integer, fecha_actual As Date
Dim hoja As String, conteo_click As Integer, hoja_tabla As String
Dim fecha_reporte As String
Dim porcentaje As Double
Dim fila As Integer, limite As Integer, i As Integer, limite_conteo As Integer
Sub informe_gantt()

' Abre BD y crea filtros
Workbooks.Open ("\\vmedsis03\Suministros\Plantillas\FICHEROS\contratos.xlsx")
Workbooks("contratos.xlsx").Sheets(1).Range("U1").Value = "Activo"
Workbooks("contratos.xlsx").Sheets(1).Range("J2", Range("J2").End(xlDown)).Select
fecha_actual = Date
For Each celda In Selection
  If celda >= fecha_actual Then
      celda.Offset(0, 11).Value = 1
  Else
      celda.Offset(0, 11).Value = 0
  End If
Next
ActiveWorkbook.Sheets.Add
ActiveSheet.Name = "2"
ActiveSheet.Range("A1").Value = "Activo"
ActiveSheet.Range("A2").Value = 1
ActiveWorkbook.Sheets.Add
ActiveSheet.Name = "3"
Sheets("Sheet1").Range("A1:U999999").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Sheets("2").Range("A1:A2"), _
CopyToRange:=ActiveWorkbook.Sheets("3").Range("A1"), Unique:=False
With Workbooks("contratos.xlsx").Sheets("3")
    .Columns("B:G").Delete
    .Columns("C").Delete
    .Columns("D").Delete
    .Columns("G:M").Delete
End With

' Por Monto (20%)
Sheets("3").Range("F2", Range("F2").End(xlDown)).Select
For Each celda In Selection
    On Error Resume Next
    porcentaje = celda.Value / celda.Offset(0, -1).Value
    If porcentaje <= 0.2 Then
        celda.Offset(0, 1).Value = 1
    Else
        celda.Offset(0, 1).Value = 0
    End If
Next

' Calcula días
Sheets("3").Range("C2", Range("C2").End(xlDown)).Select
For Each celda In Selection
  celda.Offset(0, 5).Value = celda.Value - Date
Next

' Por Fecha (90 dias)
Sheets("3").Range("H2", Range("H2").End(xlDown)).Select
For Each celda In Selection
  If celda.Value <= 90 Then
      celda.Offset(0, 1).Value = 1
  Else
      celda.Offset(0, 1).Value = 0
  End If
Next
      
' Valida fecha y monto
Sheets("3").Range("I2", Range("I2").End(xlDown)).Select
For Each celda In Selection
  celda.Offset(0, 1).Value = celda.Value + celda.Offset(0, -2).Value
Next
Range("J1").Value = "Validacion"
ActiveWorkbook.Sheets.Add
ActiveSheet.Name = "4"
ActiveSheet.Range("A1").Value = "Validacion"
ActiveSheet.Range("A2").Value = 1
ActiveSheet.Range("A3").Value = 2
ActiveWorkbook.Sheets.Add
ActiveSheet.Name = "5"
Sheets("3").Range("A1:J10000").AdvancedFilter Action:=xlFilterCopy, _
  CriteriaRange:=Sheets("4").Range("A1:A3"), CopyToRange:=ActiveWorkbook.Sheets("5").Range("A1"), _
  Unique:=False
Sheets("5").Columns("E").Delete

' Copia datos del archivo descargado y pega en archivo de macro
Range("A1").CurrentRegion.Copy
Workbooks("macro_contratos.xlsm").Worksheets.Add(Before:=Worksheets(1)).Name = "Tabla"
Workbooks("macro_contratos.xlsm").Sheets("Tabla").Range("A1").PasteSpecial
Application.DisplayAlerts = False
Workbooks("contratos.xlsx").Close
Application.DisplayAlerts = True

' Crea los datos para generar la tabla, esta ok
Workbooks("macro_contratos.xlsm").Sheets("Tabla").Columns("C").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
limite_conteo = Workbooks("macro_contratos.xlsm").Sheets("Tabla").Range("A1", Range("A1").End(xlDown)).Count
If limite_conteo <= 2 Then
    Workbooks("macro_contratos.xlsm").Sheets("Tabla").Range("C2").Value = Sheets("Tabla").Range("A2").Value & "//" & Sheets("Tabla").Range("B2").Value
Else
    limite = Sheets("Tabla").Range("A2", Range("A2").End(xlDown)).Count
    fila = 2
    For i = 0 To limite - 1
        Sheets("Tabla").Range("C" & fila).Value = Sheets("Tabla").Range("A" & fila).Value & "//" & Sheets("Tabla").Range("B" & fila).Value
        fila = fila + 1
    Next
End If

' Organiza las fechas en orden descendente
ActiveSheet.Range("A1").AutoFilter
Worksheets("Tabla").AutoFilter.Sort.SortFields.Clear
Worksheets("Tabla").AutoFilter.Sort.SortFields.Add Key:=Range("D1:D100000"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("Tabla").AutoFilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
Sheets("Tabla").Range("C1").Value = "Contrato"
Sheets("Tabla").Columns("A:B").Delete

' Crea la tabla en forma de gantt
limite_conteo = Sheets("Tabla").Range("A1", Range("A1").End(xlDown)).Count
If limite_conteo <= 2 Then
    ActiveSheet.Shapes.AddChart2(216, xlBarClustered).Select
    With ActiveChart
        .SeriesCollection.NewSeries
        .FullSeriesCollection(1).Name = "Fecha Fin Contrato"
        .FullSeriesCollection(1).Values = Sheets("Tabla").Range("B2")
        .FullSeriesCollection(1).XValues = Sheets("Tabla").Range("A2")
        .Location xlLocationAsNewSheet, "Gráfico"
    End With
Else
    ActiveSheet.Shapes.AddChart2(216, xlBarClustered).Select
    With ActiveChart
        .SeriesCollection.NewSeries
        .FullSeriesCollection(1).Name = "Fecha Fin Contrato"
        .FullSeriesCollection(1).Values = Sheets("Tabla").Range("B2", Range("B2").End(xlDown))
        .FullSeriesCollection(1).XValues = Sheets("Tabla").Range("A2", Range("A2").End(xlDown))
        .Location xlLocationAsNewSheet, "Gráfico"
    End With
End If
ActiveChart.ApplyDataLabels Type:=xlDataLabelsShowValue
ActiveChart.HasTitle = True
ActiveChart.ChartTitle.Text = "Fecha Finalización Contrato"
End Sub
Sub prueba2()

' buscar crear un filtro para dear solo lo que esté en en mes a cosnultar
Dim fecha_año As Integer, fecha_mes As Integer, celda As Range, valor As Date, fecha_total As Date

    fecha_año = InputBox("Introducir año")
    fecha_mes = InputBox("Introducir mes")
    
    For Each celda In Selection
    
        If fecha_mes = 4 Or fecha_mes = 6 Or fecha_mes = 9 Or fecha_mes = 11 Then
            
            fecha_total = "30" & "/" & fecha_mes & "/" & fecha_año
            
            valor = celda.Value
            
            If valor >= fecha_total Then
                celda.ClearContents
            ElseIf valor <= fecha_total Then
                celda.ClearContents
            End If
            
        End If
        
    Next
    
        For Each celda In Selection

            valor = celda.Value

            If valor >= fecha Then
                celda.ClearContents
            ElseIf valor <= fecha Then
                celda.ClearContents
            End If
            
        Next

End Sub
Sub Macro2()
ActiveWindow.DisplayWorkbookTabs = False 'Oculta las fichas de las hohas
ActiveWindow.DisplayHeadings = False 'Oculta títulos
Application.DisplayFormulaBar = False 'Oculta la barra de formulas
ActiveWindow.DisplayGridlines = False 'Oculta las lineas de la cuadricula
Application.DisplayStatusBar = False 'Oculta la barra de estado
Application.DisplayFullScreen = True 'Ves pantalla completa
End Sub

Sub cerrar()
ActiveWindow.DisplayWorkbookTabs = True 'Oculta las fichas de las hohas
'ActiveWindow.DisplayHeadings = True 'Oculta títulos
Application.DisplayFormulaBar = True 'Oculta la barra de formulas
'ActiveWindow.DisplayGridlines = True 'Oculta las lineas de la cuadricula
Application.DisplayStatusBar = True 'Oculta la barra de estado
Application.DisplayFullScreen = False 'Ves pantalla completa

End Sub


