Attribute VB_Name = "Tasa_Real"
Dim nom_hoja As String, fecha_texto As String, año As String, nom_hoja2 As String
Dim rango_campos As Range, celda As Range
Dim conteo As Long, n As Long


Sub tasa_servicio_acumulada()

mes = Month(Date) - 1
If mes = 0 Then
    año = Year(Date) - 1
Else
    año = Year(Date)
End If
Select Case mes
    Case 0: mes = "Diciembre"
    Case 1: mes = "Enero"
    Case 2: mes = "Febrero"
    Case 3: mes = "Marzo"
    Case 4: mes = "Abril"
    Case 5: mes = "Mayo"
    Case 6: mes = "Junio"
    Case 7: mes = "Julio"
    Case 8: mes = "Agosto"
    Case 9: mes = "Septiembre"
    Case 10: mes = "Octubre"
    Case 11: mes = "Noviembre"
    Case 12: mes = "Diciembre"
End Select
año_consulta = InputBox("Introduce año a consultar", "Año Consulta", año)
Path = "\\vmedsis03\Suministros\Indicadores Compras\" & año_consulta & "\"
archivo = "tasa_servicio(" & año_consulta & ").xlsx"

If Dir(Path, vbDirectory) <> "" Then
    If Dir(Path & archivo, vbNormal) = "" Then
tasa_real:
        ' Trae la BD filtrada por la fecha indtroducida y la prepara para pasar la info al archivo ppal
        Workbooks.OpenXML ("\\vmedsis03\Suministros\Plantillas\FICHEROS\indicadores_entregas.xls")
        With Workbooks("indicadores_entregas.xls").Sheets(1)
            .Rows("1:3").delete
            .Rows("2").delete
            .Columns("A:B").delete
            .Columns("K:L").delete
            .Columns("M:S").delete
        End With
        Range("N2", Range("N2").End(xlDown)).Select
        Range("O1").value = "Mes"
        Range("P1").value = "Año"
        For Each celda In Selection
            fecha_texto = celda.value
            mes_entrega = Month(fecha_texto)
            año = Year(fecha_texto)
            celda.Offset(0, 1).value = mes_entrega
            celda.Offset(0, 2).value = año
        Next
        ActiveWorkbook.Sheets.Add
        ActiveSheet.Name = "2"
        ActiveSheet.Range("A1").value = "Año"
        ActiveSheet.Range("A2").value = año_consulta
        ActiveWorkbook.Sheets.Add
        ActiveSheet.Name = "3"
        Sheets("indicadores_entregas").Range("A1:P999999").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Sheets("2").Range("A1:B2"), _
            CopyToRange:=ActiveWorkbook.Sheets("3").Range("A1"), Unique:=False
        ' Copia y pega en la BDatos el archivo ppal y corrige el formato de las fechas (entrega y Migo)
        Application.DisplayAlerts = False
        Workbooks("indicadores_entregas.xls").Sheets("3").Range("A2:N2", Range("A2:N2").End(xlDown)).Copy
        Workbooks.OpenXML ("C:\Documentos Empresa\OneDrive - MINEROS\Desktop\Automatizaciones\formatos\tasa_real.xlsx")
        Workbooks("tasa_real.xlsx").Sheets("BD").Range("A2").PasteSpecial
        Workbooks("indicadores_entregas.xls").Close
        Application.DisplayAlerts = True
        ' Crea tabla dinámica (consolida las OC)
        Sheets.Add
        nom_hoja = ActiveSheet.Name
        ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Tabla1", Version:=xlPivotTableVersion15).CreatePivotTable TableDestination:= _
            Sheets(nom_hoja).Cells(2, 1), TableName:="Tabla dinámica2", DefaultVersion:=xlPivotTableVersion15
        ' Agrega campos en tabla
        With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Nombre Proveedor")
            .Orientation = xlRowField
            .Position = 1
            .LayoutForm = xlTabular
            .RepeatLabels = True
            .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        End With
        With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Proveedor")
            .Orientation = xlRowField
            .Position = 2
            .LayoutForm = xlTabular
            .RepeatLabels = True
            .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        End With
        With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("OC UNIFICADA")
            .Orientation = xlRowField
            .Position = 3
            .LayoutForm = xlTabular
            .RepeatLabels = True
            .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        End With
        With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Mes")
            .Orientation = xlRowField
            .Position = 4
        End With
        With ActiveSheet.PivotTables("Tabla dinámica2")
            .AddDataField ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Cumple"), "OC a Tiempo", xlSum
            .AddDataField ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Entrega"), "OC Entregadas", xlSum
        End With
        ' Copia en formato plano la TD generada
        Range("A2:D2", Range("A2:D2").End(xlDown)).Copy
        With Range("H2")
            .PasteSpecial Paste:=xlPasteValues
            .value = "Nombre Proveedor"
            .End(xlDown).ClearContents
        End With
        ' Condicional suma cumple
        Cells(2, 12).value = "OC a Tiempo"
        Cells(2, 13).value = "OC Entregadas"
        conteo = Range("J3", Range("J3").End(xlDown)).Count
        n = 3
        For i = 1 To conteo Step 1
            If Cells(n, 5) < 1 Then Cells(n, 12).value = 0 Else Cells(n, 12).value = 1
            If Cells(n, 6) < 1 Then Cells(n, 13).value = 0 Else Cells(n, 13).value = 1
            n = n + 1
        Next
        n = 3
        ' TD Tasa De Servicio
        Set rango_campos = Range("H2:M2", Range("H2:M2").End(xlDown))
        Sheets.Add
        ActiveSheet.Name = "ts_mes"
        nom_hoja2 = ActiveSheet.Name
        ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rango_campos, Version:=xlPivotTableVersion15).CreatePivotTable TableDestination:= _
            Sheets(nom_hoja2).Cells(2, 1), TableName:="Tabla dinámica2", DefaultVersion:=xlPivotTableVersion15
        ' Agrega campos
        With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Nombre Proveedor")
            .Orientation = xlRowField
            .Position = 1
            .LayoutForm = xlTabular
            .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        End With
        With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Mes")
            .Orientation = xlRowField
            .Position = 2
        End With
        With ActiveSheet.PivotTables("Tabla dinámica2")
            .AddDataField ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("OC a Tiempo"), "Cuenta de OC a Tiempo", xlSum
            .AddDataField ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("OC Entregadas"), "Cuenta de OC Entregadas", xlSum
        End With
        ' Campo calculado para %
        With ActiveSheet.PivotTables("Tabla dinámica2")
            .CalculatedFields.Add "%", "='OC a Tiempo' /'OC Entregadas'", True
            .PivotFields("%").Orientation = xlDataField
            .PivotFields("Suma de %").NumberFormat = "0%"
        End With
        Application.DisplayAlerts = False
        Sheets(nom_hoja).Visible = xlHidden
        ' Crea segmentador
        Workbooks("tasa_real.xlsx").Sheets("ts_mes").Activate
        ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("Tabla dinámica2"), "Nombre Proveedor").Slicers.Add ActiveSheet, , "Nombre Proveedor", _
            "Nombre Proveedor", 79.5, 334.5, 144, 198.75
        ActiveSheet.Shapes.Range(Array("Nombre Proveedor")).Select
        ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_Nombre_Proveedor").Slicers("Nombre Proveedor").NumberOfColumns = 2
        With ActiveSheet.Shapes("Nombre Proveedor")
            .Top = 10
            .Left = 500
            .Width = 450
            .Height = 365
        End With
        ' Crea gráfico
        ActiveSheet.Shapes.AddChart2(297, xlColumnClustered).Select
        With ActiveChart
            .SetSourceData Source:=Range("'ts_mes'!$A$2:$E$513")
            .ShowAllFieldButtons = False
            .FullSeriesCollection(1).ChartType = xlLine
            .FullSeriesCollection(1).AxisGroup = 2
            .FullSeriesCollection(2).ChartType = xlLine
            .FullSeriesCollection(2).AxisGroup = 2
            .FullSeriesCollection(3).ChartType = xlColumnClustered
            .FullSeriesCollection(3).AxisGroup = 1
            .Legend.delete
            .Location Where:=xlLocationAsNewSheet
        End With
        ' Desaparecer Líneas
        ActiveChart.ChartArea.Select
        ActiveChart.FullSeriesCollection(1).Select
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 1
        End With
        ActiveChart.FullSeriesCollection(2).Select
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 1
        End With
        ' Tabla y datos
        With ActiveChart
            .SetElement (msoElementDataTableWithLegendKeys)
            .SetElement (msoElementDataLabelNone)
            .SetElement (msoElementChartTitleAboveChart)
            .Axes(xlValue).MaximumScale = 1
            .ChartTitle.Text = "Tasa De Servicio"
            .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 18
            .ChartArea.Format.TextFrame2.TextRange.Font.Bold = msoTrue
            .ChartArea.Format.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorText1
        End With
        Workbooks("tasa_real.xlsx").Sheets("ts_mes").Activate
        With ActiveSheet.PivotTables("Tabla dinámica2").DataPivotField
            .PivotItems("Cuenta de OC a Tiempo").Caption = "Entregas a Tiempo"
            .PivotItems("Cuenta de OC Entregadas").Caption = "Entregas Totales"
            .PivotItems("Suma de %").Caption = "TS"
        End With
        Workbooks.SaveAs (Path & archivo)
        Application.DisplayAlerts = True
    ElseIf Dir(Path & archivo, vbNormal) <> "" And mes <> "Diciembre" Then
        GoTo tasa_real
    Else
        Call Shell("explorer.exe " & Path & archivo, vbNormalFocus)
    End If
End If
End Sub

