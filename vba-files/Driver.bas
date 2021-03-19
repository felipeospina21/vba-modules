Attribute VB_Name = "Driver"
Dim filas As Integer
Dim libro As String
Dim total As Range, limite As Integer, fila As Long, numerador As Range, ejecucion As Double, porcentaje As Double
Dim Path As String, NombreCarpeta As String, fecha_actual As String, mes As String

Sub inf_driver()
' Genera el informe del driver de distribución

' Abre archivo Driver y Borrar datos de la tabla
Workbooks.OpenXML ("\\vmedsis03\Suministros\Plantillas\formatos\Driver.xlsx")
Workbooks("Driver.xlsx").Sheets("ME2N(Driver)").Activate
Workbooks("Driver.xlsx").Sheets("ME2N(Driver)").Rows("3:99999").delete

' Abre base ME2N(driver) y ajusta la base para borrar marcados con L en indicador de borrado
Workbooks.OpenXML ("\\vmedsis03\Suministros\Plantillas\FICHEROS\me2n_consolidado.xlsx")
Workbooks("me2n_consolidado.xlsx").Sheets(1).Rows(1).AutoFilter
With Workbooks("me2n_consolidado.xlsx").Sheets(1).AutoFilter.Sort
    .SortFields.Clear
    .SortFields.Add Key:=Range("R1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

' Borra las filas marcadas
Set wb = Workbooks("me2n_consolidado.xlsx")
wb.Sheets(1).Range("R2", Range("R2").End(xlDown)).Select
fila = 2
For Each celda In Selection
    If celda.FormulaR1C1 = "L" Or celda.FormulaR1C1 = "S" Then
        Rows(fila).ClearContents
    End If
    fila = fila + 1
Next
wb.Sheets(1).Cells.Select
Selection.AutoFilter
Call organizar_tabla

' Agrega Año, Mes y validador (1) si es del año actual, En caso de estar en enero, toma el año anterior.
fila = 2
Range("L2", Range("L2").End(xlDown)).Select
For Each celda In Selection
    año_indicador = Year(celda.value)
    mes_indicador = Month(celda.value)
    Range("S" & fila).value = año_indicador
    Range("T" & fila).value = mes_indicador
    fila = fila + 1
Next
Range("S2", Range("S2").End(xlDown)).Select
If Month(Date) = 1 Then
    For Each celda In Selection
        If celda.value = Year(Date) - 1 Then
            celda.Offset(0, 2).value = 1
        End If
    Next
Else
    For Each celda In Selection
        If celda.value = Year(Date) Then
            celda.Offset(0, 2).value = 1
        End If
    Next
End If

' Verifica que sea el mes y año correcto para analizar
Range("T2", Range("T2").End(xlDown)).Select
If Month(Date) = 1 Then
    For Each celda In Selection
        If celda.value = 12 And celda.Offset(0, 1).value = 1 Then
            celda.Offset(0, 2).value = 1
        End If
    Next
Else
    For Each celda In Selection
        If celda.value = Month(Date) - 1 And celda.Offset(0, 1).value = 1 Then
            celda.Offset(0, 2).value = 1
        End If
    Next
End If

' Borra todo lo que no sea del mes y año del reporte
Range("T2", Range("T2").End(xlDown)).Select
conteo = 2
For Each celda In Selection
    If celda.Offset(0, 2) <> 1 Then
        Rows(conteo).ClearContents
    End If
    conteo = conteo + 1
Next
Call organizar_tabla

' Borra las clases de documento ZMTT, ZPTR, ZNB y ZUB
Range("F2", Range("F2").End(xlDown)).Select
conteo = 2
For Each celda In Selection
    If celda.value = "ZMTT" Or celda.value = "ZPTR" Or celda.value = "ZNB" Or celda.value = "ZUB" Then
        Rows(conteo).ClearContents
    End If
    conteo = conteo + 1
Next
Call organizar_tabla

' Copia la base y la pega en el informe
Workbooks("me2n_consolidado.xlsx").Activate
Workbooks("me2n_consolidado.xlsx").Sheets(1).Range("A2:Q2", Range("A2:Q2").End(xlDown)).Copy
Workbooks("Driver.xlsx").Sheets("ME2N(Driver)").Range("A2").PasteSpecial xlAll
Application.DisplayAlerts = False
Workbooks("me2n_consolidado.xlsx").Close
Application.DisplayAlerts = True

' Crea tabla
Workbooks("Driver.xlsx").PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
    "Tabla1", Version:=xlPivotTableVersion15).CreatePivotTable TableDestination _
    :="informe_driver!R1C1", TableName:="Tabla Driver", DefaultVersion:= _
    xlPivotTableVersion15

' crea campos
With Sheets("informe_driver").PivotTables("Tabla Driver")
    .PivotFields("Organización compras").Orientation = xlRowField
    .PivotFields("Organización compras").Position = 1
    .AddDataField Sheets("informe_driver").PivotTables("Tabla Driver").PivotFields("Documento compras"), "Cantidad OC", xlCount
    .PivotFields("Cl.documento compras").Orientation = xlPageField
    .PivotFields("Cl.documento compras").Position = 1
    .PivotFields("Cl.documento compras").CurrentPage = "(All)"
    .PivotFields("Cl.documento compras").EnableMultiplePageItems = True
End With

' Copia los datos de la tabla dinamica y los pega plano
Workbooks("Driver.xlsx").Sheets("informe_driver").Activate
Workbooks("Driver.xlsx").Sheets("informe_driver").Range("A3:B15").Copy
Workbooks("Driver.xlsx").Sheets("informe_driver").Range("D3").PasteSpecial xlPasteAll

' Agrega ejecucion presupuestal y valores para grafico
Workbooks("Driver.xlsx").Sheets("informe_driver").Activate
Set total = Workbooks("Driver.xlsx").Sheets("informe_driver").Range("B3").End(xlDown)
limite = Workbooks("Driver.xlsx").Sheets("informe_driver").Range("E4", Range("E4").End(xlDown)).Count
ejecucion = InputBox("Introducir ejecución Presupuestal del mes")
With Workbooks("Driver.xlsx").Sheets("informe_driver")
    .Range("D1").value = "Ejecución Presupuestal"
    .Range("D3").value = "Organización"
    .Range("F3").value = "%"
    .Range("G3").value = "Driver"
    .Range("E1").value = ejecucion
    .Range("E1").NumberFormat = "#,##0"
End With
limite = Range("E4", Range("E4").End(xlDown)).Count
fila = 4
For i = 2 To limite
    Range("F" & fila).Formula = "=E" & fila & "/ B8"
    Range("G" & fila).Formula = "=F" & fila & "* E1"
    fila = fila + 1
Next
Range("D3").End(xlDown).ClearContents
Range("E3").End(xlDown).ClearContents
With Range("D1:E1")
    .Interior.Color = 65535
    .Font.Bold = True
    .Font.Size = 14
End With
Columns("F").NumberFormat = "0.00%"
Columns("G").NumberFormat = "#,##0"
Columns("D:I").AutoFit
Worksheets("informe_driver").ListObjects.Add(xlSrcRange, Range("D3:G3", Range("D3:G3").End(xlDown)), , xlYes).Name = "Tabla2"
Range("D3:G3", Range("D3:G3").End(xlDown)).HorizontalAlignment = xlCenter

' Guarda copia del reporte en la carpeta de suministros
Path = "\\vmedsis03\Suministros\Indicadores Compras\"
mes = Month(Date) - 1
If mes = 0 Then
    año = Year(Date) - 1
Else
    año = Year(Date)
End If
NombreCarpeta = año
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
fecha_actual = "Driver " & mes
If Dir(Path & NombreCarpeta, vbDirectory) = "" Then
    MkDir Path & NombreCarpeta
    MkDir Path & NombreCarpeta & "\" & "Driver Distribución"
    Workbooks("Driver.xlsx").SaveAs (Path & NombreCarpeta & "\" & "Driver Distribución\" & fecha_actual & ".xlsx")
Else
    Workbooks("Driver.xlsx").SaveAs (Path & NombreCarpeta & "\" & "Driver Distribución\" & fecha_actual & ".xlsx")
End If
MsgBox ("Validar antes de enviar:" + vbCrLf + "Fecha Documento en la BD" + vbCrLf + "Organizaciones de compra" + vbCrLf + "Sum Driver = Ej Presupuestal")
End Sub

Sub organizar_tabla()
'Oganiza la tabla
Set wb = Workbooks("me2n_consolidado.xlsx")
wb.Sheets(1).Cells.Select
Selection.AutoFilter
With wb.Sheets(1).AutoFilter.Sort
    .SortFields.Add Key:=Range("A1:A9999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
wb.Sheets(1).Cells.AutoFilter
End Sub

Sub grafica_driver()
' Crea la gráfica
Workbooks("Driver.xlsx").Sheets("informe_driver").Shapes.AddChart2(201, xlColumnClustered).Select
With ActiveChart
    .SeriesCollection.NewSeries
    .FullSeriesCollection(1).Name = "=informe_driver!$G$3"
    .FullSeriesCollection(1).Values = Range("G4", Range("G4").End(xlDown))
    .FullSeriesCollection(1).XValues = Range("D4", Range("D4").End(xlDown))
    .SetElement (msoElementDataLabelOutSideEnd)
    .Axes(xlValue).DisplayUnit = xlMillions
    .HasTitle = True
    .ChartTitle.Caption = "Driver"
    .ChartArea.Height = 300
    .ChartArea.Width = 600
    .ChartArea.Left = 250
    .ChartArea.Top = 20
End With
End Sub





