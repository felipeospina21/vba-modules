Attribute VB_Name = "Seguimiento_facturas"
Dim celda As Range, filax As Integer
Dim lookupvalue As Variant, mimatriz(1000000) As Variant, clave As Variant
Dim mirango As Range, lookuprange As Range, micelda As Range
Dim wb As Workbook
Dim fila As Long, i As Integer, indice As Variant
Dim fecha_actual As String, fecha_dia As String, fecha_mes As String, fecha_reporte As String
Dim limite As Long, limitex As Long

Sub informe_facturas()

Application.ScreenUpdating = False
' Abre BD, concatena OC y posición, y filtra los pedidos 45 excluyendo sociedad 4000
Workbooks.OpenXML ("\\vmedsis03\Suministros\Plantillas\FICHEROS\MB5S(facturas).xlsx")
Set wb = Workbooks("MB5S(facturas).xlsx")
With wb.Sheets(1)
    .Columns("E:F").NumberFormat = "@"
    .Columns("N").NumberFormat = "@"
    .Range("E2", Range("E2").End(xlDown)).Select
End With
fila = 2
For Each celda In Selection
    Cells(fila, 14).FormulaR1C1 = Range("E" & fila).FormulaR1C1 & Range("F" & fila).FormulaR1C1
    If Left(celda, 2) = "45" And Range("A" & fila).value <> "4000" Then
        Cells(fila, 15).FormulaR1C1 = 1
    Else
        Cells(fila, 15).FormulaR1C1 = 0
    End If
    fila = fila + 1
Next

' Borra las filas marcadas con 0
wb.Sheets(1).Range("O2", Range("O2").End(xlDown)).Select
fila = 2
For Each celda In Selection
    If celda.FormulaR1C1 = 0 Then
        Rows(fila).ClearContents
    End If
    fila = fila + 1
Next

'Organiza el archivo de manera Tabula
wb.Sheets(1).Cells.AutoFilter
wb.Sheets(1).AutoFilter.Sort.SortFields.Add Key:=Range("A1:A99999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With wb.Sheets(1).AutoFilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
wb.Sheets(1).Cells.AutoFilter

' Trae la fecha MIGO del archivo consol_pedidos
indice = 0
Workbooks.OpenXML ("\\vmedsis03\Suministros\Plantillas\FICHEROS\indicadores_entregas.xls")
With ActiveSheet
    .Columns("E:F").NumberFormat = "@"
    .Columns("Y").Insert
    .Columns("Y").NumberFormat = "@"
    .Range("E6", Range("E6").End(xlDown)).Select
End With
fecha_reporte = Workbooks("indicadores_entregas.xls").Sheets(1).Range("A1").value
fila = 6
For Each celda In Selection
    Cells(fila, 25).FormulaR1C1 = Range("E" & fila).FormulaR1C1 & Range("F" & fila).FormulaR1C1
    fila = fila + 1
Next

Set lookuprange = Workbooks("indicadores_entregas.xls").Sheets(1).Range("Y6:Z999999")
With Workbooks("MB5S(facturas).xlsx").Sheets(1)
    .Activate
    .Range("R1").value = fecha_reporte
    .Range("N2", Range("N2").End(xlDown)).Select
    .Columns("P").NumberFormat = "m/d/yyyy"
End With
For Each celda In Selection
    mimatriz(indice) = celda.value
    indice = indice + 1
Next
Workbooks("MB5S(facturas).xlsx").Sheets(1).Select
limite = Sheets(1).Range("N2", Range("N2").End(xlDown)).Count
fila = 0
For clave = 0 To limite - 1 Step 1
    lookupvalue = Application.VLookup(mimatriz(clave), lookuprange, 2, False)
    Sheets(1).Range("P2").Offset(fila, 0).value = lookupvalue
    fila = fila + 1
Next
Application.DisplayAlerts = False
Workbooks("indicadores_entregas.xls").Close
Application.DisplayAlerts = True

' Ajusta el formato final a trabajar
With Workbooks("MB5S(facturas).xlsx").Sheets(1)
    .Columns("N:O").delete
    .Range("N1").value = "Fecha MIGO"
    .Columns("L:M").delete
    .Columns("B:C").delete
    .Range("A2", Range("A2").End(xlDown)).Select
End With
fila = 2
For Each celda In Selection
    If celda.value = "1100" Then
        Cells(fila, 11).FormulaR1C1 = "Operadora Minera"
    ElseIf celda.value = "1200" Then
        Cells(fila, 11).FormulaR1C1 = "Negocios Agroforestales"
    ElseIf celda.value = "1300" Then
        Cells(fila, 11).FormulaR1C1 = "Mineros Aluvial"
    Else
        Cells(fila, 11).FormulaR1C1 = "Mineros S.A"
    End If
    fila = fila + 1
Next
With Workbooks("MB5S(facturas).xlsx").Sheets(1)
    .Columns("A").delete
    .Range("J1").value = "Sociedad"
End With

' Copia la BD al formato
Workbooks.OpenXML ("\\vmedsis03\Suministros\Plantillas\formatos\seguimiento_facturas.xlsx")
With Workbooks("seguimiento_facturas.xlsx").Sheets("hoja_rango")
    .Activate
    .Range("A2:L2", Range("A2:L2").End(xlDown)).ClearContents
End With
Workbooks("MB5S(facturas).xlsx").Sheets(1).Activate
fecha_reporte = Range("K1").value
Range("A2:J2", Range("A2:J2").End(xlDown)).Copy
With Workbooks("seguimiento_facturas.xlsx").Sheets("hoja_rango")
    .Activate
    .Range("A2").PasteSpecial xlPasteAll
End With
With Workbooks("seguimiento_facturas.xlsx").Sheets("criterio")
    .Activate
    .Range("E2").value = fecha_reporte
    .Range("D2").value = "Fecha Reporte Entregas"
    .Range("D2:E2").Font.Bold = True
    .Range("D2:E2").Font.Size = 14
    .Range("D2:E2").Interior.Color = 65535
    .Columns("D:E").AutoFit
End With
Application.DisplayAlerts = False
Workbooks("MB5S(facturas).xlsx").Close
Application.DisplayAlerts = True

' Calcula los días entre fecha migo y hoy
With Workbooks("seguimiento_facturas.xlsx").Sheets("hoja_rango")
    .Activate
    .Range("I2", Range("I2").End(xlDown)).Select
End With
fila = 2
For Each celda In Selection
    On Error Resume Next
    Cells(fila, 11).value = Date - celda.value
    fila = fila + 1
Next

' Elimina las que tengan menos de 8 dias
limite = Range("I2", Range("I2").End(xlDown)).Count
fila = 2
For i = 0 To limite
    If Range("K" & fila).value < 8 Then
        Rows(fila).ClearContents
    End If
    fila = fila + 1
Next

' Organiza filas
Set wb_seguimiento = Workbooks("seguimiento_facturas.xlsx")
wb_seguimiento.Sheets("hoja_rango").Cells.AutoFilter
wb_seguimiento.Sheets("hoja_rango").AutoFilter.Sort.SortFields.Add Key:=Range("K1:K999999"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
With wb_seguimiento.Sheets("hoja_rango").AutoFilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
wb_seguimiento.Sheets("hoja_rango").Cells.AutoFilter

' Se crean los filtros y se prepara el archivo para generar los correos automaticos
Workbooks("seguimiento_facturas.xlsx").Sheets("criterio").Range("A2:XFD1048576").delete
Workbooks("seguimiento_facturas.xlsx").Sheets("hoja_rango").Range("A1:A100000").AdvancedFilter Action:=xlFilterCopy, _
    CopyToRange:=ActiveWorkbook.Sheets("criterio").Range("A2"), Unique:=True
Workbooks("seguimiento_facturas.xlsx").Sheets("criterio").Activate
Range("A3", Range("A3").End(xlDown)).Select
For Each celda In Selection
    celda.value = celda.FormulaR1C1
Next
Columns("A:B").EntireColumn.AutoFit
Range("B2").value = "Nombre Proveedor"
Range("C2").value = "Correos"
Range("A2").Copy
Range("B2:C2").PasteSpecial xlPasteFormats
    
' Verificar Correos
Workbooks.Open ("\\vmedsis03\Suministros\Plantillas\formatos\correos_proveedores.xlsx")
Workbooks("correos_proveedores.xlsx").Sheets("correos").Select
Set lookuprange = ActiveWorkbook.Sheets("correos").Range("A2:C999999")
With Workbooks("seguimiento_facturas.xlsx").Sheets("criterio")
    .Activate
    .Range("A3", Range("A3").End(xlDown)).Select
End With
indice = 0
For Each celda In Selection
    mimatriz(indice) = celda.value
    indice = indice + 1
Next
limite = Workbooks("seguimiento_facturas.xlsx").Sheets("criterio").Range("A3", Selection.End(xlDown)).Count
fila = 0
For clave = 0 To limite - 1 Step 1
    lookupvalue = Application.VLookup(mimatriz(clave), lookuprange, 2, False)
    Workbooks("seguimiento_facturas.xlsx").Sheets("criterio").Range("B3").Offset(fila, 0).value = lookupvalue
    lookupvalue = Application.VLookup(mimatriz(clave), lookuprange, 3, False)
    Workbooks("seguimiento_facturas.xlsx").Sheets("criterio").Range("C3").Offset(fila, 0).value = lookupvalue
    fila = fila + 1
Next
Application.DisplayAlerts = False
Workbooks("correos_proveedores.xlsx").Close
Application.DisplayAlerts = True
End Sub





