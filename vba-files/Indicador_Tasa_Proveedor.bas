Attribute VB_Name = "Indicador_Tasa_Proveedor"
Dim nom_hoja2 As String, rango_campos As Range
Dim conteo As Integer, n As Integer
Dim nom_hoja As String, fecha_entrega As Date
Dim mes_analizado As Date, ultimo_mes As Date
Dim limite As Long
Dim fila As Long, indice As Long
Dim lookupvalue As Variant, mimatriz(1000000) As Variant, clave As Variant
Dim lookuprange As Range, micelda As Range
Dim valor As Date, celda As Range
Dim fecha_año As String, fecha As String, valor_celda As String

Sub ts_proveedor()
' Layout SAP --> Ts_Proveedor,
' Funciona, pero toma mucho tiempo el proceso revisar diferentes bucles para simplificar
' Revisar como traer las fechas de solped (ME5A), de las OC que no las trae en el reporte.

Application.ScreenUpdating = False

' Abre y organiza la BD
Workbooks.OpenXML ("\\vmedsis03\Suministros\Plantillas\FICHEROS\indicadores_entregas.xls")
Set wb_bd = Workbooks("indicadores_entregas.xls")
With wb_bd.Sheets(1)
    .Rows("5").delete
    .Rows("1:3").delete
    .Columns("A").delete
End With
wb_bd.Sheets(1).Columns("X").value = wb_bd.Sheets(1).Columns("X").FormulaR1C1
Cells.AutoFilter
wb_bd.Sheets(1).AutoFilter.Sort.SortFields.Add Key:=Range("X1:X99999"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
With wb_bd.Sheets(1).AutoFilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
Cells.AutoFilter

' Agrega Año y Mes de la transacción
fila = 2
wb_bd.Sheets(1).Range("X2", Range("X2").End(xlDown)).Select
For Each celda In Selection
    año_indicador = Year(celda.value)
    mes_indicador = Month(celda.value)
    Range("Y" & fila).value = año_indicador
    Range("Z" & fila).value = mes_indicador
    fila = fila + 1
Next

' Valida y marca si es el año a revisar
wb_bd.Sheets(1).Range("Y2", Range("Y2").End(xlDown)).Select
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
wb_bd.Sheets(1).Range("Z2", Range("Z2").End(xlDown)).Select
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
Call arrange_rows2

' Borra todo lo que no sea del mes y año del reporte
wb_bd.Sheets(1).Range("AB2").End(xlDown).Offset(1, 0).Select
fila = ActiveCell.Row
wb_bd.Sheets(1).Range("A" & fila & ":AA" & fila, Range("A" & fila & ":AA" & fila).End(xlDown)).ClearContents

' Elimina las compras intercompany (sociedades de mineros)
wb_bd.Sheets(1).Range("B2", Range("B2").End(xlDown)).Select
conteo = 2
For Each celda In Selection
codigo = celda.value
    If codigo = "1000" Or codigo = "1001" Or codigo = "1002" Or codigo = "1003" Or codigo = "1100" Or codigo = "1200" Or codigo = "1300" Then
        Rows(conteo).ClearContents
    End If
conteo = conteo + 1
Next
Call arrange_rows2

' Copia y pega en la BDatos ela archivo
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
Range("A2:X2", Range("A2:X2").End(xlDown)).Copy
Application.DisplayAlerts = False
Workbooks.OpenXML ("\\vmedsis03\Suministros\Plantillas\formatos\tasa_proveedor.xlsx")
Workbooks("tasa_proveedor.xlsx").Sheets("BDATOS").Range("A2").PasteSpecial
Workbooks("tasa_proveedor.xlsx").Sheets("RESUMEN ENTREGAS").Range("A1").value = mes
Workbooks("tasa_proveedor.xlsx").SaveAs ("C:\Documentos Empresa\OneDrive - MINEROS\Desktop\INDICADORES\" & año & "\" & mes & "\Ts_Proveedor(" & mes & ").xlsx")
Workbooks("indicadores_entregas.xls").Close

'----------Formulas--------------

' Trae el tipo de proveedor de la BD Correos
Workbooks.OpenXML ("\\vmedsis03\Suministros\Plantillas\formatos\correos_proveedores.xlsx")
Set wb_bv = Workbooks("correos_proveedores.xlsx")

wb_bv.Sheets(1).Select
Set lookuprange = wb_bv.Sheets(1).Range("A2:E99999")
Workbooks("Ts_Proveedor(" & mes & ").xlsx").Activate
ActiveWorkbook.Sheets("BDATOS").Select
Sheets("BDATOS").Range("B2", Range("B2").End(xlDown)).Select
indice = 0
For Each micelda In Selection
    mimatriz(indice) = micelda.value
    indice = indice + 1
Next micelda
limite = Sheets("BDATOS").Range("B2", Range("B2").End(xlDown)).Count
filax = 0
For clave = 0 To limite - 1 Step 1
    lookupvalue = Application.VLookup(mimatriz(clave), lookuprange, 5, False)
    Sheets("BDATOS").Range("AA2").Offset(filax, 0).value = lookupvalue
    filax = filax + 1
Next
wb_bv.Close
Erase mimatriz

' Trae Lead Time
' Este archivo se le debe insertar una columna para concat material y centro. Esto no está automatizado
Workbooks.OpenXML ("\\vmedsis03\Suministros\Plantillas\FICHEROS\zmm011(lead time).xlsx")
Set wb_bv = Workbooks("zmm011(lead time).xlsx")

wb_bv.Sheets(1).Select
Set lookuprange = wb_bv.Sheets(1).Range("C2:E99999")
Workbooks("Ts_Proveedor(" & mes & ").xlsx").Activate
ActiveWorkbook.Sheets("BDATOS").Select
Sheets("BDATOS").Range("AC2", Range("AC2").End(xlDown)).Select
indice = 0
For Each micelda In Selection
    mimatriz(indice) = micelda.value
    indice = indice + 1
Next
limite = Sheets("BDATOS").Range("AC2", Range("AC2").End(xlDown)).Count
filax = 0
For clave = 0 To limite - 1 Step 1
    lookupvalue = Application.VLookup(mimatriz(clave), lookuprange, 3, False)
    Sheets("BDATOS").Range("AL2").Offset(filax, 0).value = lookupvalue
    filax = filax + 1
Next
wb_bv.Close
Erase mimatriz

' Refrescar Tablas
Sheets("CUMPLIMIENTO").Activate
Sheets("CUMPLIMIENTO").PivotTables("Tabla dinámica1").PivotCache.Refresh
limite = Workbooks("Ts_Proveedor(" & mes & ").xlsx").Sheets("CUMPLIMIENTO").Range("E1", Range("E1").End(xlDown)).Count
Rows(limite + 1 & ":10000").delete
Sheets("TS").Select
Sheets("TS").PivotTables("Tabla dinámica2").PivotCache.Refresh

' Pareto
Workbooks.OpenXML ("\\vmedsis03\Suministros\Plantillas\formatos\proveedores_pareto.xlsx")
Set wb_bv = Workbooks("proveedores_pareto.xlsx")
wb_bv.Sheets(1).Select
Set lookuprange = wb_bv.Sheets(1).Range("A2:B99999")
Workbooks("Ts_Proveedor(" & mes & ").xlsx").Activate
ActiveWorkbook.Sheets("TS").Select
Sheets("TS").Range("H4", Range("H4").End(xlDown)).Select
indice = 0
For Each micelda In Selection
    If micelda = "Total general" Then
        Exit For
    Else
        mimatriz(indice) = micelda.value
    indice = indice + 1
    End If
Next
limite = Sheets("TS").Range("H4", Range("H4").End(xlDown)).Count
filax = 0
For clave = 0 To limite - 1 Step 1
    lookupvalue = Application.VLookup(mimatriz(clave), lookuprange, 2, False)
    If IsError(lookupvalue) Then
        Sheets("TS").Range("N4").Offset(filax, 0).value = 0
    Else
        Sheets("TS").Range("N4").Offset(filax, 0).value = 1
    End If
    filax = filax + 1
Next
wb_bv.Close
Erase mimatriz

' Traer Festivos
Workbooks.OpenXML ("\\vmedsis03\Suministros\Plantillas\formatos\festivos.xlsx")
Workbooks("festivos.xlsx").Sheets(1).Range("A2", Range("A2").End(xlDown)).Copy
Workbooks("Ts_Proveedor(" & mes & ").xlsx").Sheets("festivos").Range("A2").PasteSpecial
Workbooks("festivos.xlsx").Close
Application.DisplayAlerts = True
End Sub
Sub arrange_rows2()
' Organiza las filas de forma tabular
    
Set wb_bd = Workbooks("indicadores_entregas.xls")
wb_bd.Sheets(1).Cells.Select
Selection.AutoFilter
wb_bd.Sheets(1).AutoFilter.Sort.SortFields.Add Key:=Range("AB1:AB999999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

With wb_bd.Sheets(1).AutoFilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
wb_bd.Sheets(1).Cells.Select
Selection.AutoFilter
End Sub
