Attribute VB_Name = "Indicadores_Servicios"
Dim año_indicador As Integer, mes_indicador As Integer, fila As Long, limite As Long, cant_filas As Long
Dim celda As Range, mes As String
Dim lookupvalue As Variant, mimatriz(1000000) As Variant, clave As Variant
Dim mirango As Range, lookuprange As Range, micelda As Range
Sub Ind_Servicios()

Application.ScreenUpdating = False
' Abre y organiza la BD
Workbooks.OpenXML ("\\vmedsis03\Suministros\Plantillas\FICHEROS\consol_compras(servicios).xls")
With ActiveWorkbook.Sheets(1)
    .Rows("5").delete
    .Rows("1:3").delete
    .Columns("A").delete
End With
Selection.AutoFilter
ActiveWorkbook.Sheets(1).AutoFilter.Sort.SortFields.Add Key:=Range("X1:X9999"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal

With ActiveWorkbook.Sheets(1).AutoFilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

' Agrega Año, Mes y validador (1) si es del año actual.
fila = 2
Range("X2", Range("X2").End(xlDown)).Select
For Each celda In Selection
    año_indicador = Year(celda.value)
    mes_indicador = Month(celda.value)

    Range("Z" & fila).value = año_indicador
    Range("AA" & fila).value = mes_indicador
    fila = fila + 1
Next
Range("Z2", Range("Z2").End(xlDown)).Select
For Each celda In Selection
    If celda.value = Year(Date) Then
        celda.Offset(0, 2).value = 1
    End If
Next

' Deja solo las filas que sean del mes actual y anterior.
cant_filas = Range("AB2", Range("AB2").End(xlDown)).Count
cant_filas = cant_filas + 2
Range("A" & cant_filas, Range("A" & cant_filas).End(xlDown)).Select
Range(Selection, Range("AA" & cant_filas, Range("AA" & cant_filas).End(xlDown))).ClearContents
cant_filas = 0
Range("AA2", Range("AA2").End(xlDown)).Select
For Each celda In Selection
    If celda.value < Month(Date) - 1 Then
        celda.Offset(0, 2).value = 1
    End If
Next
cant_filas = Range("AC2", Range("AC2").End(xlDown)).Count
cant_filas = cant_filas + 1
Range("A" & cant_filas, Range("A" & cant_filas).End(xlDown)).Select
Range(Selection, Range("AD" & cant_filas, Range("AD" & cant_filas).End(xlDown))).ClearContents

' Elimina las filas que sean del mes actual
Set wb = Workbooks("consol_compras(servicios).xls")
Range("AA2", Range("AA2").End(xlDown)).Select
conteo = 2
For Each celda In Selection
    If celda.value = Month(Date) Then
        Rows(conteo).ClearContents
    End If
conteo = conteo + 1
Next
wb.Sheets(1).Cells.Select
Selection.AutoFilter

Call arrange_rows

' Elimina las compras intercompany (sociedades de mineros)
Set wb = Workbooks("consol_compras(servicios).xls")
Columns("B").Select
conteo = 1
For Each celda In Selection
codigo = celda.value
    If codigo = "1000" Or codigo = "1001" Or codigo = "1002" Or codigo = "1003" Or codigo = "1100" Or codigo = "1200" Or codigo = "1300" Then
        Rows(conteo).ClearContents
    End If
conteo = conteo + 1
Next

Call arrange_rows

' Copia la info procesada y pega en el formato del indicador.
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

Set wb = Workbooks("consol_compras(servicios).xls")
wb.Sheets(1).Range("A2:Y2", Range("A2:Y2").End(xlDown)).Copy
Application.DisplayAlerts = False
Workbooks.OpenXML ("\\vmedsis03\Suministros\Plantillas\formatos\indicadores_servicios.xlsx")
Workbooks("indicadores_servicios.xlsx").Sheets("BaseDatos").Activate
Workbooks("indicadores_servicios.xlsx").Sheets("BaseDatos").Range("A2").PasteSpecial xlAll
Workbooks("indicadores_servicios.xlsx").SaveAs ("C:\Documentos Empresa\OneDrive - MINEROS\Desktop\INDICADORES\" & año & "\" & mes & "\indicadores_servicios(" & mes & ").xlsx")
wb.Close
Application.DisplayAlerts = True

' Trae los contratos de cada orden de la ME2N
Workbooks.OpenXML ("\\vmedsis03\Suministros\Plantillas\FICHEROS\me2n_consolidado.xlsx")
With Workbooks("me2n_consolidado.xlsx").Sheets(1)
    .Activate
    .Columns("C").Insert
    .Columns("C").NumberFormat = "@"
    .Range("A2", Range("A2").End(xlDown)).Select
End With
For Each celda In Selection
    concatenado = celda.value & celda.Offset(0, 1).value
    celda.Offset(0, 2) = concatenado
Next
    ' BuscarV
Workbooks("me2n_consolidado.xlsx").Sheets(1).Activate
Set lookuprange = Workbooks("me2n_consolidado.xlsx").Sheets(1).Range("C2:D999999")
Workbooks("indicadores_servicios(" & mes & ").xlsx").Sheets("BaseDatos").Activate
Workbooks("indicadores_servicios(" & mes & ").xlsx").Sheets("BaseDatos").Range("AF2", Range("AF2").End(xlDown)).Select
indice = 0
For Each micelda In Selection
    mimatriz(indice) = micelda.value
    indice = indice + 1
Next micelda
limite = Workbooks("indicadores_servicios(" & mes & ").xlsx").Sheets("BaseDatos").Range("AF2", Range("AF2").End(xlDown)).Count
fila = 0
For clave = 0 To limite - 1 Step 1
    lookupvalue = Application.VLookup(mimatriz(clave), lookuprange, 2, False)
    Workbooks("indicadores_servicios(" & mes & ").xlsx").Sheets("BaseDatos").Range("AH2").Offset(fila, 0).value = lookupvalue
    fila = fila + 1
Next
Application.DisplayAlerts = False
Workbooks("me2n_consolidado.xlsx").Close
Application.DisplayAlerts = True

' Trae el tipo de proveedor de la BD Correos
Workbooks.OpenXML ("\\vmedsis03\Suministros\Plantillas\formatos\correos_proveedores.xlsx")
Set wb_bv = Workbooks("correos_proveedores.xlsx")
wb_bv.Sheets(1).Select
Set lookuprange = wb_bv.Sheets(1).Range("A2:E99999")
Workbooks("indicadores_servicios(" & mes & ").xlsx").Activate
ActiveWorkbook.Sheets("BaseDATOS").Select
Sheets("BaseDATOS").Range("B2", Range("B2").End(xlDown)).Select
indice = 0
For Each micelda In Selection
    mimatriz(indice) = micelda.value
    indice = indice + 1
Next micelda
limite = Sheets("BaseDatos").Range("B2", Range("B2").End(xlDown)).Count
filax = 0
For clave = 0 To limite - 1 Step 1
    lookupvalue = Application.VLookup(mimatriz(clave), lookuprange, 5, False)
    Sheets("BaseDatos").Range("Z2").Offset(filax, 0).value = lookupvalue
    filax = filax + 1
Next
wb_bv.Close
Erase mimatriz
End Sub

Sub arrange_rows()
' Organiza las filas de forma tabular
Set wb = Workbooks("consol_compras(servicios).xls")
wb.Sheets(1).Cells.Select
Selection.AutoFilter
wb.Sheets(1).AutoFilter.Sort.SortFields.Add Key:=Range("B1:B9999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With wb.Sheets(1).AutoFilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
wb.Sheets(1).Cells.Select
Selection.AutoFilter
End Sub


Sub CORREOS()
' Trae el tipo de proveedor de la BD Correos
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
Workbooks.OpenXML ("\\vmedsis03\Suministros\Plantillas\formatos\correos_proveedores.xlsx")
Set wb_bv = Workbooks("correos_proveedores.xlsx")
wb_bv.Sheets(1).Select
Set lookuprange = wb_bv.Sheets(1).Range("A2:E99999")
Workbooks("indicadores_servicios(" & mes & ").xlsx").Activate
ActiveWorkbook.Sheets("BaseDATOS").Select
Sheets("BaseDATOS").Range("B2", Range("B2").End(xlDown)).Select
indice = 0
For Each micelda In Selection
    mimatriz(indice) = micelda.value
    indice = indice + 1
Next micelda
limite = Sheets("BaseDatos").Range("B2", Range("B2").End(xlDown)).Count
filax = 0
For clave = 0 To limite - 1 Step 1
    lookupvalue = Application.VLookup(mimatriz(clave), lookuprange, 5, False)
    Sheets("BaseDatos").Range("Z2").Offset(filax, 0).value = lookupvalue
    filax = filax + 1
Next
wb_bv.Close
Erase mimatriz
End Sub

