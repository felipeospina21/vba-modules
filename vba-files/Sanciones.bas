Attribute VB_Name = "Sanciones"
Dim mes As String
Dim conteo As Integer, filas As Integer
Dim email As String, prov_name As String, milibro As String, mihoja As String, hoja_rango As String, hoja_filtro As String, nombre_libro_informe As String
Dim lookupvalue As Variant, mimatriz(1000000) As Variant, clave As Variant
Dim mirango As Range, lookuprange As Range, micelda As Range
Dim fila As Integer, i As Integer, indice As Integer
Dim fecha_actual As String, fecha_dia As String, fecha_mes As String
Dim limite As Long, limitex As Long
Sub informe_multas()
' Genera informe de sanciones mes anterior

' Abre el archivo de indicadores, para sacar los retrasos del mes.
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
Workbooks.OpenXML ("\\vmedsis03\Suministros\Plantillas\formatos\informe_sanciones.xlsx")
With Workbooks("informe_sanciones.xlsx").Sheets("hoja_rango")
    .Range("A8:M8", Range("A8:M8").End(xlDown)).ClearContents
    .Rows("9:1000").delete
End With
Workbooks.OpenXML ("\\vmedsis03\Suministros\Indicadores Compras\" & año & "\Indicadores Mensuales\" & mes & "\resumen_indicadores.xlsx")
   
' Copia los retrasos del proveedor y los pega en el formato informe_sanciones
Workbooks("resumen_indicadores.xlsx").Sheets("Análisis_Entrega").Activate
Workbooks("resumen_indicadores.xlsx").Sheets("Análisis_Entrega").Range("A1", Range("X1000").End(xlDown)).Copy
Workbooks("resumen_indicadores.xlsx").Sheets.Add
ActiveSheet.Range("A1").PasteSpecial xlAll
With ActiveSheet
    .Columns("A").delete
    .Columns("D:E").delete
    .Columns("H:J").delete
    .Columns("I").delete
    .Columns("L:O").delete
    .Range("A2:M2", Range("A2:M2").End(xlDown)).Copy
End With
With Workbooks("informe_sanciones.xlsx").Sheets("hoja_rango")
    .Activate
    .Range("A8").PasteSpecial xlAll
    .Range("B2").value = mes
End With
conteo = Range("A8", Range("A8").End(xlDown)).Count
ActiveWorkbook.Sheets("hoja_rango").Range("N8:R8").Copy
fila = 9
For i = 2 To conteo Step 1
    Range("N" & fila).PasteSpecial (xlPasteAll)
    fila = fila + 1
Next
Application.DisplayAlerts = False
Workbooks("resumen_indicadores.xlsx").Close

' Limpia el filtro
Workbooks("informe_sanciones.xlsx").Sheets("criterio").Range("A2:XFD1048576").delete
    
' Se crean los filtros y se prepara el archivo para generar los correos automaticos
hoja_rango = Workbooks("informe_sanciones.xlsx").Sheets("hoja_rango").Name
Workbooks("informe_sanciones.xlsx").Sheets(hoja_rango).Range("A7:B100000").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ActiveWorkbook.Sheets("criterio").Range("A2"), Unique:=True
Columns("A:A").EntireColumn.AutoFit
Columns("B:B").EntireColumn.AutoFit
    
' Verificar Correos
Workbooks.Open ("\\vmedsis03\Suministros\Plantillas\formatos\correos_proveedores.xlsx")
Workbooks("correos_proveedores.xlsx").Sheets("correos").Select
Set lookuprange = ActiveWorkbook.Sheets("correos").Range("A2:C10000")
Workbooks("informe_sanciones.xlsx").Sheets("criterio").Activate
Workbooks("informe_sanciones.xlsx").Sheets("criterio").Range("A3", Range("A3").End(xlDown)).Select
indice = 0
For Each micelda In Selection
    mimatriz(indice) = micelda.value
    indice = indice + 1
Next
Workbooks("informe_sanciones.xlsx").Sheets("criterio").Select
limite = Workbooks("informe_sanciones.xlsx").Sheets("criterio").Range("A3", Selection.End(xlDown)).Count
fila = 0
For clave = 0 To limite - 1 Step 1
    lookupvalue = Application.VLookup(mimatriz(clave), lookuprange, 3, False)
    Workbooks("informe_sanciones.xlsx").Sheets("criterio").Range("C3").Offset(fila, 0).value = lookupvalue
    fila = fila + 1
Next
Workbooks("correos_proveedores.xlsx").Close
Application.DisplayAlerts = True
End Sub


