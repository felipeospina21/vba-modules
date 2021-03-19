Attribute VB_Name = "Seguimiento_OC"
Dim wb_bd As Workbook
Dim OutApp As Outlook.Application
Dim OutMail As Outlook.MailItem
Dim celda As Range, mirango As Range, lookuprange As Range, micelda As Range
Dim lookupvalue As Variant, mimatriz(1000000) As Variant, clave As Variant, indice As Variant
Dim email As String, prov_name As String, milibro As String, mihoja As String, hoja_rango As String, hoja_filtro As String, nombre_libro_informe As String
Dim fecha_actual As String, fecha_dia As String, fecha_mes As String, texto As String, y As String, codigo As String
Dim limite As Long, limitex As Long, conteo As Long, filas As Long, fila As Long, i As Long, filax As Long
Dim x As Date

Sub informe_seguimiento()
' Crea reporte en excel con las OC pendientes. Extrae la información del archivo Consol_compras

' Abre la bd de seguimiento
Workbooks.OpenXML ("\\vmedsis03\Suministros\Plantillas\FICHEROS\consol_compras(seguimiento).xls")
Set wb_bd = Workbooks("consol_compras(seguimiento).xls")

' Borra las filas que trae el archivo por defecto
With wb_bd.Sheets(1)
    .Columns("A").delete
    .Rows("1:3").delete
    .Rows("2").delete
    .Cells.Select
End With
Call organizar_filas

' Concatena la oc con posición
ActiveSheet.Columns("D:E").NumberFormat = "@"
ActiveSheet.Columns("O").NumberFormat = "@"
ActiveSheet.Range("D2", Range("D2").End(xlDown)).Select
fila = 2
For Each celda In Selection
    Cells(fila, 15).FormulaR1C1 = Range("D" & fila).FormulaR1C1 & Range("E" & fila).FormulaR1C1
    fila = fila + 1
Next

' Busca las Oc que se debe borrar (Reorganización) y las Borra
Workbooks.OpenXML ("\\vmedsis03\Suministros\Plantillas\formatos\oc_borrar.xlsx")
Workbooks("oc_borrar.xlsx").Sheets(1).Select
Set lookuprange = ActiveSheet.Range("A2:B231")
wb_bd.Activate
ActiveWorkbook.Sheets(1).Select
Sheets(1).Range("O2", Range("O2").End(xlDown)).Select
indice = 0
For Each micelda In Selection
    mimatriz(indice) = micelda.value
    indice = indice + 1
Next
Sheets(1).Select
limite = Sheets(1).Range("O2", Range("O2").End(xlDown)).Count
filax = 0
For clave = 0 To limite - 1 Step 1
    lookupvalue = Application.VLookup(mimatriz(clave), lookuprange, 2, False)
    Sheets(1).Range("P2").Offset(filax, 0).value = lookupvalue
    filax = filax + 1
Next
ActiveSheet.Range("P2", Range("P2").End(xlDown)).Select
fila = 2
For Each celda In Selection
    If celda.FormulaR1C1 = 1 Then
        Rows(fila).ClearContents
    End If
    fila = fila + 1
Next
Application.DisplayAlerts = False
Workbooks("oc_borrar.xlsx").Close
Application.DisplayAlerts = True
Call organizar_filas

' Borra los Nlag (codigo 5)
'-----------Mirar si en vez de borrarlos los copia en otra hoja-----------
wb_bd.Sheets(1).Cells.Select
Selection.AutoFilter
wb_bd.Sheets(1).AutoFilter.Sort.SortFields.Add Key:=Range("F1:F9999"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
With wb_bd.Sheets(1).AutoFilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
wb_bd.Sheets(1).Cells.Select
Selection.AutoFilter
wb_bd.Sheets(1).Range("F2", Range("F2").End(xlDown)).Select
conteo = 2
For Each celda In Selection
    texto = celda.value
    y = Left(texto, 1)
    If y = "5" Then
        Rows(conteo).ClearContents
    End If
    conteo = conteo + 1
Next
Call organizar_filas

' Borra compras Intercompany
wb_bd.Sheets(1).Range("B2", Range("B2").End(xlDown)).Select
conteo = 2
For Each celda In Selection
    codigo = celda.value
    If codigo = "1000" Or codigo = "1001" Or codigo = "1002" Or codigo = "1003" Or codigo = "1100" Or codigo = "1200" Or codigo = "1300" Then
        Rows(conteo).ClearContents
    End If
    conteo = conteo + 1
Next
Call organizar_filas

' Borra OC Internacionales
'--------revisar si tamb se copia y pega en otra hoja---------
wb_bd.Sheets(1).Range("B2", Range("B2").End(xlDown)).Select
conteo = 2
For Each celda In Selection
    texto = celda.value
    y = Left(texto, 1)
    If y = "2" Then
        Rows(conteo).ClearContents
    End If
    conteo = conteo + 1
Next
Call organizar_filas

' Abre formato seguimiento y Borra información anterior
Workbooks.Open ("\\vmedsis03\Suministros\Plantillas\formatos\seguimiento_oc.xlsx")
Workbooks("seguimiento_oc.xlsx").Sheets("hoja_rango").Activate
Workbooks("seguimiento_oc.xlsx").Sheets("hoja_rango").Rows("9:10000").delete

' Pega las BD en el formato
wb_bd.Sheets(1).Activate
wb_bd.Sheets(1).Range("A2:M2", Range("A2:M2").End(xlDown)).Copy
With Workbooks("seguimiento_oc.xlsx").Sheets("hoja_rango")
    .Activate
    .Range("A8").PasteSpecial xlPasteAll
End With

' Extiende las celdas formuladas a toda la BD
conteo = Range("A7", Range("A7").End(xlDown)).Count
ActiveWorkbook.Sheets("hoja_rango").Range("N8:Q8").Copy
fila = 8
For i = 2 To conteo Step 1
    Range("N" & fila).PasteSpecial (xlPasteAll)
    fila = fila + 1
Next
Workbooks("seguimiento_oc.xlsx").Sheets("hoja_rango").Range("A8").Select

' Ordenar el formato (dias restantes: Menor a Mayor, Dias faltantes: Mayor a Menor)
Workbooks("seguimiento_oc.xlsx").Sheets("hoja_rango").Range("A7:Q7").AutoFilter
With Workbooks("seguimiento_oc.xlsx").Sheets("hoja_rango").AutoFilter.Sort
    .SortFields.Clear
    .SortFields.Add Key:=Range("O7:O9999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
With Workbooks("seguimiento_oc.xlsx").Worksheets("hoja_rango").AutoFilter.Sort
    .SortFields.Clear
    .SortFields.Add Key:=Range("P7:P9999"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
Range("C2").value = Date
Rows("7").AutoFilter

' Limpia el filtro
Workbooks("seguimiento_oc.xlsx").Sheets("criterio").Range("a2:xfd1048576").delete

' Se crean los filtros y se prepara el archivo para generar los correos automaticos
hoja_rango = Sheets("hoja_rango").Name
Workbooks("seguimiento_oc.xlsx").Sheets("hoja_rango").Range("B7:C100000").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Workbooks("seguimiento_oc.xlsx").Sheets("criterio").Range("a2"), Unique:=True
Columns("A:A").EntireColumn.AutoFit
Columns("B:B").EntireColumn.AutoFit

' Verificar Correos
Workbooks.OpenXML ("\\vmedsis03\Suministros\Plantillas\formatos\correos_proveedores.xlsx")
Workbooks("correos_proveedores.xlsx").Sheets("correos").Select
Set lookuprange = Workbooks("correos_proveedores.xlsx").Sheets("correos").Range("A2:C20000")

Workbooks("seguimiento_oc.xlsx").Sheets("criterio").Activate
Workbooks("seguimiento_oc.xlsx").Sheets("criterio").Range("A3", Range("A3").End(xlDown)).Select
indice = 0
For Each micelda In Selection
    mimatriz(indice) = micelda.value
    indice = indice + 1
Next micelda
    
Sheets("criterio").Select
limite = Sheets("criterio").Range("A3", Selection.End(xlDown)).Count
fila = 0
For clave = 0 To limite - 1 Step 1
    lookupvalue = Application.VLookup(mimatriz(clave), lookuprange, 3, False)
    Sheets("criterio").Range("C3").Offset(fila, 0).value = lookupvalue
    fila = fila + 1
Next
Application.DisplayAlerts = False
Workbooks("correos_proveedores.xlsx").Close
wb_bd.Close
Application.DisplayAlerts = True
End Sub

Sub organizar_filas()
' Organiza el archivo de manera Tabular
    
Set wb_bd = Workbooks("consol_compras(seguimiento).xls")
wb_bd.Sheets(1).Cells.Select
Selection.AutoFilter
wb_bd.Sheets(1).AutoFilter.Sort.SortFields.Add Key:=Range("A1:A9999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
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

'------ Estos 4 Modulos de abajo, se llaman en el botón de enviar correos, y no cuando se genera el archivo ------

Sub Informe_LuisE_Suministros()
' Guarda copia en suministros y envía correo a LuisE

Application.ScreenUpdating = False
fecha_actual = Year(Date)
num_semana = Format(Now, "ww")
Workbooks("seguimiento_oc.xlsx").Sheets("hoja_rango").Copy
ActiveWorkbook.SaveCopyAs ("\\vmedsis03\Suministros\Seguimientos\OC\" & fecha_actual & "\Nacionales\Seguimiento_Nacional_Semana_" & num_semana & ".xlsx")
Application.DisplayAlerts = False
ActiveWorkbook.Close (False)
Application.DisplayAlerts = True
    
' Envía el correo
Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItemFromTemplate("\\vmedsis03\Suministros\Plantillas\outlook\informe_luise.oft")

On Error Resume Next

With OutMail
    .To = "luis.sarmiento@mineros.com.co"
    .Subject = "Informe"
    .Attachments.Add ("\\vmedsis03\Suministros\Seguimientos\OC\" & fecha_actual & "\Nacionales\Seguimiento_Nacional_Semana_" & num_semana & ".xlsx")
    .Display
End With
End Sub
Sub aviso_seguimiento_internacional()
'Envía el correo

Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItemFromTemplate("\\vmedsis03\Suministros\Plantillas\outlook\aviso_seguimiento_internacional.oft")

On Error Resume Next

With OutMail
    .To = "sandra.mejia@mineros.com.co"
    .Subject = "Seguimiento Internacional"
   '.Attachments.Add ("\\vmedsis03\Suministros\Seguimientos\OC\2019\Seguimiento_Nacional_Semana_" & num_semana & ".xlsx")
    .Display
End With
End Sub
Sub crear_carpeta()
' Crea carpeta para almacenar correos de respuesta proveedores de la semana

Application.ScreenUpdating = False
fecha_actual = Year(Date)
Path = "\\vmedsis03\Suministros\Seguimientos\OC\" & fecha_actual & "\Nacionales\Respuestas\"
NombreCarpeta = "Semana_" & Format(Now, "ww")

If Dir(Path, vbDirectory) <> "" Then
    If Dir(Path & NombreCarpeta, vbDirectory) = "" Then
        MkDir Path & NombreCarpeta
'    Else
'        Call Shell("explorer.exe " & Path & NombreCarpeta, vbNormalFocus)
    End If
End If
End Sub

Sub crear_subcarpetas()
' Crea carpeta por proveedor dentro de la carpeta de la semana

' Borra datos de Retrasos anteriores
fecha_actual = Year(Date)
Workbooks("seguimiento_oc.xlsx").Sheets("prov_retrasados").Activate
Workbooks("seguimiento_oc.xlsx").Sheets("prov_retrasados").Range("A2:B2", Range("A2:B2").End(xlDown)).ClearContents

' Filtra retrasos
Workbooks("seguimiento_oc.xlsx").Sheets("hoja_rango").Activate
Sheets("hoja_rango").Rows("7").AutoFilter
Workbooks("seguimiento_oc.xlsx").Sheets("hoja_rango").Range("$A$7:$Q$999999").AutoFilter Field:=16, Criteria1:=">0"

' Filtro avanzado
Workbooks("seguimiento_oc.xlsx").Sheets("hoja_rango").Activate
hoja_rango = Sheets("hoja_rango").Name
Workbooks("seguimiento_oc.xlsx").Sheets("hoja_rango").Range("B7:C7", Range("B7:C7").End(xlDown)).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Workbooks("seguimiento_oc.xlsx").Sheets("prov_retrasados").Range("A1:B1"), Unique:=True
Columns("A:A").EntireColumn.AutoFit
Columns("B:B").EntireColumn.AutoFit

' Crea Sub Carpetas
num_semana = "Semana_" & Format(Now, "ww")
Path = "\\vmedsis03\Suministros\Seguimientos\OC\" & fecha_actual & "\Nacionales\Respuestas\" & num_semana & "\"
Workbooks("seguimiento_oc.xlsx").Sheets("prov_retrasados").Activate
Workbooks("seguimiento_oc.xlsx").Sheets("prov_retrasados").Range("B2", Range("B2").End(xlDown)).Select
For Each celda In Selection
    On Error Resume Next
    NombreCarpeta = celda.value
    If Dir(Path, vbDirectory) <> "" Then
        If Dir(Path & NombreCarpeta, vbDirectory) = "" Then
            MkDir Path & NombreCarpeta
        End If
    End If
Next
End Sub
