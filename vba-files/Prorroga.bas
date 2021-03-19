Attribute VB_Name = "Prorroga"

Dim Rango As Range, Dato As Double, celda As Range
Dim Path As String, NombreCarpetaAño As String, NombreCarpetaOc As String, fecha_actual As String, mes As String, NombreCarpeta As String

Sub cambio_fecha()
' Ruta carpeta = \\vmedsis03\Suministros\Cambio Fechas\

' Crear carpeta
Application.ScreenUpdating = False
Path = "\\vmedsis03\Suministros\Cambio Fechas\"
NombreCarpeta = InputBox("# OC")

If Dir(Path, vbDirectory) <> "" Then
    If Dir(Path & NombreCarpeta, vbDirectory) = "" Then
        MkDir Path & NombreCarpeta
        Call Shell("explorer.exe " & Path & NombreCarpeta, vbNormalFocus)
        Set OutApp = CreateObject("Outlook.Application")
        Set OutMail = OutApp.CreateItemFromTemplate("\\vmedsis03\Suministros\Plantillas\outlook\retrasos_entregas.oft")
        With OutMail
            .To = "jorge.alvis@mineros.com.co"
            '.CC = "felipe.ospina@mineros.com.co"
            .Subject = "Retraso OC " & NombreCarpeta
            .Display
        End With
    Else
        Call Shell("explorer.exe " & Path & NombreCarpeta, vbNormalFocus)
    End If
End If

fecha_actual = Date
Workbooks.OpenXML ("\\vmedsis03\Suministros\Cambio Fechas\Reporte_Cambio_Fecha.xlsx")
Set wb = Workbooks("Reporte_Cambio_Fecha.xlsx")
If wb.Sheets("BD").Range("A2").value = "" Then
    wb.Sheets("BD").Range("A2").FormulaR1C1 = NombreCarpeta
    wb.Sheets("BD").Range("B2").FormulaR1C1 = fecha_actual
Else
    wb.Sheets("BD").Range("A1").End(xlDown).Select
    ActiveCell.Offset(1, 0).FormulaR1C1 = NombreCarpeta
    ActiveCell.Offset(1, 1).FormulaR1C1 = fecha_actual
End If
Application.DisplayAlerts = False
Workbooks("Reporte_Cambio_Fecha.xlsx").Save
Workbooks("Reporte_Cambio_Fecha.xlsx").Close
Application.DisplayAlerts = True
End Sub

Sub informe_cambio_fecha()
' Cruza el proveedor con cada OC, eliminca OC duplicadas y crea TD con el conteo de solicitudes de cambio de fecha

Application.ScreenUpdating = False
Workbooks.OpenXML ("\\vmedsis03\Suministros\Plantillas\FICHEROS\consol_compras (indicadores).xls")
Workbooks.OpenXML ("\\vmedsis03\Suministros\Cambio Fechas\Reporte_Cambio_Fecha.xlsx")
Workbooks("Reporte_Cambio_Fecha.xlsx").Sheets("BD").Range("A1:D1", Range("A1:D1").End(xlDown)).RemoveDuplicates Columns:=1, Header:=xlYes
fila = 2
Workbooks("Reporte_Cambio_Fecha.xlsx").Sheets("BD").Range("A2", Range("A2").End(xlDown)).Select
For Each celda In Selection
    ' Valor a buscar
    Dato = celda.value
    ' Rango Función Coincidir
    Workbooks("consol_compras (indicadores).xls").Activate
    Set Rango = Workbooks("consol_compras (indicadores).xls").Sheets(1).Range("E6", Range("E6").End(xlDown))
    ' Función Coincidir
    valor = Application.Match(Dato, Rango, 0)
    ' Función Indice
    proveedor = Application.Index(Workbooks("consol_compras (indicadores).xls").Sheets(1).Range("C6:E99999"), valor, 2)
    Workbooks("Reporte_Cambio_Fecha.xlsx").Sheets("BD").Range("C" & fila).FormulaR1C1 = proveedor
    proveedor = Application.Index(Workbooks("consol_compras (indicadores).xls").Sheets(1).Range("C6:E99999"), valor, 1)
    Workbooks("Reporte_Cambio_Fecha.xlsx").Sheets("BD").Range("D" & fila).FormulaR1C1 = proveedor
    fila = fila + 1
Next

' Crear tabla dinámica
Workbooks("Reporte_Cambio_Fecha.xlsx").Sheets("Informe").Activate
Workbooks("Reporte_Cambio_Fecha.xlsx").PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Tabla1", _
Version:=xlPivotTableVersion15).CreatePivotTable TableDestination:=ActiveWorkbook.Sheets("Informe").Cells(2, 1), _
TableName:="Conteo", DefaultVersion:=xlPivotTableVersion15
With ActiveSheet.PivotTables("Conteo")
    .PivotFields("Proveedor").Orientation = xlRowField
    .PivotFields("Proveedor").Position = 1
    .AddDataField ActiveSheet.PivotTables("Conteo").PivotFields("OC"), "Cuenta", xlCount
    .PivotFields("Proveedor").AutoSort xlDescending, "Cuenta"
End With
Application.DisplayAlerts = False
Workbooks("consol_compras (indicadores).xls").Close
Application.DisplayAlerts = True
End Sub

Sub consultar_cotizacion()
' Busca en NASTRELLA la carpeta de la OC y la abre. De no existir arroja un mensaje de alerta.

Application.ScreenUpdating = False
Path = "\\Nasestrella\oc\"
fecha_actual = Year(Date)
NombreCarpetaAño = InputBox("Introduce año de creación de la OC", "Año Consulta", fecha_actual)
NombreCarpetaOc = InputBox("Introduce el #OC")

If Left(NombreCarpetaOc, 2) = "45" Then
    If Dir(Path & NombreCarpetaAño & "\Nacionales\" & NombreCarpetaOc, vbDirectory) = "" Then
        MsgBox ("La carpeta consultada no existe")
    Else
        Call Shell("explorer.exe " & Path & NombreCarpetaAño & "\Nacionales\" & NombreCarpetaOc, vbNormalFocus)
    End If
ElseIf Left(NombreCarpetaOc, 2) = "55" Then
    If Dir(Path & NombreCarpetaAño & "\Importaciones\" & NombreCarpetaOc, vbDirectory) = "" Then
        MsgBox ("La carpeta consultada no existe")
    Else
        Call Shell("explorer.exe " & Path & NombreCarpetaAño & "\Importaciones\" & NombreCarpetaOc, vbNormalFocus)
    End If
ElseIf Left(NombreCarpetaOc, 2) = "62" Then
    If Dir(Path & NombreCarpetaAño & "\Servicios\" & NombreCarpetaOc, vbDirectory) = "" Then
        MsgBox ("La carpeta consultada no existe")
    Else
        Call Shell("explorer.exe " & Path & NombreCarpetaAño & "\Servicios\" & NombreCarpetaOc, vbNormalFocus)
    End If
End If
End Sub
