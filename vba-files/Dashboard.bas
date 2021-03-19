Attribute VB_Name = "Dashboard"
Dim año As String, libro As String, hoja_temporal As String
Sub crear_tablas()
Application.ScreenUpdating = False
    Call abrir_consolidado
'    Call pareto_monto
'    Call pareto_ga
'    Call consolidado
    Call me2n
    Call abrir_dashboard
    
End Sub
Sub abrir_consolidado()

Workbooks.OpenXML ("C:\Documentos Empresa\OneDrive - MINEROS\Desktop\INDICADORES\dashboard\consolidado.xlsx")
Application.DisplayAlerts = False
With Workbooks("consolidado.xlsx")
    Sheets("consolidado").Activate
    Sheets("consolidado").Range("B2:E13").ClearContents
'    Sheets("pareto_monto").Activate
'    Sheets("pareto_monto").Range("A2:C21").ClearContents
'    Sheets("pareto_ga").Activate
'    Sheets("pareto_ga").Range("A2:C21").ClearContents
'    Sheets("me2n").Activate
    Sheets("me2n").Rows("3:99999").delete
'    Sheets("me2n").Range("A2:Q2", Range("A2:Q2").End(xlDown)).ClearContents
End With
Application.DisplayAlerts = True
End Sub
Sub pareto_monto()
' Genera la tabla Pareto Monto

año = Year(Date)
libro = "Proveedores Pareto " & año & ".xlsx"
Workbooks.OpenXML ("C:\Documentos Empresa\OneDrive - MINEROS\Desktop\INDICADORES\" & año & "\" & libro)
Workbooks(libro).Sheets("pareto_monto_total").Range("A5:C24").Copy
With Workbooks("consolidado.xlsx").Sheets("pareto_monto")
    .Activate
    .Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End With
Application.DisplayAlerts = False
Workbooks(libro).Close
Application.DisplayAlerts = True
End Sub

Sub pareto_ga()
' Genera la tabla Pareto Grupo Artículo

año = Year(Date)
libro = "Pareto por GA.xlsx"
Application.DisplayAlerts = False
Workbooks.OpenXML ("C:\Documentos Empresa\OneDrive - MINEROS\Desktop\INDICADORES\" & año & "\" & libro)
Application.DisplayAlerts = True
Workbooks(libro).Sheets("Materiales").Range("A4:C23").Copy
With Workbooks("consolidado.xlsx").Sheets("pareto_ga")
    .Activate
    .Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End With
Application.DisplayAlerts = False
Workbooks(libro).Close
Application.DisplayAlerts = True
End Sub

Sub consolidado()
' Genera la tabla Consolidado

año = Year(Date)
mes = Month(Date) - 1
Select Case mes
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
libro = "resumen_indicadores.xlsx"
With Workbooks("consolidado.xlsx")
    .Activate
    .Sheets.Add
End With
hoja_temporal = ActiveSheet.Name
Workbooks.OpenXML ("\\vmedsis03\Suministros\Indicadores Compras\" & año & "\Indicadores Mensuales\" & mes & "\" & libro)

' Monto
Workbooks(libro).Sheets("Consolidado").Range("B3:M4").Copy
With Workbooks("consolidado.xlsx").Sheets(hoja_temporal)
    .Activate
    .Range("A1").PasteSpecial xlPasteAll
    .Range("A1:J2").Style = "Currency [0]"
End With

' Tasa Compradores
Workbooks(libro).Sheets("Consolidado").Activate
Workbooks(libro).Sheets("Consolidado").Range("B13:M13").Copy
With Workbooks("consolidado.xlsx").Sheets(hoja_temporal)
    .Activate
    .Range("A3").PasteSpecial xlPasteValues
End With

' Tasa Proveedores
Workbooks(libro).Sheets("Consolidado").Activate
Workbooks(libro).Sheets("Consolidado").Range("B27:M27").Copy
With Workbooks("consolidado.xlsx").Sheets(hoja_temporal)
    .Activate
    .Range("A4").PasteSpecial xlPasteValues
    .Range("A3:J4").Style = "Percent"
End With
Workbooks("consolidado.xlsx").Sheets(hoja_temporal).Range("A1:J4").Copy
With Workbooks("consolidado.xlsx").Sheets("consolidado")
    .Activate
    .Range("B2").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    .Range("J1").value = mes
End With
Application.DisplayAlerts = False
Workbooks("consolidado.xlsx").Sheets(hoja_temporal).delete
Workbooks(libro).Close
Application.DisplayAlerts = True
End Sub

Sub me2n()
' libro me2n_dashboard, en sumninistros. De este archivo se puede sacar paretos, montos, etc.
' Este script, separa el codigo del nombre del proveedor.
libro = "me2n_consolidado.xlsx"
Workbooks.OpenXML ("\\vmedsis03\Suministros\Plantillas\FICHEROS\" & libro)

' Borra compras a 0 valor
Workbooks(libro).Sheets(1).Range("P2", Range("P2").End(xlDown)).Select
conteo = 2
For Each celda In Selection
    If celda.value = 0 Then
        Rows(conteo).ClearContents
    End If
    conteo = conteo + 1
Next
Call organizar_filas

' Separa el codigo y nombre del proveedor
Application.DisplayAlerts = False
With Workbooks(libro).Sheets(1)
    .Columns("L").Insert
    .Range("K1", Range("K1").End(xlDown)).Select
End With
Selection.TextToColumns Destination:=Range("K1"), DataType:=xlFixedWidth, _
    FieldInfo:=Array(Array(0, 1), Array(10, 1)), TrailingMinusNumbers:=True
    
' Borra Intercompany
Workbooks(libro).Sheets(1).Range("K2", Range("K2").End(xlDown)).Select
conteo = 2
For Each celda In Selection
    codigo = celda.value
    If codigo = "1000" Or codigo = "1001" Or codigo = "1002" Or codigo = "1003" Or codigo = "1100" _
    Or codigo = "1200" Or codigo = "1300" Or codigo = "9999" Then
        Rows(conteo).ClearContents
    End If
    conteo = conteo + 1
Next
Call organizar_filas

' Borrar Devoluciones
Workbooks(libro).Sheets(1).Range("F2", Range("F2").End(xlDown)).Select
conteo = 2
For Each celda In Selection
    If celda.value = "ZNB" Then
        Rows(conteo).ClearContents
    End If
    conteo = conteo + 1
Next
Call organizar_filas

' Copia, pega y guarda en la BD
Application.DisplayAlerts = False
Workbooks(libro).Sheets(1).Range("A2:R2", Range("A2:R2").End(xlDown)).Copy
Workbooks("consolidado.xlsx").Sheets("me2n").Range("A2").PasteSpecial xlPasteAll
Workbooks(libro).Close
Workbooks("consolidado.xlsx").SaveAs ("C:\Documentos Empresa\OneDrive - MINEROS\Desktop\INDICADORES\dashboard\consolidado.xlsx")
Workbooks("consolidado.xlsx").Close
Application.DisplayAlerts = True
End Sub
Sub organizar_filas()
' Organiza el archivo de manera Tabular
Workbooks(libro).Sheets(1).Cells.Select
Selection.AutoFilter
Workbooks(libro).Sheets(1).AutoFilter.Sort.SortFields.Add Key:=Range("A1:A9999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With Workbooks(libro).Sheets(1).AutoFilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
Workbooks(libro).Sheets(1).Cells.Select
Selection.AutoFilter
End Sub
Sub abrir_dashboard()
Dim Shex As Object
Set Shex = CreateObject("Shell.Application")
tgtfile = "C:\Documentos Empresa\OneDrive - MINEROS\Desktop\DASHBOARD.pbix"
Shex.Open (tgtfile)
End Sub
