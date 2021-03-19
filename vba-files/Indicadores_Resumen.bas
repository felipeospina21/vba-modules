Attribute VB_Name = "Indicadores_Resumen"
Dim wb_c As Workbook, wb_p As Workbook, wb_resumen As Workbook
Dim a As String
Dim columna As String
Dim Shex As Object
Dim PP As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Dim PPSlide As PowerPoint.Slide

Sub resumen_indicadores()
' Genera el archivo resumen para analisis compradores

' Falta grabar en suministros el archivo resumen, creando carpetas si no existe

' Abre la plantilla de resumen indicadores
Application.ScreenUpdating = False
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
Application.DisplayAlerts = False
Workbooks.OpenXML ("C:\Documentos Empresa\OneDrive - MINEROS\Desktop\INDICADORES\" & año & "\" & mes & "\Ts_Comprador(" & mes & ").xlsx")
Workbooks.OpenXML ("C:\Documentos Empresa\OneDrive - MINEROS\Desktop\INDICADORES\" & año & "\" & mes & "\Ts_Proveedor(" & mes & ").xlsx")
Workbooks.OpenXML ("\\vmedsis03\Suministros\Plantillas\formatos\resumen_indicadores.xlsx")
Application.DisplayAlerts = True

Set wb_c = Workbooks("Ts_Comprador(" & mes & ").xlsx")
Set wb_p = Workbooks("Ts_Proveedor(" & mes & ").xlsx")
Set wb_resumen = Workbooks("resumen_indicadores.xlsx")

' filtra los incumplimientos en compras, copia y pega
wb_c.Sheets("BD").Activate
wb_c.Sheets("BD").ListObjects("Tabla1").Sort.SortFields.Clear
wb_c.Sheets("BD").ListObjects("Tabla1").Sort.SortFields.Add Key:=Range("Tabla1[[#All],[cumplimiento]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
With wb_c.Sheets("BD").ListObjects("Tabla1").Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
With wb_c.Sheets("BD")
    .ListObjects("Tabla1").Range.AutoFilter Field:=30, Criteria1:="0"
    .Activate
    .Range("A1").Select
    .Range("A2:Y2", Range("A2:Y2").End(xlDown)).Copy
End With

wb_resumen.Sheets("Análisis_Compras").Activate
wb_resumen.Sheets("Análisis_Compras").Range("A2").PasteSpecial xlAll
Application.DisplayAlerts = False
wb_c.Close
Application.DisplayAlerts = True

' filtra los incumplimientos en entregas, copia y pega
wb_p.Sheets("BDATOS").Activate
wb_p.Sheets("BDATOS").ListObjects("Tabla1").Sort.SortFields.Clear
wb_p.Sheets("BDATOS").ListObjects("Tabla1").Sort.SortFields.Add Key:=Range("Tabla1[[#All],[cumplimiento]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
With wb_p.Sheets("BDATOS").ListObjects("Tabla1").Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
With wb_p.Sheets("BDATOS")
    .ListObjects("Tabla1").Range.AutoFilter Field:=30, Criteria1:="0"
    .Activate
    .Range("A1").Select
    .Range("A2:X2", Range("A2:X2").End(xlDown)).Copy
End With

wb_resumen.Sheets("Análisis_Entrega").Activate
wb_resumen.Sheets("Análisis_Entrega").Range("A2").PasteSpecial xlAll
Application.DisplayAlerts = False
wb_p.Close
Application.DisplayAlerts = True
End Sub

Sub consolidado()
' Consolida el análisis, crea plantilla power BI y pptx.

    Call consolidar_resumen
    Call plantilla_power_bi
    Call presentacion_indicadores
End Sub

Sub consolidar_resumen()
' Una vez verificado el análisis, consolida en las tablas resumen y consolidado

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

'Workbooks.OpenXML ("\\vmedsis03\Suministros\Indicadores Compras\" & año & "\Indicadores Mensuales\" & mes & "\resumen_indicadores.xlsx")
Set wb_bd = Workbooks("resumen_indicadores.xlsx")
Select Case mes
    Case "Enero": columna = "B"
    Case "Febrero": columna = "C"
    Case "Marzo": columna = "D"
    Case "Abril": columna = "E"
    Case "Mayo": columna = "F"
    Case "Junio": columna = "G"
    Case "Julio": columna = "H"
    Case "Agosto": columna = "I"
    Case "Septiembre": columna = "J"
    Case "Octubre": columna = "K"
    Case "Noviembre": columna = "L"
    Case "Diciembre": columna = "M"
End Select
Application.DisplayAlerts = False
Workbooks.OpenXML ("C:\Documentos Empresa\OneDrive - MINEROS\Desktop\INDICADORES\" & año & "\" & mes & "\Ts_Comprador(" & mes & ").xlsx")
Workbooks("Ts_Comprador(" & mes & ").xlsx").Sheets("Resumen").Range("B3", "B23").Copy
wb_bd.Sheets("Resumen_Compras").Activate
wb_bd.Sheets("Resumen_Compras").Range("B3").PasteSpecial xlValue
Workbooks("Ts_Comprador(" & mes & ").xlsx").Close

Workbooks.OpenXML ("C:\Documentos Empresa\OneDrive - MINEROS\Desktop\INDICADORES\" & año & "\" & mes & "\Ts_Proveedor(" & mes & ").xlsx")
Workbooks("Ts_Proveedor(" & mes & ").xlsx").Sheets("RESUMEN ENTREGAS").Range("B2", "B19").Copy
wb_bd.Sheets("Resumen_Entregas").Activate
wb_bd.Sheets("Resumen_Entregas").Range("B3").PasteSpecial xlValue
Workbooks("Ts_Proveedor(" & mes & ").xlsx").Close
With wb_bd
' Compras
    .Sheets("Resumen_Compras").Range("B3", "B5").Copy
    .Sheets("Consolidado").Range(columna & "3").PasteSpecial xlValue
    .Sheets("Resumen_Compras").Range("B12", "B15").Copy
    .Sheets("Consolidado").Range(columna & "6").PasteSpecial xlValue
    .Sheets("Resumen_Compras").Range("B9", "B11").Copy
    .Sheets("Consolidado").Range(columna & "11").PasteSpecial xlValue
    .Sheets("Resumen_Compras").Range("B23").Copy
    .Sheets("Consolidado").Range(columna & "10").PasteSpecial xlValue
    .Sheets("Resumen_Compras").Range("B6", "B8").Copy
    .Sheets("Resumen").Range(columna & "2").PasteSpecial xlValue
' Entregas
    .Sheets("Resumen_Entregas").Range("B12", "B15").Copy
    .Sheets("Consolidado").Range(columna & "15").PasteSpecial xlValue
    .Sheets("Resumen_Entregas").Range("B9", "B11").Copy
    .Sheets("Consolidado").Range(columna & "20").PasteSpecial xlValue
    .Sheets("Resumen_Entregas").Range("B20").Copy
    .Sheets("Consolidado").Range(columna & "19").PasteSpecial xlValue
    .Sheets("Resumen_Entregas").Range("B3").Copy
    .Sheets("Resumen").Range(columna & "5").PasteSpecial xlValue
    .Sheets("Resumen_Entregas").Range("B7", "B8").Copy
    .Sheets("Resumen").Range(columna & "6").PasteSpecial xlValue
End With
Application.DisplayAlerts = True
End Sub

Sub plantilla_power_bi()
' Actualiza la plantilla de power BI

Workbooks.OpenXML ("C:\Documentos Empresa\OneDrive - MINEROS\Desktop\INDICADORES\plantilla_PowerBI.xlsx")

Workbooks("resumen_indicadores.xlsx").Sheets("Resumen_Entregas").Activate
Workbooks("resumen_indicadores.xlsx").Sheets("Resumen_Entregas").Range("B3", "B20").Copy
Workbooks("plantilla_PowerBI.xlsx").Sheets("Entregas").Activate
Workbooks("plantilla_PowerBI.xlsx").Sheets("Entregas").Range("A2").PasteSpecial Paste:=xlPasteAll, Transpose:=True

Workbooks("resumen_indicadores.xlsx").Sheets("Resumen_Compras").Activate
Workbooks("resumen_indicadores.xlsx").Sheets("Resumen_Compras").Range("B3", "B23").Copy
Workbooks("plantilla_PowerBI.xlsx").Sheets("Compras").Activate
Workbooks("plantilla_PowerBI.xlsx").Sheets("Compras").Range("A2").PasteSpecial Paste:=xlPasteAll, Transpose:=True

Application.DisplayAlerts = False
Workbooks("plantilla_PowerBI.xlsx").SaveAs ("C:\Documentos Empresa\OneDrive - MINEROS\Desktop\INDICADORES\plantilla_PowerBI.xlsx")
Workbooks("plantilla_PowerBI.xlsx").Close
Application.DisplayAlerts = True

Set Shex = CreateObject("Shell.Application")
tgtfile = "C:\Documentos Empresa\OneDrive - MINEROS\Desktop\INDICADORES\Indicadores.pbix"
Shex.Open (tgtfile)
End Sub

Sub presentacion_indicadores()
' Crea pptx con gráficos

Set PP = CreateObject("PowerPoint.Application")
Set PPPres = PP.Presentations.Open("\\vmedsis03\Suministros\Plantillas\formatos\revision_gerencial.pptx")
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

' Copia gráficas a pptx
On Error Resume Next
' falta que abra o no el archivo resumen_indicadores
Workbooks("resumen_indicadores.xlsx").Sheets("Resumen").ChartObjects(1).Chart.CopyPicture Size:=xlScreen, Format:=xlPicture
PPPres.Slides.Range(3).Shapes.Paste
Workbooks("resumen_indicadores.xlsx").Sheets("Consolidado").ChartObjects(1).Chart.CopyPicture Size:=xlScreen, Format:=xlPicture
PPPres.Slides.Range(4).Shapes.Paste
Workbooks("resumen_indicadores.xlsx").Sheets("Grafica_C").ChartObjects(1).Chart.CopyPicture Size:=xlScreen, Format:=xlPicture
PPPres.Slides.Range(7).Shapes.Paste
Workbooks("resumen_indicadores.xlsx").Sheets("Grafica_E").ChartObjects(1).Chart.CopyPicture Size:=xlScreen, Format:=xlPicture
PPPres.Slides.Range(11).Shapes.Paste

Workbooks.OpenXML ("C:\Documentos Empresa\OneDrive - MINEROS\Desktop\INDICADORES\" & año & "\" & mes & "\Ts_Comprador(" & mes & ").xlsx")
Workbooks("Ts_Comprador(" & mes & ").xlsx").Sheets("TS_Comprador").ChartObjects(1).Chart.CopyPicture Size:=xlScreen, Format:=xlPicture
PPPres.Slides.Range(8).Shapes.Paste
Application.DisplayAlerts = False
Workbooks("Ts_Comprador(" & mes & ").xlsx").Close

Workbooks.OpenXML ("C:\Documentos Empresa\OneDrive - MINEROS\Desktop\INDICADORES\" & año & "\" & mes & "\Ts_Proveedor(" & mes & ").xlsx")
Workbooks("Ts_Proveedor(" & mes & ").xlsx").Sheets("Incumplimientos_Prov_Pareto").ChartObjects(1).Chart.CopyPicture Size:=xlScreen, Format:=xlPicture
PPPres.Slides.Range(12).Shapes.Paste
Workbooks("Ts_Proveedor(" & mes & ").xlsx").Close

Workbooks.OpenXML ("C:\Documentos Empresa\OneDrive - MINEROS\Desktop\INDICADORES\" & año & "\" & mes & "\indicadores_servicios(" & mes & ").xlsx")
Workbooks("indicadores_servicios(" & mes & ").xlsx").Sheets("Cantidad x Clasificación").ChartObjects(1).Chart.CopyPicture Size:=xlScreen, Format:=xlPicture
PPPres.Slides.Range(13).Shapes.Paste
Workbooks("indicadores_servicios(" & mes & ").xlsx").Sheets("Dias en contratar").ChartObjects(1).Chart.CopyPicture Size:=xlScreen, Format:=xlPicture
PPPres.Slides.Range(14).Shapes.Paste
Workbooks("indicadores_servicios(" & mes & ").xlsx").Sheets("Servicios x comp").ChartObjects(1).Chart.CopyPicture Size:=xlScreen, Format:=xlPicture
PPPres.Slides.Range(15).Shapes.Paste
Workbooks("indicadores_servicios(" & mes & ").xlsx").Sheets("Dias").ChartObjects(1).Chart.CopyPicture Size:=xlScreen, Format:=xlPicture
PPPres.Slides.Range(16).Shapes.Paste
Workbooks("indicadores_servicios(" & mes & ").xlsx").Sheets("Dias2").ChartObjects(1).Chart.CopyPicture Size:=xlScreen, Format:=xlPicture
PPPres.Slides.Range(17).Shapes.Paste
Workbooks("indicadores_servicios(" & mes & ").xlsx").Close

Application.DisplayAlerts = True
End Sub
