Attribute VB_Name = "Divisas"
Dim wb_divisas As Workbook, rango_tabla As Range, rango_fechas As Range, nom_grafico As String

Sub grafico_divisas()
' Crea gráficos del comportamiento de las divisas

Workbooks.OpenXML ("\\vmedsis03\Suministros\Plantillas\formatos\divisas.xlsx")
Workbooks.OpenXML ("C:\Documentos Empresa\OneDrive - MINEROS\Desktop\Automatizaciones\web_scrapping\historico.xlsx")

' USD
Set wb_divisas = Workbooks("historico.xlsx")
Set rango_tabla = wb_divisas.Sheets(1).Range("B2", Range("B2").End(xlDown))
Set rango_fechas = wb_divisas.Sheets(1).Range("A2", Range("A2").End(xlDown))

Workbooks("divisas.xlsx").Sheets("usd").Activate
Workbooks("divisas.xlsx").Sheets("usd").Shapes.AddChart2(332, xlLineMarkers).Select

With ActiveChart
    .SeriesCollection.NewSeries
    .FullSeriesCollection(1).Name = "USD"
    .FullSeriesCollection(1).Values = rango_tabla
    .FullSeriesCollection(1).XValues = rango_fechas
    .Axes(xlValue).MinimumScale = 3200
    '.Location Where:=xlLocationAsNewSheet, Name:="USD2"
End With
With ActiveSheet.ChartObjects
    .Width = 900
    .Height = 320
    .Top = 1
    .Left = 1
End With

' EUR
wb_divisas.Activate
Set rango_tabla = wb_divisas.Sheets(1).Range("C2", Range("C2").End(xlDown))
Set rango_fechas = wb_divisas.Sheets(1).Range("A2", Range("A2").End(xlDown))

Workbooks("divisas.xlsx").Sheets("eur").Activate
Workbooks("divisas.xlsx").Sheets("eur").Shapes.AddChart2(332, xlLineMarkers).Select

With ActiveChart
    .SeriesCollection.NewSeries
    .FullSeriesCollection(1).Name = "EUR"
    .FullSeriesCollection(1).Values = rango_tabla
    .FullSeriesCollection(1).XValues = rango_fechas
    .Axes(xlValue).MinimumScale = 3600
End With
With ActiveSheet.ChartObjects
    .Width = 900
    .Height = 320
    .Top = 1
    .Left = 1
End With

' AUD
wb_divisas.Activate
Set rango_tabla = wb_divisas.Sheets(1).Range("D2", Range("D2").End(xlDown))
Set rango_fechas = wb_divisas.Sheets(1).Range("A2", Range("A2").End(xlDown))

Workbooks("divisas.xlsx").Sheets("aud").Activate
Workbooks("divisas.xlsx").Sheets("aud").Shapes.AddChart2(332, xlLineMarkers).Select

With ActiveChart
    .SeriesCollection.NewSeries
    .FullSeriesCollection(1).Name = "AUD"
    .FullSeriesCollection(1).Values = rango_tabla
    .FullSeriesCollection(1).XValues = rango_fechas
    .Axes(xlValue).MinimumScale = 2200
End With
With ActiveSheet.ChartObjects
    .Width = 900
    .Height = 320
    .Top = 1
    .Left = 1
End With

' CAD
wb_divisas.Activate
Set rango_tabla = wb_divisas.Sheets(1).Range("E2", Range("E2").End(xlDown))
Set rango_fechas = wb_divisas.Sheets(1).Range("A2", Range("A2").End(xlDown))

Workbooks("divisas.xlsx").Sheets("cad").Activate
Workbooks("divisas.xlsx").Sheets("cad").Shapes.AddChart2(332, xlLineMarkers).Select

With ActiveChart
    .SeriesCollection.NewSeries
    .FullSeriesCollection(1).Name = "CAD"
    .FullSeriesCollection(1).Values = rango_tabla
    .FullSeriesCollection(1).XValues = rango_fechas
    .Axes(xlValue).MinimumScale = 2400
End With
With ActiveSheet.ChartObjects
    .Width = 900
    .Height = 320
    .Top = 1
    .Left = 1
End With
Application.DisplayAlerts = False
Workbooks("historico.xlsx").Close
Application.DisplayAlerts = True
End Sub
