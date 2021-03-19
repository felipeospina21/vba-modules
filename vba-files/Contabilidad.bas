Attribute VB_Name = "Contabilidad"
Dim Hoja1 As String, Hoja2 As String, mes1 As String, mes2 As String, libro As String
Sub Informe()
Application.ScreenUpdating = False

    Call abrir_consolc
    Call abrir_informe_cont
    Call copiar_rango
    Call cerrar_consolc
'    Call borrar_tabla
'    Call informe_cont
    
    
End Sub

Sub abrir_consolc()

'Abre base y borra filas y columnas

    Workbooks.OpenXML ("C:\Documentos Empresa\OneDrive - MINEROS\Desktop\FICHEROS\consol_compras (indicadores).xls")
    
    With Workbooks("consol_compras (indicadores)").Sheets(1)
        .Columns("A").delete
        .Rows("1:3").delete
        .Rows("2").delete
    End With
End Sub

Sub abrir_informe_cont()

'Abre formato informe

    Workbooks.OpenXML ("C:\Documentos Empresa\OneDrive - MINEROS\Desktop\Automatizaciones\formatos\informe_contabilidad.xlsx")
End Sub

Sub copiar_rango()

'Copia la base filtrada en las respectivas tablas del formato

Set wb = Workbooks("consol_compras (indicadores)")
        
        wb.Sheets(1).Activate
        wb.Sheets(1).Range("B2", Range("B2").End(xlDown)).Select
        conteo = 2
        For Each celda In Selection
        
            texto = celda.value
            y = Left(texto, 1)
            
            If y = "2" Then
                Range("Z" & conteo).value = "I"
            
            Else
                Range("Z" & conteo).value = "N"
                
            End If
            conteo = conteo + 1
        Next
        
'---- Pega Internacionales -------
'    With wb.Sheets(1)
'        .Cells.Select
'        .Rows.AutoFilter
'        .Range("$A$1:$Y$99999").AutoFilter Field:=26, Criteria1:="I"
'        .Range("A1:Y1", Range("A1:Y1").End(xlDown)).Copy
'    End With
'
'    Workbooks("informe_contabilidad").Sheets("BD (I)").Range("A6").PasteSpecial xlAll
'
''---- Pega Nacionales -------
'    With wb.Sheets(1)
''        .Cells.Select
''        .Rows.AutoFilter
'        .Range("$A$1:$Y$99999").AutoFilter Field:=26, Criteria1:="N"
'        .Range("A1:Y1", Range("A1:Y1").End(xlDown)).Copy
'    End With
'
'    Workbooks("informe_contabilidad").Sheets("BD (N)").Range("A6").PasteSpecial xlAll
End Sub

Sub cerrar_consolc()

'cierra base

    Application.DisplayAlerts = False
    Workbooks("consol_compras (indicadores)").Close
    Application.DisplayAlerts = True
    
End Sub

Sub borrar_tabla()

Set wb = Workbooks("informe_contabilidad")
    
    Application.DisplayAlerts = False
    wb.Sheets("Compras Nacionales").delete
    wb.Sheets("Compras Internacionales").delete
    wb.Sheets("BD (N)").Activate
    wb.Sheets("BD (N)").Range("A6:Y6", Range("A6:Y6").End(xlDown)).ClearContents
    wb.Sheets("BD (I)").Activate
    wb.Sheets("BD (I)").Range("A6:Y6", Range("A6:Y6").End(xlDown)).ClearContents
    Application.DisplayAlerts = True
    

End Sub
Sub informe_cont()

' genera las tablas dinamicas de cada base de datos

Set wb = Workbooks("informe_contabilidad")

    wb.Activate
    
    wb.Sheets.Add
    mes1 = "Compras Internacionales"
    ActiveSheet.Name = mes1 ' aca se usará esta variable para cambiar el nombre de acuerdo al mes
    Hoja1 = ActiveSheet.Name

    
    wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "BD_2", Version:=xlPivotTableVersion15).CreatePivotTable TableDestination:= _
        Sheets(Hoja1).Cells(2, 1), TableName:="Tabla dinámica3", DefaultVersion:= _
        xlPivotTableVersion15
    
    
    Sheets.Add
    mes2 = "Compras Nacionales"
    ActiveSheet.Name = mes2 ' aca se usará esta variable para cambiar el nombre de acuerdo al mes
    Hoja2 = ActiveSheet.Name
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "BD", Version:=xlPivotTableVersion15).CreatePivotTable TableDestination:= _
        Sheets(Hoja2).Cells(2, 1), TableName:="Tabla dinámica3", DefaultVersion:= _
        xlPivotTableVersion15
'---------------------------------------------------------------------------
'agrega los campos a la tabla generada

  With wb.Sheets(Hoja1).PivotTables("Tabla dinámica3").PivotFields( _
        "Organizaciòn Compra")
        .Orientation = xlRowField
        .Position = 1
    End With
'    With Sheets(Hoja1).PivotTables("Tabla dinámica3").PivotFields("Semana")
'        .Orientation = xlRowField
'        .Position = 2
'    End With
    Sheets(Hoja1).PivotTables("Tabla dinámica3").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica3").PivotFields("TOTAL en COP"), "Suma de TOTAL en COP", xlSum
    Sheets(Hoja1).PivotTables("Tabla dinámica3").PivotFields("Organizaciòn Compra"). _
        LayoutForm = xlTabular
    ActiveSheet.Range("A5").Select
    Sheets(Hoja1).PivotTables("Tabla dinámica3").PivotFields("Organizaciòn Compra"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    With Sheets(Hoja1).PivotTables("Tabla dinámica3").PivotFields( _
        "Suma de TOTAL en COP")
        .NumberFormat = "#,###,###"
    End With
    
    'adiciona la gráfica
    Sheets(Hoja1).Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    
    With ActiveChart
        .SetSourceData Source:=ActiveSheet.Range("$A$2:$C$30")
        .Legend.delete
        .ApplyDataLabels
        .ShowAllFieldButtons = False
        .ChartTitle.Text = mes1
        .ChartColor = 6
    '4 naranja
    End With
    
    With ActiveChart.ChartArea
        .Height = 400
        .Width = 600
        .Top = 10
        .Left = 300
    End With
    
'------
Sheets(Hoja2).Select
  With Sheets(Hoja2).PivotTables("Tabla dinámica3").PivotFields( _
        "Organizaciòn Compra")
        .Orientation = xlRowField
        .Position = 1
    End With
'    With Sheets(hoja2).PivotTables("Tabla dinámica3").PivotFields("Semana")
'        .Orientation = xlRowField
'        .Position = 2
'    End With
    Sheets(Hoja2).PivotTables("Tabla dinámica3").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica3").PivotFields("TOTAL en COP"), "Suma de TOTAL en COP", xlSum
    Sheets(Hoja2).PivotTables("Tabla dinámica3").PivotFields("Organizaciòn Compra"). _
        LayoutForm = xlTabular
    ActiveSheet.Range("A5").Select
    Sheets(Hoja2).PivotTables("Tabla dinámica3").PivotFields("Organizaciòn Compra"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    With ActiveSheet.PivotTables("Tabla dinámica3").PivotFields( _
        "Suma de TOTAL en COP")
        .NumberFormat = "#,###,###"
    End With
    
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    With ActiveChart
        .SetSourceData Source:=ActiveSheet.Range("$A$2:$C$30")
        .Legend.delete
        .ApplyDataLabels
        .ShowAllFieldButtons = False
        .ChartTitle.Text = mes2
        .ChartColor = 8
    '6 amarillo
    '8 verde
    End With
    
    With ActiveChart.ChartArea
        .Height = 400
        .Width = 600
        .Top = 10
        .Left = 300
    End With

'Guardar archivo
Application.DisplayAlerts = False
 ActiveWorkbook.SaveAs Filename:= _
        "C:\Documentos Empresa\OneDrive - MINEROS\Desktop\Automatizaciones\formatos\informe_contabilidad.xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
Application.DisplayAlerts = True

'correo

Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItemFromTemplate("\\vmedsis03\Suministros\Plantillas\outlook\Informe Contabilidad.oft")
            With OutMail
            .To = "paola.castrillon@mineros.com.co"
            .CC = "carlosfelipe.garcia@mineros.com.co"
            .Subject = "Informe Compras"
            .Attachments.Add ("C:\Documentos Empresa\OneDrive - MINEROS\Desktop\Automatizaciones\formatos\informe_contabilidad.xlsx")
            .Display
            End With
End Sub
