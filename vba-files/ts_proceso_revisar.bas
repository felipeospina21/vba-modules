Attribute VB_Name = "ts_proceso_revisar"
' EL siguiente proceso funciona bien.Revisar como implementarlo (hace todo el calculo de la TS
'------------------------------------------------
Sub copia_pega()

'Copia y pega en la BDatos ela archivo ppal y corrige el formato de las fechas (entrega y Migo)

    Range("A2:Z2", Range("A2:Z2").End(xlDown)).Copy
    
    Workbooks.OpenXML ("\\vmedsis03\Suministros\Plantillas\formatos\tasa_proveedor.xlsx")
    Workbooks("tasa_proveedor.xlsx").Sheets("BDATOS").Range("A2").PasteSpecial
    
    Application.DisplayAlerts = False
    Workbooks("pedido_compras(indicadores).xls").Close
    Application.DisplayAlerts = True

End Sub

Sub ts_mes_seleccionado()

'funciona ok ( en la formula original del campo calculado es un si.error, aca esta solo un condicional si)
' falta completar la BD (solo esta febrero) y mirar como generar la hoja de TS apartir de acá.

' Inicialmente solo debe haber 2 hojas, BD y TS 6 meses. La TS generada dependerá del periodo de la fecha de entrega en la BD.

' Macro1 Macro
'crea tabla
    
    
    Sheets.Add
    nom_hoja = ActiveSheet.Name
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Tabla1", Version:=xlPivotTableVersion15).CreatePivotTable TableDestination:= _
        Sheets(nom_hoja).Cells(2, 1), TableName:="Tabla dinámica2", DefaultVersion:= _
        xlPivotTableVersion15


'pone campos en tabla

'
    With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Nombre Proveedor")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Proveedor")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("OC UNIFICADA")
        .Orientation = xlRowField
        .Position = 3
    End With
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Nombre Proveedor"). _
        LayoutForm = xlTabular
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Proveedor").LayoutForm _
        = xlTabular

    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Proveedor").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Proveedor"). _
        RepeatLabels = True
    ActiveSheet.PivotTables("Tabla dinámica2").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica2").PivotFields("Cumple"), "Suma de Cumple", xlSum
    ActiveSheet.PivotTables("Tabla dinámica2").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica2").PivotFields("Entrega"), "Suma de Entrega", xlSum
        
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Nombre Proveedor"). _
        RepeatLabels = True
 
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Nombre Proveedor"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)



' copia en formato plano la TD generada

    Range("A2", "C2").Select

    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("G2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False



'condicional suma cumple

    
    Cells(2, 10).value = "Suma de Cumple"
    Cells(2, 11).value = "Suma de Entrega"
    
    conteo = Range("I3", Range("I3").End(xlDown)).Count
    n = 3
    For i = 1 To conteo Step 1
    
        If Cells(n, 4) < 1 Then
            Cells(n, 10).value = 0
        Else
            Cells(n, 10).value = 1
        End If
        '----
        If Cells(n, 5) < 1 Then
            Cells(n, 11).value = 0
        Else
            Cells(n, 11).value = 1
        End If
        
        
        n = n + 1
    Next
    
    n = 3



' agrega TD del rango plano generado en Macro7

 
  
    Set rango_campos = Range("G2").CurrentRegion
    Sheets.Add
    If BuscarHoja1("ts_mes") = True Then
        ActiveSheet.Name = "ts_semestre"
    Else
        ActiveSheet.Name = "ts_mes"
    End If
'    ActiveSheet.Name = "ts_mes" & fecha   'revisar para eliminar error por hoja repetida
    
    nom_hoja2 = ActiveSheet.Name
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        rango_campos, Version:=xlPivotTableVersion15).CreatePivotTable TableDestination:= _
        Sheets(nom_hoja2).Cells(2, 1), TableName:="Tabla dinámica2", DefaultVersion:= _
        xlPivotTableVersion15




'
' campos Macro
'agrega campos

    
     With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Proveedor")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Etiquetas de fila" _
        )
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("Tabla dinámica2").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica2").PivotFields("Suma de Cumple"), "Cuenta de Suma de Cumple", _
        xlSum
    ActiveSheet.PivotTables("Tabla dinámica2").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica2").PivotFields("Suma de Entrega"), "Cuenta de Suma de Entrega" _
        , xlSum

   
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Proveedor").LayoutForm _
        = xlTabular
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Etiquetas de fila"). _
        LayoutForm = xlTabular

    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Etiquetas de fila"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)

    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Proveedor").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )


'
' Macro13 Macro
'campo calculado para %

'
  
    ActiveSheet.PivotTables("Tabla dinámica2").CalculatedFields.Add "%", _
        "='Suma de Cumple' /'Suma de Entrega'", True
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("%").Orientation = _
        xlDataField
    With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Suma de %")
        .NumberFormat = "0%"
    End With

    fecha_entrega = Sheets("BDATOS").Range("W2").value ' Pone la fecha que estoy consultando para luego ser comparada y generar o no la TS
    Sheets("ts_mes").Range("A1").value = fecha_entrega
        
    Application.DisplayAlerts = False
    Sheets(nom_hoja).Visible = xlHidden
    Application.DisplayAlerts = True
End Sub

Function BuscarHoja1(nombreHoja As String) As Boolean
 
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = nombreHoja Then
            BuscarHoja1 = True
            Exit Function
        End If
    Next
     
    BuscarHoja1 = False
 
End Function

