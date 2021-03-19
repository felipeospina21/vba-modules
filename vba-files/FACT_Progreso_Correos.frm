VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FACT_Progreso_Correos 
   Caption         =   "Enviando Correos"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6570
   OleObjectBlob   =   "FACT_Progreso_Correos.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "FACT_Progreso_Correos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OutApp As Outlook.Application
Dim OutMail As Outlook.MailItem
Dim email As String, prov_name As String, milibro As String, mihoja As String, hoja_rango As String, hoja_filtro As String, nombre_libro_informe As String
Dim lookupvalue As Variant, mimatriz(1000000) As Variant, clave As Variant
Dim mirango As Range, lookuprange As Range, micelda As Range
Dim fila As Integer, i As Integer, indice As Integer
Dim fecha_actual As String, fecha_dia As String, fecha_mes As String
Dim limite As Long, limitex As Long

Sub formato_facturas()
'Formato del cuadro
                    Application.ScreenUpdating = False
                    Workbooks("seguimiento_facturas.xlsx").Activate
                    Workbooks("seguimiento_facturas.xlsx").Sheets(hoja_filtro).Columns("A:J").Select ' funciona bien
                    Selection.Columns.AutoFit
                    
                    ActiveSheet.Range("A2").CurrentRegion.Select
                    
                        With Selection.Borders(xlEdgeLeft)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlEdgeTop)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlEdgeBottom)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlEdgeRight)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlInsideVertical)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlInsideHorizontal)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        
        Workbooks("seguimiento_facturas.xlsx").Worksheets(hoja_filtro).Copy
        Application.DisplayAlerts = False
            ActiveWorkbook.SaveAs ("\\vmedsis03\Suministros\Plantillas\adjuntos\facturas_pendientes.xlsx")
            Workbooks("facturas_pendientes.xlsx").Close
        Application.DisplayAlerts = True
        
        Application.DisplayAlerts = False
            Workbooks("seguimiento_facturas.xlsx").Sheets(hoja_filtro).delete
        Application.DisplayAlerts = True
End Sub

Sub UserForm_Activate()

Application.ScreenUpdating = False

'Defino clave de busqueda
Workbooks("seguimiento_facturas.xlsx").Activate
Workbooks("seguimiento_facturas.xlsx").Sheets("criterio").Range("A3", Range("A3").End(xlDown)).Select

indice = 0
For Each micelda In Selection
    mimatriz(indice) = micelda.value
    indice = indice + 1
Next micelda

'----------------------------------
        
'Siempre verificar primero correos
Workbooks("seguimiento_facturas.xlsx").Sheets("criterio").Select
limite = Workbooks("seguimiento_facturas.xlsx").Sheets("criterio").Range("A3", Selection.End(xlDown)).Count
ProgressBar2.Min = 0
ProgressBar2.Max = limite

For clave = 0 To limite - 1 Step 1 ' genero un bucle for para darle diferentes valores a "i" que sera el indice de la matriz
    fila = 0
                   
' Application.ScreenUpdating = False
    '------o-------------------------
                
    'Crea una hoja y pega lo de cada proveedor en esta
        Workbooks("seguimiento_facturas.xlsx").Activate
        Workbooks("seguimiento_facturas.xlsx").Worksheets.Add
        hoja_filtro = ActiveSheet.Name
        ActiveWorkbook.Worksheets("hoja_rango").Range("A1:J100000").AdvancedFilter Action:=xlFilterCopy, _
        CriteriaRange:=ActiveWorkbook.Worksheets("criterio").Range("A2:A3"), CopyToRange:=ActiveWorkbook.Sheets(hoja_filtro).Range("A1"), Unique:=False

            
                ' Formato del cuadro
                  Call formato_facturas
        
        Set OutApp = CreateObject("Outlook.Application")
        Set OutMail = OutApp.CreateItemFromTemplate("\\vmedsis03\Suministros\Plantillas\outlook\facturas pendientes.oft")
        Set mirango = Workbooks("seguimiento_facturas.xlsx").Sheets("criterio").Range("C3")
        
        email = mirango.value
        prov_name = mirango.Offset(0, -1).value
            
            If email = "" Or email = "#N/A" Or email = "0" Then
                With OutMail
                    .To = "felipe.ospina@mineros.com.co"
                    .Subject = "Facturas Pendientes " & prov_name
                    .Attachments.Add ("\\vmedsis03\Suministros\Plantillas\adjuntos\facturas_pendientes.xlsx")
                    .Send
                End With
            
            Else
            
                With OutMail
'                    .To = "felipe.ospina@mineros.com.co"
                    .To = email
                    .Subject = "Facturas Pendientes " & prov_name
                    .Attachments.Add ("\\vmedsis03\Suministros\Plantillas\adjuntos\facturas_pendientes.xlsx")
                    .Send
                End With
            
            End If
                
        Workbooks("seguimiento_facturas.xlsx").Sheets("criterio").Range("A3:xfd3").delete
        Range("A2").Select
        DoEvents
        ProgressBar2.value = clave
        fila = fila + 1
Next

FACT_Progreso_Correos.Hide

MsgBox ("Proceso Finalizado")
End Sub





