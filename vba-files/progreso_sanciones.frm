VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} progreso_sanciones 
   Caption         =   "UserForm1"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6990
   OleObjectBlob   =   "progreso_sanciones.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "progreso_sanciones"
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

Sub UserForm_Activate()

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
Workbooks("informe_sanciones.xlsx").Sheets("criterio").Select
limite = Workbooks("informe_sanciones.xlsx").Sheets("criterio").Range("A3", Selection.End(xlDown)).Count
ProgressBar3.Min = 0
ProgressBar3.Max = limite

For clave = 0 To limite - 1 Step 1
    fila = 0
    
    ' Crea una hoja y pega lo de cada proveedor en esta
    Workbooks("informe_sanciones.xlsx").Activate
    Workbooks("informe_sanciones.xlsx").Worksheets.Add
    hoja_filtro = ActiveSheet.Name
    ActiveWorkbook.Worksheets("hoja_rango").Range("A7:R100000").AdvancedFilter Action:=xlFilterCopy, _
    CriteriaRange:=ActiveWorkbook.Worksheets("criterio").Range("A2:A3"), CopyToRange:=ActiveWorkbook.Sheets(hoja_filtro).Range("A1"), Unique:=False

    ' Formato del cuadro
    Call formato_facturas
        
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItemFromTemplate("\\vmedsis03\Suministros\Plantillas\outlook\multas.oft")
    Set mirango = Workbooks("informe_sanciones.xlsx").Sheets("criterio").Range("C3")
    email = mirango.value
    prov_name = mirango.Offset(0, -1).value
    If email = "" Or email = "#N/A" Or email = "0" Then
        With OutMail
            .To = "felipe.ospina@mineros.com.co"
            .Attachments.Add ("\\vmedsis03\Suministros\Plantillas\adjuntos\sanciones.xlsx")
            .Send
        End With
    Else
        With OutMail
            '.To = "felipe.ospina@mineros.com.co"
            .To = email
            .Attachments.Add ("\\vmedsis03\Suministros\Plantillas\adjuntos\sanciones.xlsx")
            .Send
        End With
    End If
    Workbooks("informe_sanciones.xlsx").Sheets("criterio").Range("A3:xfd3").delete
    Range("A2").Select
    DoEvents
    ProgressBar3.value = clave
    fila = fila + 1
Next
progreso_sanciones.Hide
Workbooks("informe_sanciones.xlsx").SaveAs ("\\vmedsis03\Suministros\Seguimientos\Sanciones\" & año & "\informe_sanciones(" & mes & ").xlsx")
MsgBox ("Proceso Finalizado")
End Sub

Sub formato_facturas()
'Formato del cuadro
Application.ScreenUpdating = False
Workbooks("informe_sanciones.xlsx").Activate
Workbooks("informe_sanciones.xlsx").Sheets(hoja_filtro).Columns("A:R").Select
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
Workbooks("informe_sanciones.xlsx").Worksheets(hoja_filtro).Copy
Application.DisplayAlerts = False
ActiveWorkbook.SaveAs ("\\vmedsis03\Suministros\Plantillas\adjuntos\sanciones.xlsx")
Workbooks("sanciones.xlsx").Close
Workbooks("informe_sanciones.xlsx").Sheets(hoja_filtro).delete
Application.DisplayAlerts = True
End Sub


