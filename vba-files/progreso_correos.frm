VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} progreso_correos 
   Caption         =   "Enviando"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6390
   OleObjectBlob   =   "progreso_correos.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "progreso_correos"
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
        
' Genero los limites para la barra de progreso
Workbooks("seguimiento_oc.xlsx").Sheets("criterio").Select
limite = Workbooks("seguimiento_oc.xlsx").Sheets("criterio").Range("A3", Range("A3").End(xlDown)).Count
ProgressBar1.Min = 0
ProgressBar1.Max = limite

' Creo archivo por proveedor y lo envía adjunto
For clave = 0 To limite - 1 Step 1
    fila = 0
    Workbooks("seguimiento_oc.xlsx").Worksheets.Add
    hoja_filtro = ActiveSheet.Name
    ActiveWorkbook.Worksheets("hoja_rango").Range("A7:Q100000").AdvancedFilter Action:=xlFilterCopy, _
    CriteriaRange:=ActiveWorkbook.Worksheets("criterio").Range("A2:A3"), CopyToRange:=ActiveWorkbook.Sheets(hoja_filtro).Range("A1"), Unique:=False
    Call formato
     
    'Envío Correos
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItemFromTemplate("\\vmedsis03\Suministros\Plantillas\outlook\Seguimiento(auto).oft")
    Set mirango = Workbooks("seguimiento_oc.xlsx").Sheets("criterio").Range("C3")
    On Error Resume Next
    
    email = mirango.value
    prov_name = mirango.Offset(0, -1).value
    If email = "" Or email = "#N/A" Or email = "0" Then
        With OutMail
            .To = "felipe.ospina@mineros.com.co"
            .Subject = "Seguimiento " & prov_name
            .Attachments.Add ("\\vmedsis03\Suministros\Plantillas\adjuntos\seguimiento.xlsx")
            .Send
        End With
    Else
        With OutMail
            '.To = "felipe.ospina@mineros.com.co"
            .To = email
            .Subject = "Seguimiento " & prov_name
            .Attachments.Add ("\\vmedsis03\Suministros\Plantillas\adjuntos\seguimiento.xlsx")
            .Send
        End With
    End If
Workbooks("seguimiento_oc.xlsx").Sheets("criterio").Range("A3:xfd3").delete
Range("A2").Select
DoEvents
ProgressBar1.value = clave
fila = fila + 1
email = ""
Next
progreso_correos.Hide
MsgBox ("Proceso Finalizado")
End Sub

Sub formato()
' Formato del cuadro
                   
' Amarillo
With Workbooks("seguimiento_oc.xlsx").Sheets(hoja_filtro)
    .Columns("A:Q").AutoFit
    .Columns("M:M").EntireColumn.Hidden = True
End With
With Workbooks("seguimiento_oc.xlsx").Sheets(hoja_filtro).Range("O2:O10000")
    .FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, Formula1:="=1", Formula2:="=8"
    .FormatConditions(Range("O2:O10000").FormatConditions.Count).SetFirstPriority
    .FormatConditions(1).Interior.PatternColorIndex = xlAutomatic
    .FormatConditions(1).Interior.Color = 6737151
    .FormatConditions(1).StopIfTrue = False
End With

' Verde
With Workbooks("seguimiento_oc.xlsx").Sheets(hoja_filtro).Range("O2:O10000")
    .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=8"
    .FormatConditions(Range("O2:O10000").FormatConditions.Count).SetFirstPriority
    .FormatConditions(1).Interior.PatternColorIndex = xlAutomatic
    .FormatConditions(1).Interior.Color = RGB(0, 238, 108)
    .FormatConditions(1).StopIfTrue = False
End With

' Rojo
With Workbooks("seguimiento_oc.xlsx").Sheets(hoja_filtro).Range("P2:P10000")
    .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
    .FormatConditions(Range("P2:P10000").FormatConditions.Count).SetFirstPriority
    .FormatConditions(1).Interior.PatternColorIndex = xlAutomatic
    .FormatConditions(1).Interior.Color = 5263615
    .FormatConditions(1).StopIfTrue = False
End With

' Margenes
Workbooks("seguimiento_oc.xlsx").Sheets(hoja_filtro).Range("A2").CurrentRegion.Select
With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .Weight = xlThin
End With
With Selection.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .Weight = xlThin
End With
With Selection.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .Weight = xlThin
End With
    
Workbooks("seguimiento_oc.xlsx").Worksheets(hoja_filtro).Copy
Application.DisplayAlerts = False
ActiveWorkbook.SaveAs ("\\vmedsis03\Suministros\Plantillas\adjuntos\seguimiento.xlsx")
Workbooks("seguimiento.xlsx").Close
Workbooks("seguimiento_oc.xlsx").Sheets(hoja_filtro).delete
Application.DisplayAlerts = True
End Sub

