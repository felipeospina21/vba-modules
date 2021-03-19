Attribute VB_Name = "Auditoria"
Dim gc As Integer

Sub consolidar_acuerdos()
' Consolida bases acuerdos en un sólo excel, y a su vez los codigos en un txt.

' Consolida bases acuerdos
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Workbooks.OpenXML ("\\Nasestrella\oc\Automatizacion\COMPRAS\ACUERDOS\consolidado.xlsx")
Workbooks("consolidado.xlsx").Sheets(1).Activate
Workbooks("consolidado.xlsx").Sheets(1).Range("A2:U2", Range("A2:U2").End(xlDown)).ClearContents
gc = 0
Select Case gc
    Case 0: gc = "101"
    Case 1: gc = "102"
    Case 2: gc = "103"
    Case 3: gc = "104"
    Case 4: gc = "105"
    Case 5: gc = "106"
    Case 6: gc = "107"
End Select

For i = 0 To 6
    Workbooks.OpenXML ("\\Nasestrella\oc\Automatizacion\COMPRAS\ACUERDOS\ACUERDOS_" & gc & ".xlsb")
    Windows("ACUERDOS_" & gc & ".xlsb").Visible = True
    Workbooks("ACUERDOS_" & gc & ".xlsb").Sheets("Precios").Range("A2:U2", Range("A2:U2").End(xlDown)).Copy
    Workbooks("consolidado.xlsx").Sheets(1).Activate
    If Range("A2").value = "" Then
        Workbooks("consolidado.xlsx").Sheets(1).Range("A2").PasteSpecial xlAll
    Else:
        Workbooks("consolidado.xlsx").Sheets(1).Range("A1").End(xlDown).Offset(1, 0).PasteSpecial xlAll
    End If
    Workbooks("ACUERDOS_" & gc & ".xlsb").Close
    gc = gc + 1
Next

' Quita acuerdos duplicados y agrega listado a archivo txt
Workbooks("consolidado.xlsx").Sheets("Hoja3").Activate
Workbooks("consolidado.xlsx").Sheets("Hoja3").Columns("A").delete
Workbooks("consolidado.xlsx").Sheets("Hoja1").Range("B1:B999999").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("A1"), Unique:=True
Columns("A:A").EntireColumn.AutoFit

ruta = "C:\Documentos Empresa\OneDrive - MINEROS\Desktop\Automatizaciones\auditoria\acuerdos\input_file.txt"
LastRow = Workbooks("consolidado.xlsx").Sheets("Hoja3").Cells(Rows.Count, 1).End(xlUp).Row
Open ruta For Output As #1
Print #1, LastRow - 1
For i = 2 To LastRow
    Print #1, Range("A" & i).value
Next i
Close #1

Workbooks("consolidado.xlsx").Save
Workbooks("consolidado.xlsx").Close
Application.DisplayAlerts = True
End Sub


Sub auditoria()
' Ejecuta script acuerdos.py a traves de archiv bat

ChDir "C:\Documentos Empresa\OneDrive - MINEROS\Desktop\Automatizaciones\auditoria\acuerdos"
Shell "acuerdo.bat", vbNormalFocus
End Sub
