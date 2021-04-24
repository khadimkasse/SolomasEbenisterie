Option Explicit

Sub createNewInvoice()
    'We save the number of the invoice if it was not yet donein case the active sheet isn't the template
    If ActiveSheet.Name <> "template" Then
        Range("invoiceNumber").Copy
        Range("invoiceNumber").PasteSpecial _
        Paste:=xlPasteValues
        Application.CutCopyMode = False
    End If

    'We copy the sheet template to a new location with a named input by the user
    Dim sheetName As String
    sheetName = Application.InputBox("Saisissez le nom de la facture", Type:=2)
    Sheets("template").Copy before:=Sheets("informations enregistrées")
    ActiveSheet.Name = sheetName
End Sub

Sub createNewFacturier()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    'We save the current facturier
    ActiveWorkbook.Save
    'We save it again under a name input by the user
    Dim wbNameProposed As String, wbName As String
    wbNameProposed = Year(Date) & " " & Format(Month(Date), "00") & " Facturier mensuel.xlsm"
    wbName = Application.InputBox("Nom du nouveau facturier : ", Type:=2, Default:=wbNameProposed)
    ActiveWorkbook.SaveAs wbName
    'Deleting the sheets except template and informations enregistrées
    Dim xWs As Worksheet
    For Each xWs In Application.ActiveWorkbook.Worksheets
        If xWs.Name <> "template" And xWs.Name <> "informations enregistrées" Then
            xWs.Delete
        End If
    Next
    Sheets("template").Select
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ActiveWorkbook.Save
End Sub



