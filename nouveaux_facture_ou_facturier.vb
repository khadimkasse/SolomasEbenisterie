Option Explicit

Sub createNewInvoice()
    Dim sheetName As String
    sheetName = Application.InputBox("Saisissez le nom de la facture", Type:=2)
    'The selected sheet is the base for the new invoice we create
    'If the selected sheet was template, we just copy it as it is ready. Else, we need to modify it a little bit
    If ActiveSheet.Name = "template" Then
        'We copy the sheet template to a new location with the name input by the user
        ActiveSheet.Copy after:=Sheets("informations enregistrées")
        ActiveSheet.Name = sheetName
    Else
        'We save the number of the invoice if it was not yet done in case the active sheet isn't the template
        Range("invoiceNumber").Copy
        Range("invoiceNumber").PasteSpecial _
        Paste:=xlPasteValues
        Application.CutCopyMode = False
        'We copy the current sheet to a new location with the name input by the user
        ActiveSheet.Copy after:=ActiveSheet
        ActiveSheet.Name = sheetName
        'We empty the content of the table Appel de fond
        'First we delete the lines above the 2nd one
        Dim startRange As Integer : startRange = Range("MontantsTVAAppeles").Row  
        Dim endRange As Integer : endRange = Range("MontantsTVAAppeles").Row + Range("MontantsTVAAppeles").Rows.Count - 1
        Range("A" & startRange + 4 & _ 
        ":S" & endRange).Delete Shift:=xlUp
        'Then we delete the content of the percentage and replace the ref. devis by XXX
        Range("F" & startRange).Value = 0
        Range("G" & startRange).Value = "XXX"
        Range("F" & startRange + 2).Value = 0
        Range("G" & startRange + 2).Value = "XXX"
    End If
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



