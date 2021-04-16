Option Explicit

Sub saveInvoiceInformation()
    Dim currentLine As Integer, invoiceNumber As String
    Dim ClientDetailsExport As Range
    Set ClientDetailsExport = Sheets("informations enregistrées").Range("ClientDetailsExport")
    currentLine = Sheets("informations enregistrées").Range("ClientDetailsExport").Rows.Count + 1
    invoiceNumber = Range("invoiceNumber").Value
    'If the current invoice has already been saved, currentLine has to refer to it
    Dim savedInvoice As Range
    Set savedInvoice = Sheets("informations enregistrées").Range("B1:B10000").Find(what:=invoiceNumber, searchorder:=xlByRows)
    If Not savedInvoice Is Nothing Then
        currentLine = savedInvoice.Row
    Else
        'We add one line to the export
        Set ClientDetailsExport = ClientDetailsExport.Resize(currentLine)
    End If

    'Saving the invoice reference
    Sheets("informations enregistrées").Range("A" & currentLine).Formula = "=$A" & currentLine - 1 & "+1"
    Sheets("informations enregistrées").Range("B" & currentLine).Value = invoiceNumber

    'Saving client informations
    'Looping through all the details of a clinet identification. To bypass the merged cells, we go from one line to another by using the Offset() function
    Dim clientInfo As Range, clientInfoRow As Integer
    Dim colClientInfosExport As Integer
    clientInfoRow = 1
    Set clientInfo = Range("ClientDetails").Rows(clientInfoRow)
    colClientInfosExport = 1
    While clientInfo.Row <= Range("ClientDetails").Rows.Count
        Dim clientInfoLibelle As String, clientInfoValue As String
        clientInfoLibelle = clientInfo.Columns(1).Value
        clientInfoValue = clientInfo.Columns(1).Offset(0, 1).Value
        'If the clientInfoLibelle matches with the header of the current column on the sheet informations enregistrees, we export the value on the table
        If clientInfoLibelle = ClientDetailsExport.Rows(1).Columns(colClientInfosExport).Value Then
            ClientDetailsExport.Rows(currentLine).Columns(colClientInfosExport).Value = clientInfoValue
        Else
            'We search the clientInfoLibelle among the headers of the ClientDetailsExport named range. If it is present, we set colClientInfosExport to the corresponding column
            'and then we export the clientInfoValue there
            Dim colClientInfosColumn As Range
            Set colClientInfosColumn = ClientDetailsExport.Rows(1).Find( _
                what:=clientInfoLibelle, searchorder:=xlByColumns, searchdirection:=xlPrevious)
            If Not colClientInfosColumn Is Nothing Then
            'Beware the colClientInfosExport is relative to the ClientDetailsExport range
                colClientInfosExport = colClientInfosColumn.Column - ClientDetailsExport.Column + 1
                ClientDetailsExport.Rows(currentLine).Columns(colClientInfosExport).Value = clientInfoValue
            Else
                'When the libelle is not present in the colClientInfosExport range header, we add a column at the end of the range with the given libelle
                Dim lastColumnClientInfoExport As Integer
                lastColumnClientInfoExport = ClientDetailsExport.Column + ClientDetailsExport.Columns.Count - 1
                Worksheets("informations enregistrées").Columns(lastColumnClientInfoExport + 1).Insert Shift:=xlToRight
                colClientInfosExport = ClientDetailsExport.Columns.Count + 1
                Set ClientDetailsExport = ClientDetailsExport.Resize(, colClientInfosExport)
                ClientDetailsExport.Rows(1).Columns(colClientInfosExport).Value = clientInfoLibelle
                ClientDetailsExport.Rows(currentLine).Columns(colClientInfosExport).Value = clientInfoValue
            End If
        End If
        clientInfoRow = Range("ClientDetails").Rows(clientInfoRow).Columns(1).Offset(1, 0).Row
        Set clientInfo = Range("ClientDetails").Rows(clientInfoRow)
        colClientInfosExport = colClientInfosExport + 1
    Wend
    ClientDetailsExport.Name = "ClientDetailsExport"

    'Once the invoice has been saved, we cannot update the reference of it 
    'Therefore, we change the value of the range invoiceNumber to keep the values as fixed and not depending on a formula
    Range("invoiceNumber").Copy
    Range("invoiceNumber").PasteSpecial _
    Paste:=xlPasteValues
End Sub
