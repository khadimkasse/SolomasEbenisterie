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

    '***************************************** Saving the invoice reference ***************************************** 
    Sheets("informations enregistrées").Range("A" & currentLine).Formula = "=$A" & currentLine - 1 & "+1"
    Sheets("informations enregistrées").Range("B" & currentLine).Value = invoiceNumber

    '***************************************** Saving client informations ***************************************** 
    'Looping through all the details of a client identification. To bypass the merged cells, we go from one line to another by using the Offset() function
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

    '***************************************** Saving the devis and DMPs **************************************
    Dim DevisExport, DMPExport As Range 
    Dim devisOrDMP As Range, devisOrDMPRow As Integer
    Dim colDevisOrDMPExport As Integer
    devisOrDMPRow = 1
    Set DevisExport = Sheets("informations enregistrées").Range("DevisExport")
    Set DMPExport = Sheets("informations enregistrées").Range("DMPExport")
    Set devisOrDMP = Range("DevisEtDMPs").Rows(devisOrDMPRow)
    colDevisExport = 1
    colDMPExport = 1
    While devisOrDMP.Row <= Range("DevisEtDMPs").Rows.Count
        Dim selonDevisOrSelon As String, refDevis As String, dateDevis As Long, montantDevis As Double
        selonDevisOrSelon = devisOrDMP.Columns(1).Value
        refDevis = devisOrDMP.Columns(1).Offset(0, 1).Value
        dateDevis = devisOrDMP.Columns(1).Offset(0, 2).Value
        montantDevis = devisOrDMP.Columns(1).Offset(0, 4).Value
        If selonDevisOrSelon = "Selon devis" Then
            Dim lastColumDevisExport As Integer
            lastColumDevisExport = DevisExport.Column + DevisExport.Columns.Count - 1
            'If the columns corresponding to devis on the export page are not completely filled, we use the current one
            If colDevisExport <= lastColumDevisExport
                DevisExport.Rows(currentLine).Column(colDevisExport).Value = refDevis
                DevisExport.Rows(currentLine).Column(colDevisExport + 1).Value = dateDevis
                DevisExport.Rows(currentLine).Column(colDevisExport + 2).Value = montantDevis
            Else
                'We have used all the available columns designed for the Devis. Then we insert 3 new columns and affect them to the range DevisExport
                Worksheets("informations enregistrées").Columns(lastColumDevisExport + 3 & ":" & lastColumDevisExport + 5).Insert Shift:=xlToRight
                lastColumDevisExport = lastColumDevisExport + 3
                Set DevisExport = DevisExport.Resize(, lastColumDevisExport)
                'Setting the right index for the newly created columns
                Dim libelleDevis As Integer
                libelleDevis = "Devis " & lastColumDevisExport / 3 + 1
                'Merging the 3 first lines and setting the title 
                Sheets("informations enregistrées").Range(Cells(1, lastColumDevisExport), Cells(1, lastColumDevisExport + 2)).Copy
                DevisExport.Rows(lastColumDevisExport - 2).PasteSpecial _
                Paste:=xlPasteFormats
                Application.CutCopyMode = False
                DevisExport.Rows(currentLine)Columns(lastColumDevisExport - 2).Value = libelleDevis
            End If
        Else 
            Dim lastColumDMPExport As Integer
            lastColumDMPExport = DMPExport.Column + DMPExport.Columns.Count - 1
             'If the columns corresponding to DMP on the export page are not completely filled, we use the current one
            If colDMPExport <= lastColumDMPExport
                DMPExport.Rows(currentLine).Column(colDMPExport).Value = refDevis
                DMPExport.Rows(currentLine).Column(colDMPExport + 1).Value = dateDevis
                DMPExport.Rows(currentLine).Column(colDMPExport + 2).Value = montantDevis
            Else
                'We have used all the available columns designed for the Devis. Then we insert 3 new columns and affect them to the range DMPExport
                Worksheets("informations enregistrées").Columns(lastColumDMPExport + 3 & ":" & lastColumDMPExport + 5).Insert Shift:=xlToRight
                lastColumDMPExport = lastColumDMPExport + 3
                Set DMPExport = DMPExport.Resize(, lastColumDMPExport)
                'Setting the right index for the newly created columns
                Dim libelleDMP As Integer
                libelleDMP = "DMP " & lastColumDMPExport / 3 + 1
                'Merging the 3 first lines and setting the title 
                Sheets("informations enregistrées").Range(Cells(1, lastColumDMPExport), Cells(1, lastColumDMPExport + 2)).Copy
                DMPExport.Rows(lastColumDMPExport - 2).PasteSpecial _
                Paste:=xlPasteFormats
                Application.CutCopyMode = False
                DMPExport.Rows(currentLine)Columns(lastColumDMPExport - 2).Value = libelleDMP
            End If
            colDMPExport = colDMPExport + 3
        End If
        devisOrDMP = Range("DevisEtDMPs").Rows(devisOrDMPRow).Columns(1).Offset(1, 0).Row
        Set devisOrDMP = Range("DevisEtDMPs").Rows(clientInfoRow)
    Wend
    DevisExport.Name = "DevisExport"
    DMPExport.Name = "DMPExport"

    'Once the invoice has been saved, we cannot update the reference of it 
    'Therefore, we change the value of the range invoiceNumber to keep the values as fixed and not depending on a formula
    Range("invoiceNumber").Copy
    Range("invoiceNumber").PasteSpecial _
    Paste:=xlPasteValues
    Application.CutCopyMode = False
End Sub

