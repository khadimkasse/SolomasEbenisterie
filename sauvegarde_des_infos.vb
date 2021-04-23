Option Explicit

Sub saveInvoiceInformation()
    Application.ScreenUpdating = False
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
        'We delete the content of the line to avoid keeping information that were deleted from the invoice if any
        ClientDetailsExport.Rows(currentLine).ClearContents
    Else
        'We add one line to the export
        Set ClientDetailsExport = ClientDetailsExport.Resize(currentLine)
    End If

    '***************************************** Saving the invoice reference *****************************************
    Sheets("informations enregistrées").Range("A" & currentLine).Formula = "=$A" & currentLine - 1 & "+1"
    Sheets("informations enregistrées").Range("B" & currentLine).Value = invoiceNumber

    '***************************************** Saving client informations *****************************************
    'Looping through all the details of a client identification. To bypass the merged cells, we go from one line to another by using the Offset() function
    Dim rangeInClientDetailsKey As Range, rangeInClientDetailsValue1 As Range, rangeInClientDetails As Range
    Set rangeInClientDetailsKey = Range("ClientDetails").Columns(1)
    Set rangeInClientDetailsValue1 = Range("ClientDetails").Columns(4)
    Set rangeInClientDetails = Range(rangeInClientDetailsKey, rangeInClientDetailsValue1)
    Call exportReferencesFromRange(rangeInClientDetails, "ClientDetailsExport", currentLine)

    '***************************************** Saving the devis and DMPs **************************************
    Dim DevisExport, DMPsExport As Range
    Dim devisOrDMP As Range, devisOrDMPRow As Integer
    Dim colDevisExport As Integer, colDMPExport As Integer
    devisOrDMPRow = 1
    Set DevisExport = Sheets("informations enregistrées").Range("DevisExport")
    Set DMPsExport = Sheets("informations enregistrées").Range("DMPsExport")
    Set devisOrDMP = Range("DevisEtDMPs").Rows(devisOrDMPRow)
    colDevisExport = 1
    colDMPExport = 1
    While devisOrDMP.Row <= Range("DevisEtDMPs").Row + Range("DevisEtDMPs").Rows.Count - 1
        Dim selonDevisOrSelon As String, refDevis As String, dateDevis As Long, montantDevis As Double
        selonDevisOrSelon = devisOrDMP.Columns(1).Offset(0, -1).Value
        refDevis = devisOrDMP.Columns(1).Value
        dateDevis = devisOrDMP.Columns(1).Offset(0, 1).Value
        montantDevis = devisOrDMP.Columns(1).Offset(0, 4).Value 'Why 4 here. Shouldn't it be 3 ?
        If selonDevisOrSelon = "Selon devis" Then
            Dim lastColumDevisExport As Integer
            lastColumDevisExport = DevisExport.Column + DevisExport.Columns.Count - 1
            'If the columns corresponding to devis on the export page are not completely filled, we use the current one
            If colDevisExport <= DevisExport.Columns.Count Then
                DevisExport.Rows(currentLine).Columns(colDevisExport).Value = refDevis
                DevisExport.Rows(currentLine).Columns(colDevisExport + 1).Value = dateDevis
                DevisExport.Rows(currentLine).Columns(colDevisExport + 2).Value = montantDevis
            Else
                'We have used all the available columns designed for the Devis. Then we insert 3 new columns and affect them to the range DevisExport
                Worksheets("informations enregistrées").Columns(lastColumDevisExport + 1).Resize(, 3).Insert Shift:=xlToRight
                lastColumDevisExport = lastColumDevisExport + 3
                Set DevisExport = DevisExport.Resize(, DevisExport.Columns.Count + 3)
                'Setting the right index for the newly created columns
                Dim libelleDevis As String
                libelleDevis = "Devis " & DevisExport.Columns.Count / 3
                'Merging the 3 first lines and setting the title
                Application.Union(DevisExport.Rows(1).Columns(DevisExport.Columns.Count - 5), DevisExport.Rows(1).Columns(DevisExport.Columns.Count - 4), DevisExport.Rows(1).Columns(DevisExport.Columns.Count - 3)).Copy
                DevisExport.Rows(1).Columns(DevisExport.Columns.Count - 2).PasteSpecial _
                Paste:=xlPasteFormats
                Application.CutCopyMode = False
                DevisExport.Rows(1).Columns(DevisExport.Columns.Count - 2).Value = libelleDevis
                DevisExport.Rows(currentLine).Columns(colDevisExport).Value = refDevis
                DevisExport.Rows(currentLine).Columns(colDevisExport + 1).Value = dateDevis
                DevisExport.Rows(currentLine).Columns(colDevisExport + 2).Value = montantDevis
            End If
            colDevisExport = colDevisExport + 3
        Else
            Dim lastColumDMPExport As Integer
            lastColumDMPExport = DMPsExport.Column + DMPsExport.Columns.Count - 1
             'If the columns corresponding to DMP on the export page are not completely filled, we use the current one
            If colDMPExport <= DMPsExport.Columns.Count Then
                DMPsExport.Rows(currentLine).Columns(colDMPExport).Value = refDevis
                DMPsExport.Rows(currentLine).Columns(colDMPExport + 1).Value = dateDevis
                DMPsExport.Rows(currentLine).Columns(colDMPExport + 2).Value = montantDevis
            Else
                'We have used all the available columns designed for the DMPs. Then we insert 3 new columns and affect them to the range DMPsExport
                Worksheets("informations enregistrées").Columns(lastColumDMPExport + 1).Resize(, 3).Insert Shift:=xlToRight
                lastColumDMPExport = lastColumDMPExport + 3
                Set DMPsExport = DMPsExport.Resize(, DMPsExport.Columns.Count + 3)
                'Setting the right index for the newly created columns
                Dim libelleDMP As String
                libelleDMP = "DMP " & DMPsExport.Columns.Count / 3
                'Merging the 3 first lines and setting the title
                Application.Union(DMPsExport.Rows(1).Columns(DMPsExport.Columns.Count - 5), DMPsExport.Rows(1).Columns(DMPsExport.Columns.Count - 4), DMPsExport.Rows(1).Columns(DMPsExport.Columns.Count - 3)).Copy
                DMPsExport.Rows(1).Columns(DMPsExport.Columns.Count - 2).PasteSpecial _
                Paste:=xlPasteFormats
                Application.CutCopyMode = False
                DMPsExport.Rows(1).Columns(DMPsExport.Columns.Count - 2).Value = libelleDMP
                DMPsExport.Rows(currentLine).Columns(colDMPExport).Value = refDevis
                DMPsExport.Rows(currentLine).Columns(colDMPExport + 1).Value = dateDevis
                DMPsExport.Rows(currentLine).Columns(colDMPExport + 2).Value = montantDevis
            End If
            colDMPExport = colDMPExport + 3
        End If
        devisOrDMPRow = devisOrDMPRow + 1
        Set devisOrDMP = Range("DevisEtDMPs").Rows(devisOrDMPRow)
    Wend
    DevisExport.Name = "DevisExport"
    DMPsExport.Name = "DMPsExport"

    'Once the invoice has been saved, we cannot update the reference of it
    'Therefore, we change the value of the range invoiceNumber to keep the values as fixed and not depending on a formula
    Range("invoiceNumber").Copy
    Range("invoiceNumber").PasteSpecial _
    Paste:=xlPasteValues
    Application.CutCopyMode = False
    Application.ScreenUpdating = True

    '***************************************** Showing the display before printing (aperçu avant impression) **************************************
    'ActiveSheet.PrintPreview
End Sub

'rangeIn : {key, value1, value2, value3, ....}
Sub exportReferencesFromRange(rangeIn As Range, rangeExportStr As String, currentLine As Integer)
    Dim rangeExport As Range
    Set rangeExport = Sheets("informations enregistrées").Range(rangeExportStr)
    Dim rangeInSubRow As Range, rowRangeInSubRow As Integer, colRangeExport As Integer
    
    rowRangeInSubRow = 1
    Set rangeInSubRow = rangeIn.Rows(rowRangeInSubRow)
    colRangeExport = 1

    While rangeInSubRow.Row <= rangeIn.Row + rangeIn.Rows.Count - 1
        Dim lastColumnRangeExport As Integer
        lastColumnRangeExport = rangeExport.Column + rangeExport.Columns.Count - 1
        'If the key of the values we are about to add doesn't exist in the export,
        'we insert as much columns as needed
        If rangeInSubRow.Columns(1).Value <> rangeExport.Rows(1).Columns(colRangeExport).Value Then
            Dim matchingColInExport As Range
            Set matchingColInExport = rangeExport.Rows(1).Find( _
                            what:=rangeInSubRow.Rows(1).Columns(1).Value, searchorder:=xlByColumns)
            If Not matchingColInExport Is Nothing Then
                colRangeExport = matchingColInExport.Column
            Else
                Dim nbColumnsToAdd As Integer
                nbColumnsToAdd = rangeInSubRow.Columns.Count
                Worksheets("informations enregistrées").Columns(lastColumnRangeExport + 1).Resize(, nbColumnsToAdd).Insert Shift:=xlToRight
                lastColumnRangeExport = lastColumnRangeExport + nbColumnsToAdd
                Set rangeExport = rangeExport.Resize(, rangeExport.Columns.Count + nbColumnsToAdd)
                'Merging the columns we just added if needed
                Range(Cells(1, rangeExport.Columns.Count - 2 * nbColumnsToAdd), Cells(1000, rangeExport.Columns.Count - nbColumnsToAdd)).Copy
                Range(Cells(1, rangeExport.Columns.Count - nbColumnsToAdd + 1), Cells(1, rangeExport.Columns.Count - nbColumnsToAdd + 1)).PasteSpecial _
                Paste:=xlPasteFormats
                Application.CutCopyMode = False
            End If
        End If
        Dim thisColumn As Integer
        thisColumn = rangeInSubRow.Columns(1).Offset(, 1).Column - rangeInSubRow.Columns(1).Column + 1
        While thisColumn <= rangeInSubRow.Columns.Count
            rangeExport.Rows(currentLine).Columns(colRangeExport).Value = rangeInSubRow.Columns(thisColumn).Value
            colRangeExport = colRangeExport + 1
            thisColumn = rangeInSubRow.Rows(1).Columns(thisColumn).Offset(0, 1).Column - rangeInSubRow.Rows(1).Columns(1).Column + 1
        Wend
        rowRangeInSubRow = rangeIn.Rows(rowRangeInSubRow).Columns(1).Offset(1, 0).Row
        Set rangeInSubRow = rangeIn.Rows(rowRangeInSubRow)
    Wend

    'Impacting the changes made on input ranges
    rangeExport.Name = rangeExportStr
End Sub

