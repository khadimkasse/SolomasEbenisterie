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
    Call exportReferencesFromRange(rangeInClientDetails, "ClientDetailsExport", currentLine, 1)

    '***************************************** Saving the devis and DMPs **************************************
    Dim rangeInDevisKey As Range, rangeInDevisValue1 As Range, rangeInDevisValue2 As Range, rangeInDevisValue3 As Range, rangeInDevis As Range
    Dim rangeInDMPsKey As Range, rangeInDMPsValue1 As Range, rangeInDMPsValue2 As Range, rangeInDMPsValue3 As Range, rangeInDMPs As Range
    Dim devisOrDMP As Range, devisOrDMPRow As Integer
    devisOrDMPRow = 1
    Set devisOrDMP = Range("DevisEtDMPs").Rows(devisOrDMPRow)
    Dim temp As Range
    While devisOrDMP.Row <= Range("DevisEtDMPs").Row + Range("DevisEtDMPs").Rows.Count - 1
        Dim selonDevisOrSelon As String, refDevis As String, dateDevis As Long, montantDevis As Double
        selonDevisOrSelon = devisOrDMP.Columns(1).Offset(0, -1).Value
        If selonDevisOrSelon = "Selon devis" Then
            'Setting the right index for the newly created columns
            Set rangeInDevisKey = CustomUnion(rangeInDevisKey, Range("AC" & 27 + devisOrDMPRow))
            'If rangeInDevisKey Is Nothing Then
                'Set rangeInDevisKey = Range("AC" & 27 + devisOrDMPRow)
            'Else
                'Set temp = rangeInDevisKey
                'Set rangeInDevisKey = Application.Union(temp, Range("AC" & 27 + devisOrDMPRow))
            'End If
            Set rangeInDevisValue1 = CustomUnion(rangeInDevisValue1, devisOrDMP)
            Set rangeInDevisValue2 = CustomUnion(rangeInDevisValue2, devisOrDMP.Offset(, 1))
            Set rangeInDevisValue3 = CustomUnion(rangeInDevisValue3, devisOrDMP.Offset(, 4))
        ElseIf selonDevisOrSelon = "Selon" Then
            'Setting the right index for the newly created columns
            Set rangeInDMPsKey = CustomUnion(rangeInDMPsKey, Range("AC" & 27 + devisOrDMPRow))
            Set rangeInDMPsValue1 = CustomUnion(rangeInDMPsValue1, devisOrDMP)
            Set rangeInDMPsValue2 = CustomUnion(rangeInDMPsValue2, devisOrDMP.Offset(, 1))
            Set rangeInDMPsValue3 = CustomUnion(rangeInDMPsValue3, devisOrDMP.Offset(, 4))
        End If
        devisOrDMPRow = devisOrDMPRow + 1
        Set devisOrDMP = Range("DevisEtDMPs").Rows(devisOrDMPRow)
    Wend
    Set rangeInDevis = Application.Union(rangeInDevisValue1, rangeInDevisValue2, rangeInDevisValue3)
    Set rangeInDMPs = Application.Union(rangeInDMPsValue1, rangeInDMPsValue2, rangeInDMPsValue3)
    Call exportReferencesFromRange(rangeInDevis, "DevisExport", currentLine, 3, "Devis")
    Call exportReferencesFromRange(rangeInDMPs, "DMPsExport", currentLine, 3, "DMP")

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

Function CustomUnion(rangeA As Range, rangeB As Range) As Range
    If rangeA Is Nothing Then
        Set CustomUnion = rangeB
    Else
        Set CustomUnion = Application.Union(rangeA, rangeB)
    End If
End Function

'rangeIn : {key, value1, value2, value3, ....}
Sub exportReferencesFromRange(rangeIn As Range, rangeExportStr As String, currentLine As Integer, nbColumnsOfValues As Integer, Optional libelle As String = "")
    Dim rangeExport As Range
    Set rangeExport = Sheets("informations enregistrées").Range(rangeExportStr)
    Dim rangeInSubRow As Range, rowRangeInSubRow As Integer, colRangeExport As Integer
    Dim libelleColumnImport As String
    
    rowRangeInSubRow = 1
    Set rangeInSubRow = rangeIn.Rows(rowRangeInSubRow)
    colRangeExport = 1

    While rangeInSubRow.Row <= rangeIn.Row + rangeIn.Rows.Count - 1
        Dim lastColumnRangeExport As Integer
        lastColumnRangeExport = rangeExport.Column + rangeExport.Columns.Count - 1
        'If the key of the values we are about to add doesn't exist in the export,
        'we insert as much columns as needed
        If libelle <> "" Then
            libelleColumnImport = libelle & " " & rowRangeInSubRow
        Else
            libelleColumnImport = rangeInSubRow.Rows(1).Columns(1).Value
        End If
        If libelleColumnImport <> rangeExport.Rows(1).Columns(colRangeExport).Value Then
            Dim matchingColInExport As Range
            Set matchingColInExport = rangeExport.Rows(1).Find( _
                            what:=rangeInSubRow.Rows(1).Columns(1).Value, searchorder:=xlByColumns)
            If Not matchingColInExport Is Nothing Then
                colRangeExport = matchingColInExport.Column
            Else
                'Dim nbColumnsToAdd As Integer, columnRangeInSubRow As Integer
                'Getting the number of columns isn't straight forward because of the merged cells
                'columnRangeInSubRow = 1
                'nbColumnsToAdd = 0
                'While columnRangeInSubRow < rangeInSubRow.Columns.Count
                    'nbColumnsToAdd = nbColumnsToAdd + 1
                    'columnRangeInSubRow = rangeInSubRow.Rows(1).Columns(columnRangeInSubRow).Offset(, 1).Column - rangeInSubRow.Column + 1
                'Wend
                Worksheets("informations enregistrées").Columns(lastColumnRangeExport + 1).Resize(, nbColumnsOfValues).Insert Shift:=xlToRight
                lastColumnRangeExport = lastColumnRangeExport + nbColumnsOfValues
                'Merging the columns we just added if needed
                'Sheets("informations enregistrées").Range(Cells(1, rangeExport.Columns.Count - 2 * nbColumnsToAdd), Cells(1000, rangeExport.Columns.Count - nbColumnsToAdd)).Copy
                
                Range(rangeExport.Columns(rangeExport.Columns.Count - nbColumnsOfValues + 1), rangeExport.Columns(rangeExport.Columns.Count)).Copy
                rangeExport.Columns(rangeExport.Column + rangeExport.Columns.Count).PasteSpecial _
                Paste:=xlPasteFormats
                Application.CutCopyMode = False
                Set rangeExport = rangeExport.Resize(, rangeExport.Columns.Count + nbColumnsOfValues)
                'Setting the title of the columns just added
                rangeExport.Rows(1).Columns(colRangeExport).Value = libelleColumnImport
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

