'endRow: the last row from wich we have to create new line after it
'lineInsert: the location at which we have to insert the incoming lines
'columnStartLine: the starting colum of the range we have to replicate
'columnEndLine: its ending column
'refToCopy: the relative position of the line whom format will be replicated. If not provided the previous one will be considered
'moreLinesToAdd: the number of lines above one we want to add. Naturaly, its value is set to 0 if not provided meaning one line will be added
Sub addOneLine(endRow, lineInsert, columnStartLine, columnEndLine, Optional refToCopy As Integer = 0, Optional moreLinesToAdd As Integer = 0)
    'Adding one or several lines to the document at the right place
    Range("A" & lineInsert & ":" & "S" & lineInsert).Resize(1 + moreLinesToAdd).Insert Shift:=xlDown
    Range("A" & lineInsert & ":" & "S" & lineInsert).ClearFormats
    'Copying the format of the latest row we already had
    Range(columnStartLine & (endRow - refToCopy) & ":" & columnEndLine & (endRow - refToCopy)).Copy
    Range(columnStartLine & endRow + 1).PasteSpecial _
    Paste:=xlPasteFormats
    'Emptying the clipboard
    Application.CutCopyMode = False
    'Clearing the content of the line we just copied if any
    Selection.ClearContents
End Sub

Sub replicateLineOnRecap(endRow, lineInsert, columnStartLine, columnEndLine, referenceLine, columnsToReplicate() As String, Optional refToCopy As Integer = 0, Optional moreLinesToAdd As Integer = 0)
    Call addOneLine(endRow, lineInsert, columnStartLine, columnEndLine, refToCopy, moreLinesToAdd)
    For Each thisColumn In columnToReplicate
         Range(columStartLine & (endRow + 1)).Formula = "=" & thisColumn & referenceLine
    Next thisColumn
End Sub

Sub copyFormulasFromLine(referenceLineStart, referenceLineStop, columnStart, columnStop, pasteLine)
    Range(columnStart & referenceLineStart & ":" & columnStop & referenceLineStop).Copy
    Range(columnStart & pasteLine).PasteSpecial _
    Paste:=xlPasteFormulas
    'Emptying the clipboard
    Application.CutCopyMode = False
End Sub

Sub addOneLineToClientIdentification()
    'Getting the last row of the client identification part or the invoice details part
    startRow = Range("ClientDetails").Row
    startRowInvoiceDetails = Range("InvoiceDetails").Row
    endRow = Range("ClientDetails").Row + Range("ClientDetails").Rows.Count - 1
    endRowInvoiceDetails = Range("InvoiceDetails").Row + Range("InvoiceDetails").Rows.Count - 1
    'If the invoice details part is longer than the client identification one, we don't need to add any new lines
    If endRowInvoiceDetails <= endRow Then 
        lineInsert = endRow + 1
        Call addOneLine(endRow, lineInsert, "H", "M")
    Else
        Range("H" & endRow & ":M" & endRow).Copy
        Range("H" & endRow + 1).PasteSpecial _
        Paste:=xlPasteFormats
        'Emptying the clipboard
        Application.CutCopyMode = False
    End If
    'Refreshing the named ranges ClientDetails, InvoiceDetails and impression_des_titres
    Range("'template'!$H$" & startRow & ":$M$" & endRow + 1).Name = "ClientDetails"
    Range("'template'!$" & startRow & ":$" & endRow + 1).Name = "impression_des_titres"
End Sub

Sub addOneLineToInvoiceDetails()
    'Getting the last row of the client identification part or the invoice details part
    startRow = Range("InvoiceDetails").Row
    startRowClientDetails = Range("ClientDetails").Row
    endRow = Range("InvoiceDetails").Row + Range("InvoiceDetails").Rows.Count - 1
    endRowClientDetails = Range("ClientDetails").Row + Range("ClientDetails").Rows.Count - 1
    lineInsert = endRow + 1
    'If the client identification details part is longer than the invoice one, we don't need to add any new lines
    If endRowClientDetails <= endRow Then 
        lineInsert = endRow + 1
        Call addOneLine(endRow, lineInsert, "P", "S")
    Else
        Range("P" & endRow & ":S" & endRow).Copy
        Range("P" & endRow + 1).PasteSpecial _
        Paste:=xlPasteFormats
        'Emptying the clipboard
        Application.CutCopyMode = False
    End If
    'Refreshing the named ranges ClientDetails, InvoiceDetails and impression_des_titres
    Range("'template'!$P$" & startRow & ":$S$" & endRow + 1).Name = "InvoiceDetails"
    Range("'template'!$" & startRowClientDetails & ":$" & endRow + 1).Name = "impression_des_titres"
End Sub

Sub addOneLineOfDevisOrDMP()
    'Getting the last line in which we defined a devis or DMP
    startRow = Range("DevisEtDMPs").Row
    endRow = Range("DevisEtDMPs").Row + Range("DevisEtDMPs").Rows.Count - 1
    lineInsert = endRow + 1
    Call addOneLine(endRow, lineInsert, "C", "J")
    'Getting the same content on Recapitulatif
    startRowRecap = Range("DevisEtDMPRecap").Row
    endRowRecap = Range("DevisEtDMPRecap").Row + Range("DevisEtDMPRecap").Rows.Count - 1
    lineInsertRecap = endRowRecap + 1
    Call addOneLine(endRowRecap, lineInsertRecap, "C", "J")
    'Replicate the formulas of the previous line
    Call copyFormulasFromLine(endRowRecap, endRowRecap, "C", "J", endRowRecap + 1)
    'Updating the DevisEtDMP range
    Range("'template'!$D$" & startRow & ":$D$" & endRow + 1).Name = "DevisEtDMPs"
    Range("'template'!$D$" & startRowRecap & ":$D$" & endRowRecap + 1).Name = "DevisEtDMPRecap"
    Range("C" & endLine + 1).Select
End Sub

Sub addOneArticle()
    'Getting the last line in which we have an article.
    startRow = Range("RefsDevisEtDMP").Row
    endRow = Range("RefsDevisEtDMP").Row + Range("RefsDevisEtDMP").Rows.Count - 1
    lineInsert = endRow + 1
    Call addOneLine(endRow, lineInsert, "C", "S", 1)
    Range("C" & lineInsert).Value = "Si la référence du devis ou DMP est nouvelle, cliquer sur le bouton ""Chercher de nouveaux devis"" en colonne X."
    'Saving the references of the different columns so that the formulas can adapt themselves
    Range("'template'!$C$" & startRow & ":$C$" & endRow + 1).Name = "RefsDevisEtDMP"
    Range("'template'!$R$" & startRow & ":$R$" & endRow + 1).Name = "MontantsArticlesHT"
    Range("'template'!$S$" & startRow & ":$S$" & endRow + 1).Name = "CodesTVA"
End Sub

Sub addNewHeaderToArticlesTable()
    Dim rngInput As Range
    'Getting the place in which we are going to insert the table
    Set rngInput = Application.InputBox( _
      Title:="Une référence à la ligne où insérer l'entête", _
      Prompt:="Sélectionner la ligne à laquelle insérer l'entête du tableau", _
      Type:=8)
    insertRow = rngInput.Row
    'Inserting 3 lines down the given row
    Call addOneLine(insertRow, insertRow + 1, "C", "S", 0, 2)
    'Getting the line in which we have the header of the articles table.
    headerRow = Range("ArticlesListingHeader").Row ' Application.Match("Ref devis et DMP", Range("C:C"), 0)
    Range("C" & headerRow & ":S" & headerRow + 2).Copy
    Range("C" & insertRow + 1).PasteSpecial _
    Paste:=xlPasteAll
    'Emptying the clipboard
    Application.CutCopyMode = False
End Sub

Sub addLinesToDoubleRowTables(endRow, lineInsert, columnStartLine, columnEndLine)
    'Adding two lines to the document at the right place
    Range("A" & lineInsert & ":" & "S" & lineInsert).Resize(2).Insert Shift:=xlDown
    'Copying the format of the last but one row we already had
    Range(columnStartLine & (endRow - 3) & ":" & columnEndLine & (endRow - 2)).Copy
    Range(columnStartLine & (endRow + 1)).PasteSpecial _
    Paste:=xlPasteFormats
    'Emptying the clipboard
    Application.CutCopyMode = False
End Sub

Sub addOneAppelDeFond()
    'Getting the last line in which he have an "appel de fond".
    'We take the named range MontantsTVAAppeles as it is one of the most used of that table
    startRow = Range("MontantsTVAAppeles").Row
    endRow = Range("MontantsTVAAppeles").Row + Range("MontantsTVAAppeles").Rows.Count - 1
    lineInsert = endRow + 1
    'Beware: we have two lines to add here
    Call addLinesToDoubleRowTables(endRow, lineInsert, "F", "O")
    'Copying the formulas from the last line
    Call copyFormulasFromLine(endRow - 1, endRow, "F", "O", endRow + 1)
    'Clearing some of the content of the line we just copied if any
    Range("F" & endRow + 1 & ":" & "H" & endRow + 2).Select
    Selection.ClearContents
    'Refreshing the named range of that table
    Range("'template'!$K$" & startRow & ":$K$" & endRow + 2).Name = "TauxTVAAppeles"
    Range("'template'!$L$" & startRow & ":$M$" & endRow + 2).Name = "MontantsTVAAppeles"
    Range("'template'!$N$" & startRow & ":$O$" & endRow + 2).Name = "MontantsTTCAppeles"
    'We check if the table containig the taxe d'ameublement is well calculated. If not we refresh it.
    Call refreshTaxeDAmeublementTable
    Call refreshTaxeDAmeublemnentOfCurrentInvoice()
End Sub

Sub addOneInvoiceToRecapitulatif()
    'Getting the starting and ending row of the "Factures d'acompte"
    startRow = Range("MontantsFacturesAppeles").Row
    endRow = Range("MontantsFacturesAppeles").Row + Range("MontantsFacturesAppeles").Rows.Count - 1
    lineInsert = endRow + 1
    'Beware: we have two lines to add here
    Call addLinesToDoubleRowTables(endRow, lineInsert, "C", "S")
    Range("C" & endRow + 1).Value = "Facture d'acompte :"
    Range("H" & endRow + 1).Formula = "=" & Range("TotalTTCFacture").Address ' "=$L$" & invoiceTotalTTC
    'Refreshing the named ranges based on the table Factures d'acompte
    Range("'template'!$H$" & startRow & ":$I$" & endRow + 2).Name = "MontantsFacturesAppeles"
    Range("'template'!$J$" & startRow & ":$K$" & endRow + 2).Name = "MontantsPayesSurFacturesAppeles"
    Range("'template'!$R$" & startRow & ":$S$" & endRow + 2).Name = "RestantsDusSurFacturesAppeles"
End Sub

Sub refreshTaxeDAmeublementTable()
    'The named range UniqueRefDevisEtDMPs should be the same length as the DevisEtDMPs one
    'If not, the table have to be extended
    If Range("UniqueRefDevisEtDMPs").Rows.Count <> Range("DevisEtDMPs").Count Then
        startRow = Range("UniqueRefDevisEtDMPs").Row
        endRowBeforeUpdate = Range("UniqueRefDevisEtDMPs").Row + Range("UniqueRefDevisEtDMPs").Rows.Count - 1
        endRow = Range("UniqueRefDevisEtDMPs").Row + Range("DevisEtDMPs").Rows.Count - 1
        Range("'template'!$Z$" & startRow & ":$Z$" & endRow).Name = "UniqueRefDevisEtDMPs"
        Range("'template'!$AA$" & startRow & ":$AA$" & endRow).Name = "TaxeAmeublementN"
        Range("'template'!$AB$" & startRow & ":$AB$" & endRow).Name = "TaxeAmeublementR"
        'Copying the formulas to the lines we have to add
        thisRow = endRowBeforeUpdate
        While thisRow < endRow
            Call copyFormulasFromLine(startRow, startRow, "AA", "AB", thisRow + 1)
            thisRow = thisRow + 1
        Wend
    End If
End Sub

Sub refreshTaxeDAmeublemnentOfCurrentInvoice()
    startRow = Range("MontantsTVAAppeles").Row
    endRow = Range("MontantsTVAAppeles").Row + Range("MontantsTVAAppeles").Rows.Count - 1
    i = startRow
    formulaTaxeDAmeublement = "=IF(TaxeAmeublementExiste=""Y"","
    formulaTaxeDAmeublement = formulaTaxeDAmeublement & _
        "$F$" & i & "*(SUMIF(UniqueRefDevisEtDMPs," & "$G$" & i & ",TaxeAmeublementN)+" _
        & "SUMIF(UniqueRefDevisEtDMPs," & "$G$" & i & ",TaxeAmeublementR))"
    i = i + 2
    Do While i <= endRow
        formulaTaxeDAmeublement = formulaTaxeDAmeublement & _
        "+$F$" & i & "*(SUMIF(UniqueRefDevisEtDMPs," & "$G$" & i & ",TaxeAmeublementN)+" _
        & "SUMIF(UniqueRefDevisEtDMPs," & "$G$" & i & ",TaxeAmeublementR))"
        i = i + 2
    Loop
    formulaTaxeDAmeublement = formulaTaxeDAmeublement & ","""")"
    'Puting the formula at the right place
    Range("TotalTaxeDAmeublementFacture").Formula = formulaTaxeDAmeublement
End Sub

Sub updateReferenceOnUserCommand(rangeName As String, rngUpdated As Range, referenceName As String)
     If Not (Range(rangeName).Address = rngUpdated.Address) Then
        If MsgBox("La référence de " & referenceName & " semble éronnée. La bonne référence semble être " & _
        rngUpdated.Address & " au lieu de l'actuel : " & Range(rangeName).Address & ". Voulez-vous le mettre à jour maintenant ?", vbYesNo, "Demande de confirmation") = vbYes Then
            rngUpdated.Name = rangeName
            MsgBox "La référence de " & referenceName & " a été mise à jour !"
        End If
    End If
End Sub
Sub refreshNamedRanges()
    'Devis et DMP
    startRow = Application.Match("Selon devis", Range("C:C"), 0)
    endRow = Range("C" & startRow).End(xlDown).Row
    Dim rngUpdated As Range
    Set rngUpdated = Range("'template'!$D$" & startRow & ":$D$" & endRow)
    Call updateReferenceOnUserCommand("DevisEtDMPs", rngUpdated, "Devis et DMP")
   
    'Devis et DMP on RECAPITULATIF
    recapPosition = Range("recapPosition").Row
    startRowRecap = recapPosition + Application.Match("Selon devis", Range("C" & recapPosition & ":C" & recapPosition + 100), 0) - 1 ' Why do we need here - 1 ?
    endRowRecap = Range("C" & startRowRecap).End(xlDown).Row
    Set rngUpdated = Range("'template'!$D$" & startRowRecap & ":$D$" & endRowRecap)
    Call updateReferenceOnUserCommand("DevisEtDMPRecap", rngUpdated, "Devis et DMP du Recapitulatif")
    'We check if the number of devis and DMP is the same compared to the ones on first page. If not we raise an alert and ask to delete the extra lines
    nbLinesDevisAndDMPs = endRow - startRow + 1
    nbLinesDevisAndDMPsRecap = endRowRecap - startRowRecap + 1
    If nbLinesDevisAndDMPsRecap > nbLinesDevisAndDMPs Then
        If MsgBox("Le nombre de devis renseigné dans le récapitulatif semble éroné. Voulez-vous supprimer les lignes en trop ?", vbYesNo, "Demande de confirmation") = vbYes Then
            Range("A" & startRowRecap + nbLinesDevisAndDMPsRecap - 1 & ":S" & endRowRecap).Delete Shift:=xlUp
            MsgBox "Les lignes en trop dans le récapitulatif ont bien été supprimées !"
        End If
    End If

    
    'Articles listing
    'The " + 3" is to avoid the problems brought by the merged cells
    startRow = Application.Match("Ref devis et DMP", Range("C:C"), 0) + 3
    'To find the last line of the table, we search from the bottom the latest header of the
    'table ( + 3 ) and then jump to the end
    endRow = Range("C1:C100").Find(what:="Ref devis et DMP", searchorder:=xlByRows, searchdirection:=xlPrevious).Row + 3
    endRow = Range("C" & endRow).End(xlDown).Row
    Set rngUpdated = Range("'template'!$C$" & startRow & ":$C$" & endRow)
    Call updateReferenceOnUserCommand("RefsDevisEtDMP", rngUpdated, "Ref devis et DMP")
    Set rngUpdated = Range("'template'!$R$" & startRow & ":$R$" & endRow)
    Call updateReferenceOnUserCommand("MontantsArticlesHT", rngUpdated, "Montant HT des différents articles")


    'Appel de fond
    'Getting the last line in which he have an "appel de fond". We add 2 to bypass the problems caused by the merged cells
    startRow = Application.Match("Appel de fond", Range("F:F"), 0) + 2
    'Let us move until the column "Base" before looking at the end. Indeed, the Range.End excel
    'function is mislead by the merged cell. Therefore we pick the first "safe" place in which we can run it
    endRow = Range("I" & startRow).End(xlDown).Row
    Set rngUpdated = Range("'template'!$K$" & startRow & ":$K$" & endRow)
    Call updateReferenceOnUserCommand("TauxTVAAppeles", rngUpdated, "Taux TVA de l'appel de fond")
    endRow = Range("I" & startRow).End(xlDown).Row
    Set rngUpdated = Range("'template'!$L$" & startRow & ":$M$" & endRow)
    Call updateReferenceOnUserCommand("MontantsTVAAppeles", rngUpdated, "Montants TVA de l'appel de fond")
    endRow = Range("I" & startRow).End(xlDown).Row
    Set rngUpdated = Range("'template'!$N$" & startRow & ":$O$" & endRow)
    Call updateReferenceOnUserCommand("MontantsTTCAppeles", rngUpdated, "Montants TTC de l'appel de fond")

    'Factures d'acompte
    startRow = Range("C1:D1000").Find(what:="Facture d'acompte :", searchorder:=xlByRows).Row
    endRow = Range("C1:D1000").Find(what:="Facture d'acompte :", searchorder:=xlByRows, searchdirection:=xlPrevious).Row + 1
    Set rngUpdated = Range("'template'!$H$" & startRow & ":$I$" & endRow)
    Call updateReferenceOnUserCommand("MontantsFacturesAppeles", rngUpdated, "Montants Factures appelees")
    Set rngUpdated = Range("'template'!$J$" & startRow & ":$K$" & endRow)
    Call updateReferenceOnUserCommand("MontantsPayesSurFacturesAppeles", rngUpdated, "Montants payes sur Factures appelees")
    Set rngUpdated = Range("'template'!$R$" & startRow & ":$S$" & endRow)
    Call updateReferenceOnUserCommand("RestantsDusSurFacturesAppeles", rngUpdated, "Montants restants sur Factures appelees")

    'Table d'aggregation de la taxe d'ameublement
    'Getting the position of the table "Aggrégation de la taxe d'ameublement"
    startRow = Application.Match("Aggrégation de la taxe d'ameublement", Range("Z:Z"), 0) + 3
    endRow = Range("Z" & startRow).End(xlDown).Row
    Set rngUpdated = Range("'template'!$Z$" & startRow & ":$Z$" & endRow)
    Call updateReferenceOnUserCommand("UniqueRefDevisEtDMPs", rngUpdated, "Références des devis et DMP dans la table d'aggrégation de la taxe d'ameublement")
    Set rngUpdated = Range("'template'!$AA$" & startRow & ":$AA$" & endRow)
    Call updateReferenceOnUserCommand("TaxeAmeublementN", rngUpdated, "Taxe d'ameublement N dans la table d'aggrégation de la taxe d'ameublement")
    Set rngUpdated = Range("'template'!$AB$" & startRow & ":$AB$" & endRow)
    Call updateReferenceOnUserCommand("TaxeAmeublementR", rngUpdated, "Taxe d'ameublement R dans la table d'aggrégation de la taxe d'ameublement")

    'Impression des titres
    startRow = Range("ClientDetails").Row
    endRow = Range("ClientDetails").Row + Range("ClientDetails").Rows.Count - 1
    endRowInvoiceDetails = Range("InvoiceDetails").Row + Range("InvoiceDetails").Rows.Count - 1
    endRow = WorksheetFunction.Max(endRow, endRowInvoiceDetails)
    Set rngUpdated = Range("'template'!$" & startRow & ":$" & endRow)
    Call updateReferenceOnUserCommand("impression_des_titres", rngUpdated, "Impression des titres")
End Sub

