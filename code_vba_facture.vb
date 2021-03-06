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
        Call addOneLine(endRow, lineInsert, "H", "M", 1)
    Else
        Range("H" & endRow & ":M" & endRow).Copy
        Range("H" & endRow + 1).PasteSpecial _
        Paste:=xlPasteFormats
        'Emptying the clipboard
        Application.CutCopyMode = False
    End If
    'Refreshing the named ranges ClientDetails, InvoiceDetails and impression_des_titres
    Range("'" & ActiveSheet.Name & "'" & "!$H$" & startRow & ":$M$" & endRow + 1).Name = "ClientDetails"
    Range("'" & ActiveSheet.Name & "'" & "!$" & startRow & ":$" & endRow + 1).Name = "impression_des_titres"
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
        Call addOneLine(endRow, lineInsert, "P", "S", 1)
    Else
        Range("P" & endRow - 1 & ":S" & endRow - 1).Copy
        Range("P" & endRow + 1).PasteSpecial _
        Paste:=xlPasteFormats
        'Emptying the clipboard
        Application.CutCopyMode = False
    End If
    'Refreshing the named ranges ClientDetails, InvoiceDetails and impression_des_titres
    Range("'" & ActiveSheet.Name & "'" & "!$P$" & startRow & ":$S$" & endRow + 1).Name = "InvoiceDetails"
    Range("'" & ActiveSheet.Name & "'" & "!$" & startRowClientDetails & ":$" & endRow + 1).Name = "impression_des_titres"
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
    Range("'" & ActiveSheet.Name & "'" & "!$D$" & startRow & ":$D$" & endRow + 1).Name = "DevisEtDMPs"
    Range("'" & ActiveSheet.Name & "'" & "!$D$" & startRowRecap & ":$D$" & endRowRecap + 1).Name = "DevisEtDMPRecap"
    Range("C" & endRow + 1).Value = Range("C" & endRow).Value
    Range("C" & endRow + 1).Select
End Sub

Sub addOneArticle()
    'Getting the last line in which we have an article.
    startRow = Range("RefsDevisEtDMP").Row
    endRow = Range("RefsDevisEtDMP").Row + Range("RefsDevisEtDMP").Rows.Count - 1
    lineInsert = endRow + 1
    Call addOneLine(endRow, lineInsert, "C", "S", 1)
    Range("C" & lineInsert).Value = "Si la r??f??rence du devis ou DMP est nouvelle, cliquer sur le bouton ""Chercher de nouveaux devis"" en colonne X."
    'Saving the references of the different columns so that the formulas can adapt themselves
    Range("'" & ActiveSheet.Name & "'" & "!$C$" & startRow & ":$C$" & endRow + 1).Name = "RefsDevisEtDMP"
    Range("'" & ActiveSheet.Name & "'" & "!$R$" & startRow & ":$R$" & endRow + 1).Name = "MontantsArticlesHT"
    Range("'" & ActiveSheet.Name & "'" & "!$S$" & startRow & ":$S$" & endRow + 1).Name = "CodesTVA"
End Sub

Sub addNewHeaderToArticlesTable()
    Dim rngInput As Range
    'Getting the place in which we are going to insert the table
    Set rngInput = Application.InputBox( _
      Title:="Une r??f??rence ?? la ligne o?? ins??rer l'ent??te", _
      Prompt:="S??lectionner la ligne ?? laquelle ins??rer l'ent??te du tableau", _
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
    Range("'" & ActiveSheet.Name & "'" & "!$I$" & startRow & ":$J$" & endRow + 2).Name = "BasesAppel??es"
    Range("'" & ActiveSheet.Name & "'" & "!$K$" & startRow & ":$K$" & endRow + 2).Name = "TauxTVAAppeles"
    Range("'" & ActiveSheet.Name & "'" & "!$L$" & startRow & ":$M$" & endRow + 2).Name = "MontantsTVAAppeles"
    Range("'" & ActiveSheet.Name & "'" & "!$N$" & startRow & ":$O$" & endRow + 2).Name = "MontantsTTCAppeles"
    'We check if the table containig the taxe d'ameublement is well calculated. If not we refresh it.
    Call refreshTaxeDAmeublementTable
    Call refreshTaxeDAmeublementOfCurrentInvoice
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
    Range("'" & ActiveSheet.Name & "'" & "!$H$" & startRow & ":$I$" & endRow + 2).Name = "MontantsFacturesAppeles"
    Range("'" & ActiveSheet.Name & "'" & "!$J$" & startRow & ":$K$" & endRow + 2).Name = "MontantsPayesSurFacturesAppeles"
    Range("'" & ActiveSheet.Name & "'" & "!$R$" & startRow & ":$S$" & endRow + 2).Name = "RestantsDusSurFacturesAppeles"
End Sub

Sub refreshTaxeDAmeublementTable()
    'The named range UniqueRefDevisEtDMPs should be the same length as the DevisEtDMPs one
    'If not, the table have to be extended
    If Range("UniqueRefDevisEtDMPs").Rows.Count <> Range("DevisEtDMPs").Count Then
        startRow = Range("UniqueRefDevisEtDMPs").Row
        endRowBeforeUpdate = Range("UniqueRefDevisEtDMPs").Row + Range("UniqueRefDevisEtDMPs").Rows.Count - 1
        endRow = Range("UniqueRefDevisEtDMPs").Row + Range("DevisEtDMPs").Rows.Count - 1
        Range("'" & ActiveSheet.Name & "'" & "!$Z$" & startRow & ":$Z$" & endRow).Name = "UniqueRefDevisEtDMPs"
        Range("'" & ActiveSheet.Name & "'" & "!$AA$" & startRow & ":$AA$" & endRow).Name = "TaxeAmeublementN"
        Range("'" & ActiveSheet.Name & "'" & "!$AB$" & startRow & ":$AB$" & endRow).Name = "TaxeAmeublementR"
        Range("'" & ActiveSheet.Name & "'" & "!$AC$" & startRow & ":$AC$" & endRow).Name = "AggregationMontantsDevisTTC"
        'Copying the formulas to the lines we have to add
        thisRow = endRowBeforeUpdate
        While thisRow < endRow
            Call copyFormulasFromLine(startRow, startRow, "AA", "AC", thisRow + 1)
            thisRow = thisRow + 1
        Wend
    End If
End Sub

Sub refreshTaxeDAmeublementOfCurrentInvoice()
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

Private Function CustomMin(a As Range, b As Range) As Integer
    If b Is Nothing Then
        CustomMin = a.Row
    ElseIf a Is Nothing Then
        CustomMin = b.Row
    Else
        CustomMin = Application.Min(a.Row, b.Row)
    End If
End Function

Private Function CustomMax(a As Range, b As Range) As Integer
    If b Is Nothing Then
        CustomMax = a.Row
    ElseIf a Is Nothing Then
        CustomMax = b.Row
    Else
        CustomMax = WorksheetFunction.Max(a.Row, b.Row)
    End If
End Function

Sub checkMontantDevis()
    Dim devis As Range, montantDevis As Double, sommeArticlesDevis As Double
    Dim indexDevis As Integer: indexDevis = 1
    For Each devis In Range("DevisEtDMPs")
        sommeArticlesDevis = Range("AggregationMontantsDevisTTC").Rows(indexDevis).Value
        montantDevis = devis.Columns(1).Offset(, 4).Value
        'We compare the two values without the decimals as the rounding isn't always the same
        If CLng(sommeArticlesDevis) <> CLng(montantDevis) Then
            MsgBox "Attention ! Le montant TTC du devis " & devis.Value & " d??clar?? en " & devis.Offset(, 4).Address & " semble erron??. V??rifiez que tous les articles relatifs ?? ce devis ont ??t?? list??s."
        End If
        indexDevis = indexDevis + 1
    Next devis
End Sub
Sub updateReferenceOnUserCommand(rangeName As String, rngUpdated As Range, referenceName As String)
     If Not (Range(rangeName).Address = rngUpdated.Address) Then
        If MsgBox("La r??f??rence de " & referenceName & " semble ??ronn??e. La bonne r??f??rence semble ??tre " & _
        rngUpdated.Address & " au lieu de l'actuel : " & Range(rangeName).Address & ". Voulez-vous le mettre ?? jour maintenant ?", vbYesNo, "Demande de confirmation") = vbYes Then
            rngUpdated.Name = rangeName
            MsgBox "La r??f??rence de " & referenceName & " a ??t?? mise ?? jour !"
        End If
    End If
End Sub
Sub refreshNamedRanges()
    'Devis et DMP
    recapPosition = Range("recapPosition").Row
    startRow = CustomMin(Range("C1:C" & recapPosition).Find(what:="Selon devis", searchorder:=xlByRows, lookat:=xlWhole), _
    Range("C1:C" & recapPosition).Find(what:="Selon", searchorder:=xlByRows, lookat:=xlWhole))
    'As we don't want to take the Selon devis of the RECAPITULATIF page, we search our endRow above it
    endRow = CustomMax(Range("C1:C" & recapPosition).Find(what:="Selon devis", searchorder:=xlByRows, searchdirection:=xlPrevious, lookat:=xlWhole), _
    Range("C1:C" & recapPosition).Find(what:="Selon", searchorder:=xlByRows, searchdirection:=xlPrevious, lookat:=xlWhole))
    Dim rngUpdated As Range
    Set rngUpdated = Range("'" & ActiveSheet.Name & "'" & "!$D$" & startRow & ":$D$" & endRow)
    Call updateReferenceOnUserCommand("DevisEtDMPs", rngUpdated, "Devis et DMP")
   
    'Devis et DMP on RECAPITULATIF
    
    startRowRecap = CustomMin(Range("C" & recapPosition & ":C" & recapPosition + 1000).Find(what:="Selon devis", searchorder:=xlByRows, LookIn:=xlValues, lookat:=xlWhole), _
    Range("C" & recapPosition & ":C" & recapPosition + 1000).Find(what:="Selon", searchorder:=xlByRows, LookIn:=xlValues, lookat:=xlWhole))
    endRowRecap = CustomMax(Range("C" & recapPosition & ":C" & recapPosition + 1000).Find(what:="Selon devis", searchorder:=xlByRows, searchdirection:=xlPrevious, lookat:=xlWhole), _
    Range("C" & recapPosition & ":C" & recapPosition + 1000).Find(what:="Selon", searchorder:=xlByRows, searchdirection:=xlPrevious, lookat:=xlWhole))
    Set rngUpdated = Range("'" & ActiveSheet.Name & "'" & "!$D$" & startRowRecap & ":$D$" & endRowRecap)
    Call updateReferenceOnUserCommand("DevisEtDMPRecap", rngUpdated, "Devis et DMP du Recapitulatif")
    'We check if the number of devis and DMP is the same compared to the ones on first page. If not we raise an alert and ask to delete the extra lines
    nbLinesDevisAndDMPs = endRow - startRow + 1
    nbLinesDevisAndDMPsRecap = endRowRecap - startRowRecap + 1
    If nbLinesDevisAndDMPsRecap > nbLinesDevisAndDMPs Then
        If MsgBox("Le nombre de devis renseign?? dans le r??capitulatif semble ??ron??. Voulez-vous supprimer les lignes en trop ?", vbYesNo, "Demande de confirmation") = vbYes Then
            Range("A" & startRowRecap + nbLinesDevisAndDMPsRecap - 1 & ":S" & endRowRecap).Delete Shift:=xlUp
            MsgBox "Les lignes en trop dans le r??capitulatif ont bien ??t?? supprim??es !"
        End If
    End If

    
    'Articles listing
    'The " + 3" is to avoid the problems brought by the merged cells
    startRow = Application.Match("Ref devis et DMP", Range("C:C"), 0) + 3
    'To find the last line of the table, we search from the bottom the latest header of the
    'table ( + 3 ) and then jump to the end
    endRow = Range("C1:C100").Find(what:="Ref devis et DMP", searchorder:=xlByRows, searchdirection:=xlPrevious).Row + 3
    endRow = Range("C" & endRow).End(xlDown).Row
    Set rngUpdated = Range("'" & ActiveSheet.Name & "'" & "!$C$" & startRow & ":$C$" & endRow)
    Call updateReferenceOnUserCommand("RefsDevisEtDMP", rngUpdated, "Ref devis et DMP")
    Set rngUpdated = Range("'" & ActiveSheet.Name & "'" & "!$R$" & startRow & ":$R$" & endRow)
    Call updateReferenceOnUserCommand("MontantsArticlesHT", rngUpdated, "Montant HT des diff??rents articles")


    'Appel de fond
    'Getting the last line in which he have an "appel de fond". We add 2 to bypass the problems caused by the merged cells
    startRow = Application.Match("Appel de fond", Range("F:F"), 0) + 2
    'Let us move until the column "Base" before looking at the end. Indeed, the Range.End excel
    'function is mislead by the merged cell. Therefore we pick the first "safe" place in which we can run it
    endRow = Range("I" & startRow).End(xlDown).Row
    Set rngUpdated = Range("'" & ActiveSheet.Name & "'" & "!$K$" & startRow & ":$K$" & endRow)
    Call updateReferenceOnUserCommand("TauxTVAAppeles", rngUpdated, "Taux TVA de l'appel de fond")
    endRow = Range("I" & startRow).End(xlDown).Row
    Set rngUpdated = Range("'" & ActiveSheet.Name & "'" & "!$L$" & startRow & ":$M$" & endRow)
    Call updateReferenceOnUserCommand("MontantsTVAAppeles", rngUpdated, "Montants TVA de l'appel de fond")
    endRow = Range("I" & startRow).End(xlDown).Row
    Set rngUpdated = Range("'" & ActiveSheet.Name & "'" & "!$N$" & startRow & ":$O$" & endRow)
    Call updateReferenceOnUserCommand("MontantsTTCAppeles", rngUpdated, "Montants TTC de l'appel de fond")

    'Factures d'acompte
    startRow = Range("C1:D1000").Find(what:="Facture d'acompte :", searchorder:=xlByRows).Row
    endRow = Range("C1:D1000").Find(what:="Facture d'acompte :", searchorder:=xlByRows, searchdirection:=xlPrevious).Row + 1
    Set rngUpdated = Range("'" & ActiveSheet.Name & "'" & "!$H$" & startRow & ":$I$" & endRow)
    Call updateReferenceOnUserCommand("MontantsFacturesAppeles", rngUpdated, "Montants Factures appelees")
    Set rngUpdated = Range("'" & ActiveSheet.Name & "'" & "!$J$" & startRow & ":$K$" & endRow)
    Call updateReferenceOnUserCommand("MontantsPayesSurFacturesAppeles", rngUpdated, "Montants payes sur Factures appelees")
    Set rngUpdated = Range("'" & ActiveSheet.Name & "'" & "!$R$" & startRow & ":$S$" & endRow)
    Call updateReferenceOnUserCommand("RestantsDusSurFacturesAppeles", rngUpdated, "Montants restants sur Factures appelees")

    'Table d'aggregation de la taxe d'ameublement
    'Getting the position of the table "Aggr??gation de la taxe d'ameublement"
    startRow = Application.Match("Aggr??gation de la taxe d'ameublement", Range("Z:Z"), 0) + 3
    'We know that the table contains max 16 lines. Therefore we do our xlUp using that
    endRow = Range("Z" & startRow + 17).End(xlUp).Row
    Set rngUpdated = Range("'" & ActiveSheet.Name & "'" & "!$Z$" & startRow & ":$Z$" & endRow)
    Call updateReferenceOnUserCommand("UniqueRefDevisEtDMPs", rngUpdated, "R??f??rences des devis et DMP dans la table d'aggr??gation de la taxe d'ameublement")
    Set rngUpdated = Range("'" & ActiveSheet.Name & "'" & "!$AA$" & startRow & ":$AA$" & endRow)
    Call updateReferenceOnUserCommand("TaxeAmeublementN", rngUpdated, "Taxe d'ameublement N dans la table d'aggr??gation de la taxe d'ameublement")
    Set rngUpdated = Range("'" & ActiveSheet.Name & "'" & "!$AB$" & startRow & ":$AB$" & endRow)
    Call updateReferenceOnUserCommand("TaxeAmeublementR", rngUpdated, "Taxe d'ameublement R dans la table d'aggr??gation de la taxe d'ameublement")
    'Aggregation des montants TTC des devis
    Set rngUpdated = Range("'" & ActiveSheet.Name & "'" & "!$AC$" & startRow & ":$AC$" & endRow)
    Call updateReferenceOnUserCommand("AggregationMontantsDevisTTC", rngUpdated, "Montants TTC des devis dans la table d'aggr??gation de la taxe d'ameublement")
    'Checking the montant devis TTC announced are right
    Call checkMontantDevis


    'Impression des titres
    startRow = Range("ClientDetails").Row
    endRow = Range("ClientDetails").Row + Range("ClientDetails").Rows.Count
    endRowInvoiceDetails = Range("InvoiceDetails").Row + Range("InvoiceDetails").Rows.Count
    endRow = WorksheetFunction.Max(endRow, endRowInvoiceDetails)
    Set rngUpdated = Range("'" & ActiveSheet.Name & "'" & "!$" & startRow & ":$" & endRow)
    Call updateReferenceOnUserCommand("impression_des_titres", rngUpdated, "Impression des titres")
End Sub



