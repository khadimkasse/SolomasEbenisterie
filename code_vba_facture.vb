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
    endRow = Range("ClientDetails").Row + Range("ClientDetails").Rows.Count - 1 ' Application.Match("Adresse 1", Range("H1:H18"), 0)
    endRowInvoiceDetails = Range("InvoiceDetails").Row + Range("InvoiceDetails").Rows.Count - 1
    endRow = WorksheetFunction.Max(endRow, endRowInvoiceDetails)
    lineInsert = endRow + 1
    Call addOneLine(endRow, lineInsert, "H", "M")
    'Refreshing the named ranges ClientDetails, InvoiceDetails and impression_des_titres
    Range("'template'!$H$" & startRow & ":$M$" & endRow + 1).Name = "ClientDetails"
    Range("'template'!$P$" & startRowInvoiceDetails & ":$S$" & endRow + 1).Name = "InvoiceDetails"
    Range("'template'!$" & startRow & ":$" & endRow + 1).Name = "impression_des_titres"
End Sub

Sub addOneLineToInvoiceDetails()
    'Getting the last row of the client identification part or the invoice details part
    startRow = Range("InvoiceDetails").Row
    startRowClientDetails = Range("ClientDetails").Row
    endRow = Range("InvoiceDetails").Row + Range("InvoiceDetails").Rows.Count - 1 ' Application.Match("Adresse 1", Range("H1:H18"), 0)
    endRowClientDetails = Range("ClientDetails").Row + Range("ClientDetails").Rows.Count - 1
    endRow = WorksheetFunction.Max(endRow, endRowClientDetails)
    lineInsert = endRow + 1
    Call addOneLine(endRow, lineInsert, "P", "S")
    'Refreshing the named ranges ClientDetails, InvoiceDetails and impression_des_titres
    Range("'template'!$P$" & startRow & ":$S$" & endRow + 1).Name = "InvoiceDetails"
    Range("'template'!$H$" & startRowClientDetails & ":$M$" & endRow + 1).Name = "ClientDetails"
    Range("'template'!$" & startRowClientDetails & ":$" & endRow + 1).Name = "impression_des_titres"
End Sub

Sub addOneLineOfDevisOrDMP()
    'Getting the last line in which we defined a devis or DMP
    startRow = Range("DevisEtDMPs").Row
    endRow = Range("DevisEtDMPs").Row + Range("DevisEtDMPs").Rows.Count - 1
    lineInsert = endRow + 1
    Call addOneLine(endRow, lineInsert, "C", "H")
    'Getting the same content on Recapitulatif
    startRowRecap = Range("DevisEtDMPRecap").Row
    endRowRecap = Range("DevisEtDMPRecap").Row + Range("DevisEtDMPRecap").Rows.Count - 1
    lineInsertRecap = endRowRecap + 1
    Call addOneLine(endRowRecap, lineInsertRecap, "C", "H")
    'Replicate the formulas of the previous line
    Call copyFormulasFromLine(endRowRecap, endRowRecap, "C", "H", endRowRecap + 1)
    'Updating the DevisEtDMP range
    Range("'template'!$D$" & startRow & ":$D$" & endRow + 1).Name = "DevisEtDMPs"
    Range("'template'!$D$" & startRowRecap & ":$D$" & endRowRecap + 1).Name = "DevisEtDMPRecap"
    Range("C" & endLine + 1).Select
End Sub
'Not used for now as the direct formula seems to be more convenient
Sub refreshAFormula(cellTo, columnFromStart, columnFromEnd, startRow, endRow)
    Range(cellTo).Value = "$" & columnFrom & "$" & startRow & "$" & columnFrom & "$" & endRow
End Sub
Sub refreshFormulasDependingOnArticlesListing(startRow, endRow)
    Range("Z36").Value = "$C$" & startRow & ":$C$" & endRow
    Range("AA36").Value = "$D$" & startRow & ":$D$" & endRow
    Range("AB36").Value = "$E$" & startRow & ":$Q$" & endRow
    Range("AC36").Value = "$R$" & startRow & ":$R$" & endRow
    Range("AD36").Value = "$S$" & startRow & ":$S$" & endRow
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

Sub refreshFormulasDependingOnAppelDeFondTable(startRow, endRow)
    Range("Z40").Value = "$K$" & startRow & ":$K$" & endRow
    Range("AA40").Value = "$L$" & startRow & ":$M$" & endRow
    Range("AB40").Value = "$N$" & startRow & ":$O$" & endRow
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

End Sub

Sub addOneInvoiceToRecapitulatif()
    'Getting the starting and ending row of the "Factures d'acompte"
    startRow = Range("MontantsFacturesAppeles").Row 
    endRow =  Range("MontantsFacturesAppeles").Row + Range("MontantsFacturesAppeles").Rows.Count - 1 ' Range("C1:D1000").Find(what:="Facture d'acompte :", searchorder:=xlByRows, searchdirection:=xlPrevious).Row + 1
    lineInsert = endRow + 1
    'Beware: we have two lines to add here
    Call addLinesToDoubleRowTables(endRow, lineInsert, "C", "S")
    Range("C" & endRow + 1).Value = "Facture d'acompte :"
    Range("H" & endRow + 1).Formula = "=" & Range("TotalTTCFacture").Address ' "=$L$" & invoiceTotalTTC
    'Refreshing the named ranges based on the table Factures d'acompte
    'Call refreshFormulasDependingOnFacturesDAcomptesTable(startRow, endRow + 2)
    Range("'template'!$H$" & startRow & ":$I$" & endRow + 2).Name = "MontantsFacturesAppeles"
    Range("'template'!$J$" & startRow & ":$K$" & endRow + 2).Name = "MontantsPayesSurFacturesAppeles"
    Range("'template'!$R$" & startRow & ":$S$" & endRow + 2).Name = "RestantsDusSurFacturesAppeles"
End Sub
Private Sub oneArticleAdded(ByVal currentCell As Range, currentLine)
    If currentCell.Address = "C" & currentLine Then
        
    End If
End Sub
Sub addLineToTaxeDAmeublementTable()
    'Getting the position of the table "Aggrégation de la taxe d'ameublement"
    startRow = Application.Match("Aggrégation de la taxe d'ameublement", Range("Z:Z"), 0) + 3
    endRow = Range("Z" & startRow).End(xlDown).Row
    'Testing whether the current devis references are all different or not
    i = startRow
    tableNeedsToBeErased = False
    Do While i <= endRow
        j = i + 1
        Do While j <= endRow
            If Range("Z" & i).Value = Range("Z" & j).Value Then
                MsgBox ("La référence " & Range("Z" & i).Value & " apparait deux fois dans le tableau d'aggrégation. Le tableau va être réinitialisé.")
                tableNeedsToBeErased = True
            End If
            j = j + 1
        Loop
        i = i + 1
    Loop
    If tableNeedsToBeErased Then
        Range("Z" & startRow + 1 & ":AC" & endRow).Select
        Selection.ClearContents
        endRow = startRow
    End If
    'Looping through all the articles and comparing it to the entries in the Tablea d'aggrégation t decide whether we need an extra line or not
    'Getting the starting row and the ending row of the articles listing
    artStartRow = Application.Match("Ref devis et DMP", Range("C:C"), 0) + 3
    artEndRow = Range("C" & startRow).End(xlDown).Row
    thisDevisReferenceRow = startRow
    For Each thisArticle In Range("C" & artStartRow & ":C" & artEndRow)
        While thisDevisReferenceRow <= endRow
            If thisArticle.Value = Range("Z" & thisDevisReferenceRow).Value Then
                'The current devis exist in the agreagation table. We can then take another one
                GoTo ContinueFor
            End If
            thisDevisReferenceRow = thisDevisReferenceRow + 1
        Wend
        'Saving the devis reference that is not present in the aggragation table
        Call copyFormulasFromLine(startRow, startRow, "Z", "AC", endRow + 1)
        endRow = endRow + 1
        Range("Z" & endRow).Value = thisArticle.Value
ContinueFor:
    Next thisArticle
    
    Call refreshFormulasDependingOnTaxeDAmeublementTable(startRow, endRow)
End Sub
Sub refreshFormulasDependingOnTaxeDAmeublementTable(startRow, endRow)
    Range("Z23").Value = "$Z$" & startRow & ":$AA$" & endRow
    Range("AA23").Value = "$AB$" & startRow & ":$AB$" & endRow
    Range("AB23").Value = "$AC$" & startRow & ":$AC$" & endRow
End Sub

Sub refreshTaxeDAmeublemnentOfCurrentInvoice(startRow, endRow)
    i = startRow
    formulaTaxeDAmeublement = "=IF($AA$11=""Y"","
    formulaTaxeDAmeublement = formulaTaxeDAmeublement & _
        "$F$" & i & "*(SUMIF(INDIRECT($Z$23)," & "$G$" & i & ",INDIRECT($AA$23))+" _
        & "SUMIF(INDIRECT($Z$23)," & "$G$" & i & ",INDIRECT($AB$23)))"
    i = i + 2
    Do While i <= endRow
        formulaTaxeDAmeublement = formulaTaxeDAmeublement & _
        "+$F$" & i & "*(SUMIF(INDIRECT($Z$23)," & "$G$" & i & ",INDIRECT($AA$23))+" _
        & "SUMIF(INDIRECT($Z$23)," & "$G$" & i & ",INDIRECT($AB$23)))"
        i = i + 2
    Loop
    formulaTaxeDAmeublement = formulaTaxeDAmeublement & ","""")"
    'Puting the formula at the right place
    lineOfTaxeDAmeublement = Application.Match("** dont Taxe d'ameublement (A) 0,18% : ", Range("H:H"), 0)
    Range("L" & lineOfTaxeDAmeublement).Formula = formulaTaxeDAmeublement
End Sub

Sub refreshFormulasDependingOnFacturesDAcomptesTable(startRow, endRow)
    Range("Z93").Value = "$H$" & startRow & ":$I$" & endRow
    Range("AA93").Value = "$R$" & startRow & ":$S$" & endRow
End Sub

Sub refreshAllReferences()
    'Refreshing formulas based on the table of Articles
    'Getting the last line in which we have an article. We add 3 to bypass the problems caused by the merged cells
    startRow = Application.Match("Ref devis et DMP", Range("C:C"), 0) + 3
    'As it is possible to have multiple headers for the table, we search the last one
    endRow = Range("C1:C100").Find(what:="Ref devis et DMP", searchorder:=xlByRows, searchdirection:=xlPrevious).Row + 3
    endRow = Range("C" & endRow).End(xlDown).Row
    Call refreshFormulasDependingOnArticlesListing(startRow, endRow)
    
    'Refreshing formulas based on the table of Appel de fond
    startRow = Application.Match("Appel de fond", Range("F:F"), 0) + 2
    'Let us move until the column "Base" before looking at the end. Indeed, the Range.End excel
    'function is mislead by the merged cell. Therefore we pick the first "safe" place in which we can run it
    endRow = Range("I" & startRow).End(xlDown).Row
    
    'Refreshing the formulas of taxe d'ameublement of current invoice
    Call refreshTaxeDAmeublemnentOfCurrentInvoice(startRow, endRow)
    Call refreshFormulasDependingOnAppelDeFondTable(startRow, endRow)
    
    'Refreshing the formulas based on the table Taxe d'ameublement
    startRow = Application.Match("Aggrégation de la taxe d'ameublement", Range("Z:Z"), 0) + 3
    endRow = Range("Z" & startRow).End(xlDown).Row
    Call refreshFormulasDependingOnTaxeDAmeublementTable(startRow, endRow)
    
    'Refreshing the formulas based on the table Factures d'acompte
    startRow = Range("C1:D1000").Find(what:="Facture d'acompte :", searchorder:=xlByRows).Row
    endRow = Range("C1:D1000").Find(what:="Facture d'acompte :", searchorder:=xlByRows, searchdirection:=xlPrevious).Row + 1
    Call refreshFormulasDependingOnFacturesDAcomptesTable(startRow, endRow)
    
    'Refreshing the rest of the range names
    
End Sub

Sub updateReferenceOnUserCommand(rangeName As String, rngUpdated As Range, referenceName As String)
     If Not (Range(rangeName).Address = rngUpdated.Address) Then
        If MsgBox("La référence de " & referenceName & " semble éronné. La bonne référence sembl être " & _
        rngUpdated.Address & " au lieu de l'actuel : " & Range(rangeName).Address & ". Voulez vous le mettre à jour automatiquement ?", vbYesNo, "Demande de confirmation") = vbYes Then
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
    startRowRecap = recapPosition + Application.Match("Selon devis", Range("C" & recapPosition & ":C" & recapPosition + 100), 0)
    endRowRecap = Range("C" & startRowRecap).End(xlDown).Row
    Set rngUpdated = Range("'template'!$D$" & startRowRecap & ":$D$" & endRowRecap)
    Call updateReferenceOnUserCommand("DevisEtDMPRecap", rngUpdated, "Devis et DMP du Recapitulatif")
    
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

    'Impression des titres
    startRow = Range("ClientDetails").Row
    endRow = Range("ClientDetails").Row + Range("ClientDetails").Rows.Count - 1
    endRowInvoiceDetails = Range("InvoiceDetails").Row + Range("InvoiceDetails").Rows.Count - 1
    endRow = WorksheetFunction.Max(endRow, endRowInvoiceDetails)
    Set rngUpdated = Range("'template'!$" & startRow & ":$" & endRow)
    Call updateReferenceOnUserCommand("impression_des_titres", rngUpdated, "Impression des titres")
End Sub

