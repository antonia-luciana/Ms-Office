Attribute VB_Name = "Module1"
Sub lista()
    Dim wa As Word.Application
    Dim wd As Word.Document
    Set wa = New Word.Application
    wa.Visible = True
    Set wd = wa.Documents.Add
    With wd.PageSetup
        .TopMargin = CentimetersToPoints(2)
        .LeftMargin = CentimetersToPoints(2)
        .BottomMargin = CentimetersToPoints(2)
        .RightMargin = CentimetersToPoints(2)
    End With
    wa.Selection.TypeText ("Cãrþi")
    wa.Selection.TypeParagraph
    wd.Paragraphs(1).Range.Font.Bold = True
End Sub
