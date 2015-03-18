Sub ExportCommentsToNewDoc()
  ' Basic macro, without section number, table reformatting, and save-as, is thanks to
  ' http://www.wordbanter.com/showthread.php?t=129695
  Dim Source As Document, Target As Document
  Dim TTable As Table
  Dim TRow As Row
  Dim acomment As Comment
  Set Source = ActiveDocument
  Set Target = Documents.Add
  Set TTable = Target.Tables.Add(Target.Range, 1, 5)
  TTable.AutoFitBehavior (wdAutoFitContent)
  With TTable.Rows(1)
    .Cells(1).Range.Text = "Comment"
    .Cells(2).Range.Text = "Commenter"
    .Cells(3).Range.Text = "Section"
    .Cells(4).Range.Text = "Page Number"
    .Cells(5).Range.Text = "Response"
  End With
 
  For Each acomment In Source.Comments
    Set TRow = TTable.Rows.Add
   
    ' Comment and author
    TRow.Cells(1).Range.Text = acomment.Range
    TRow.Cells(2).Range.Text = acomment.Author
 
    ' Page number
    acomment.Scope.Select
    TRow.Cells(4).Range.Text = Selection.Information(wdActiveEndAdjustedPageNumber)
   
    ' Section number -- i.e. the number of the first preceding line in heading style
    Dim cPara As Paragraph
    Set cPara = Selection.Range.Paragraphs(1)
    Dim Counter As Integer
    Counter = 0
 
    Do While True = True
      Counter = Counter + 1
      ' Check for heading -- would be better to do this with Outline Level but this will do
      If Left(cPara.Range.Style, Len("Heading")) = "Heading" Then
        ' use of ListString thanks to http://www.word.mvps.org/faqs/numbering/ListString.htm
        TRow.Cells(3).Range.Text = cPara.Range.ListFormat.ListString
        Exit Do
      End If
      ' Check for start of document
      If (ActiveDocument.Range(0, cPara.Range.End).Paragraphs.Count) = 1 Then
        Exit Do
      End If
      ' Avoid infinite loops
      If Counter = 50 Then
        Exit Do
      End If
      Set cPara = cPara.Previous(1)
    Loop
 
  Next acomment
 
  ' Uncomment this next line to skip the table formatting
  ' GoTo SkipFormattingTable
 
  ' Fix the cell borders -- thanks to
  ' http://word.tips.net/T000880_Setting_a_Default_Table_Border_Width.html
  ' Work through all cells in each table
  Dim objCell As Cell
  For Each objCell In TTable.Range.Cells
    ' Work through all borders in each cell
    With objCell.Borders(wdBorderLeft)
      .Color = wdColorBlack
      .LineStyle = wdLineStyleSingle
      .LineWidth = wdLineWidth075pt
    End With
    With objCell.Borders(wdBorderRight)
      .Color = wdColorBlack
      .LineStyle = wdLineStyleSingle
      .LineWidth = wdLineWidth075pt
    End With
    With objCell.Borders(wdBorderTop)
      .Color = wdColorBlack
      .LineStyle = wdLineStyleSingle
      .LineWidth = wdLineWidth075pt
    End With
    With objCell.Borders(wdBorderBottom)
      .Color = wdColorBlack
      .LineStyle = wdLineStyleSingle
      .LineWidth = wdLineWidth075pt
    End With
    Next objCell
 
SkipFormattingTable:
    Dim TargetName As String
    TargetName = Source.FullName
    TargetStartExtension = InStrRev(TargetName, ".") - 1
    Dim TargetExtension As String
    TargetExtension = Right(TargetName, Len(TargetName) - TargetStartExtension)
    TargetBaseName = Left(TargetName, TargetStartExtension)
    Target.SaveAs (TargetBaseName & "-ExtractedComments" & TargetExtension)
 
End Sub
