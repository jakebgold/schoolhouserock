Attribute VB_Name = "SHR"
Sub SchoolHouseRock()
Application.Templates.LoadBuildingBlocks

Dim objDoc As Document
 
Set objDoc = ActiveDocument
 
  With objDoc.Styles("Line Number").Font
    .Name = "Times New Roman"
    .Size = 12
    .ColorIndex = wdBlack
  End With
  
    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
    ActiveDocument.AttachedTemplate.BuildingBlockEntries("Page X of Y").Insert Where:=Selection.Range, _
        RichText:=True
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument

With Selection.PageSetup
        With .LineNumbering
            .Active = True
            .StartingNumber = 1
            .CountBy = 1
            .RestartMode = wdRestartContinuous
            .DistanceFromText = wdAutoPosition
        End With
        .Orientation = wdOrientPortrait
        .TopMargin = InchesToPoints(1)
        .BottomMargin = InchesToPoints(1)
        .LeftMargin = InchesToPoints(1)
        .RightMargin = InchesToPoints(1)
        .Gutter = InchesToPoints(0)
        .HeaderDistance = InchesToPoints(0.5)
        .FooterDistance = InchesToPoints(0.5)
        .PageWidth = InchesToPoints(8.5)
        .PageHeight = InchesToPoints(11)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .BookFoldPrintingSheets = 1
        .GutterPos = wdGutterPosLeft
    End With
    
    
With ActiveDocument.Styles("Footer")
  .ParagraphFormat.Alignment = wdAlignParagraphCenter
End With


Selection.Font.Color = RGB(0, 0, 0)
Selection.Font.Size = 12
Selection.Font.Name = "Times New Roman"
Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.Find.ClearFormatting
    Selection.Find.Font.Italic = True
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Italic = True
        .Underline = True
    End With
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     With Selection.Find.Font
        .StrikeThrough = True
        .DoubleStrikeThrough = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Superscript = False
        .Subscript = False
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Underline = wdUnderlineNone
        .StrikeThrough = True
        .DoubleStrikeThrough = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Superscript = False
        .Subscript = False
    End With
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
      If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Selection.TypeText Text:="NOT PREPARED BY DLS"
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    With ActiveDocument.Styles("Header").Font
        .Name = "Times New Roman"
        .Size = 12
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 0
        .Animation = wdAnimationNone
    End With
    With ActiveDocument.Styles("Header")
        .AutomaticallyUpdate = False
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Header"
    End With
    
 
Dim revs As Word.Revisions
  Dim rev As Word.Revision, revOld As Word.Revision
  Dim rngDoc As Word.Range
  Dim rngRevNew As Word.Range, rngRevOld As Word.Range
  Dim authMain As String, authNew As String, authOld As String
  Dim bReject As Boolean

  bReject = False
  Set rngDoc = ActiveDocument.Content
  Set revs = rngDoc.Revisions
  If revs.Count > 0 Then
    authMain = revs(1).Author
  Else 'No revisions so...
    Exit Sub
  End If

  For Each rev In revs
    'rev.Range.Select  'for debugging, only
    authNew = rev.Author
    If rev.Type = wdRevisionInsert Or wdRevisionDelete Then
        Set rngRevNew = rev.Range
        'There's only something to compare if an Insertion
        'or Deletion have been made prior to this
        If Not rngRevOld Is Nothing Then
            'The last revision was rejected, so we need to check
            'whether the next revision (insertion for a deletion, for example)
            'is adjacent and reject it, as well
            If bReject Then
                If rngRevNew.Start - rngRevOld.End <= 1 And authNew <> authMain Then
                    rev.Reject
                End If
                bReject = False 'reset in any case
            End If

            'If the authors are the same there's no conflict
            If authNew <> authOld Then
                'If the current revision is not the main author
                'and his revision is in the same range as the previous
                'this means his revision has replaced that
                'of the main author and must be rejected.
                If authNew <> authMain And rngRevNew.InRange(rngRevOld) Then
                    rev.Reject
                    bReject = True
                'If the previous revision is not the main author
                'and the new one is in the same range as the previous
                'this means that revision has replaced this one
                'of the main author and the previous must be rejected.
                ElseIf authOld <> authMain And rngRevOld.InRange(rngRevNew) Then
                    revOld.Reject
                    bReject = True
                End If
            End If
        End If
        Set rngRevOld = rngRevNew
        Set revOld = rev
        authOld = authNew
    End If

  Next
  
Dim chgAdd As Word.Revision

If ActiveDocument.Revisions.Count = 0 Then
    MsgBox "There are no revisions in this document", vbOKOnly
Else
    ActiveDocument.TrackRevisions = False
    For Each chgAdd In ActiveDocument.Revisions
        If chgAdd.Type = wdRevisionDelete Then
            chgAdd.Range.Font.StrikeThrough = True
            chgAdd.Range.Font.Color = wdColorBlack
            chgAdd.Reject
        ElseIf chgAdd.Type = wdRevisionInsert Then
            chgAdd.Range.Font.Color = wdColorBlack
            chgAdd.Range.Font.Underline = wdUnderlineSingle
            chgAdd.Range.Font.Bold = True
            chgAdd.Range.Font.Italic = True
            chgAdd.Accept
        Else
            MsgBox ("Unexpected Change Type Found"), vbOKOnly + vbCritical
            chgAdd.Range.Select ' move insertion point
        End If
    Next chgAdd
End If




End Sub

