Attribute VB_Name = "MHD Pruefung nach Path verschieben"
Sub result()

    ' copy and change the form
    Dim rowcount As Integer
    Dim FL As Integer
    zeit = format(now(), "yyyymmdd_hh")
    All_Rows = Sheets("Tabelle2").UsedRange.Rows.Count
    'Rows = Sheets("Tabelle1").Range("A:A").Cells.Count
    Sheets("Tabelle2").Select
    Columns("P:P").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets.Add.Name = "FL_List"
    Sheets("FL_List").Select
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.Range("$A:$A").RemoveDuplicates Columns:=1, Header:=xlYes
    Dim fil As Integer
    FL = Sheets("FL_List").UsedRange.Rows.Count
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],Filialen!C[-1]:C,2,0)"
    Selection.AutoFill Destination:=Range("B2:B" & FL)
    For r = 2 To FL
        If Sheets("FL_List").Cells(r, 1).Value = "" Then
            Exit For
        End If
        FL_name = Sheets("FL_List").Cells(r, 1).Value  ' select the actualle sheet
        FilePath = Sheets("FL_List").Cells(r, 2).Value
        On Error Resume Next
            Sheets(FL_name).Select
            If Err <> 0 Then
                Sheets("Tabelle1").Select
                row_sheet = Sheets("Tabelle1").UsedRange.Rows.Count
                Rows("3:" & row_sheet).Select
                Selection.Delete Shift:=xlUp
                Sheets("Tabelle2").Select
                Range("P1").Select
                Selection.AutoFilter
                ActiveSheet.Range("$A$1:$P$" & All_Rows).AutoFilter Field:=16, Criteria1:=FL_name
                'Sheets.Add.Name = FL_name
                Sheets("Tabelle1").Select
                Sheets("Tabelle1").Cells(1, 7).Value = FL_name
                Sheets("Tabelle2").Select
                Range("A2:O" & All_Rows).Select
                Application.CutCopyMode = False
                Selection.Copy
                Sheets("Tabelle1").Select
                Range("A3").Select
                ActiveSheet.Paste
                With ActiveSheet.PageSetup
                    .PrintTitleRows = "$1:$2"
                    .PrintTitleColumns = ""
                End With
                Application.PrintCommunication = True
                ActiveSheet.PageSetup.PrintArea = ""
                Application.PrintCommunication = False
                With ActiveSheet.PageSetup
                    .LeftHeader = ""
                    .CenterHeader = ""
                    .RightHeader = ""
                    .LeftFooter = ""
                    .CenterFooter = ""
                    .RightFooter = ""
                    .LeftMargin = Application.InchesToPoints(0.15)
                    .RightMargin = Application.InchesToPoints(0.15)
                    .TopMargin = Application.InchesToPoints(0.75)
                    .BottomMargin = Application.InchesToPoints(0.75)
                    .HeaderMargin = Application.InchesToPoints(0.3)
                    .FooterMargin = Application.InchesToPoints(0.3)
                    .PrintHeadings = False
                    .PrintGridlines = False
                    .PrintComments = xlPrintNoComments
                    .PrintQuality = 600
                    .CenterHorizontally = False
                    .CenterVertically = False
                    .Orientation = xlLandscape
                    .Draft = False
                    .PaperSize = xlPaperA4
                    .FirstPageNumber = xlAutomatic
                    .Order = xlDownThenOver
                    .BlackAndWhite = False
                    .Zoom = 100
                    .PrintErrors = xlPrintErrorsDisplayed
                    .OddAndEvenPagesHeaderFooter = False
                    .DifferentFirstPageHeaderFooter = False
                    .ScaleWithDocHeaderFooter = True
                    .AlignMarginsHeaderFooter = True
                    .EvenPage.LeftHeader.Text = ""
                    .EvenPage.CenterHeader.Text = ""
                    .EvenPage.RightHeader.Text = ""
                    .EvenPage.LeftFooter.Text = ""
                    .EvenPage.CenterFooter.Text = ""
                    .EvenPage.RightFooter.Text = ""
                    .FirstPage.LeftHeader.Text = ""
                    .FirstPage.CenterHeader.Text = ""
                    .FirstPage.RightHeader.Text = ""
                    .FirstPage.LeftFooter.Text = ""
                    .FirstPage.CenterFooter.Text = ""
                    .FirstPage.RightFooter.Text = ""
                End With
                Application.PrintCommunication = True
                Columns("A:A").ColumnWidth = 8
                Columns("B:B").ColumnWidth = 27
                Columns("C:C").ColumnWidth = 7
                Columns("D:D").ColumnWidth = 6
                Columns("E:E").ColumnWidth = 11
                Columns("F:F").ColumnWidth = 6
                Columns("G:G").ColumnWidth = 8
                Columns("H:H").ColumnWidth = 5
                Columns("I:I").ColumnWidth = 8
                Columns("J:J").ColumnWidth = 10
                Columns("K:K").ColumnWidth = 6
                Columns("L:L").ColumnWidth = 6
                Columns("M:M").ColumnWidth = 6
                Columns("N:N").ColumnWidth = 6
                Columns("O:O").ColumnWidth = 12
                ' �ϲ���Ԫ��
                row1 = Sheets("Tabelle1").UsedRange.Rows.Count
                Rows("3:" & row1).Select
                Selection.RowHeight = 40
                With Selection
                    .HorizontalAlignment = xlLeft
                    .VerticalAlignment = xlCenter
                    .WrapText = True
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = False
                End With
                With Selection.Font
                    .Name = "Calibri"
                    .Size = 12
                    .Strikethrough = False
                    .Superscript = False
                    .Subscript = False
                    .OutlineFont = False
                    .Shadow = False
                    .Underline = xlUnderlineStyleNone
                    .Color = -16777216
                    .TintAndShade = 0
                    .ThemeFont = xlThemeFontNone
                End With
                Range("A2:O" & row1).Select
                ActiveWorkbook.Worksheets("Tabelle1").Sort.SortFields.Clear
                ActiveWorkbook.Worksheets("Tabelle1").Sort.SortFields.Add2 Key:=Range( _
                    "E3:E" & row1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
                    xlSortNormal
                With ActiveWorkbook.Worksheets("Tabelle1").Sort
                    .SetRange Range("A2:O" & row1)
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
                Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                With Selection.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlInsideVertical)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlInsideHorizontal)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                'Call DoppeltMarker
                'MyTime = Range("L1").Value
                MyTime = zeit & "Uhr"
                Set ws = ActiveSheet
                Set mainWorkbook = ThisWorkbook
                Filename = FL_name & "_" & MyTime & ".pdf"
                On Error GoTo SaveErr
                        savePath = FilePath & "\" & Filename
                        ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=savePath, _
                                           Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                                           IgnorePrintAreas:=False, OpenAfterPublish:=False
                End If
                Next
                Sheets("FL_List").Select
                ActiveWindow.SelectedSheets.Delete
                Exit Sub
SaveErr:
                        savePath = mainWorkbook.Path & "\" & FL_name & "_" & MyTime & ".pdf"
                        ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=savePath, _
                                           Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                                           IgnorePrintAreas:=False, OpenAfterPublish:=False
                        Resume Next
End Sub


Sub DoppeltMarker()
    Columns("A:A").Select
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub














