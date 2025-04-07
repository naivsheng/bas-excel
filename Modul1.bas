Attribute VB_Name = "Angebotelist"
Sub result()

    ' copy and change the form
    Dim rowcount As Integer
    rowcount = Sheets("Sheet1").UsedRange.Rows.Count
    
    Sheets.Add.Name = "Result"
    Sheets("Sheet1").Select
    Columns("A:G").Select   'all of the data
    Selection.Copy
    Sheets("Result").Select
    ActiveSheet.Paste
    ' column 'gueltig zu' split
    Columns("G:G").Select
    Selection.TextToColumns Destination:=Range("G1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=".", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-1],""-"",RC[-2],""-"",RC[-3])"
    Range("J1").Select
    Selection.AutoFill Destination:=Range("J1:J" & rowcount)
    Range("J1:J" & rowcount).Select
    Selection.Copy
    Range("K1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("G:J").Select
    Selection.ClearContents
    ' column 'gueltig ab' split
    Columns("F:F").Select
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("F1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=".", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-1],""-"",RC[-2],""-"",RC[-3])"
    Range("I1").Select
    Selection.AutoFill Destination:=Range("I1:I" & rowcount)
    Range("I1:I" & rowcount).Select
    Selection.Copy
    Range("J1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("J:K").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("F:G").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("H:K").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "G�ltig zu"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "G�ltig ab"
    
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Kategorie"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNA(VLOOKUP(RC[-6],Kategorie!C[-7]:C[-6],2,FALSE)),""MHD"",VLOOKUP(RC[-6],Kategorie!C[-7]:C[-6],2,FALSE))"
    Range("H2").Select
    Selection.AutoFill Destination:=Range("H2:H" & rowcount)
    
    ' ����
    Cells.Select
    ActiveWorkbook.Worksheets("Result").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Result").Sort.SortFields.Add2 Key:=Range( _
        "H2:H1189"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Result").Sort.SortFields.Add2 Key:=Range( _
        "G2:G1189"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Result").Sort
        .SetRange Range("A1:J" & rowcount)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' ��������Ա����
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "NR."
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=RC[1]&COUNTIF(R1C2:RC[1],RC[1])"
    Range("A2").Select
    Selection.AutoFill Destination:=Range("A2:A" & rowcount)
    
    ' ������ҳ
    
    Dim FL As Integer
    FL = Sheets("Filialen").UsedRange.Rows.Count
    Dim fil As Integer
    For r = 2 To FL
        If Sheets("Filialen").Cells(r, 1).Value = "" Then
            Exit For
        End If
        FL_name = Sheets("Filialen").Cells(r, 2).Value  ' select the actualle sheet
        FL_id = Sheets("Filialen").Cells(r, 3).Value
        On Error Resume Next
            Sheets(FL_name).Select
            If Err <> 0 Then
                Sheets.Add.Name = FL_name
                Sheets(FL_name).Select
                Sheets(FL_name).Cells(1, 1).Value = Sheets("Filialen").Cells(r, 1).Value
                Sheets("Result").Select
                Range("C1:I1").Select
                Application.CutCopyMode = False
                Selection.Copy
                Sheets(FL_name).Select
                Range("A2").Select
                ActiveSheet.Paste
                Range("A3").Select
                Application.CutCopyMode = False
                ActiveCell.FormulaR1C1 = _
                    "=IFERROR(VLOOKUP(R1C1&ROW(R[-2]C),Result!C:C[2],3,0),"""")"
                Range("A3").Select
                Application.CutCopyMode = False
                Range("B3").Select
                ActiveCell.FormulaR1C1 = _
                    "=IFERROR(VLOOKUP(R1C1&ROW(R[-2]C[-1]),Result!C[-1]:C[2],4,0),"""")"
                Range("C2").Select
                ActiveCell.FormulaR1C1 = "Preis"
                Range("C3").Select
                ActiveCell.FormulaR1C1 = _
                    "=IFERROR(VLOOKUP(R1C1&ROW(R[-2]C[-2]),Result!C[-2]:C[2],5,0),"""")"
                Range("D3").Select
                ActiveCell.FormulaR1C1 = _
                    "=IFERROR(VLOOKUP(R1C1&ROW(R[-2]C[-3]),Result!C[-3]:C[2],6,0),"""")"
                Range("E3").Select
                ActiveCell.FormulaR1C1 = _
                    "=IFERROR(VLOOKUP(R1C1&ROW(R[-2]C[-4]),Result!C[-4]:C[2],7,0),"""")"
                Range("F3").Select
                ActiveCell.FormulaR1C1 = _
                    "=IFERROR(VLOOKUP(R1C1&ROW(R[-2]C[-5]),Result!C[-5]:C[2],8,0),"""")"
                Range("G3").Select
                ActiveCell.FormulaR1C1 = _
                    "=IFERROR(VLOOKUP(R1C1&ROW(R[-2]C[-6]),Result!C[-6]:C[2],9,0),"""")"
                Range("A3:G3").Select
                Selection.AutoFill Destination:=Range("A3:G105"), Type:=xlFillDefault
                Range("H2").Select
                ActiveCell.FormulaR1C1 = "Menge"
                Range("I2").Select
                ActiveCell.FormulaR1C1 = "Werbung"
                Application.PrintCommunication = False
                With ActiveSheet.PageSetup
                    .PrintTitleRows = ""
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
                    .LeftMargin = Application.InchesToPoints(0.25)
                    .RightMargin = Application.InchesToPoints(0.25)
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
                Columns("A:A").ColumnWidth = 10
                Columns("B:B").ColumnWidth = 60
                Columns("C:C").ColumnWidth = 8
                Columns("D:D").ColumnWidth = 8
                Columns("E:E").ColumnWidth = 10
                Columns("F:F").ColumnWidth = 10
                Columns("G:G").ColumnWidth = 14
                Columns("H:H").ColumnWidth = 8
                Columns("I:I").ColumnWidth = 8
                ' �ϲ���Ԫ��
                Range("A1:I1").Select
                With Selection
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = False
                End With
                Selection.Merge
               ' copy from Sheet2
                Range("J2").Select
                ' ActiveCell.FormulaR1C1 = "=COUNT(R[1]C[-9]:R[103]C[-9])"
                For row_1 = 3 To Sheets(FL_name).UsedRange.Rows.Count
                    If Sheets(FL_name).Cells(row_1, 1) = "" Then
                        Exit For
                    End If
                Next
                If row_1 = 3 Then
                    row_1 = 3
                Else
                    row_1 = row_1 + 1
                End If
                'ActiveCell.FormulaR1C1 = "=COUNTBLANK(R[1]C[-9]:R[103]C[-9])"
                'row_1 = Sheets(FL_name).Cells(2, 10).Value
                row_2 = Sheets("Sheet2").UsedRange.Rows.Count
                If row_2 > 1 Then
                    Sheets("Sheet2").Select
                    range_2 = "A2:I" & row_2
                    Range(range_2).Select
                    Selection.Copy
                    range_1 = "A" & row_1
                    'If row_1 = 103 Then
                    '    range_1 = "A3"
                    'Else
                    '    range_1 = "A" & 105 - row_1 + 2
                    'End If
                    Sheets(FL_name).Select
                    Range(range_1).Select
                    ActiveSheet.Paste
                    row_1 = row_1 + row_2
                    'end func
                End If
                'row_1_neu = Sheets(FL_name).Cells(2, 10).Value
                'Sheets(FL_name).Cells(2, 10).Value = ""
                'If row_1_neu <> row_1 Then
                '    range_1 = 105 - row_1_neu + 3
                'Else
                '    range_1 = 105 - row_1_neu + 2
                'End If
                
                Call from_cockpit(row_1, FL_id, FL_name)
                                
                ' ���ӱ߿�
                Sheets(FL_name).Select
                Range("A1:I105").Select
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

                Rows("1:1").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Range("A1").Select
                ActiveCell.FormulaR1C1 = _
                    "Datum:_________________       Unterschrift:_____________________"
                For row_1_neu = row_1 To Sheets(FL_name).UsedRange.Rows.Count
                    If Sheets(FL_name).Cells(row_1_neu, 1) = "" Then
                        Exit For
                    End If
                Next
                range_1 = row_1_neu + 1 & ":105"
                Rows(range_1).Select
                Selection.Delete Shift:=xlUp
            End If
    Next
End Sub

Sub get_FL()
    On Error Resume Next
    Sheets("Filialen").Select
    If Err <> 0 Then
        Sheets.Add.Name = "Filialen"
        Sheets("Sheet1").Select
        Dim rowcount As Integer
        rowcount = Sheets("Sheet1").UsedRange.Rows.Count
        Columns("A:A").Select
        Selection.Copy
        Sheets("Filialen").Select
        Range("A1").Select
        ActiveSheet.Paste
        Columns("A:A").Select
        Application.CutCopyMode = False
        ActiveSheet.Range("$A$1:$A$" & rowcount).RemoveDuplicates Columns:=1, Header:= _
            xlYes
    End If
End Sub

Sub from_cockpit(row, FL_id, FL_name)
    On Error Resume Next
    'If Sheets("Tabelle1").Cells(1, 1) = "BranchId" Then
    '    r_start = 2
    'Else
    '    r_start = 1
    'End If
    For i = 1 To Sheets("Tabelle1").UsedRange.Columns.Count
        If Sheets("Tabelle1").Cells(1, i) = "KategorieId" Then
            KategorieId = i
        ElseIf Sheets("Tabelle1").Cells(1, i) = "Kommentar" Then
            Kommentar = i
        ElseIf Sheets("Tabelle1").Cells(1, i) = "Beschreibung" Then
            Beschreibung = i
        ElseIf Sheets("Tabelle1").Cells(1, i) = "AP" Then
            AP = i
        ElseIf Sheets("Tabelle1").Cells(1, i) = "EndDate" Then
            EDate = i
        ElseIf Sheets("Tabelle1").Cells(1, i) = "StartDate" Then
            SDate = i
        ElseIf Sheets("Tabelle1").Cells(1, i) = "BranchId" Then
            BranchId = i
        End If
    Next i
    r_start = 2
    For r_1 = r_start To Sheets("Tabelle1").UsedRange.Rows.Count
        kategorie = Sheets("Tabelle1").Cells(r_1, KategorieId).Value
        EndDate = CDate(Sheets("Tabelle1").Cells(r_1, EDate).Value)
        today = Date
        If kategorie = 3 And EndDate >= today Then
            p = -1
            p = InStr(1, Sheets("Tabelle1").Cells(r_1, BranchId).Value, ",")
            If p >= 0 Then
                branch = Split(Sheets("Tabelle1").Cells(r_1, BranchId).Value, ",")
            Else:
                branch = Array(Sheets("Tabelle1").Cells(r_1, BranchId).Value)
            End If
            lens = UBound(branch)
            For Items = 0 To lens
                If FL_id = CInt(branch(Items)) Then
            'p = WorksheetFunction.Match(Str(FL_id), branch, -1)
                    Sheets(FL_name).Cells(row, 1).Value = Sheets("Tabelle1").Cells(r_1, Kommentar)
                    Sheets(FL_name).Cells(row, 2).Value = Sheets("Tabelle1").Cells(r_1, Beschreibung)
                    Sheets(FL_name).Cells(row, 3).Value = Sheets("Tabelle1").Cells(r_1, AP)
                    date1 = Sheets("Tabelle1").Cells(r_1, SDate).Value
                    date1 = Split(CStr(date1), ".")
                    date2 = date1(2) & "-" & date1(1) & "-" & date1(0)
                    Sheets(FL_name).Cells(row, 5).Value = date2
                    date1 = Sheets("Tabelle1").Cells(r_1, EDate).Value
                    date1 = Split(CStr(date1), ".")
                    date2 = date1(2) & "-" & date1(1) & "-" & date1(0)
                    Sheets(FL_name).Cells(row, 6).Value = date2
                    Sheets(FL_name).Cells(row, 7).Value = "MHD"
                    row = row + 1
                    Exit For
                End If
            Next
        End If
        
    Next
End Sub



