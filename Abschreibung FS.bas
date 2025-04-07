Attribute VB_Name = "Abschreibung"
Sub Result()
'
' Makro1 Makro
'

'
    Sheets("Tabelle1").Select
    Range("A1").Select
    Selection.AutoFilter
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "=CONCAT(RC[-10],"" "",RC[-9])"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],FL!C[-11]:C[-10],2,0)"
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(SEARCH(""FS"",RC[-5]),"""")"
    All_Rows = Sheets("Tabelle1").UsedRange.Rows.Count
    Range("K2:M2").Select
    Selection.AutoFill Destination:=Range("K2:M" & All_Rows)
    Range("L2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add.Name = "FL_List"
    Sheets("FL_List").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    'Rows = Sheets("FL_List").UsedRange.Rows.Count
    ActiveSheet.Range("$A:$A").RemoveDuplicates Columns:=1, Header:=xlNo
    FL = Sheets("FL_List").UsedRange.Rows.Count
    
    Sheets("FL_List").Select
    FL_update = Range("a1:a" & FL).Value
    FL = Sheets("FL").UsedRange.Rows.Count
    Sheets.Add.Name = "Bemerkung"
    
    For r = 1 To FL
        Sheets("Tabelle2").Select
        row_sheet = Sheets("Tabelle2").UsedRange.Rows.Count
        If row_sheet >= 3 Then
            Rows("3:" & row_sheet).Select
            Selection.Delete Shift:=xlUp
        End If
        FL_name = Sheets("FL_List").Cells(r, 1).Value
        Range("A1").Select
        ActiveCell.FormulaR1C1 = "Filiale: " & FL_name
        Item_Match = Application.Match(FL_name, FL_update, False)
        If Not IsEmpty(FL_name) Then
            If Item_Match = False Then ' the FL has no update
                Range("A2:D57").Select
            Else
                Sheets("Tabelle1").Select
                ActiveSheet.Range("$A$1:$L$" & All_Rows).AutoFilter Field:=8
                ActiveSheet.Range("$A$1:$L$" & All_Rows).AutoFilter Field:=12, Criteria1:=FL_name
                Range("H2:H" & Sheets("Tabelle1").UsedRange.Rows.Count).Select
                'Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
                Sheets("Bemerkung").Select
                Range("A1").Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                Application.CutCopyMode = False
                ActiveSheet.Range("$A:$A").RemoveDuplicates Columns:=1, Header:=xlNo
                Sheets("Bemerkung").Select
                B_rows = Sheets("Bemerkung").UsedRange.Rows.Count + 1
                Dim Berm() As Variant
                ReDim Berm(1)
                Berm(1) = "test"
                For fs = 1 To B_rows
                    If InStr(1, UCase(Range("A" & fs).Value), "FS") Then
                        If Berm(1) = "test" Then
                            Berm(1) = Range("A" & fs).Value
                        Else
                            ReDim Preserve Berm(UBound(Berm) + 1)
                            Berm(LBound(Berm)) = Range("A" & fs).Value
                        End If
                    End If
                Next fs
                If Berm(1) = "test" Then
                    Sheets("Tabelle2").Select
                    Range("A2:D57").Select ' not match
                Else
                    Sheets("Tabelle1").Select
                    ActiveSheet.Range("$A$1:$L$" & All_Rows).AutoFilter Field:=8, Criteria1:=Berm, Operator:=xlFilterValues
                    Range("C2:E" & All_Rows).Select
                    Selection.Copy
                    Sheets("Tabelle2").Select
                    Range("A3").Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                        :=False, Transpose:=False
                    Range("A1:B1").Select
                    Application.CutCopyMode = False
                    
                    row_sheet = Sheets("Tabelle2").UsedRange.Rows.Count
                    Range("A2:D" & row_sheet).Select
                End If
            End If
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            ' Set Layout
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
            ' save as .pdf file
            Set ws = ActiveSheet
            MyTime = Range("D1").Value
            Set mainWorkbook = ThisWorkbook
            savePath = mainWorkbook.Path & "\" & "FS " & FL_name & " KW" & MyTime & ".pdf"
            ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=savePath, _
                               Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                               IgnorePrintAreas:=False, OpenAfterPublish:=False
            'ws.Range("A1:D" & lastRow).ExportAsFixedFormat Type:=xlTypePDF, Filename:=mainWorkbook.Path & "\" & FL_name & ".pdf"
            End If
    Next r
    
    Sheets("Tabelle1").Select
    Selection.AutoFilter
    Rows("2:" & All_Rows).Select
    Selection.Delete Shift:=xlUp
    Sheets("FL_List").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("Bemerkung").Select
    ActiveWindow.SelectedSheets.Delete
    
End Sub





