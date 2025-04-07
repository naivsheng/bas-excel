Attribute VB_Name = "COHE"
Sub Makro2()
Attribute Makro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Makro2 Makro
'

'
    Cells.Select
    Range("A1:AD" & ActiveSheet.Rows.Count).Activate
    Selection.UnMerge
    For i = 1 To Rows
        If InStr(cell(i, 2).Value, "Art.") Then
            Exit For
    Next i
    UsedRange = "B" & Str(i) & "AF"
    For i = 100 To Rows
        If InStr(cell(i, 3).Value, "BTW") Then
            Exit For
    Next i
    UsedRange = UsedRange & Str(i - 1)
    Range(UsedRange).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$AE$" & ActiveSheet.Rows.Count).AutoFilter Field:=1, Criteria1:="<>"
    Columns("B:O").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:E").Select
    Selection.Delete Shift:=xlToLeft
    Range("D1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-2]*RC[-1]"
    Range("D1").Select
    Selection.FillDown
    Columns("E:K").Select
    Selection.Delete Shift:=xlToLeft
    Range("$A$1:$E$" & ActiveSheet.Rows.Count).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:C").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Ja"
    Range("C1").Select
    Selection.AutoFill Destination:=Range("C1:C" & ActiveSheet.Rows.Count)

End Sub
