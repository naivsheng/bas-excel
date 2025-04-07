Attribute VB_Name = "Format_WhsCode"
Sub set_Format()
'   Format the Value of WhsCode
    Firma = InputBox("GA or OM for this time?")
    Range("BH3").Select
    ActiveCell.FormulaR1C1 = "=TEXT(RC[-37],""00"")"
    row_stop = Sheets("Tabelle1").UsedRange.Rows.Count
    Selection.AutoFill Destination:=Range("BH3:BH" & row_stop)
    Range("BH3:BH" & row_stop).Select
    Selection.Copy
    Range("W3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("BH:BH").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Range("BG3").Select
    Selection.AutoFill Destination:=Range("BG3:BG" & row_stop)
    Set mainWorkbook = ThisWorkbook
    zeit = Format(Now(), "yyyymmdd")
    ActiveWorkbook.SaveAs Filename:= _
        mainWorkbook.Path & "\ItemWarehouseInfo1_" & zeit & " " & Firma & ".txt" _
        , FileFormat:=xlText, CreateBackup:=False
End Sub


