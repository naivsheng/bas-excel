Attribute VB_Name = "模块1"
Sub mark()
    Dim arr_FL(), arr_Artikel()
    Dim str1 As String
    Dim FL As Integer
    FL = Sheets("Filialen").UsedRange.Rows.Count
    Dim fil As Integer
    Sheets("Filialen").Select
    arr_FL = Range("A2:A" & FL)
    arr_FL = Excel.WorksheetFunction.Transpose(Excel.WorksheetFunction.Index(arr_FL, 0, 1))
    str1 = Application.ActiveWorkbook.FullName
    Dim dic As Object
    Set dic = CreateObject("scripting.dictionary")
    
    Filename = Split(str1, "KW")(0) & "KW" & Replace(Str(Int(Split(Split(str1, "KW")(1), ".")(0)) - 1), " ", "") & ".xlsx"
    Workbooks.Open Filename:=Filename
        
    ' 遍历每家店获取数据
    For r = 1 To FL - 1
        Sheets(arr_FL(r)).Select
        arr_Artikel = Range("A3:A105")
        arr_Artikel = Excel.WorksheetFunction.Transpose(Excel.WorksheetFunction.Index(arr_Artikel, 0, 1))
        'arr_Artikel = Split(Application.Trim(Join(Application.Transpose(arr_Artikel))))
        dic(arr_FL(r)) = arr_Artikel

    Next
    ActiveWindow.Close ' close the window
    
    For r = 1 To FL - 1
        Sheets(arr_FL(r)).Select
        For rr = 3 To Sheets(arr_FL(r)).UsedRange.Rows.Count
            On Error Resume Next
                flag = Excel.WorksheetFunction.Match(Cells(rr, 1).Value, dic(arr_FL(r)), 0)
                If Err <> 0 Then
                'If flag = 0 Then
                    Range("A" & rr).Select
                    With Selection.Font
                        .Color = -16776961
                        .TintAndShade = 0
                    End With
                End If
        Next
    Next
            
End Sub

Sub ConvertPDF()
strPath = ThisWorkbook.Path & "\"
For Each s In Sheets
If s.Name <> "Sheet1" And s.Name <> "Sheet2" And s.Name <> "Tabelle1" And s.Name <> "Filialen" And s.Name <> "Result" And s.Name <> "Tabelle1" Then 'w为当前excel表的名称
s.ExportAsFixedFormat Type:=xlTypePDF, Filename:=strPath & s.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
End If
Next
End Sub

