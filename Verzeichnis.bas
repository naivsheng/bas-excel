'
' @Author: naivsheng naivsheng@outlook.com
' @Date: 1979-12-31 23:00:00
' @LastEditors: naivsheng naivsheng@outlook.com
' @LastEditTime: 2025-04-07 13:41:43
' @FilePath: \vba\Verzeichnis.bas
' @Description: 这是默认设置,请设置`customMade`, 打开koroFileHeader查看配置 进行设置: https://github.com/OBKoro1/koro1FileHeader/wiki/%E9%85%8D%E7%BD%AE
'
Attribute VB_Name = "Verzeichnis"
' 添加工作表目录
Sub CreateDirectory()
    Dim ws As Worksheet
    Dim directorySheet As Worksheet
    Dim i As Integer
    ' 检查是否已经存在名为 "Verzeichnis" 的工作表，如果存在则删除
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("Verzeichnis").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    ' 添加一个新的工作表用于目录
    Set directorySheet = Worksheets.Add
    directorySheet.Name = "Verzeichnis"

    ' 在目录工作表中添加标题
    directorySheet.Cells(1, 1).Value = "Sheet Name"
    directorySheet.Cells(1, 2).Value = "Sheet Number"

    ' 循环遍历所有工作表，并将其名称添加到目录工作表
    i = 2
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Verzeichnis" Then
            directorySheet.Cells(i, 1).Value = ws.Name
            directorySheet.Cells(i, 2).Value = i - 1
            directorySheet.Hyperlinks.Add Anchor:=directorySheet.Cells(i, 1), Address:="", SubAddress:="'" & ws.Name & "'!A1", TextToDisplay:=ws.Name
            i = i + 1
            ws.Hyperlinks.Add Anchor:=ws.Range("M1"), Address:="", SubAddress:= _
                "Verzeichnis!A1", TextToDisplay:="GoBack"
        End If
    Next ws
End Sub
