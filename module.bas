Attribute VB_Name = "模块1"
Rem 检查是否有此 ID
Function CheckID(id As Variant, range As Variant)
    If id.Text = "" Or Application.CountIf(range, id) > 0 Then
        CheckID = ""
    Else
        CheckID = "[ERROR] " + Str(id)
    End If
End Function


Rem 检查是否 ID 列表都有效
Function CheckIDs(ids As Variant, delimiter As String, range As Variant)
    If ids.Text = "" Then
        CheckIDs = ""
    Else
        Dim lst() As String
        lst = Split(ids.Text, delimiter)
        For Each id In lst
            If Application.CountIf(range, id) = 0 Then
                CheckIDs = "[ERROR] " + id
                Exit Function
            End If
        Next id
        CheckIDs = ""
    End If
End Function

Rem 单击检查
Sub Check_Click()
    For r = 2 To Worksheets("check").UsedRange.Rows.Count
        For c = 1 To Worksheets("check").UsedRange.Columns.Count
            v = Worksheets("check").Cells(r, c).Value
            If Not v = "" Then
                MsgBox "Row: " + Str(r) + " Col: " + Str(c) + Chr(13) + Chr(10) + "检查不通过，值为：" + v
                Exit Sub
            End If
        Next
    Next
    MsgBox "检查通过"
End Sub

