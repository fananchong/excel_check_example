Attribute VB_Name = "ģ��1"
Rem ����Ƿ��д� ID
Function CheckID(id As Variant, range As Variant)
    If id.Text = "" Or Application.CountIf(range, id) > 0 Then
        CheckID = ""
    Else
        CheckID = "[ERROR] " + Str(id)
    End If
End Function


Rem ����Ƿ� ID �б���Ч
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

Rem �������
Sub Check_Click()
    For r = 2 To Worksheets("check").UsedRange.Rows.Count
        For c = 1 To Worksheets("check").UsedRange.Columns.Count
            v = Worksheets("check").Cells(r, c).Value
            If Not v = "" Then
                MsgBox "Row: " + Str(r) + " Col: " + Str(c) + Chr(13) + Chr(10) + "��鲻ͨ����ֵΪ��" + v
                Exit Sub
            End If
        Next
    Next
    MsgBox "���ͨ��"
End Sub

