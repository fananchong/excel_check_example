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
    msg = Check()
    If msg = "" Then
        MsgBox "���ͨ��"
    Else
        MsgBox msg
    End If
End Sub

Rem ���
Function Check()
    For r = 2 To Worksheets("check").UsedRange.Rows.Count
        For c = 1 To Worksheets("check").UsedRange.Columns.Count
            v = Worksheets("check").Cells(r, c).Value
            If Not v = "" Then
                Check = "Row: " + Str(r) + " Col: " + Str(c) + Chr(13) + Chr(10) + "��鲻ͨ����ֵΪ��" + v
                Exit Function
            End If
        Next
    Next
    Check = ""
End Function


Rem ���������������
Sub CheckAll_Click()
    f = Dir(ThisWorkbook.Path + "\*.xlsm")
    Do While f <> ""
        If f = "" Then
            Exit Do
        End If
        
        Rem �Լ�
        Dim msg As String
        If Application.ActiveWorkbook.Name = f Then
            msg = Check()
        Else
            
            Dim xlApp
            Dim xlBook

            Set xlApp = CreateObject("Excel.Application")
            Set xlBook = xlApp.Workbooks.Open(ThisWorkbook.Path & "\" & f, 3)
            xlApp.Application.Visible = False
            xlApp.DisplayAlerts = False
            xlBook.SaveLinkValues = True
            msg = xlApp.Application.Run("Check")
            xlBook.Save
            xlBook.Close
            xlApp.Quit

            Set xlBook = Nothing
            Set xlApp = Nothing
            
        End If
        
        If msg <> "" Then
            MsgBox "�ļ� " + f + " ��鲻ͨ��" + Chr(13) + Chr(10) + msg
            Exit Sub
        End If
        f = Dir
    Loop
    MsgBox "���ͨ��"
End Sub
