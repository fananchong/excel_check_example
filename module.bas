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
    msg = Check()
    If msg = "" Then
        MsgBox "检查通过"
    Else
        MsgBox msg
    End If
End Sub

Rem 检查
Function Check()
    For r = 2 To Worksheets("check").UsedRange.Rows.Count
        For c = 1 To Worksheets("check").UsedRange.Columns.Count
            v = Worksheets("check").Cells(r, c).Value
            If Not v = "" Then
                Check = "Row: " + Str(r) + " Col: " + Str(c) + Chr(13) + Chr(10) + "检查不通过，值为：" + v
                Exit Function
            End If
        Next
    Next
    Check = ""
End Function


Rem 单击检查所有配置
Sub CheckAll_Click()

    Dim errnum As Integer
    f = Dir(ThisWorkbook.Path + "\*.xlsm")
    Do While f <> ""
        If f = "" Then
            Exit Do
        End If
         
        If checkOneFile(f) = False Then
            errnum = errnum + 1
        End If
       
        f = Dir
    Loop
    MsgBox "错误文件数：" + Str(errnum)
End Sub



Function checkOneFile(f As Variant)

    Dim xlApp
    Dim xlBook
    Dim msg As String
    If Application.ActiveWorkbook.Name = f Then
        msg = Check()
    Else

        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Open(ThisWorkbook.Path & "\" & f, 3)
        xlApp.Application.Visible = False
        xlApp.DisplayAlerts = False
        xlBook.SaveLinkValues = True
        
        On Error GoTo MyErr
        msg = xlApp.Application.Run("Check")
        On Error GoTo 0
        
        xlBook.Save
        xlBook.Close
        xlApp.Quit

        Set xlBook = Nothing
        Set xlApp = Nothing
            
    End If
    
    checkOneFile = True
    If msg <> "" Then
        MsgBox "文件 " + f + " 检查不通过" + Chr(13) + Chr(10) + msg
        checkOneFile = False
        Exit Function
    End If
    Exit Function
    
MyErr:

    msg = "文件 " + f + " 错误 " & Err.Number & " ： " & Err.Description
    MsgBox msg
    
    xlBook.Close
    xlApp.Quit
    Set xlBook = Nothing
    Set xlApp = Nothing
    
    checkOneFile = False
    Exit Function
    
End Function





