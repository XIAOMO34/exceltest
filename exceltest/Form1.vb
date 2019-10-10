Public Class Form1
    Dim myexcel As Microsoft.Office.Interop.Excel.Application
    Dim myworkbook As Microsoft.Office.Interop.Excel.Workbook
    Dim myworksheet As Microsoft.Office.Interop.Excel.Worksheet
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            myexcel = CType(GetObject(, "Excel.Application"), Microsoft.Office.Interop.Excel.Application)
            ''获得已经打开的EXCEL对象
            Step21()
            Step3()
        Catch ex As Exception
            Step22()
            Step3()
        End Try
    End Sub
    Function Step21()  ''对已经打开的表格操作,此处用函数不用进程来选择性调用！！
        myexcel.WindowState = -4137 ''全屏
        myworkbook = myexcel.ActiveWorkbook
        myworksheet = myworkbook.Worksheets("原始数据1")
    End Function
    Function Step22() ''重新打开EXCEL顶级对象APPLICATION
        myexcel = CreateObject("Excel.Application")
        myexcel.Visible = True
        myexcel.WindowState = -4137 ''全屏
        myworkbook = myexcel.Workbooks.Open("D:\腾讯文件\原始数据\文件1.xlsx")
        myworksheet = myworkbook.Worksheets("原始数据1")
    End Function
    Function Step3() ''进入EXCEL工作簿操作
        ''数学竞赛处理数据
        Dim i As Long
        Dim j As Long
        For i = 7081 To 150000
            If (myexcel.Range("O" & (i + 1)).Value - myexcel.Range("O" & i).Value) <> 1 And
                (myexcel.Range("O" & (i + 1)).Value - myexcel.Range("O" & i).Value) <> -86399 Then
                myexcel.Range("O" & (i + 1)).Select()
                For j = i + 1 To i + 1000
                    If myexcel.Range("B" & j).Value = 0 Then
                        myexcel.Rows((i) & ":" & (j - 1)).Delete(-4162)
                        ''myexcel.Range("B" & j).EntireRow.Delete()
                        ''j = j - 1

                        Exit For
                    End If
                Next
            End If
        Next
        Me.Close()
    End Function
End Class
