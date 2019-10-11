Public Class Form1
    Dim myexcel As Microsoft.Office.Interop.Excel.Application
    Dim myworkbook As Microsoft.Office.Interop.Excel.Workbook
    Dim myworksheet As Microsoft.Office.Interop.Excel.Worksheet
    Dim myword As Microsoft.Office.Interop.Word.Application
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            myexcel = CType(GetObject(, "Excel.Application"), Microsoft.Office.Interop.Excel.Application)
            ''获得已经打开的EXCEL对象
            Step21()
            Wmassout()
            'Step3()
            Me.Close()
        Catch ex As Exception
            Step22()
        End Try
    End Sub
    Function Step21()  ''对已经打开的表格操作,此处用函数不用进程来选择性调用！！
        myexcel.WindowState = -4137 ''全屏
        myworkbook = myexcel.ActiveWorkbook
        myworksheet = myworkbook.Worksheets("Sheet1")
    End Function
    Function Step22() ''重新打开EXCEL顶级对象APPLICATION
        myexcel = CreateObject("Excel.Application")
        myexcel.Visible = True
        myexcel.WindowState = -4137 ''全屏
        Wmassout()
    End Function
    Function Wmassout() ''读取Wmass文件中信息
        'myworkbook = myexcel.Workbooks.Open("C:\Users\LJX\Desktop\新建 Microsoft Excel 工作表.xlsx")
        'myworksheet = myworkbook.Worksheets("Sheet1")
        Dim REG As Microsoft.Office.Interop.Excel.Range
        REG = myworksheet.Range("A1:A500")
        For Each i In REG
            If i.VALUE Like "* 恒载质量    活载质量*" Then ''关键词和通配符
                'MessageBox.Show(i.ROW + 2)
                myworksheet.Range("A" & （i.ROW + 2) & ": A" & (i.ROW + 3)).Copy()
                myworksheet.Range("B1").Select()
                myworkbook.ActiveSheet.PASTE
            End If
        Next
        myword = CreateObject("Word.application")

    End Function
End Class
