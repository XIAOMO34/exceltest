Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Word
Public Class Form1
    Dim myexcel As Microsoft.Office.Interop.Excel.Application
    Dim myworkbook As Microsoft.Office.Interop.Excel.Workbook
    Dim myworksheet As Microsoft.Office.Interop.Excel.Worksheet
    Dim myword As Microsoft.Office.Interop.Word.Application
    Dim myworddoc As Microsoft.Office.Interop.Word.Document
    Dim REG As Microsoft.Office.Interop.Excel.Range
    Dim C As Integer ''楼层数
    ''窗口移动
    ''Private Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)暂停程序，单位：毫秒
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As IntPtr,
         ByVal wMsg As Integer,
         ByVal wParam As Integer,
         ByVal lParam As Integer) As Boolean
    Public Declare Function ReleaseCapture Lib "user32" Alias "ReleaseCapture" () As Boolean
    Public Const WM_SYSCOMMAND = &H112
    Public Const SC_MOVE = &HF010&
    Public Const HTCAPTION = 2
    Private Sub Panel1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) _
        Handles Panel1.MouseDown
        ReleaseCapture()
        SendMessage(Me.Handle, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0)
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Private Sub BunifuFlatButton2_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton2.Click
        'Try
        '    myexcel = CType(GetObject(, "Excel.Application"), Microsoft.Office.Interop.Excel.Application)
        '    ''获得已经打开的EXCEL对象
        '    Step21()
        '    Wmassout()
        '    'Duquword()
        '    'Step3()
        '    Me.Close()
        'Catch ex As Exception
        Try
            myexcel = CType(GetObject(, "Excel.Application"), Microsoft.Office.Interop.Excel.Application)
            MessageBox.Show("EXCEL已经在运行!") ''提示已经打开EXCEL时应关闭进程
            Exit Sub
        Catch ex As Exception
            If OpenFileDialog1.FileName = "OpenFileDialog1" Then
                MessageBox.Show("请先选择文件！")
            Else
                Openexcel()
                Wmassout()
                'Wzq()
            End If
        End Try
    End Sub
    'Function Step21()  ''对已经打开的表格操作,此处用函数不用进程来选择性调用！！
    '    myexcel.WindowState = -4137 ''全屏
    '    myworkbook = myexcel.ActiveWorkbook
    '    myworksheet = myworkbook.Worksheets("myworkbook.Sheet1")
    'End Function
    Function Openexcel() ''重新打开EXCEL顶级对象APPLICATION
        If OpenFileDialog1.FileName = "OpenFileDialog1" Then
            MessageBox.Show("请先选择文件！")
        End If
        myexcel = CreateObject("Excel.Application")
        ''myexcel.ReferenceStyle = -4150 ''启用R1C1引用模式
        myexcel.Visible = True
        myexcel.WindowState = -4137 ''全屏
        myworkbook = myexcel.Workbooks.Open(OpenFileDialog1.FileName)

    End Function
    Function Wmassout() ''读取Wmass文件中信息
        ''myworkbook = myexcel.Workbooks.Open("C:\Users\LJX\Desktop\新建 Microsoft Excel 工作表.xlsx")
        myworksheet = myworkbook.Worksheets("Sheet1")
        myexcel.ReferenceStyle = -4150
        REG = myworksheet.UsedRange ''EXCEL所用区域（类似于有数据的区域）
        For Each i In REG ''REG为所有单元格，遍历单元格
            If i.VALUE Like "* 恒载质量    活载质量*" Then ''当出现* 恒载质量    活载质量*时
                ''停止遍历,*为通配符
                Dim wc As Object = i.column
                myworksheet.Range(Nts(wc) & (i.row + 2) & ":" & Nts(wc) & (i.row + 3)).Copy()
                myworksheet.Range("A1000").Select()
                myworkbook.ActiveSheet.PASTE ''复制所需的数据并进行处理
                myexcel.Selection.texttocolumns(,,,,,,, True,,,,,,)
                myworksheet.Range("K1000").Select()
                myexcel.ActiveCell.FormulaR1C1 = "=RC[-5]+RC[-4]"
                myworksheet.Range("K1000").Select()
                Exit For ''跳出循环
            End If
        Next
        C = 0
        REG = myworksheet.Range("F1000:F2000")
        For Each j In REG ''C为楼层数，对任意层数具有通用性
            If j.VALUE IsNot Nothing Then
                C = C + 1
            End If
        Next
        myexcel.Selection.AutoFill(myworksheet.Range("K1000:K" & (999 + C)))
    End Function

    Private Sub BunifuImageButton1_Click(sender As Object, e As EventArgs) Handles BunifuImageButton1.Click
        Me.Close()
    End Sub
    Private Sub BunifuFlatButton3_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton3.Click
        ''关闭excel进程，释放内存
        Dim p As Process() = Process.GetProcessesByName("EXCEL")
        For Each pr As Process In p
            pr.Kill()
        Next
        Dim q As Process() = Process.GetProcessesByName("WINWORD")
        For Each i As Process In q
            i.Kill()
        Next
    End Sub

    Private Sub BunifuFlatButton4_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton4.Click
        OpenFileDialog1.Filter = "所有文件|*.*" ''文件筛选器
        OpenFileDialog1.ShowDialog()
        If OpenFileDialog1.FileName <> "OpenFileDialog1" Then
            TextBox1.Text = "文件已选择：" & OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Panel3_Paint(sender As Object, e As PaintEventArgs) Handles Panel3.Paint

    End Sub
    Function Wzq()
        REG = myworksheet.Range("B1:B500")
        For Each i In REG ''REG为所有单元格，遍历单元格
            If i.VALUE Like "*周期      转角*" Then
                ''停止遍历,*为通配符
                myworksheet.Range("B" & （i.ROW + 2) & ": B" & (i.ROW + 3)).Copy()
                myworksheet.Range("C1").Select()
                myworkbook.ActiveSheet.PASTE ''复制所需的数据并进行处理
                myexcel.Selection.texttocolumns(,,,,,,, True,,,,,,)
                myworksheet.Range("M1").Select()
                myexcel.ActiveCell.FormulaR1C1 = "=RC[-5]+RC[-4]"
                myworksheet.Range("M1").Select()
                Exit For ''跳出循环
            End If
        Next
        Dim reg2 As Microsoft.Office.Interop.Excel.Range
        reg2 = myworksheet.Range("B1:B500")
        For Each i In REG
            If i.value Like "*Fx           Vx (分塔剪重比)*" Then
                myworksheet.Range("B" & （i.ROW + 2) & ": B" & (i.ROW + 6)).Copy()
                myworksheet.Range("D1").Select()
                myworkbook.ActiveSheet.PASTE
                myexcel.Selection.texttocolumns(,,,,,,, True,,,,,,)
            End If
        Next

    End Function
    Function Nts(n) ''转换数字为字母
        Dim s As String, i As Integer
        For i = 1 To Len(CStr(n))
            s = s & Chr(Val(Mid(n, i, 1)) + Asc("a") - 1)
        Next
        Nts = s
    End Function

End Class