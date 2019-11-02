Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Word
Public Class Form1
    Dim myexcel As Microsoft.Office.Interop.Excel.Application ''excel顶级对象
    Dim myworkbook As Microsoft.Office.Interop.Excel.Workbook ''数据文件
    Dim myworkbook2 As Microsoft.Office.Interop.Excel.Workbook ''表格文件
    Dim myworksheet As Microsoft.Office.Interop.Excel.Worksheet ''数据文件工作簿
    Dim myworksheet2 As Microsoft.Office.Interop.Excel.Worksheet ''表格文件工作簿
    Dim myword As Microsoft.Office.Interop.Word.Application ''word顶级对象
    Dim myworddoc As Microsoft.Office.Interop.Word.Document ''ETABS非减震
    Dim myworddoc2 As Microsoft.Office.Interop.Word.Document ''ETABS减震
    Dim reg As Microsoft.Office.Interop.Excel.Range ''excel的range对象
    Dim c As Integer ''楼层数
    Dim cc As Integer ''循环用
    Dim ccc As Integer ''循环用
    Dim x2 As Integer ''楼层角所在行数
    Dim x1 As Integer ''楼层力所在行数
    Dim y As Integer ''列数
    Dim P As Process（） ''excel进程
    Dim q As Process（） ''word进程
    Dim er As Integer ''错误处理
    ''窗口移动
    Private Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long) ''暂停程序，单位：毫秒
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

    'Private Sub BunifuFlatButton2_Click(sender As Object, e As EventArgs)
    '    'Try
    '    '    myexcel = CType(GetObject(, "Excel.Application"), Microsoft.Office.Interop.Excel.Application)
    '    '    ''获得已经打开的EXCEL对象
    '    '    Step21()
    '    '    Wmassout()
    '    '    'Duquword()
    '    '    'Step3()
    '    '    Me.Close()
    '    'Catch ex As Exception
    '    Try
    '        myexcel = CType(GetObject(, "Excel.Application"), Microsoft.Office.Interop.Excel.Application)
    '        MessageBox.Show("EXCEL已经在运行!") ''提示已经打开EXCEL时应关闭进程
    '        Exit Sub
    '    Catch ex As Exception
    '        If OpenFileDialog1.FileName = "OpenFileDialog1" Then
    '            MessageBox.Show("请先选择文件！")
    '        Else
    '            Openexcel()
    '        End If
    '    End Try
    'End Sub
    Private Sub BunifuFlatButton5_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton5.Click
        myexcel = CreateObject("Excel.application")
        myexcel.Visible = True
        myword = CreateObject("Word.application")
        myword.Visible = True ''常规窗口
        er = 0
        myworkbook = myexcel.Workbooks.Open("C:\Users\LJX\Desktop\报告程序\数据文件.xlsx")
        myworkbook2 = myexcel.Workbooks.Open("C:\Users\LJX\Desktop\报告程序\表格文件.xlsx")
        'myworkbook = myexcel.Workbooks.Open(OpenFileDialog2.FileName)''''''''''''''''''''''''''
        'myworkbook2 = myexcel.Workbooks.Open(OpenFileDialog3.FileName)''''''''''''''''''''''''''''
        If er <> 1 Then
            Openetabs("C:\Users\LJX\Desktop\报告程序\ETABS文件.Docx")
            'Openetabs(OpenFileDialog1.FileName)''''''''''''''''''''''''''''''''''''''
            Opensheet()
            Disizhang()
            Openetabs("C:\Users\LJX\Desktop\报告程序\ETABS文件.Docx")
            'Openetabs(OpenFileDialog4.FileName)'''''''''''''''''''''''''''''''''''''
            'Opensheet()
            Diwuzhang()
        End If
    End Sub
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

    Function Openetabs(ByVal a As String) ''操作ETABS文档到WORD中并且处理数据
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''处理层间力程序
        myworddoc = myword.Documents.Open(a) ''aword路径,b是excel工作簿
        myword.ActiveWindow.WindowState = 2 ''最小化窗口
        'myworddoc.Tables(1).Select()
        myword.Selection.WholeStory() ''选中所有内容
        myword.Selection.Copy()
        myworksheet = myworkbook.Worksheets("ETABS") ''copy到ETABS的sheet中
        myworksheet.Activate()
        myexcel.WindowState = -4137
        myworksheet.Cells.Select()
        myexcel.Selection.ClearContents
        myworksheet.Activate()
        myworksheet.Cells(1, 1).Select()
        Sleep(1000)
        myexcel.ActiveSheet.PASTESPECIAL("HTML",,,,,,) ''粘贴
        myworddoc.Close()
        reg = myworksheet.Range("A:A")
        For Each i In reg
            If i.value Like "*楼层力" Then ''找到楼层力所在区域
                x1 = i.ROW
                y = i.COLUMN
                c = 2
                'c = TextBox4.Text'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If myworksheet.Range("A" & x1 + 1&).Value Like "*Story Forces" Then
                    'MessageBox.Show(1)                   
                    Exit For
                End If
            End If
        Next
        reg = myexcel.Range("B" & x1 + 4 & ":B" & x1 + 20000)
        For Each I In reg ''处理楼层力数据
            If I.VALUE Like "X*" Then
                x1 = I.ROW
                myworksheet.Range("J" & x1).Select()
                myexcel.ActiveCell.FormulaR1C1 = "=MAX(RC[-5],ABS(R[" & c & "]C[-5])，ABS(RC[-4]),ABS(R[" & c & "]C[-4]))"
                ''字符串连接应当加空格
                myexcel.Selection.AutoFill(myworksheet.Range("J" & x1 & ":j" & x1 + c - 1), 0)
                myexcel.Range("J" & x1 + c & ":J" & x1 + 2 * c - 1).Value = 0
                myworksheet.Range("J" & x1 & ":J" & x1 + 2 * c - 1).Select()
                myexcel.Selection.AutoFill(myworksheet.Range("J" & x1 & ":J" & x1 + 32 * c - 1), 0) ''填充包含原区域
                Dim x3 As Integer
                Dim C1 As Integer = 11 ''c1表示层间力表格索引
                x3 = x1
                For cc = 1 To 16
                    myexcel.Range("J" & x3 & ":J" & x3 + c - 1).Copy()
                    myexcel.Cells(x1, C1).Select()
                    myexcel.Selection.PASTESPECIAL(-4163,,,)
                    x3 = x3 + 2 * c
                    C1 = C1 + 1
                Next
                Exit For
            End If
        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''处理位移角程序
        reg = myworksheet.Range("A:A")
        For Each i In reg
            If i.value Like "*层间位移角" Then ''找到楼层角所在区域
                x2 = i.ROW
                y = i.COLUMN
                If myworksheet.Range("A" & x2 + 1&).Value Like "*Story Drifts" Then
                    'MessageBox.Show(1)                   
                    Exit For
                End If
            End If
        Next
        reg = myexcel.Range("B" & x2 + 3 & ":B" & x2 + 20000)
        For Each I In reg ''处理楼层角数据
            If I.VALUE Like "X*" Then
                x2 = I.ROW ''此时X为第一个地震波数据行数
                Exit For
            End If
        Next
        For i = x2 To x2 + 100 * c ''不处理数据，直接根据单元格内容导出位移角数据
            If myworksheet.Range("B" & i).Value Like "*Max" Then
                If myworksheet.Range("B" & i).Value Like myworksheet.Range("C" & i).Value & "*" Then
                    myworksheet.Range("E" & i).Copy()
                    myworksheet.Cells(i, 7).Select()
                    myexcel.Selection.PASTESPECIAL(-4163,,,)
                End If
            End If
        Next
        cc = 0 ''行数
        ccc = 0 ''列数
        reg = myworksheet.Range("G" & x2 & ":G" & x2 + 64 * c)
        For Each I In reg
            If I.VALUE IsNot Nothing Then
                I.COPY
                myworksheet.Cells(x2 + cc, 8 + ccc).select
                myexcel.Selection.PASTESPECIAL(-4163,,,)
                cc = cc + 1
            End If
            If cc = c Then
                ccc = ccc + 1
                cc = 0
            End If
        Next
    End Function
    Function Opensheet()
        c = 2
        'c = TextBox4.Text''''''''''''''''''
        myworksheet2 = myworkbook2.Worksheets("Sheet1")
        myworksheet2.Activate()
        cc = 1
        For cc = 1 To c - 1
            myworksheet2.Range("F3:T3").Select() ''根据层数扩充表格4.2~4.5
            myexcel.Selection.EntireRow.Insert(0)
        Next
        cc = 0
        For cc = 0 To 48 Step 16
            myworksheet2.Range("F2:T" & (2 + c)).Copy()
            myworksheet2.Cells(c + 4, 6 + cc).select()
            myexcel.ActiveSheet.paste()
        Next
    End Function
    Function Disizhang()
        cc = 0
        ccc = 0
        For cc = 0 To 8 Step 8
            myworksheet.Range(Nts(11 + cc) & x1 & ":" & Nts(11 + cc + 7) & x1 + c - 1).Copy() ''通用复制到表格文件的代码
            myworksheet2.Cells(3, 7 + ccc).Select()
            myexcel.Selection.PASTESPECIAL(-4163,,,)
            ccc = 16
        Next
        cc = 0
        ccc = 32
        For cc = 0 To 8 Step 8
            myworksheet.Range(Nts(8 + cc) & x2 & ":" & Nts(8 + cc + 7) & x2 + c - 1).Copy() ''通用复制到表格文件的代码
            myworksheet2.Cells(3, 7 + ccc).Select()
            myexcel.Selection.PASTESPECIAL(-4163,,,)
            ccc = 48
        Next
    End Function
    Function Diwuzhang()
        cc = 0
        ccc = 0
        For cc = 0 To 8 Step 8
            myworksheet.Range(Nts(11 + cc) & x1 & ":" & Nts(11 + cc + 7) & x1 + c - 1).Copy() ''通用复制到表格文件的代码
            myworksheet2.Activate()
            myworksheet2.Cells(5 + c, 7 + ccc).Select()
            myexcel.Selection.PASTESPECIAL(-4163,,,)
            ccc = 16
        Next
        cc = 0
        ccc = 32
        For cc = 0 To 8 Step 8
            myworksheet.Range(Nts(8 + cc) & x2 & ":" & Nts(8 + cc + 7) & x2 + c - 1).Copy() ''通用复制到表格文件的代码
            myworksheet2.Cells(5 + c, 7 + ccc).Select()
            myexcel.Selection.PASTESPECIAL(-4163,,,)
            ccc = 48
        Next
        myworkbook.Close()
    End Function

    Private Sub BunifuFlatButton6_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton6.Click
        OpenFileDialog2.Filter = "所有文件|*.*" ''文件筛选器
        OpenFileDialog2.ShowDialog()
        If OpenFileDialog2.FileName <> "OpenFileDialog2" Then
            TextBox2.Text = "文件已选择：" & OpenFileDialog2.FileName '1
        End If
    End Sub
    Private Sub BunifuFlatButton7_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton7.Click
        OpenFileDialog3.Filter = "所有文件|*.*" ''文件筛选器
        OpenFileDialog3.ShowDialog()
        If OpenFileDialog3.FileName <> "OpenFileDialog3" Then
            TextBox3.Text = "文件已选择：" & OpenFileDialog3.FileName
        End If
    End Sub
    Private Sub BunifuFlatButton4_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton4.Click
        OpenFileDialog1.Filter = "所有文件|*.*" ''文件筛选器
        OpenFileDialog1.ShowDialog()
        If OpenFileDialog1.FileName <> "OpenFileDialog1" Then
            TextBox1.Text = "文件已选择：" & OpenFileDialog1.FileName
        End If
    End Sub


    Private Sub BunifuFlatButton1_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton1.Click
        Panel3.Show()
        Panel4.Hide()
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub BunifuFlatButton2_Click_1(sender As Object, e As EventArgs) Handles BunifuFlatButton2.Click
        OpenFileDialog4.Filter = "所有文件|*.*" ''文件筛选器
        OpenFileDialog4.ShowDialog()
        If OpenFileDialog4.FileName <> "OpenFileDialog4" Then
            TextBox5.Text = "文件已选择：" & OpenFileDialog4.FileName
        End If
    End Sub
    Private Sub BunifuFlatButton3_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton3.Click
        ''关闭excel进程，释放内存
        P = Process.GetProcessesByName("EXCEL")
        For Each pr As Process In P
            pr.Kill()
        Next
        q = Process.GetProcessesByName("WINWORD")
        For Each i As Process In q
            i.Kill()
        Next
    End Sub
    Function Nts(n) ''转换数字为字母
        Dim s As String, i As Integer
        'For i = 1 To Len(CStr(n))
        i = 1
        If n < 27 Then
            s = s & Chr(Val(Mid(n, i, 2)) + Asc("a") - 1)
        Else
            n = n - 26
            s = "a" & Chr(Val(Mid(n, i, 2)) + Asc("a") - 1)
        End If
        Nts = s
    End Function
    Private Sub BunifuImageButton1_Click(sender As Object, e As EventArgs) Handles BunifuImageButton1.Click
        Me.Close()
    End Sub



End Class
