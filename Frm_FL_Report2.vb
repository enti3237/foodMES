Imports Excel = Microsoft.Office.Interop.Excel

Public Class Frm_FL_Report2
    Dim xlapp As Excel.Application
    Dim xlbook As Excel.Workbook
    Dim xlsheet As Excel.Worksheet
    Dim Excel_Check As Boolean
    Private Sub Frm_FL_Report2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()

        S_Text_Setup()
        Grid_Main_Setup()

        Panel_Main.Top = 136
        Panel_Main.Left = 12
        Panel_Excel.Top = 136
        Panel_Excel.Left = 12

        Panel_Main.Visible = True
        Panel_Excel.Visible = False

        Com_Excel_Print.Visible = False
        Search_Com.PerformClick()
    End Sub
    Public Function S_Text_Setup() As Boolean
        Dim DBT As New DataTable
        Dim i As Integer

        ComboBox1.Items.Clear()
        ComboBox1.Text = "전체"
        StrSQL = "Select * FROM SI_MACHINE with(NOLOCK) Order By PD_Code"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            For i = 0 To Re_Count - 1
                ComboBox1.Items.Add(DBT.Rows(i)("PD_Code"))
            Next i
            '   ComboBox1.Text = DBT.Rows(0)("PL_Code")

        Else
        End If

        ComboBox2.Items.Clear()
        ComboBox2.Text = "전체"
        StrSQL = "Select * FROM SI_MACHINE with(NOLOCK) Order By PD_Code"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            For i = 0 To Re_Count - 1
                ComboBox2.Items.Add(DBT.Rows(i)("PD_Name"))
            Next i
            '    ComboBox2.Text = DBT.Rows(0)("PL_Name")
        Else
        End If

        S_Text_Setup = True

    End Function

    Public Function Grid_Main_Setup() As Boolean
        Grid_Color(Grid_Main)

        Grid_Main.AllowUserToAddRows = False
        Grid_Main.RowTemplate.Height = 24
        Grid_Main.ColumnCount = 9
        Grid_Main.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Main.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Main.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Main.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.RowHeadersWidth = 10


        Grid_Main.Columns(0).HeaderText = "설비이력코드"
        Grid_Main.Columns(1).HeaderText = "등록일자"
        Grid_Main.Columns(2).HeaderText = "등록시각"
        Grid_Main.Columns(3).HeaderText = "담당자"
        Grid_Main.Columns(4).HeaderText = "설비코드"
        Grid_Main.Columns(5).HeaderText = "설비명"
        Grid_Main.Columns(6).HeaderText = "내역"
        Grid_Main.Columns(7).HeaderText = "금액"
        Grid_Main.Columns(8).HeaderText = "비고"

        Dim i As Integer


        'For i = 0 To Grid_Main.ColumnCount - 1
        '    Grid_Main.Columns(i).Width = 180
        'Next i


        Grid_Main_Setup = True
    End Function

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click

        Log_Display()
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer



        Grid_Main.RowCount = 0


        If ComboBox1.Text = "전체" Then
            StrSQL = "select * from FC_RECORD"
        Else

            StrSQL = "select * from FC_RECORD where FL_PD_CODE='" & ComboBox1.Text & "'"

        End If

        Re_Count = DB_Select(DBT)

        If Re_Count <> 0 Then
            Grid_Main.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To 8
                    Grid_Main.Item(j, i).Value = DBT.Rows(i).Item(j).ToString
                    Grid_Main.Item(j, i).ReadOnly = True
                Next j
            Next i
        Else
            Grid_Main.RowCount = 1
            For j = 0 To Grid_Main.ColumnCount - 1
                Grid_Main.Item(j, 0).Value = ""
            Next j
        End If
    End Sub
    Public Function Log_Setup() As Boolean
        Grid_Color(DataGridView1)

        DataGridView1.AllowUserToAddRows = False
        DataGridView1.RowTemplate.Height = 20
        DataGridView1.ColumnCount = 5
        DataGridView1.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.RowHeadersWidth = 10

        DataGridView1.Columns(0).HeaderText = "작업일자"
        DataGridView1.Columns(1).HeaderText = "작업화면"
        DataGridView1.Columns(2).HeaderText = "작업"

        DataGridView1.Columns(3).HeaderText = "작업자"
        DataGridView1.Columns(4).HeaderText = "작업비고"


        Dim i As Integer
        For i = 0 To DataGridView1.ColumnCount - 1
            DataGridView1.Columns(i).ReadOnly = True
        Next i
        Log_Setup = True
    End Function

    Public Function Log_Display() As Boolean
        Log_Setup()
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer

        StrSQL = "Select * FROM MAN_LOG with(NOLOCK) order By ACT_DATE DESC"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            DataGridView1.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To DataGridView1.ColumnCount - 1
                    DataGridView1.Item(j, i).Value = DBT.Rows(i).Item(j)
                Next j
            Next i
        Else
            DataGridView1.RowCount = 1
            For j = 0 To DataGridView1.ColumnCount - 1
                DataGridView1.Item(j, 0).Value = ""
            Next j

        End If

        Log_Display = True
        DataGridView1.ClearSelection()


    End Function
    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
        Dim DBT As New DataTable
        StrSQL = "Select * FROM SI_MACHINE with(NOLOCK) Where PD_Code = '" & ComboBox1.SelectedItem.ToString() & "'"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            ComboBox2.Text = DBT.Rows(0)("PD_Name")
        End If
    End Sub



    Private Sub ComboBox2_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox2.SelectionChangeCommitted
        Dim DBT As New DataTable
        StrSQL = "Select * FROM SI_MACHINE with(NOLOCK) Where PD_Name = '" & ComboBox2.SelectedItem.ToString() & "'"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            ComboBox1.Text = DBT.Rows(0)("PD_Code")
        End If

    End Sub
    Private Sub Com_Contract_Click(sender As Object, e As EventArgs) Handles Com_Contract.Click
        If Excel_Check = True Then
            xlbook.Close()
            xlapp.Quit()
            xlsheet = Nothing
            xlbook = Nothing
            xlapp = Nothing
            releaseObject(xlsheet)
            releaseObject(xlbook)
            releaseObject(xlapp)
            Excel_Check = False
        End If
        Label_Color(sender, Color_Label, Di_Panel2)
        '판넬 보이기
        Panel_Main.Visible = True
        Panel_Excel.Visible = False
        Com_Excel_Print.Visible = False
    End Sub

    Private Sub Com_Excel_Click(sender As Object, e As EventArgs) Handles Com_Excel.Click
        Label_Color(sender, Color_Label, Di_Panel2)
        '판넬 보이기
        Panel_Main.Visible = False
        Panel_Excel.Visible = True
        Com_Excel_Print.Visible = True

        Cursor = Cursors.AppStarting


        '여기서 처리

        If Excel_Check = True Then
            xlbook.Close()
            xlapp.Quit()
            xlsheet = Nothing
            xlbook = Nothing
            xlapp = Nothing
            releaseObject(xlsheet)
            releaseObject(xlbook)
            releaseObject(xlapp)
            Excel_Check = False

        End If

        Dim i As Integer
        Dim j As Integer


        xlapp = New Excel.Application
        xlbook = xlapp.Workbooks.Open(Application.StartupPath & "\excel\null.xlsx")
        xlsheet = xlbook.ActiveSheet
        xlapp.DisplayAlerts = False
        xlapp.WindowState = Excel.XlWindowState.xlMinimized
        xlapp.Visible = False
        xlapp.DisplayFormulaBar = False
        xlapp.DisplayStatusBar = False




        'PInvoke 에러시 프로젝트 속성에서 컴파일을 x64로 하거나 Any CPU일 경우 32비트 기본 사용 체크를 해제한다.
        'SetWindowLong(xlapp.Hwnd, GWL_STYLE, GetWindowLong(xlapp.Hwnd, GWL_STYLE) And Not WS_CAPTION)

        xlapp.ActiveWindow.DisplayWorkbookTabs = False
        xlapp.ActiveWindow.DisplayHeadings = False

        'xlapp.ActiveWindow.DisplayWorkbookTabs = True
        'xlapp.ActiveWindow.DisplayHeadings = True


        SetParent(xlapp.Hwnd, Panel_Excel.Handle)
        SendMessage(xlapp.Hwnd, WM_SYSCOMMAND, SC_MAXIMIZE, 0)



        '제목

        For j = 0 To Grid_Main.ColumnCount - 1
            xlsheet.Cells(1, j + 1).Value = Grid_Main.Columns(j).HeaderText
            xlsheet.Columns(j + 1).ColumnWidth = Grid_Main.Columns(j).Width / 10

        Next j

        For i = 0 To Grid_Main.RowCount - 1
            For j = 0 To Grid_Main.ColumnCount - 1
                xlsheet.Cells(i + 2, j + 1).Value = "'" & Grid_Main.Item(j, i).Value
            Next j
        Next i



        xlbook.SaveAs(Application.StartupPath & "\Excel\내역서\" & Format(Now, "yyyyMMddHHmmss") & ".xlsx")
        Cursor = Cursors.Default
        Excel_Check = True
    End Sub

    Private Sub Com_Excel_Print_Click(sender As Object, e As EventArgs) Handles Com_Excel_Print.Click
        If Excel_Check = True Then
            xlsheet.PrintOut(From:=1, To:=1, Copies:=1, Collate:=True)
        End If
    End Sub

    Private Sub Insert_Com_Click(sender As Object, e As EventArgs) Handles Insert_Com.Click

        Dim form As New FLReport1_Insert
        form.Label1.Text = "insert"
        form.ShowDialog()

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '해당 칼럼 삭제
        If Len(Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        If MsgBox("순번 '" & Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value & "'을(를) 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "삭제") = vbNo Then
            Exit Sub
        End If

        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Delete FC_RECODE Where FL_CODE = '" & Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

    End Sub

    Private Sub Grid_Main_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Main.CellClick
        Dim form As New FLReport1_Insert
        form.Label1.Text = "update"


        form.FL_CODE.Text = Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value



        form.ShowDialog()

        Grid_Main.CurrentRow.HeaderCell.Value = "*"

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If

    End Sub
End Class
