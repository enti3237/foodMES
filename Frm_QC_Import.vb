Imports Excel = Microsoft.Office.Interop.Excel

Public Class Frm_QC_Import
    Dim QC_Label As String = "수입검사"
    Dim Grid_Display_Ch As Boolean

    Dim Grid_Info_Row As Integer
    Dim Grid_Info_Col As Integer

    Dim Grid_Main_Row As Integer
    Dim Grid_Main_Col As Integer



    Dim xlapp As Excel.Application
    Dim xlbook As Excel.Workbook
    Dim xlsheet As Excel.Worksheet

    Dim Excel_Check As Boolean


    Private Sub Frm_QC_Travel_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()




        Com_QC.Text = QC_Label
        Grid_Code_Setup()
        Log_Setup()

        Panel2.Visible = True
        Excel_Panel.Visible = False

        Print_Com.Visible = False
        Search_Com.PerformClick()
    End Sub
    Public Function Grid_Code_Setup() As Boolean
        Grid_Color(Grid_Code)

        Grid_Code.AllowUserToAddRows = False
        Grid_Code.RowTemplate.Height = 24
        Grid_Code.ColumnCount = 17
        Grid_Code.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Code.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Code.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Code.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Code.RowHeadersWidth = 10
        Grid_Code.Columns(0).Width = 100
        Grid_Code.Columns(1).Width = 110
        Grid_Code.Columns(2).Width = 110

        Grid_Code.Columns(0).HeaderText = "리스트코드"
        Grid_Code.Columns(1).HeaderText = "검사일"
        Grid_Code.Columns(2).HeaderText = "품질검사구분"
        Grid_Code.Columns(3).HeaderText = "담당자"
        Grid_Code.Columns(4).HeaderText = "코드"
        Grid_Code.Columns(5).HeaderText = "제품코드"
        Grid_Code.Columns(6).HeaderText = "제품명"
        Grid_Code.Columns(7).HeaderText = "검사항목"
        Grid_Code.Columns(8).HeaderText = "검사규격"
        Grid_Code.Columns(9).HeaderText = "검사수준"
        Grid_Code.Columns(10).HeaderText = "AQL"
        Grid_Code.Columns(11).HeaderText = "시료1"
        Grid_Code.Columns(12).HeaderText = "시료2"
        Grid_Code.Columns(13).HeaderText = "시료3"
        Grid_Code.Columns(14).HeaderText = "시료4"
        Grid_Code.Columns(15).HeaderText = "시료5"
        Grid_Code.Columns(16).HeaderText = "비고"


        Dim i As Integer
        For i = 0 To Grid_Code.ColumnCount - 1
            Grid_Code.Columns(0).ReadOnly = True
            Grid_Code.Columns(1).ReadOnly = True
            Grid_Code.Columns(2).ReadOnly = True
        Next i
        Grid_Code_Setup = True
    End Function


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
    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        Grid_Code_Display()
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
            Panel2.Visible = True

        End If
        Log_Display()
    End Sub
    Public Function Log_Display() As Boolean

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
    Public Function Grid_Code_Display() As Boolean

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Code.RowCount = 1
        If Search_Text.Text = "" Then
            StrSQL = "Select * from QUALITY Where QC_TYPE= '" & QC_Label & "'  Order By QC_Date"
        Else
            StrSQL = "Select * from QUALITY Where QC_TYPE= '" & QC_Label & "' and QC_PL_NAME like '%" & Search_Text.Text & "%' Order By QC_Date"
        End If

        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Code.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To Grid_Code.ColumnCount - 1
                    Grid_Code.Item(j, i).Value = DBT.Rows(i).Item(j)
                Next j
            Next i
        Else
            Grid_Code.RowCount = 1
            For j = 0 To Grid_Code.ColumnCount - 1
                Grid_Code.Item(j, 0).Value = ""
            Next j
        End If
        Grid_Code_Display = True

        Print_Com.Visible = False
        Grid_Code.ClearSelection()
    End Function

    Private Sub Insert_Com_Click(sender As Object, e As EventArgs) Handles Insert_Com.Click
        ''새로운 코드 추가

        Dim form As New Import_Insert
        form.Label1.Text = "insert"
        form.ShowDialog()

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If


    End Sub


    Private Sub Com_QC_Click(sender As Object, e As EventArgs) Handles Com_QC.Click

        Print_Com.Visible = False


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

        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        Label_Color(sender, Color_Label, Di_Panel2)
        '판넬 보이기
        Panel2.Visible = True
        Excel_Panel.Visible = False
    End Sub

    Private Sub Com_Excel_Click(sender As Object, e As EventArgs) Handles Com_Excel.Click
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        Label_Color(sender, Color_Label, Di_Panel2)
        '판넬 보이기
        Panel2.Visible = False
        Excel_Panel.Visible = True
        Print_Com.Visible = True


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
        '여기서 처리


        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer

        xlapp = New Excel.Application
        xlbook = xlapp.Workbooks.Open(Application.StartupPath & "\excel\수입검사.xlsx")
        xlsheet = xlbook.ActiveSheet
        xlapp.DisplayAlerts = False
        xlapp.WindowState = Excel.XlWindowState.xlMinimized
        xlapp.Visible = False
        xlapp.DisplayFormulaBar = False
        xlapp.DisplayStatusBar = False

        xlapp.ActiveWindow.DisplayWorkbookTabs = False
        xlapp.ActiveWindow.DisplayHeadings = False

        SetParent(xlapp.Hwnd, Excel_Panel.Handle)
        SendMessage(xlapp.Hwnd, WM_SYSCOMMAND, SC_MAXIMIZE, 0)




        'xlbook.SaveAs(Application.StartupPath & "\Excel\QC\" & Format(Now, "yyyyMMddHHmmss") & ".xlsx")

        Excel_Panel.Visible = True
    End Sub

    Private Sub Excel_Save_Com_Click(sender As Object, e As EventArgs)
        xlbook.SaveAs(Application.StartupPath & "\Excel\" & QC_Label & Format(Now, "yyyyMMddHHmmss") & ".xlsx")
        xlsheet = Nothing
        xlbook.Close()
        xlbook = Nothing
        xlapp.Quit()
        xlapp = Nothing
    End Sub

    Private Sub Print_Com_Click(sender As Object, e As EventArgs) Handles Print_Com.Click
        If Excel_Check = True Then
            xlsheet.PrintOut(From:=1, To:=1, Copies:=1, Collate:=True)
        End If
    End Sub





    Private Sub Grid_Code_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Code.CellClick
        Dim form As New Import_Insert
        form.Label1.Text = "update"


        form.QC_CODE.Text = Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value
        form.number = Grid_Code.CurrentCell.RowIndex


        form.ShowDialog()

        Grid_Code.CurrentRow.HeaderCell.Value = "*"

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If




    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '해당 칼럼 삭제
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        If MsgBox("순번 '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "'을(를) 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "삭제") = vbNo Then
            Exit Sub
        End If

        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Delete QUALITY Where QC_CODE = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        Grid_Code_Display()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

    Private Sub Panel_Menu_Paint(sender As Object, e As PaintEventArgs) Handles Panel_Menu.Paint

    End Sub

    Private Sub Search_Text_TextChanged(sender As Object, e As EventArgs) Handles Search_Text.TextChanged

    End Sub

    Private Sub Di_Panel2_Paint(sender As Object, e As PaintEventArgs) Handles Di_Panel2.Paint

    End Sub

    Private Sub Color_Label_Click(sender As Object, e As EventArgs) Handles Color_Label.Click

    End Sub

    Private Sub Excel_Panel_Paint(sender As Object, e As PaintEventArgs) Handles Excel_Panel.Paint

    End Sub

    Private Sub Di_Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Di_Panel1.Paint

    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Panel2_Paint(sender As Object, e As PaintEventArgs) Handles Panel2.Paint

    End Sub

    Private Sub Grid_Code_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Code.CellContentClick

    End Sub

    Private Sub Save_Com_Click(sender As Object, e As EventArgs) Handles Save_Com.Click

    End Sub
End Class
