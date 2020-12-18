Public Class Frm_Claim
    Dim Grid_Display As Boolean
    Dim Grid_Display_Ch As Boolean

    Dim Grid_Bom_Row As Integer
    Dim Grid_Bom_Col As Integer

    Private Sub Frm_Device2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()
        Process_Line_Grid_Setup()
        Search_Combo.Items.Add("날짜")
        '  Search_Combo.Items.Add("상태")
        Search_Combo.Text = "전체"

        ' Insert_Com.Visible = False

        Search_Com.PerformClick()
        Process_Line_Grid.ClearSelection()
    End Sub
    Public Function Process_Line_Grid_Setup() As Boolean

        Grid_Color(Process_Line_Grid)


        Process_Line_Grid.AllowUserToAddRows = False
        Process_Line_Grid.RowTemplate.Height = 20


        '열의 갯수는 하나 더 많아야 함.
        Process_Line_Grid.ColumnCount = 6
        Process_Line_Grid.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Process_Line_Grid.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Process_Line_Grid.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Process_Line_Grid.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Process_Line_Grid.RowHeadersWidth = 40
        Process_Line_Grid.Columns(0).Width = 100
        Process_Line_Grid.Columns(1).Width = 100
        Process_Line_Grid.Columns(2).Width = 100
        Process_Line_Grid.Columns(3).Width = 100
        Process_Line_Grid.Columns(4).Width = 100


        Process_Line_Grid.Columns(0).HeaderText = "코드"
        Process_Line_Grid.Columns(1).HeaderText = "거래처명"
        Process_Line_Grid.Columns(2).HeaderText = "내용"
        Process_Line_Grid.Columns(3).HeaderText = "담당자"
        Process_Line_Grid.Columns(4).HeaderText = "일자"
        Process_Line_Grid.Columns(5).HeaderText = "비고"


        Process_Line_Grid_Setup = True
    End Function

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        Process_Line_Grid_Display()
        Log_Display()

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

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Log_Setup()

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
    Public Function Process_Line_Grid_Display() As Boolean

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display = False
        Process_Line_Grid.RowCount = 1
        Select Case Search_Combo.Text
            Case "전체"
                StrSQL = "Select * FROM QC_CLAIM with(NOLOCK) Order By qc_CODE"

        End Select
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Process_Line_Grid.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To DBT.Columns.Count - 1
                    Process_Line_Grid.Item(j, i).Value = DBT.Rows(i).Item(j)
                Next j
            Next i
        Else
            Process_Line_Grid.RowCount = 1
            For j = 0 To Process_Line_Grid.ColumnCount - 1
                Process_Line_Grid.Item(j, 0).Value = ""
            Next j

        End If
        Grid_Display = True
        Process_Line_Grid_Display = True
        Process_Line_Grid.ClearSelection()

    End Function
    Private Sub Process_Line_Grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Process_Line_Grid.CellValueChanged
        If Grid_Display = True Then
            Process_Line_Grid.CurrentRow.HeaderCell.Value = "*"
        End If

    End Sub

    Private Sub Save_Com_Click(sender As Object, e As EventArgs) Handles Save_Com.Click
        Dim i As Integer
        Grid_Display = False
        For i = 0 To Process_Line_Grid.Rows.Count - 1
            If Process_Line_Grid.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE QC_CLAIM Set
                                     QC_CUSTOMER = '" & Process_Line_Grid.Item(1, i).Value & "',
                                     QC_CONTENT = '" & Process_Line_Grid.Item(2, i).Value & "',
                                     QC_MAN = '" & Process_Line_Grid.Item(3, i).Value & "',
                                     QC_DATE = '" & Process_Line_Grid.Item(4, i).Value & "'.
                                     QC_BIGO = '" & Process_Line_Grid.Item(5, i).Value & "'
                                   where qc_CODE = '" & Process_Line_Grid.Item(0, i).Value & "'"
                Re_Count = DB_Execute()
                Process_Line_Grid.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display = True
    End Sub

    Private Sub Insert_Com_Click(sender As Object, e As EventArgs) Handles Insert_Com.Click
        ''Frm_Insert.Text = "클레임 코드 추가"
        ''Frm_Insert.Show()
        ''Frm_Main.Enabled = False

        'Dim DBT As New DataTable
        'Dim JP_Code As Long

        'StrSQL = "Select QC_CODE FROM QC_CLAIM with(NOLOCK)  Order By Convert(Decimal,QC_CODE) Desc "
        'Re_Count = DB_Select(DBT)
        'If Re_Count = 0 Then
        '    JP_Code = JP_Code & "0001"
        'Else
        '    JP_Code = Val(DBT.Rows(0).Item(0)) + 1
        'End If

        ''추가한다
        'StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        'StrSQL = StrSQL & "Insert INTO QC_CLAIM (QC_CODE)  Values ('" & JP_Code & "')"
        'Re_Count = DB_Execute()

        'Search_Com.PerformClick()


        Dim form As New Claim_Insert
        form.Label1.Text = "insert"
        form.ShowDialog()

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If
    End Sub

    Private Sub Delete_Com_Click(sender As Object, e As EventArgs) Handles Delete_Com.Click
        If Len(Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value) < 1 Then
            MsgBox("코드를 선택하세요.")
            Exit Sub
        End If
        If MsgBox("코드  '" & Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value & "'를 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "코드 삭제") = vbNo Then
            Exit Sub
        End If
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE QC_CLAIM where qc_CODE = '" & Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()
        Process_Line_Grid_Display()
        MsgBox("삭제 되었습니다.")
    End Sub

    Private Sub Panel_Menu_Paint(sender As Object, e As PaintEventArgs) Handles Panel_Menu.Paint

    End Sub

    Private Sub Process_Line_Grid_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Process_Line_Grid.CellClick
        Dim form As New Claim_Insert
        form.Label1.Text = "update"


        form.QC_CODE.Text = Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value



        form.ShowDialog()

        Process_Line_Grid.CurrentRow.HeaderCell.Value = "*"

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If


    End Sub
End Class
