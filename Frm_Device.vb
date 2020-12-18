Public Class Frm_Device
    Dim Grid_Display As Boolean
    Dim Grid_Display_Ch As Boolean
    Dim CellClick As Boolean

    Dim Grid_Bom_Row As Integer
    Dim Grid_Bom_Col As Integer

    Private Sub Frm_Device_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()
        Process_Line_Grid_Setup()
        Process_Line_QC_Setup()
        'Search_Combo.Items.Add("설비코드")
        'Search_Combo.Items.Add("설비명")
        'Search_Combo.Text = "설비코드"
        Process_Line_Grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        Process_Line_QC.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        Process_Line_Grid_Display()
        Process_Line_Grid.ClearSelection()
        CellClick = False
    End Sub
    Public Function Process_Line_Grid_Setup() As Boolean

        Grid_Color(Process_Line_Grid)

        Process_Line_Grid.AllowUserToAddRows = False
        Process_Line_Grid.RowTemplate.Height = 20


        '열의 갯수는 하나 더 많아야 함.
        Process_Line_Grid.ColumnCount = 5
        Process_Line_Grid.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Process_Line_Grid.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Process_Line_Grid.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Process_Line_Grid.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Process_Line_Grid.RowHeadersWidth = 40
        Process_Line_Grid.Columns(0).Width = 60
        Process_Line_Grid.Columns(1).Width = 100
        Process_Line_Grid.Columns(2).Width = 100
        Process_Line_Grid.Columns(3).Width = 60
        Process_Line_Grid.Columns(4).Width = 400

        Process_Line_Grid.Columns(0).HeaderText = "설비코드"
        Process_Line_Grid.Columns(1).HeaderText = "설비명"
        Process_Line_Grid.Columns(2).HeaderText = "ID"
        Process_Line_Grid.Columns(3).HeaderText = "Port"
        Process_Line_Grid.Columns(4).HeaderText = "비고"

        'For i = 3 To 7
        '    Process_Line_Grid.Columns(i).Visible = False
        'Next



        Process_Line_Grid.Columns(0).ReadOnly = True
        Process_Line_Grid_Setup = True
    End Function
    Public Function Process_Line_QC_Setup() As Boolean

        Grid_Color(Process_Line_QC)

        Process_Line_QC.AllowUserToAddRows = False
        Process_Line_QC.RowTemplate.Height = 24


        '열의 갯수는 하나 더 많아야 함.
        Process_Line_QC.ColumnCount = 3
        Process_Line_QC.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Process_Line_QC.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Process_Line_QC.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Process_Line_QC.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Process_Line_QC.RowHeadersWidth = 40
        Process_Line_QC.Columns(0).Width = 60
        Process_Line_QC.Columns(1).Width = 150
        Process_Line_QC.Columns(2).Width = 500
        Process_Line_QC.Columns(0).HeaderText = "순번"
        Process_Line_QC.Columns(1).HeaderText = "점검내역"
        Process_Line_QC.Columns(2).HeaderText = "비고"


        Process_Line_QC.Columns(0).ReadOnly = True
        Process_Line_QC_Setup = True
    End Function

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        Process_Line_Grid_Display()
        Process_Line_QC_Display(Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value)
    End Sub
    Public Function Process_Line_Grid_Display() As Boolean

        Dim i As Integer
        Dim j As Integer
        Dim DBT As New DataTable
        Grid_Display = False
        Process_Line_Grid.RowCount = 1

        StrSQL = "Select * FROM SI_MACHINE with(NOLOCK) Where PD_Name Like '%" & Search_Text.Text & "%' Or PD_Name Is Null Order By PD_Name"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Process_Line_Grid.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To 4
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
        Process_Line_QC.ClearSelection()

    End Function

    Private Sub Insert_Com_Click(sender As Object, e As EventArgs) Handles Insert_Com.Click
        'Frm_Insert.Text = "생산설비 코드 추가"
        'Frm_Insert.Show()
        'Frm_Main.Enabled = False
        Dim form As New Device_Insert
        form.Label1.Text = "insert"
        form.ShowDialog()

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If
    End Sub

    Public Function Process_Line_QC_Display(PD_Code As String) As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display_Ch = False
        Process_Line_QC.RowCount = 1
        StrSQL = "Select PQ_Sun, PQ_History, PQ_Bigo   FROM SI_MACHINE_QC with(NOLOCK)  Where PQ_Code = '" & PD_Code & "' Order By Convert(Decimal,PQ_Sun)"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Process_Line_QC.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                Process_Line_QC.Item(0, i).Value = DBT.Rows(i).Item(0).ToString
                Process_Line_QC.Item(1, i).Value = DBT.Rows(i).Item(1).ToString
                Process_Line_QC.Item(2, i).Value = DBT.Rows(i).Item(2).ToString
            Next i
        Else
            Process_Line_QC.RowCount = 1
            For j = 0 To Process_Line_QC.ColumnCount - 1
                Process_Line_QC.Item(j, 0).Value = ""
            Next j

        End If
        Grid_Display_Ch = True
        Process_Line_QC_Display = True
    End Function

    'Private Sub Save_Com_Click(sender As Object, e As EventArgs) Handles Save_Com.Click
    '    Dim i As Integer
    '    Grid_Display = False
    '    For i = 0 To Process_Line_Grid.Rows.Count - 1
    '        If Process_Line_Grid.Rows(i).HeaderCell.Value = "*" Then
    '            'Update
    '            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
    '            StrSQL = StrSQL & "UPDATE SI_MACHINE Set PD_Name = '" & Process_Line_Grid.Item(1, i).Value & "', PD_Group = '" & Process_Line_Grid.Item(2, i).Value & "', PD_ID = '" & Process_Line_Grid.Item(3, i).Value & "', PD_Port = '" & Process_Line_Grid.Item(4, i).Value & "', PD_Sel = '" & Process_Line_Grid.Item(5, i).Value & "', PD_Count = '" & Process_Line_Grid.Item(6, i).Value & "', PD_Bigo = '" & Process_Line_Grid.Item(7, i).Value & "' where PD_Code = '" & Process_Line_Grid.Item(0, i).Value & "'"
    '            Re_Count = DB_Execute()
    '            Process_Line_Grid.Rows(i).HeaderCell.Value = ""
    '        End If
    '    Next i
    '    Grid_Display = True
    '    MsgBox("저장되었습니다")
    '    Search_Com.PerformClick()
    'End Sub

    Private Sub Process_Line_Grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
        If Grid_Display = True Then
            Process_Line_Grid.CurrentRow.HeaderCell.Value = "*"
        End If

    End Sub

    Private Sub Delete_Com_Click(sender As Object, e As EventArgs) Handles Delete_Com.Click

        If Len(Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value) < 1 Then
            MsgBox("설비코드를 선택하세요.")
            Exit Sub
        End If
        If MsgBox("설비코드  '" & Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value & "'를 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "설비코드 삭제") = vbNo Then
            Exit Sub
        End If
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_MACHINE where PD_Code = '" & Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()
        Process_Line_Grid_Display()
        MsgBox("삭제 되었습니다.")


    End Sub
    Private Sub Process_Line_QC_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
        If Grid_Display_Ch = True Then
            Process_Line_QC.CurrentRow.HeaderCell.Value = "*"
        End If
    End Sub

    'Private Sub Save_Bom_Click(sender As Object, e As EventArgs)
    '    If Len(Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value) > 0 Then
    '    Else
    '        Exit Sub
    '    End If
    '    Dim i As Integer
    '    Grid_Display_Ch = False
    '    For i = 0 To Process_Line_QC.Rows.Count - 1
    '        If Process_Line_QC.Rows(i).HeaderCell.Value = "*" Then
    '            'Update
    '            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
    '            StrSQL = StrSQL & "UPDATE SI_MACHINE_QC Set PQ_History = '" & Process_Line_QC.Item(1, i).Value & "', PQ_Bigo = '" & Process_Line_QC.Item(2, i).Value & "' where PQ_Code = '" & Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value & "' And PQ_Sun = '" & Process_Line_QC.Item(0, i).Value & "'"
    '            Re_Count = DB_Execute()
    '            Process_Line_QC.Rows(i).HeaderCell.Value = ""
    '        End If
    '    Next i
    '    Grid_Display_Ch = True
    '    MsgBox("저장되었습니다")
    'End Sub

    Private Sub Insert__Bom_Click(sender As Object, e As EventArgs)
        If Len(Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        Grid_Display_Ch = False

        Dim DBT As New DataTable
        Dim Db_Sun As Long
        StrSQL = "Select PQ_Sun FROM SI_MACHINE_QC with(NOLOCK) Where PQ_Code = '" & Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value & "' Order By Convert(Decimal,PQ_Sun) Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            Db_Sun = 1
        Else
            Db_Sun = DBT.Rows(0).Item(0) + 1
        End If

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO SI_MACHINE_QC (PQ_Code, PQ_Sun)  Values ('" & Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value & "', '" & Db_Sun & "')"
        Re_Count = DB_Execute()
        Process_Line_QC_Display(Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True

    End Sub

    Private Sub Put__Bom_Click(sender As Object, e As EventArgs)
        If Len(Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        Grid_Display_Ch = False
        Dim i As Integer
        For i = Process_Line_QC.RowCount - 1 To Process_Line_QC.CurrentCell.RowIndex Step -1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "UPDATE SI_MACHINE_QC Set PQ_Sun = '" & i + 2 & "' Where PQ_Sun = '" & i + 1 & "' And PQ_Code = '" & Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value & "'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO SI_MACHINE_QC (PQ_Code, PQ_Sun)  Values ('" & Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value & "', '" & Process_Line_QC.CurrentCell.RowIndex + 1 & "')"
        Re_Count = DB_Execute()
        Process_Line_QC_Display(Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True

    End Sub

    Private Sub Process_Line_Grid_CellEnter(sender As Object, e As DataGridViewCellEventArgs)
        If Len(Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value) > 0 Then
            Process_Line_QC_Display(Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value)
        End If

    End Sub

    Private Sub Grid_Bom_Combo_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Process_Line_Grid_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Process_Line_Grid.CellDoubleClick
        '유효성 검사
        If Len(Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value) < 0 Then
            MsgBox("해당 행은 공백이므로 처리할 수 없습니다")
            Exit Sub
        End If

        '해당 행 내용 수정
        Device_Insert.Label1.Text = "update"
        Device_Insert.device_code.Text = Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value
        Dim form As New Device_Insert
        form.Label1.Text = "update"
        form.device_code.Text = Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value
        form.ShowDialog()

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If
    End Sub

    Private Sub Insert__Bom_Click_1(sender As Object, e As EventArgs) Handles Insert__Bom.Click

        If CellClick = False Then
            MsgBox("설비를 선택해주세요")
            Exit Sub
        End If


        Dim form As New Device_Insert
        form.Label1.Text = "insert2"
        form.device_code.Text = Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value
        form.Device_name.Text = Process_Line_Grid.Item(1, Process_Line_Grid.CurrentCell.RowIndex).Value
        form.Device_ID.Text = Process_Line_Grid.Item(2, Process_Line_Grid.CurrentCell.RowIndex).Value
        form.Device_Port.Text = Process_Line_Grid.Item(3, Process_Line_Grid.CurrentCell.RowIndex).Value
        form.Device_bigo.Text = Process_Line_Grid.Item(4, Process_Line_Grid.CurrentCell.RowIndex).Value

        form.ShowDialog()

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If
    End Sub

    Private Sub Delete__Bom_Click_1(sender As Object, e As EventArgs) Handles Delete__Bom.Click

        If Len(Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If


        If MsgBox("순번 '" & Process_Line_QC.Item(0, Process_Line_QC.CurrentCell.RowIndex).Value & "'을(를) 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "삭제") = vbNo Then
            Exit Sub
        End If


        Grid_Display_Ch = False
        Dim i As Integer
        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_MACHINE_QC Where PQ_Code = '" & Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value & "' And   PQ_Sun = '" & Process_Line_QC.Item(0, Process_Line_QC.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        For i = Process_Line_QC.CurrentCell.RowIndex + 1 To Process_Line_QC.RowCount - 1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "UPDATE SI_MACHINE_QC Set PQ_Sun = '" & i & "' Where PQ_Code = '" & Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value & "' And  PQ_Sun = '" & i + 1 & "'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        Process_Line_QC_Display(Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True
    End Sub

    Private Sub Process_Line_QC_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Process_Line_QC.CellDoubleClick
        '유효성 검사
        If Len(Process_Line_QC.Item(0, Process_Line_QC.CurrentCell.RowIndex).Value) < 0 Then
            MsgBox("해당 행은 공백이므로 처리할 수 없습니다")
            Exit Sub
        End If

        '해당 행 내용 수정
        Device_Insert.Label1.Text = "update2"
        'Device_Insert.device_code.Text = Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value
        Dim form As New Device_Insert
        form.Label1.Text = "update2"
        form.device_code.Text = Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value
        form.Device_name.Text = Process_Line_Grid.Item(1, Process_Line_Grid.CurrentCell.RowIndex).Value
        form.Device_ID.Text = Process_Line_Grid.Item(2, Process_Line_Grid.CurrentCell.RowIndex).Value
        form.Device_Port.Text = Process_Line_Grid.Item(3, Process_Line_Grid.CurrentCell.RowIndex).Value
        form.Device_bigo.Text = Process_Line_Grid.Item(4, Process_Line_Grid.CurrentCell.RowIndex).Value

        form.TextBox3.Text = Process_Line_QC.Item(0, Process_Line_QC.CurrentCell.RowIndex).Value
        form.TextBox4.Text = Process_Line_QC.Item(1, Process_Line_QC.CurrentCell.RowIndex).Value
        form.TextBox6.Text = Process_Line_QC.Item(2, Process_Line_QC.CurrentCell.RowIndex).Value

        'form.device_code.Text = Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value
        form.ShowDialog()

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If
    End Sub

    Private Sub Process_Line_Grid_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Process_Line_Grid.CellClick
        If Len(Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value) > 0 Then
            CellClick = True
            Process_Line_QC_Display(Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value)
        End If
    End Sub
End Class
