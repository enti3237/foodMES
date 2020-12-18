Public Class Frm_Process
    Dim CellClick As Boolean
    Dim Grid_Display As Boolean
    Dim Process_List_Line_Grid_Row As Integer
    Dim Process_List_Line_Grid_Col As Integer

    Private Sub Frm_Process_List_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()

        'Search_Combo.Items.Add("공정코드")
        'Search_Combo.Items.Add("공정명")
        'Search_Combo.Text = "공정코드"
        'Line_Panel.Visible = False
        Process_List_Grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        Process_List_Line_Grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        Process_List_Grid_Setup()
        Process_List_Line_Grid_Setup()
        Process_List_Grid_Display()
        Process_List_Grid.ClearSelection()


        CellClick = False
    End Sub
    Public Sub Process_List_Line_Grid_Setup()
        Grid_Color(Process_List_Line_Grid)

        Process_List_Line_Grid.AllowUserToAddRows = False
        Process_List_Line_Grid.RowTemplate.Height = 24

        Process_List_Line_Grid.ColumnCount = 4
        Process_List_Line_Grid.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Process_List_Line_Grid.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Process_List_Line_Grid.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Process_List_Line_Grid.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Process_List_Line_Grid.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Process_List_Line_Grid.RowHeadersWidth = 40
        Process_List_Line_Grid.Columns(0).Width = 35
        Process_List_Line_Grid.Columns(1).Width = 80
        Process_List_Line_Grid.Columns(2).Width = 150
        Process_List_Line_Grid.Columns(3).Width = 500

        Process_List_Line_Grid.Columns(0).HeaderText = "순번"
        Process_List_Line_Grid.Columns(1).HeaderText = "설비코드"
        Process_List_Line_Grid.Columns(2).HeaderText = "설비명"
        Process_List_Line_Grid.Columns(3).HeaderText = "비고"


        Process_List_Line_Grid.Item(0, 0).Value = ""
        Process_List_Line_Grid.Item(1, 0).Value = ""
        Process_List_Line_Grid.Item(2, 0).Value = ""
        Process_List_Line_Grid.Item(3, 0).Value = ""

        Process_List_Line_Grid.Columns(0).ReadOnly = True

    End Sub
    Public Function Process_List_Grid_Setup() As Boolean
        Grid_Color(Process_List_Grid)


        Process_List_Grid.AllowUserToAddRows = False
        Process_List_Grid.RowTemplate.Height = 20


        '열의 갯수는 하나 더 많아야 함.
        Process_List_Grid.ColumnCount = 3
        Process_List_Grid.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Process_List_Grid.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Process_List_Grid.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Process_List_Grid.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Process_List_Grid.RowHeadersWidth = 40
        Process_List_Grid.Columns(0).Width = 80
        Process_List_Grid.Columns(1).Width = 185
        Process_List_Grid.Columns(2).Width = 400
        Process_List_Grid.Columns(0).HeaderText = "공정코드"
        Process_List_Grid.Columns(1).HeaderText = "공정명"
        Process_List_Grid.Columns(2).HeaderText = "비고"


        Process_List_Grid.Columns(0).ReadOnly = True
        Process_List_Grid_Setup = True
    End Function

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        Process_List_Grid_Display()
        Process_List_Line_Grid_Display(Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value)

    End Sub
    Public Function Process_List_Grid_Display() As Boolean

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display = False
        Process_List_Grid.RowCount = 1
        StrSQL = "Select * FROM SI_PROCESS with(NOLOCK) Where PC_Name Like '%" & Search_Text.Text & "%' Or PC_Name Is Null Order By PC_Name"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Process_List_Grid.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To 2
                    Process_List_Grid.Item(j, i).Value = DBT.Rows(i).Item(j)
                Next j
            Next i
        Else
            Process_List_Grid.RowCount = 1
            For j = 0 To Process_List_Grid.ColumnCount - 1
                Process_List_Grid.Item(j, 0).Value = ""
            Next j

        End If
        Grid_Display = True
        Process_List_Grid_Display = True
        Process_List_Grid.ClearSelection()
        Process_List_Line_Grid.ClearSelection()


    End Function

    Private Sub Process_List_Grid_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub Process_List_Grid_CellEnter(sender As Object, e As DataGridViewCellEventArgs)
        'If Len(Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value) > 0 Then
        '    Line_Panel.Visible = True
        '    Process_List_Line_Grid_Display(Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value)
        'End If

    End Sub
    Public Function Process_List_Line_Grid_Display(Process_List_Code As String) As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display = False
        Process_List_Line_Grid.RowCount = 1
        StrSQL = "Select PC_Sun, PC_PL_CODE,  PC_Bigo  FROM SI_Process_Device With(NOLOCK)  Where PC_Code = '" & Process_List_Code & "' Order By  Convert(Decimal,PC_Sun) "
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Process_List_Line_Grid.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                Process_List_Line_Grid.Rows(i).HeaderCell.Value = ""
                Process_List_Line_Grid.Item(0, i).Value = DBT.Rows(i).Item(0)
                Process_List_Line_Grid.Item(1, i).Value = DBT.Rows(i).Item(1)
                Process_List_Line_Grid.Item(2, i).Value = ""
                Process_List_Line_Grid.Item(3, i).Value = DBT.Rows(i).Item(2)
            Next i
        Else
            Process_List_Line_Grid.RowCount = 1
            Process_List_Line_Grid.Rows(0).HeaderCell.Value = ""
            For j = 0 To Process_List_Line_Grid.ColumnCount - 1
                Process_List_Line_Grid.Item(j, 0).Value = ""
            Next j

        End If
        For i = 0 To Process_List_Line_Grid.RowCount - 1
            StrSQL = "Select PD_Name  FROM SI_MACHINE with(NOLOCK)  Where PD_Code = '" & Process_List_Line_Grid.Item(1, i).Value & "' "
            Re_Count = DB_Select(DBT)
            If Re_Count <> 0 Then
                Process_List_Line_Grid.Item(2, i).Value = DBT.Rows(0).Item(0)
            End If
        Next
        Grid_Display = True
        Process_List_Line_Grid_Display = True
    End Function

    'Private Sub Save_Com_Click(sender As Object, e As EventArgs) Handles Save_Com.Click
    '    Dim i As Integer
    '    Grid_Display = False
    '    For i = 0 To Process_List_Grid.Rows.Count - 1
    '        If Process_List_Grid.Rows(i).HeaderCell.Value = "*" Then
    '            'Update
    '            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
    '            StrSQL = StrSQL & "UPDATE SI_PROCESS Set PC_Name = '" & Process_List_Grid.Item(1, i).Value & "', PC_Bigo = '" & Process_List_Grid.Item(2, i).Value & "' where PC_Code = '" & Process_List_Grid.Item(0, i).Value & "'"
    '            Re_Count = DB_Execute()
    '            Process_List_Grid.Rows(i).HeaderCell.Value = ""
    '        End If
    '    Next i
    '    Grid_Display = True


    '    Search_Com.PerformClick()
    '    Process_List_Grid.ClearSelection()
    '    MsgBox("저장되었습니다")
    'End Sub

    Private Sub Insert_Com_Click(sender As Object, e As EventArgs) Handles Insert_Com.Click
        'DB를 추가 한다.
        '코드 번호 입력
        '새로운 푬에서 입력 받는다.
        'Frm_Insert.Text = "공정 코드 추가"
        'Frm_Insert.Show()
        'Frm_Main.Enabled = False

        'Search_Com.PerformClick()
        Dim form As New Process_Insert
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

        If Len(Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value) < 1 Then
            MsgBox("공정코드를 선택하세요.")
            Exit Sub
        End If
        If MsgBox("공정코드  '" & Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value & "'를 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "공정코드 삭제") = vbNo Then
            Exit Sub
        End If

        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_PROCESS where PC_Code = '" & Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()
        Grid_Display = False
        Process_List_Grid_Display()
        MsgBox("삭제 되었습니다.")
        Grid_Display = True

    End Sub

    Private Sub Process_List_Line_Grid_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub Process_List_Line_Grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
        If Grid_Display = True Then
            Process_List_Line_Grid.CurrentRow.HeaderCell.Value = "*"
        End If

    End Sub

    Private Sub Process_List_Line_Grid_MouseClick(sender As Object, e As MouseEventArgs)
        'Process_List_Line_Grid_Row = Process_List_Line_Grid.CurrentCell.RowIndex
        'Process_List_Line_Grid_Col = Process_List_Line_Grid.CurrentCell.ColumnIndex
        'Process_List_Line_Grid_Combo.Visible = False
        'If Process_List_Line_Grid_Row < 0 Then
        '    Exit Sub
        'End If
        'If Process_List_Line_Grid_Col < 0 Then
        '    Exit Sub
        'End If
        'Dim DBT As New DataTable
        'Dim i As Long


        'Select Case Process_List_Line_Grid_Col
        '    Case 1
        '        StrSQL = "Select PD_Code, PD_Name  FROM SI_MACHINE with(NOLOCK)  Order By PD_Code"
        '        Re_Count = DB_Select(DBT)
        '        Process_List_Line_Grid_Combo.Items.Clear()
        '        If Re_Count <> 0 Then
        '            For i = 0 To Re_Count - 1
        '                Process_List_Line_Grid_Combo.Items.Add(DBT.Rows(i).Item(0))
        '            Next i
        '        End If

        '        Process_List_Line_Grid_Combo.Size = Process_List_Line_Grid.CurrentCell.Size
        '        Process_List_Line_Grid_Combo.Top = Process_List_Line_Grid.GetCellDisplayRectangle(Process_List_Line_Grid_Col, Process_List_Line_Grid_Row, True).Top + Process_List_Line_Grid.Top
        '        Process_List_Line_Grid_Combo.Left = Process_List_Line_Grid.GetCellDisplayRectangle(Process_List_Line_Grid_Col, Process_List_Line_Grid_Row, True).Left + Process_List_Line_Grid.Left
        '        Process_List_Line_Grid_Combo.Text = Process_List_Line_Grid.CurrentCell.Value.ToString
        '        Process_List_Line_Grid_Combo.BackColor = Process_List_Line_Grid.RowsDefaultCellStyle.BackColor
        '        Process_List_Line_Grid_Combo.Visible = True
        '        Process_List_Line_Grid_Combo.Focus.ToString()
        '    Case 2
        '        StrSQL = "Select PD_Code, PD_Name  FROM SI_MACHINE with(NOLOCK)  Order By PD_Name"
        '        Re_Count = DB_Select(DBT)
        '        Process_List_Line_Grid_Combo.Items.Clear()
        '        If Re_Count <> 0 Then
        '            For i = 0 To Re_Count - 1
        '                Process_List_Line_Grid_Combo.Items.Add(DBT.Rows(i).Item(1))
        '            Next i
        '        End If
        '        Process_List_Line_Grid_Combo.Size = Process_List_Line_Grid.CurrentCell.Size
        '        Process_List_Line_Grid_Combo.Top = Process_List_Line_Grid.GetCellDisplayRectangle(Process_List_Line_Grid_Col, Process_List_Line_Grid_Row, True).Top + Process_List_Line_Grid.Top
        '        Process_List_Line_Grid_Combo.Left = Process_List_Line_Grid.GetCellDisplayRectangle(Process_List_Line_Grid_Col, Process_List_Line_Grid_Row, True).Left + Process_List_Line_Grid.Left
        '        Process_List_Line_Grid_Combo.Text = Process_List_Line_Grid.CurrentCell.Value.ToString
        '        Process_List_Line_Grid_Combo.BackColor = Process_List_Line_Grid.RowsDefaultCellStyle.BackColor
        '        Process_List_Line_Grid_Combo.Visible = True
        '        Process_List_Line_Grid_Combo.Focus.ToString()
        'End Select

    End Sub

    Private Sub Process_List_Line_Grid_SelectionChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Process_List_Line_Grid_Scroll(sender As Object, e As ScrollEventArgs)
        'Process_List_Line_Grid_Combo.Visible = False

    End Sub

    Private Sub Process_List_Line_Grid_Combo_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Process_List_Line_Grid_Combo_SelectionChangeCommitted(sender As Object, e As EventArgs)
        'Dim DBT As New DataTable
        'Process_List_Line_Grid.Item(Process_List_Line_Grid_Col, Process_List_Line_Grid_Row).Value = Process_List_Line_Grid_Combo.SelectedItem.ToString()
        'Select Case Process_List_Line_Grid_Col
        '    Case 1
        '        StrSQL = "Select PD_Code, PD_Name  FROM SI_MACHINE with(NOLOCK) Where PD_Code = '" & Process_List_Line_Grid_Combo.SelectedItem.ToString() & "'"
        '        Re_Count = DB_Select(DBT)
        '        If Re_Count <> 0 Then
        '            Process_List_Line_Grid.Item(2, Process_List_Line_Grid_Row).Value = DBT.Rows(0).Item(1)
        '        End If
        '    Case 2
        '        StrSQL = "Select PD_Code, PD_Name  FROM SI_MACHINE with(NOLOCK) Where PD_Name = '" & Process_List_Line_Grid_Combo.SelectedItem.ToString() & "'"
        '        Re_Count = DB_Select(DBT)
        '        If Re_Count <> 0 Then
        '            Process_List_Line_Grid.Item(1, Process_List_Line_Grid_Row).Value = DBT.Rows(0).Item(0)
        '        End If
        'End Select

    End Sub

    Private Sub Process_List_Line_Grid_Combo_LostFocus(sender As Object, e As EventArgs)
        'Process_List_Line_Grid_Combo.Visible = False
    End Sub

    Private Sub Save_Line_Click(sender As Object, e As EventArgs)
        If Len(Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value) < 1 Then
            Exit Sub
        End If
        Dim i As Integer
        Grid_Display = False
        For i = 0 To Process_List_Line_Grid.Rows.Count - 1
            If Process_List_Line_Grid.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SI_PROCESS_Device Set PC_PL_CODE = '" & Process_List_Line_Grid.Item(1, i).Value & "', PC_Bigo = '" & Process_List_Line_Grid.Item(3, i).Value & "' where PC_Code = '" & Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value & "' And PC_Sun = '" & Process_List_Line_Grid.Item(0, i).Value & "'"
                Re_Count = DB_Execute()
                Process_List_Line_Grid.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display = True

    End Sub

    Private Sub Insert_Line_Click(sender As Object, e As EventArgs)
        If Len(Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value) < 1 Then
            Exit Sub
        End If
        Dim DBT As New DataTable
        Dim Db_Sun As Long
        StrSQL = "Select PC_Sun FROM SI_Process_Device With(NOLOCK) Where PC_Code = '" & Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value & "' Order By Convert(Decimal,PC_Sun) Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            Db_Sun = 1
        Else
            Db_Sun = DBT.Rows(0).Item(0) + 1
        End If

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO SI_PROCESS_Device (PC_Code, PC_Sun)  Values ('" & Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value & "', '" & Db_Sun & "')"
        Re_Count = DB_Execute()
        Process_List_Line_Grid_Display(Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value)

    End Sub

    Private Sub Put_Line_Click(sender As Object, e As EventArgs)
        If Len(Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value) < 1 Then
            Exit Sub
        End If
        If Len(Process_List_Line_Grid.Item(0, Process_List_Line_Grid.CurrentCell.RowIndex).Value) < 1 Then
            Exit Sub
        End If
        Dim i As Integer
        For i = Process_List_Line_Grid.RowCount - 1 To Process_List_Line_Grid.CurrentCell.RowIndex Step -1

            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "UPDATE SI_PROCESS_Device Set PC_Sun = '" & i + 2 & "' Where PC_Sun = '" & i + 1 & "'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO SI_PROCESS_Device (PC_Code, PC_Sun)  Values ('" & Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value & "', '" & Process_List_Line_Grid.CurrentCell.RowIndex + 1 & "')"
        Re_Count = DB_Execute()

        Process_List_Line_Grid_Display(Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value)

    End Sub

    Private Sub Delete_Line_Click(sender As Object, e As EventArgs)
        '해당 칼럼 삭제

        If Len(Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value) < 1 Then
            Exit Sub
        End If
        If Len(Process_List_Line_Grid.Item(0, Process_List_Line_Grid.CurrentCell.RowIndex).Value) < 1 Then
            Exit Sub
        End If


        If MsgBox("순번 '" & Process_List_Line_Grid.Item(0, Process_List_Line_Grid.CurrentCell.RowIndex).Value & "'을(를) 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "삭제") = vbNo Then
            Exit Sub
        End If

        Dim i As Integer

        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_PROCESS_Device Where PC_Code = '" & Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value & "' And   PC_Sun = '" & Process_List_Line_Grid.Item(0, Process_List_Line_Grid.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        For i = Process_List_Line_Grid.CurrentCell.RowIndex + 1 To Process_List_Line_Grid.RowCount - 1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "UPDATE SI_PROCESS_Device Set PC_Sun = '" & i & "' Where PC_Code = '" & Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value & "' And  PC_Sun = '" & i + 1 & "'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        Process_List_Line_Grid_Display(Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value)

    End Sub

    Private Sub Process_List_Grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
        If Grid_Display = True Then
            Process_List_Grid.CurrentRow.HeaderCell.Value = "*"
        End If
    End Sub

    Private Sub Process_List_Grid_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Process_List_Grid.CellContentDoubleClick
        '유효성 검사
        If Len(Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value) < 0 Then
            MsgBox("해당 행은 공백이므로 처리할 수 없습니다")
            Exit Sub
        End If

        '해당 행 내용 수정
        Process_Insert.Label1.Text = "update"
        Process_Insert.process_code.Text = Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value
        Dim form As New Process_Insert
        form.Label1.Text = "update"
        form.process_code.Text = Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value
        form.ShowDialog()

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If
    End Sub

    Private Sub Insert__Bom_Click(sender As Object, e As EventArgs) Handles Insert__Bom.Click

        If CellClick = False Then
            MsgBox("공정을 선택해주세요")
            Exit Sub
        End If

        Dim form As New Process_Insert
        form.Label1.Text = "insert2"
        form.process_code.Text = Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value
        form.process_name.Text = Process_List_Grid.Item(1, Process_List_Grid.CurrentCell.RowIndex).Value
        form.Process_bigo.Text = Process_List_Grid.Item(2, Process_List_Grid.CurrentCell.RowIndex).Value

        form.ShowDialog()

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If
    End Sub

    Private Sub Delete__Bom_Click(sender As Object, e As EventArgs) Handles Delete__Bom.Click
        If Len(Process_List_Line_Grid.Item(0, Process_List_Line_Grid.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If


        If MsgBox("순번 '" & Process_List_Line_Grid.Item(0, Process_List_Line_Grid.CurrentCell.RowIndex).Value & "'을(를) 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "삭제") = vbNo Then
            Exit Sub
        End If


        'Grid_Display_Ch = False
        Dim i As Integer
        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_MACHINE_QC Where PQ_Code = '" & Process_List_Line_Grid.Item(0, Process_List_Line_Grid.CurrentCell.RowIndex).Value & "' And   PQ_Sun = '" & Process_List_Line_Grid.Item(0, Process_List_Line_Grid.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        For i = Process_List_Line_Grid.CurrentCell.RowIndex + 1 To Process_List_Line_Grid.RowCount - 1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "UPDATE SI_MACHINE_QC Set PQ_Sun = '" & i & "' Where PQ_Code = '" & Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value & "' And  PQ_Sun = '" & i + 1 & "'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        Process_List_Line_Grid_Display(Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value)
        'Grid_Display_Ch = True
    End Sub

    Private Sub Process_List_Grid_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Process_List_Grid.CellClick
        If Len(Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value) > 0 Then
            CellClick = True
            Process_List_Line_Grid_Display(Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value)
        End If
    End Sub
End Class
