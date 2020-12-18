Public Class Frm_Working
    Dim Grid_Display As Boolean
    Dim Grid_Display_Ch As Boolean

    Dim Grid_Main_Row As Integer
    Dim Grid_Main_Col As Integer

    Private Sub Frm_Device2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()
        Process_Line_Grid_Setup()
        Search_Combo.Items.Add("사원코드")
        Search_Combo.Items.Add("사원명")
        Search_Combo.Text = "사원코드"

        ComboBox1.Visible = False
    End Sub
    Public Function Process_Line_Grid_Setup() As Boolean

        Grid_Color(Process_Line_Grid)


        Process_Line_Grid.AllowUserToAddRows = False
        Process_Line_Grid.RowTemplate.Height = 20


        '열의 갯수는 하나 더 많아야 함.
        Process_Line_Grid.ColumnCount = 39
        Process_Line_Grid.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Process_Line_Grid.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Process_Line_Grid.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Process_Line_Grid.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Process_Line_Grid.RowHeadersWidth = 40
        Process_Line_Grid.Columns(0).Width = 60
        Process_Line_Grid.Columns(1).Width = 250
        Process_Line_Grid.Columns(2).Width = 100
        Process_Line_Grid.Columns(3).Width = 100
        Process_Line_Grid.Columns(4).Width = 100
        ' Process_Line_Grid.Columns(5).Width = 200

        Process_Line_Grid.Columns(0).HeaderText = "사원코드"
        Process_Line_Grid.Columns(1).HeaderText = "연월"
        Process_Line_Grid.Columns(2).HeaderText = "구분"
        Process_Line_Grid.Columns(3).HeaderText = "등록시간"
        Process_Line_Grid.Columns(4).HeaderText = "비고"

        Dim J As Integer = 0
        For J = 1 To 31
            Process_Line_Grid.Columns(J + 4).HeaderText = "DAY" & J
        Next

        Process_Line_Grid.Columns(36).HeaderText = "연장시간"
        Process_Line_Grid.Columns(37).HeaderText = "조퇴시간"
        Process_Line_Grid.Columns(38).HeaderText = "특근시간"


        Process_Line_Grid.Columns(0).ReadOnly = True
        Process_Line_Grid.Columns(1).ReadOnly = True

        Process_Line_Grid_Setup = True
    End Function

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        Process_Line_Grid_Display()

    End Sub
    Public Function Process_Line_Grid_Display() As Boolean

        Dim i As Integer
        Dim j As Integer
        Dim DBT As New DataTable
        Grid_Display = False
        Process_Line_Grid.RowCount = 1
        Select Case Search_Combo.Text
            Case "사원코드"
                StrSQL = "SELECT * 
                    FROM MAN_WORKING_LIST A
                   Where MAN_Code Like '%" & Search_Text.Text & "%' Or MAN_Code Is Null Order By MAN_Code"
            Case "출근"
                StrSQL = "SELECT * 
                    FROM MAN_WORKING_LIST A
                   Where MAN_NAME Like '%" & Search_Text.Text & "%' AND MAN_SEL ='출근' Order By MAN_NAME"

            Case "퇴근"
                StrSQL = "SELECT * 
                    FROM MAN_WORKING_LIST A
                   Where MAN_NAME Like '%" & Search_Text.Text & "%' AND MAN_SEL = '퇴근' Order By MAN_NAME"
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
    End Function
    Private Sub Process_Line_Grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Process_Line_Grid.CellValueChanged
        If Grid_Display = True Then
            Process_Line_Grid.CurrentRow.HeaderCell.Value = "*"
        End If

    End Sub

    Private Sub Save_Com_Click(sender As Object, e As EventArgs) Handles Save_Com.Click
        Dim i As Integer
        Dim DBT As New DataTable
        Grid_Display = False
        For i = 0 To Process_Line_Grid.Rows.Count - 1
            If Process_Line_Grid.Rows(i).HeaderCell.Value = "*" Then


                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "SELECT MAN_CODE, MAN_YMONTH FROM MAN_WORKING_LIST
                                   where MAN_Code = '" & Process_Line_Grid.Item(0, i).Value & "' AND MAN_YMONTH = '" & Process_Line_Grid.Item(1, i).Value & "' AND MAN_SEL = '" & Process_Line_Grid.Item(2, i).Value & "'"
                Re_Count = DB_Select(DBT)

                If Re_Count <> 0 Then
                    '업데이트
                    'MAN_WORKING_LISTUpdate
                    StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                    StrSQL = StrSQL & "UPDATE MAN_WORKING_LIST Set
                                
                                        MAN_TIME = '" & DateTime.Now.ToString("HH:mm") & "'
                                      , MAN_BIGO = '" & Process_Line_Grid.Item(4, i).Value & "'
                                      , DAY1 = '" & Process_Line_Grid.Item(5, i).Value & "'
                                      , DAY2 = '" & Process_Line_Grid.Item(6, i).Value & "'
                                      , DAY3 = '" & Process_Line_Grid.Item(7, i).Value & "'
                                      , DAY4 = '" & Process_Line_Grid.Item(8, i).Value & "'
                                      , DAY5 = '" & Process_Line_Grid.Item(9, i).Value & "'
                                      , DAY6 = '" & Process_Line_Grid.Item(10, i).Value & "'
                                      , DAY7 = '" & Process_Line_Grid.Item(11, i).Value & "'
                                      , DAY8 = '" & Process_Line_Grid.Item(12, i).Value & "'
                                      , DAY9 = '" & Process_Line_Grid.Item(13, i).Value & "'
                                      , DAY10 = '" & Process_Line_Grid.Item(14, i).Value & "'
                                      , DAY11 = '" & Process_Line_Grid.Item(15, i).Value & "'
                                      , DAY12 = '" & Process_Line_Grid.Item(16, i).Value & "'
                                      , DAY13 = '" & Process_Line_Grid.Item(17, i).Value & "'

                                      , DAY14 = '" & Process_Line_Grid.Item(18, i).Value & "'
                                      , DAY15 = '" & Process_Line_Grid.Item(19, i).Value & "'
                                      , DAY16 = '" & Process_Line_Grid.Item(20, i).Value & "'
                                      , DAY17 = '" & Process_Line_Grid.Item(21, i).Value & "'
                                      , DAY18 = '" & Process_Line_Grid.Item(22, i).Value & "'
                                      , DAY19 = '" & Process_Line_Grid.Item(23, i).Value & "'
                                      , DAY20 = '" & Process_Line_Grid.Item(24, i).Value & "'
                                      , DAY21 = '" & Process_Line_Grid.Item(25, i).Value & "'
                                      , DAY22 = '" & Process_Line_Grid.Item(26, i).Value & "'

                                      , DAY23 = '" & Process_Line_Grid.Item(27, i).Value & "'
                                      , DAY24 = '" & Process_Line_Grid.Item(28, i).Value & "'
                                      , DAY25 = '" & Process_Line_Grid.Item(29, i).Value & "'
                                      , DAY26 = '" & Process_Line_Grid.Item(30, i).Value & "'
                                      , DAY27 = '" & Process_Line_Grid.Item(31, i).Value & "'
                                      , DAY28 = '" & Process_Line_Grid.Item(32, i).Value & "'
                                      , DAY29 = '" & Process_Line_Grid.Item(33, i).Value & "'
                                      , DAY30 = '" & Process_Line_Grid.Item(34, i).Value & "'
                                      , DAY31 = '" & Process_Line_Grid.Item(35, i).Value & "'

                                      , MAN_CONTINUE = '" & Process_Line_Grid.Item(36, i).Value & "'
                                      , MAN_EARLY = '" & Process_Line_Grid.Item(37, i).Value & "'
                                      , MAN_WEEK = '" & Process_Line_Grid.Item(38, i).Value & "'
                                        where MAN_Code = '" & Process_Line_Grid.Item(0, i).Value & "' And  MAN_YMONTH = '" & Process_Line_Grid.Item(1, i).Value & "'
                                      and MAN_SEL = '" & Process_Line_Grid.Item(2, i).Value & "'"

                    Re_Count = DB_Execute()

                Else
                    '추가
                    StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                    StrSQL = StrSQL & "INSERT INTO MAN_WORKING_LIST VALUES
                                        ( '" & Process_Line_Grid.Item(0, i).Value & "',
                                          '" & DateTime.Now.ToString("yyyy-MM") & "',
                                         '" & Process_Line_Grid.Item(2, i).Value & "',
                                         '" & DateTime.Now.ToString("HH:mm") & "',
                                         '" & Process_Line_Grid.Item(4, i).Value & "',
                                         '" & Process_Line_Grid.Item(5, i).Value & "',
                                         '" & Process_Line_Grid.Item(6, i).Value & "',
                                         '" & Process_Line_Grid.Item(7, i).Value & "',
                                         '" & Process_Line_Grid.Item(8, i).Value & "',
                                         '" & Process_Line_Grid.Item(9, i).Value & "',
                                         '" & Process_Line_Grid.Item(10, i).Value & "',
                                         '" & Process_Line_Grid.Item(11, i).Value & "',
                                         '" & Process_Line_Grid.Item(12, i).Value & "',
                                         '" & Process_Line_Grid.Item(13, i).Value & "',
                                         '" & Process_Line_Grid.Item(14, i).Value & "',
                                         '" & Process_Line_Grid.Item(15, i).Value & "',
                                         '" & Process_Line_Grid.Item(16, i).Value & "',
                                         '" & Process_Line_Grid.Item(17, i).Value & "',
                                         '" & Process_Line_Grid.Item(18, i).Value & "',
                                         '" & Process_Line_Grid.Item(19, i).Value & "',
                                         '" & Process_Line_Grid.Item(20, i).Value & "',
                                         '" & Process_Line_Grid.Item(21, i).Value & "',
                                          '" & Process_Line_Grid.Item(22, i).Value & "',
                                         '" & Process_Line_Grid.Item(23, i).Value & "',
                                         '" & Process_Line_Grid.Item(24, i).Value & "',
                                         '" & Process_Line_Grid.Item(25, i).Value & "',
                                         '" & Process_Line_Grid.Item(26, i).Value & "',
                                         '" & Process_Line_Grid.Item(27, i).Value & "',
                                         '" & Process_Line_Grid.Item(28, i).Value & "',
                                         '" & Process_Line_Grid.Item(29, i).Value & "',
                                         '" & Process_Line_Grid.Item(30, i).Value & "',
                                         '" & Process_Line_Grid.Item(31, i).Value & "',
                                         '" & Process_Line_Grid.Item(32, i).Value & "',
                                         '" & Process_Line_Grid.Item(33, i).Value & "',
                                         '" & Process_Line_Grid.Item(34, i).Value & "',
                                         '" & Process_Line_Grid.Item(35, i).Value & "',
                                         '" & Process_Line_Grid.Item(36, i).Value & "',
                                         '" & Process_Line_Grid.Item(37, i).Value & "',
                                        '" & Process_Line_Grid.Item(38, i).Value & "')"


                    Re_Count = DB_Execute()
                End If

                Process_Line_Grid.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display = True
        MsgBox("저장되었습니다")
        Search_Com.PerformClick()

    End Sub

    Private Sub Insert_Com_Click(sender As Object, e As EventArgs) Handles Insert_Com.Click
        ' Frm_Insert.Text = "근태사원 코드 추가"
        ' Frm_Insert.Show()
        ' Frm_Main.Enabled = False
        Process_Line_Grid.Rows.Add()

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
        StrSQL = StrSQL & "DELETE MAN_WORKING_LIST where MAN_Code = '" & Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value & "' AND MAN_YMONTH = '" & Process_Line_Grid.Item(1, Process_Line_Grid.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        'StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        'StrSQL = StrSQL & "DELETE MAN_WORKING where MAN_Code = '" & Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value & "' AND MAN_YMONTH = '" & Process_Line_Grid.Item(1, Process_Line_Grid.CurrentCell.RowIndex).Value & "'"
        'Re_Count = DB_Execute()

        Process_Line_Grid_Display()
        MsgBox("삭제 되었습니다.")
    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
        Dim DBT As New DataTable
        Process_Line_Grid.Item(Grid_Main_Col, Grid_Main_Row).Value = ComboBox1.SelectedItem.ToString()
        Select Case Grid_Main_Col
            Case 0
                StrSQL = "Select MAN_CODE FROM SI_MAN with(NOLOCK) Where MAN_Code = '" & ComboBox1.SelectedItem.ToString() & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Process_Line_Grid.Item(0, Grid_Main_Row).Value = DBT.Rows(0).Item(0)
                End If

        End Select

        ComboBox1.Visible = False
    End Sub

    Private Sub Process_Line_Grid_MouseClick(sender As Object, e As MouseEventArgs) Handles Process_Line_Grid.MouseClick


        Grid_Main_Row = Process_Line_Grid.CurrentCell.RowIndex
        Grid_Main_Col = Process_Line_Grid.CurrentCell.ColumnIndex
        ComboBox1.Visible = False
        If Grid_Main_Row < 0 Then
            Exit Sub
        End If
        If Grid_Main_Col < 0 Then
            Exit Sub
        End If
        Dim DBT As New DataTable

        Select Case Grid_Main_Col
            Case 0
                StrSQL = "Select MAN_CODE FROM SI_MAN "
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Process_Line_Grid.Item(0, Grid_Main_Row).Value = DBT.Rows(0).Item(0)
                End If
                Re_Count = DB_Select(DBT)
                ComboBox1.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        ComboBox1.Items.Add(DBT.Rows(i).Item(0))
                    Next i
                End If
                ComboBox1.Top = Process_Line_Grid.Top + Process_Line_Grid.ColumnHeadersHeight + (Grid_Main_Row) * 24
                ComboBox1.Left = Process_Line_Grid.Left + Process_Line_Grid.RowHeadersWidth
                ComboBox1.Width = Process_Line_Grid.Columns(Grid_Main_Col).Width
                '   ComboBox1.Text = Process_Line_Grid.CurrentCell.Value.ToString
                ComboBox1.BackColor = Process_Line_Grid.RowsDefaultCellStyle.BackColor
                ComboBox1.Visible = True
                ComboBox1.Focus.ToString()


        End Select
        Grid_Display = True
    End Sub

End Class
