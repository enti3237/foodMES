Public Class Frm_Temper
    Dim Grid_Display As Boolean
    Dim Grid_Display_Ch As Boolean

    Dim Grid_Bom_Row As Integer
    Dim Grid_Bom_Col As Integer

    Private Sub Frm_Device2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()
        Process_Line_Grid_Setup()
        Setup()
        S_Text_Setup()


        Search_Combo.Items.Add("전체")
        Search_Combo.Text = "전체"

        Insert_Com.Visible = False
        Delete_Com.Visible = False
        Save_Com.Visible = False

        DataGridView1.Visible = False
        Search_Combo.Visible = False
        Search_Text.Visible = False

        Process_Line_Grid_Display()
        'Search_Com.PerformClick()
        Process_Line_Grid.ClearSelection()
    End Sub

    Public Function S_Text_Setup() As Boolean
        E_Day.Text = Format(Now, "yyyy-MM-dd")
        '  S_Day.Text = Format(Now, "yyyyMM01")
        S_Day.Text = Mid(E_Day.Text, 1, 8) & "01"
        S_Time.Text = "0000"
        E_Time.Text = "2359"

        S_Text_Setup = True

    End Function

    Public Function Process_Line_Grid_Setup() As Boolean

        Grid_Color(Process_Line_Grid)


        Process_Line_Grid.AllowUserToAddRows = False
        Process_Line_Grid.RowTemplate.Height = 20


        '열의 갯수는 하나 더 많아야 함.
        Process_Line_Grid.ColumnCount = 2
        Process_Line_Grid.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Process_Line_Grid.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Process_Line_Grid.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        '  Process_Line_Grid.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Process_Line_Grid.RowHeadersWidth = 40
        Process_Line_Grid.Columns(0).Width = 40
        Process_Line_Grid.Columns(1).Width = 40
        '   Process_Line_Grid.Columns(2).Width = 50



        ' Process_Line_Grid.Columns(0).HeaderText = "코드"
        Process_Line_Grid.Columns(0).HeaderText = "창고명"
        Process_Line_Grid.Columns(1).HeaderText = "기준온도"
        ' Process_Line_Grid.Columns(2).HeaderText = "실제온도"
        ' Process_Line_Grid.Columns(3).HeaderText = "내역보기"

        Process_Line_Grid.Columns(0).ReadOnly = True


        Process_Line_Grid_Setup = True
    End Function

    Public Function Setup() As Boolean

        Grid_Color(DataGridView1)


        DataGridView1.AllowUserToAddRows = False
        DataGridView1.RowTemplate.Height = 20


        '열의 갯수는 하나 더 많아야 함.
        DataGridView1.ColumnCount = 5
        DataGridView1.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        DataGridView1.RowHeadersWidth = 80
        DataGridView1.Columns(0).Width = 80
        DataGridView1.Columns(1).Width = 80
        DataGridView1.Columns(2).Width = 100

        ' datagridview1.Columns(0).HeaderText = "코드"
        DataGridView1.Columns(0).HeaderText = "창고명"
        DataGridView1.Columns(1).HeaderText = "기준온도"
        DataGridView1.Columns(2).HeaderText = "실제온도"
        DataGridView1.Columns(3).HeaderText = "일자"
        DataGridView1.Columns(4).HeaderText = "시간"

        ' datagridview1.Columns(3).HeaderText = "내역보기"

        DataGridView1.Columns(0).ReadOnly = True


        '  datagridview1_Setup = True
    End Function

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        'Process_Line_Grid_Display()
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display = False
        DataGridView1.RowCount = 1

        'StrSQL = "Select TOP 10000 FL_NAME, FL_ST, C_DATA, N_TIME, C_TIME
        '                    FROM SI_STOCK with(NOLOCK) join iotlog b on SI_STOCK.FL_IP = b.C_IP 
        '                   Where FL_NAME = '" & Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value & "' 
        '                     AND C_DATE BETWEEN '" & S_Day.Text & "' and '" & E_Day.Text & "'
        '                     AND C_TIME BETWEEN '" & S_Time.Text & "%' and '" & E_Time.Text & "'
        '                    Order By N_TIME DESC, C_TIME DESC"

        StrSQL = "Select FL_NAME, FL_ST, C_DATA, N_TIME, C_TIME
                    FROM SI_STOCK with(NOLOCK) join iotlog b on SI_STOCK.FL_IP = b.C_IP 
                   Where FL_NAME = '" & Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value & "' 
                           AND N_TIME BETWEEN '" & S_Day.Text & "' and '" & E_Day.Text & "'
                     
                            Order By N_TIME DESC, C_TIME DESC"

        '        AND C_TIME LIKE '" & S_Time.Text.Substring(0, 2) & "%' OR C_TIME LIKE'" & E_Time.Text.Substring(0, 2) & "%'

        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            DataGridView1.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To DBT.Columns.Count - 1
                    DataGridView1.Item(j, i).Value = DBT.Rows(i).Item(j)
                Next j
            Next i
        Else
            DataGridView1.RowCount = 1
            For j = 0 To DataGridView1.ColumnCount - 1
                DataGridView1.Item(j, 0).Value = ""
            Next j

        End If
    End Sub
    Public Function Process_Line_Grid_Display() As Boolean

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display = False
        Process_Line_Grid.RowCount = 1
        Select Case Search_Combo.Text
            Case "전체"
                StrSQL = "Select FL_NAME, FL_ST 
                            FROM SI_STOCK with(NOLOCK) Order By FL_CODE"
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


        '실제 온도


        'StrSQL = "Select top 1 c_data
        '            FROM iotlog with(NOLOCK)
        '            where
        '            Order By c_date desc, c_time desc"

        'Re_Count = DB_Select(DBT)
        'If Re_Count <> 0 Then
        '    Process_Line_Grid.RowCount = Re_Count
        '    For i = 0 To 1
        '        For j = 0 To 1
        '            Process_Line_Grid.Item(j, i).Value = DBT.Rows(i).Item(j)
        '        Next j
        '    Next i
        'Else
        '    Process_Line_Grid.RowCount = 1
        '    For j = 0 To Process_Line_Grid.ColumnCount - 1
        '        Process_Line_Grid.Item(j, 0).Value = ""
        '    Next j

        'End If
        'Dim a As Integer

        'For a = 0 To Process_Line_Grid.Rows.Count - 1
        '    Process_Line_Grid.Item(4, a).Value = Val(Process_Line_Grid.Item(2, a).Value) - Val(Process_Line_Grid.Item(3, a).Value)
        '    '  Grid_Main.Item(9, Grid_Main.CurrentCell.RowIndex).Value = Val(Grid_Main.Item(6, Grid_Main.CurrentCell.RowIndex).Value) + Val(Grid_Main.Item(7, Grid_Main.CurrentCell.RowIndex).Value) + Val(Grid_Main.Item(8, Grid_Main.CurrentCell.RowIndex).Value)
        'Next

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
                StrSQL = StrSQL & "UPDATE SI_STOCK Set
                                     FL_AD = '" & Process_Line_Grid.Item(4, i).Value & "'
                                   where FL_CODE= '" & Process_Line_Grid.Item(0, i).Value & "'"
                Re_Count = DB_Execute()
                Process_Line_Grid.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display = True
        MsgBox("저장되었습니다")
    End Sub

    Private Sub Insert_Com_Click(sender As Object, e As EventArgs) Handles Insert_Com.Click

        Frm_Insert.Text = "생산라인 코드 추가"
        Frm_Insert.Show()
        Frm_Main.Enabled = False
    End Sub

    Private Sub Delete_Com_Click(sender As Object, e As EventArgs) Handles Delete_Com.Click
        If Len(Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value) < 1 Then
            MsgBox("코드를 선택하세요.")
            Exit Sub
        End If
        If MsgBox("코드  '" & Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value & "'를 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "라인코드 삭제") = vbNo Then
            Exit Sub
        End If
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_STOCK where FL_CODE = '" & Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()
        Process_Line_Grid_Display()
        MsgBox("삭제 되었습니다.")
    End Sub

    Private Sub Panel_Menu_Paint(sender As Object, e As PaintEventArgs) Handles Panel_Menu.Paint

    End Sub

    Private Sub Process_Line_Grid_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Process_Line_Grid.CellClick

        Dim row As Integer
        Dim DBT As New DataTable

        Dim i, j As Integer

        row = e.RowIndex

        If e.ColumnIndex = 1 Then

            DataGridView1.Visible = True

            StrSQL = "Select TOP 50 FL_NAME, FL_ST, N_TIME, C_TIME, C_DATA
                            FROM SI_STOCK with(NOLOCK) join iotlog b on SI_STOCK.FL_IP = b.C_IP 
                           Where FL_NAME = '" & Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value & "' Order By N_TIME DESC, C_TIME DESC"


            Re_Count = DB_Select(DBT)

            If Re_Count <> 0 Then
                DataGridView1.RowCount = Re_Count


                For i = 0 To Re_Count - 1
                    If IsDBNull(DBT.Rows(i)("FL_NAME")) Then
                        MsgBox("값이 없습니다")
                        Continue For
                    End If
                    DataGridView1.Item(0, i).Value = DBT.Rows(i)("FL_NAME")
                    DataGridView1.Item(1, i).Value = DBT.Rows(i)("FL_ST")
                    DataGridView1.Item(3, i).Value = DBT.Rows(i)("N_TIME")
                    DataGridView1.Item(4, i).Value = DBT.Rows(i)("C_TIME")
                    DataGridView1.Item(2, i).Value = DBT.Rows(i)("C_DATA")
                Next i
            End If
        End If


        'If Re_Count <> 0 Then
        '    DataGridView1.RowCount = Re_Count
        '    For i = 0 To Re_Count - 1
        '        For j = 0 To DBT.Columns.Count - 1
        '            DataGridView1.Item(j, i).Value = DBT.Rows(i).Item(j)
        '            If j = 4 Then
        '                DataGridView1.Item(j, i).Value = "내역보기"
        '            End If
        '        Next j
        '    Next i
        'Else
        '    DataGridView1.RowCount = 1
        '    For j = 0 To DataGridView1.ColumnCount - 1
        '        DataGridView1.Item(j, 0).Value = ""
        '    Next j

        'End If
    End Sub
End Class
