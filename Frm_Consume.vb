Public Class Frm_Consume
    Dim Grid_Display As Boolean
    Dim Process_List_Line_Grid_Row As Integer
    Dim Process_List_Line_Grid_Col As Integer

    Private Sub Frm_Process_List_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()
        Process_List_Grid_Setup()
        Search_Combo.Items.Add("예비품코드")
        Search_Combo.Items.Add("예비품명")
        Search_Combo.Text = "예비품코드"
        'Line_Panel.Visible = False
        Search_Com.PerformClick()
        Process_List_Grid.ClearSelection()
    End Sub

    Public Function Process_List_Grid_Setup() As Boolean
        Grid_Color(Process_List_Grid)


        Process_List_Grid.AllowUserToAddRows = False
        Process_List_Grid.RowTemplate.Height = 20


        '열의 갯수는 하나 더 많아야 함.
        Process_List_Grid.ColumnCount = 6
        Process_List_Grid.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Process_List_Grid.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Process_List_Grid.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Process_List_Grid.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Process_List_Grid.RowHeadersWidth = 40
        Process_List_Grid.Columns(0).Width = 80
        Process_List_Grid.Columns(1).Width = 100
        Process_List_Grid.Columns(2).Width = 120
        Process_List_Grid.Columns(3).Width = 230
        Process_List_Grid.Columns(4).Width = 80
        Process_List_Grid.Columns(5).Width = 230



        Process_List_Grid.Columns(0).HeaderText = "예비품코드"
        Process_List_Grid.Columns(1).HeaderText = "예비품명"
        Process_List_Grid.Columns(2).HeaderText = "예비품 구매일"
        Process_List_Grid.Columns(3).HeaderText = "내용"
        Process_List_Grid.Columns(4).HeaderText = "담당자"
        Process_List_Grid.Columns(5).HeaderText = "비고"

        Process_List_Grid.Columns(0).ReadOnly = True
        Process_List_Grid_Setup = True
    End Function

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        Process_List_Grid_Display()
    End Sub
    Public Function Process_List_Grid_Display() As Boolean

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display = False
        Process_List_Grid.RowCount = 1
        Select Case Search_Combo.Text
            Case "예비품코드"
                StrSQL = "Select * FROM FC_PART with(NOLOCK) Where FL_Code Like '%" & Search_Text.Text & "%' Or FL_Code Is Null Order By FL_Code"
            Case "예비품명"
                StrSQL = "Select * FROM FC_PART with(NOLOCK) Where FL_Name Like '%" & Search_Text.Text & "%' Or FL_Name Is Null Order By FL_Name"
        End Select
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



    End Function

    Private Sub Process_List_Grid_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Process_List_Grid.CellContentClick

    End Sub



    Private Sub Save_Com_Click(sender As Object, e As EventArgs) Handles Save_Com.Click
        Dim i As Integer
        Grid_Display = False
        For i = 0 To Process_List_Grid.Rows.Count - 1
            If Process_List_Grid.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE FC_PART Set FL_Name = '" & Process_List_Grid.Item(1, i).Value & "', FL_Day = '" & Process_List_Grid.Item(2, i).Value & "' , FL_Content = '" & Process_List_Grid.Item(3, i).Value & "' , FL_Man = '" & Process_List_Grid.Item(4, i).Value & "' , FL_Bigo = '" & Process_List_Grid.Item(5, i).Value & "' where FL_Code = '" & Process_List_Grid.Item(0, i).Value & "'"
                Re_Count = DB_Execute()
                Process_List_Grid.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display = True
        MsgBox("저장되었습니다")
    End Sub

    Private Sub Insert_Com_Click(sender As Object, e As EventArgs) Handles Insert_Com.Click
        'DB를 추가 한다.
        '코드 번호 입력
        '새로운 푬에서 입력 받는다.
        Frm_Insert.Text = "예비품코드 추가"
        Frm_Insert.Show()
        Frm_Main.Enabled = False



    End Sub

    Private Sub Delete_Com_Click(sender As Object, e As EventArgs) Handles Delete_Com.Click

        If Len(Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value) < 1 Then
            MsgBox("예비품코드를 선택하세요.")
            Exit Sub
        End If
        If MsgBox("예비품코드  '" & Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value & "'를 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "예비품코드 삭제") = vbNo Then
            Exit Sub
        End If

        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE FC_PART where FL_Code = '" & Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()
        Grid_Display = False
        Process_List_Grid_Display()
        MsgBox("삭제 되었습니다.")
        Grid_Display = True

    End Sub

    Private Sub Process_List_Line_Grid_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub Process_List_Grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Process_List_Grid.CellValueChanged
        If Grid_Display = True Then
            Process_List_Grid.CurrentRow.HeaderCell.Value = "*"
        End If
    End Sub

End Class
