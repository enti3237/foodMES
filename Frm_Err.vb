Public Class Frm_Err
    Dim Grid_Display As Boolean
    Dim Process_List_Line_Grid_Row As Integer
    Dim Process_List_Line_Grid_Col As Integer

    Private Sub Frm_Process_List_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()
        Process_List_Grid_Setup()
        Search_Combo.Items.Add("순번")
        Search_Combo.Items.Add("장소")
        Search_Combo.Text = "장소"
        'Line_Panel.Visible = False
        Search_Com.PerformClick()
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
        '  Process_List_Grid.Columns(5).Width = 230



        Process_List_Grid.Columns(0).HeaderText = "순번"
        Process_List_Grid.Columns(1).HeaderText = "장소"
        Process_List_Grid.Columns(2).HeaderText = "기준온도"
        Process_List_Grid.Columns(3).HeaderText = "범위온도"
        Process_List_Grid.Columns(4).HeaderText = "내용"
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
            Case "순번"
                StrSQL = "Select * FROM SI_ERR with(NOLOCK) Where ER_CODE Like '%" & Search_Text.Text & "%' Or ER_CODE Is Null Order By ER_CODE "
            Case "장소"
                StrSQL = "Select * FROM SI_ERR with(NOLOCK) Where ER_Place Like '%" & Search_Text.Text & "%' Or ER_Place Is Null Order By ER_Place"
        End Select
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Process_List_Grid.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To 5
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
                StrSQL = StrSQL & "UPDATE SI_ERR Set ER_Place = '" & Process_List_Grid.Item(1, i).Value & "', ER_TEMPER = '" & Process_List_Grid.Item(2, i).Value & "' , ER_RANGE = '" & Process_List_Grid.Item(3, i).Value & "' , ER_TIME = '" & Process_List_Grid.Item(4, i).Value & "' , ER_BIGO = '" & Process_List_Grid.Item(5, i).Value & "' where ER_CODE = '" & Process_List_Grid.Item(0, i).Value & "'"
                Re_Count = DB_Execute()
                Process_List_Grid.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display = True

    End Sub

    Private Sub Insert_Com_Click(sender As Object, e As EventArgs) Handles Insert_Com.Click
        Dim DBT As New DataTable
        Dim JP_Code As Long

        StrSQL = "Select ER_CODE FROM SI_ERR with(NOLOCK)  Order By Convert(Decimal,ER_CODE) Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            JP_Code = JP_Code & "001"
        Else
            JP_Code = Val(DBT.Rows(0).Item(0)) + 1
        End If

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO SI_ERR (ER_CODE)  Values ('" & JP_Code & "')"
        Re_Count = DB_Execute()

        Search_Com.PerformClick()

    End Sub

    Private Sub Delete_Com_Click(sender As Object, e As EventArgs) Handles Delete_Com.Click

        If Len(Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value) < 1 Then
            MsgBox("알림코드를 선택하세요.")
            Exit Sub
        End If
        If MsgBox("알림코드  '" & Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value & "'를 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "알림코드 삭제") = vbNo Then
            Exit Sub
        End If

        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_ERR where ER_CODE = '" & Process_List_Grid.Item(0, Process_List_Grid.CurrentCell.RowIndex).Value & "'"
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
