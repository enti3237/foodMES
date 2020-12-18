Public Class popup_process
    'Dim sel_Row As Integer
    'Dim sel_Col As Integer

    Public PROCESS_Name As String '품명
    Public PROCESS_Code As String '코드


    Private Sub X_Click(sender As Object, e As EventArgs) Handles X.Click
        Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub viewbox1_TextChanged(sender As Object, e As EventArgs)

    End Sub


    Public Function DataGridView1_Setup() As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer

        Grid_Color(DataGridView1)

        DataGridView1.AllowUserToAddRows = False
        DataGridView1.RowTemplate.Height = 24
        DataGridView1.ColumnCount = 3
        DataGridView1.RowCount = 0

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft


        DataGridView1.RowHeadersWidth = 10
        DataGridView1.Columns(0).Width = 60

        DataGridView1.Columns(0).HeaderText = "공정코드"
        DataGridView1.Columns(0).Visible = False
        DataGridView1.Columns(1).HeaderText = "공정명"
        DataGridView1.Columns(2).HeaderText = "비고"

        DataGridView1.Columns(0).Width = 140
        DataGridView1.Columns(1).Width = 100

        StrSQL = "select PC_CODE, PC_NAME, PC_BIGO from SI_PROCESS "

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

        DataGridView1_Setup = True
    End Function

    
    '더블클릭시 값이 전해지도록 변경
    Private Sub DataGridView1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles DataGridView1.MouseDoubleClick

        PROCESS_Code = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value.ToString()
        PROCESS_Name = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value.ToString()

        DialogResult = DialogResult.OK
        Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer

        StrSQL = "select PC_CODE, PC_NAME, PC_BIGO from SI_PROCESS WHERE PC_Name Like'%" + viewbox1.Text + "%' OR PC_CODE Like'%" + viewbox1.Text + "%' "

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
    End Sub

    Private Sub popup_process_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        DataGridView1_Setup()
        Button3.PerformClick()
    End Sub

    'Private Sub Button1_Click(sender As Object, e As EventArgs)
    '    clear()
    'End Sub
End Class