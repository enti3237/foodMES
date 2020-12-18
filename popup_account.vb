Public Class popup_account

    Public Customer_Name As String
    Public Customer_Man_Name As String
    Public Customer_Man_Tel As String
    Public Customer_Code As String

    Private Sub popup_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1_Setup()
    End Sub

    Private Sub X_Click(sender As Object, e As EventArgs) Handles X.Click
        Close()
    End Sub
    Public Function DataGridView1_Setup() As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Color(DataGridView1)

        DataGridView1.AllowUserToAddRows = False
        DataGridView1.RowTemplate.Height = 24
        DataGridView1.ColumnCount = 4
        DataGridView1.RowCount = 0

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        DataGridView1.RowHeadersWidth = 10
        DataGridView1.Columns(0).Width = 60

        DataGridView1.Columns(0).HeaderText = "거래처명"
        DataGridView1.Columns(1).HeaderText = "담당자명"
        DataGridView1.Columns(2).HeaderText = "담당자번호"
        DataGridView1.Columns(3).HeaderText = "거래처코드"

        DataGridView1.Columns(0).Width = 210
        DataGridView1.Columns(1).Width = 150
        DataGridView1.Columns(2).Width = 220
        DataGridView1.Columns(3).Width = 220

        DataGridView1.Columns(3).Visible = False

        StrSQL = "select CM_Name, CM_MAN_NAME, CM_Man_Tel, CM_Code from SI_CUSTOMER "

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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer

        StrSQL = "select CM_CODE, CM_NAME from SI_CUSTOMER where CM_Name Like '%" + viewbox1.Text + "%' "

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
        viewbox1.Text = ""
    End Sub

    Private Sub DataGridView1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles DataGridView1.MouseDoubleClick
        Customer_Name = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value
        Customer_Man_Name = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value
        Customer_Man_Tel = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(2).Value
        Customer_Code = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(3).Value

        DialogResult = DialogResult.OK
        Close()
    End Sub
    'Private Sub Button1_Click(sender As Object, e As EventArgs)
    'clear()
    'End Sub
End Class