Public Class popup_FL
    'Dim sel_Row As Integer
    'Dim sel_Col As Integer

    Public Code As String '품명
    Public Name As String '구분
    Public ID As String '규격
    Public Port As String '단위
    Public Bigo As String '단가


    Private Sub popup_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        DataGridView1_Setup()
        'textbox1.Items.Add("전체")
        'textbox1.Items.Add("반제품")
        'textbox1.Items.Add("완제품")
        'TextBox1.Text = "원부재료"


    End Sub

    Private Sub X_Click(sender As Object, e As EventArgs) Handles X.Click
        Close()
    End Sub

    Private Sub textbox1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub viewbox1_TextChanged(sender As Object, e As EventArgs) Handles viewbox1.TextChanged

    End Sub


    Public Function DataGridView1_Setup() As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer

        Grid_Color(DataGridView1)

        DataGridView1.AllowUserToAddRows = False
        DataGridView1.RowTemplate.Height = 24
        DataGridView1.ColumnCount = 5
        DataGridView1.RowCount = 0

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft


        DataGridView1.RowHeadersWidth = 10
        DataGridView1.Columns(0).Width = 60

        DataGridView1.Columns(0).HeaderText = "설비코드"
        DataGridView1.Columns(1).HeaderText = "설비명"
        DataGridView1.Columns(2).HeaderText = "IP"
        DataGridView1.Columns(3).HeaderText = "Port"
        DataGridView1.Columns(4).HeaderText = "비고"


        DataGridView1.Columns(0).Width = 140
        DataGridView1.Columns(1).Width = 100
        DataGridView1.Columns(2).Width = 100

        StrSQL = "Select * FROM SI_MACHINE"

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

        'StrSQL = "select PL_Name from SI_PRODUCT where PL_Name Like'%" + viewbox1.Text + "%' OR PL_CODE Like'%" + viewbox1.Text + "%' "

        'StrSQL = "select PL_Name, PL_Sel2, PL_Unit, PL_Unit_Gold, PL_Code from SI_PRODUCT where PL_Name Like'%" + viewbox1.Text + "%' UNION
        'Select PL_Name, PL_Sel2, PL_Unit, PL_Unit_Gold, PL_Code from SI_PRODUCT where PL_CODE Like'%" + viewbox1.Text + "%' "
        If viewbox1.Text = "" Then
            StrSQL = "Select * FROM SI_MACHINE"
        Else
            StrSQL = "Select * FROM SI_MACHINE where PD_Name='" & viewbox1.Text & "' "

        End If



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

    '더블클릭시 값이 전해지도록 변경
    Private Sub DataGridView1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles DataGridView1.MouseDoubleClick

        Code = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value.ToString()
        Name = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value.ToString()
        ID = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(2).Value.ToString()
        Port = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(3).Value.ToString()
        Bigo = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(4).Value.ToString()

        DialogResult = DialogResult.OK
        Close()
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs)
    End Sub

    'Private Sub Button1_Click(sender As Object, e As EventArgs)
    '    clear()
    'End Sub
End Class