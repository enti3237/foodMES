Public Class popup_order

    Public MO_CODE, MO_CM_CODE, MO_CM_NAME, MO_PL_CODE, MO_PL_NAME, MO_STANDARD, MO_UNIT_AMOUNT, MO_QTY, MO_TOTAL, MO_GOLD As String

    Private Sub popup_order_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        DataGridView2_Setup()
        DataGridView2_Display2()

    End Sub

    Public Function DataGridView2_Setup() As Boolean
        Grid_Color(DataGridView2)

        DataGridView2.AllowUserToAddRows = False
        DataGridView2.RowTemplate.Height = 24
        DataGridView2.ColumnCount = 10
        DataGridView2.RowCount = 0




        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        DataGridView2.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView2.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView2.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView2.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView2.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView2.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView2.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView2.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        DataGridView2.RowHeadersWidth = 10
        DataGridView2.Columns(0).Width = 60


        DataGridView2.Columns(0).HeaderText = "발주코드"
        DataGridView2.Columns(1).HeaderText = "거래처코드"
        DataGridView2.Columns(1).Visible = False
        DataGridView2.Columns(2).HeaderText = "거래처명"
        DataGridView2.Columns(3).HeaderText = "원부재료코드"
        DataGridView2.Columns(3).Visible = False
        DataGridView2.Columns(4).HeaderText = "원부재료명"
        DataGridView2.Columns(5).HeaderText = "규격"
        DataGridView2.Columns(6).HeaderText = "단가"
        DataGridView2.Columns(7).HeaderText = "수량"
        DataGridView2.Columns(8).HeaderText = "금액"
        DataGridView2.Columns(8).Visible = False
        DataGridView2.Columns(9).HeaderText = "부가세"
        DataGridView2.Columns(9).Visible = False

        'DataGridView2.Columns(0).Width = 20
        'DataGridView2.Columns(1).Width = 60
        'DataGridView2.Columns(2).Width = 40
        'DataGridView2.Columns(3).Width = 60
        'DataGridView2.Columns(4).Width = 60
        'DataGridView2.Columns(5).Width = 60
        'DataGridView2.Columns(6).Width = 60
        'DataGridView2.Columns(7).Width = 40

        DataGridView2_Setup = True


    End Function


    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer

        StrSQL = "select MO_CODE, MO_CM_CODE, MO_CM_NAME, MO_PL_CODE, MO_PL_NAME, MO_STANDARD, MO_UNIT_AMOUNT, MO_QTY, MO_TOTAL, MO_GOLD
                    FROM MC_STOCK_ORDER where MO_PL_Name Like '%" + TextBox2.Text + "%' AND MO_ING='진행중'"

        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            DataGridView2.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To DataGridView2.ColumnCount - 1
                    DataGridView2.Item(j, i).Value = DBT.Rows(i).Item(j)
                Next j
            Next i
        Else
            DataGridView2.RowCount = 1
            For j = 0 To DataGridView2.ColumnCount - 1
                DataGridView2.Item(j, 0).Value = ""
            Next j
        End If
        Panel1.Visible = True
        TextBox2.Text = ""
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs)
        Panel1.Visible = False
    End Sub

    Private Sub DataGridView2_DoubleClick(sender As Object, e As EventArgs) Handles DataGridView2.DoubleClick
        MO_CODE = DataGridView2.Rows(DataGridView2.CurrentRow.Index).Cells(0).Value
        MO_CM_CODE = DataGridView2.Rows(DataGridView2.CurrentRow.Index).Cells(1).Value
        MO_CM_NAME = DataGridView2.Rows(DataGridView2.CurrentRow.Index).Cells(2).Value
        MO_PL_CODE = DataGridView2.Rows(DataGridView2.CurrentRow.Index).Cells(3).Value
        MO_PL_NAME = DataGridView2.Rows(DataGridView2.CurrentRow.Index).Cells(4).Value
        MO_STANDARD = DataGridView2.Rows(DataGridView2.CurrentRow.Index).Cells(5).Value
        MO_UNIT_AMOUNT = DataGridView2.Rows(DataGridView2.CurrentRow.Index).Cells(6).Value
        MO_QTY = DataGridView2.Rows(DataGridView2.CurrentRow.Index).Cells(7).Value
        MO_TOTAL = DataGridView2.Rows(DataGridView2.CurrentRow.Index).Cells(8).Value
        MO_GOLD = DataGridView2.Rows(DataGridView2.CurrentRow.Index).Cells(9).Value

        DialogResult = DialogResult.OK
        Close()
    End Sub

    Public Function DataGridView2_Display2() As Boolean
        '단일X

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer


        StrSQL = "select MO_CODE, MO_CM_CODE, MO_CM_NAME, MO_PL_CODE, MO_PL_NAME, MO_STANDARD, MO_UNIT_AMOUNT, MO_QTY, MO_TOTAL, MO_GOLD
                    FROM MC_STOCK_ORDER where MO_PL_Name Like '%" + TextBox2.Text + "%' AND MO_ING='진행중'"

        Re_Count = DB_Select(DBT)

        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            DataGridView2.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To DataGridView2.ColumnCount - 1
                    DataGridView2.Item(j, i).Value = DBT.Rows(i).Item(j)
                    ' DataGridView2.Item(0, i).Value = i + 1
                Next j
            Next i
        Else
            DataGridView2.RowCount = 1
            For j = 0 To DataGridView2.ColumnCount - 1
                DataGridView2.Item(j, 0).Value = ""
            Next j
        End If
        Panel1.Visible = True

    End Function
End Class