Public Class Frm_Bal



    Private Sub Setup()
        If Me.Text = "발주정보 불러오기" Then
            Grid_Color(DG)
            DG.AllowUserToAddRows = False
            DG.RowTemplate.Height = 50
            DG.ColumnCount = 9
            DG.RowCount = 0

            '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
            DG.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            DG.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            DG.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            DG.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DG.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DG.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

            DG.RowHeadersWidth = 10

            DG.Columns(0).HeaderText = "발주코드"
            DG.Columns(1).HeaderText = "일자"
            DG.Columns(2).HeaderText = "거래처명"
            DG.Columns(3).HeaderText = "제품명"
            DG.Columns(4).HeaderText = "규격"
            DG.Columns(5).HeaderText = "단위"
            DG.Columns(6).HeaderText = "단가"
            DG.Columns(7).HeaderText = "수량"
            DG.Columns(8).HeaderText = "금액"

            Dim i As Integer

            DG.Columns(0).Width = 100
            DG.Columns(1).Width = 80
            DG.Columns(2).Width = 120

            For i = 3 To DG.ColumnCount - 1
                DG.Columns(i).Width = 80
            Next i
        Else
            Grid_Color(DG)
            DG.AllowUserToAddRows = False
            DG.RowTemplate.Height = 50
            DG.ColumnCount = 9
            DG.RowCount = 0

            '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
            DG.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            DG.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            DG.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            DG.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DG.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DG.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

            DG.RowHeadersWidth = 10


            DG.Columns(0).HeaderText = "수주코드"
            DG.Columns(1).HeaderText = "일자"
            DG.Columns(2).HeaderText = "거래처명"
            DG.Columns(3).HeaderText = "제품명"
            DG.Columns(4).HeaderText = "규격"
            DG.Columns(5).HeaderText = "단위"
            DG.Columns(6).HeaderText = "단가"
            DG.Columns(7).HeaderText = "수량"
            DG.Columns(8).HeaderText = "금액"

            Dim i As Integer

            DG.Columns(0).Width = 100
            DG.Columns(1).Width = 80
            DG.Columns(2).Width = 120

            For i = 3 To DG.ColumnCount - 1
                DG.Columns(i).Width = 80
            Next i
        End If
    End Sub



    Private Sub Frm_Bal_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim DBT As New DataTable
        If Me.Text = "발주정보 불러오기" Then

            Setup()

            StrSQL = "Select CO_CODE, CO_DATE, CO_CUSTOMER_NAME, C0_PL_NAME, CO_STANDARD, CO_UNIT, CO_UNIT_AMOUNT, CO_TOTAL, CO_GOLD
                        FROM MC_STOCK_ORDER AS A with(NOLOCK)
                             JOIN MC_STOCK_ORDER_LIST AS B ON A.CO_CODE = B.CO_CODE"
            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                Me.Close()

            End If
        End If

        If Me.Text = "수주정보 불러오기" Then

            Setup()

            StrSQL = "Select CR_CODE, CR_DATE, CR_CUSTOMER_NAME, CL_PL_NAME, CL_STANDARD, CL_UNIT, CL_UNIT_AMOUNT, CL_TOTAL, CL_GOLD
                        From SC_SALES As A with(NOLOCK)
                             Join SC_SALES_LIST As B On A.CR_CODE = B.CL_CODE"
            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                Me.Close()


            End If
        End If
    End Sub

    Private Sub DG_DoubleClick(sender As Object, e As EventArgs) Handles DG.DoubleClick
        If Len(DG.Item(0, DG.CurrentCell.RowIndex).Value) > 0 Then
            Dim code = DG.Item(0, DG.CurrentCell.RowIndex).Value

        End If
    End Sub
End Class