Public Class popup_suju
    Public Product_Name As String '품명
    Public Product_Sel2 As String '구분
    Public Product_Standard As String '규격
    Public Product_Unit As String '단위
    Public Product_Gold As String '단가
    Public Product_Code As String '코드


    Public a1 As String '순번
    Public a2 As String '수주코드
    Public a3 As String '제품코드
    Public a4 As String '날짜
    Public a5 As String '거래처이름
    Public a6 As String '제품명
    Public a7 As String '규격
    Public a8 As String '단위
    Public a9 As String '단위수
    Public a10 As String '총수량
    Public a11 As String '이전납품수량
    Public a12 As String '납품수량


    Public Function DataGridView1_Setup() As Boolean

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer

        Grid_Color(DataGridView1)

        DataGridView1.AllowUserToAddRows = False
        DataGridView1.RowTemplate.Height = 24
        DataGridView1.ColumnCount = 12
        DataGridView1.RowCount = 0




        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft



        DataGridView1.RowHeadersWidth = 10
        DataGridView1.Columns(0).Width = 60


        DataGridView1.Columns(0).HeaderText = "순번"
        DataGridView1.Columns(1).HeaderText = "수주코드"
        DataGridView1.Columns(2).HeaderText = "제품코드"
        DataGridView1.Columns(3).HeaderText = "날짜"
        DataGridView1.Columns(4).HeaderText = "거래처이름"
        DataGridView1.Columns(5).HeaderText = "제품명"
        DataGridView1.Columns(6).HeaderText = "규격"
        DataGridView1.Columns(7).HeaderText = "단위"
        DataGridView1.Columns(8).HeaderText = "단위수"
        DataGridView1.Columns(9).HeaderText = "총수량"
        DataGridView1.Columns(10).HeaderText = "이전납품수량"
        DataGridView1.Columns(11).HeaderText = "납품수량"

        DataGridView1.Columns(0).Visible = False
        DataGridView1.Columns(2).Visible = False
        DataGridView1.Columns(6).Visible = False
        DataGridView1.Columns(7).Visible = False
        DataGridView1.Columns(8).Visible = False

        DataGridView1.Columns(0).Width = 10
        DataGridView1.Columns(1).Width = 50
        DataGridView1.Columns(3).Width = 40
        DataGridView1.Columns(4).Width = 50
        DataGridView1.Columns(5).Width = 50
        DataGridView1.Columns(9).Width = 30
        DataGridView1.Columns(10).Width = 40
        DataGridView1.Columns(11).Width = 50


        StrSQL = "select  SC_SALES_LIST.CL_Sun,SC_SALES_LIST.CL_Code ,SC_SALES_LIST.CL_PL_Code, SC_SALES.CR_Date , SC_SALES.CR_Customer_Name  , 
SC_SALES_LIST.CL_PL_Name,SC_SALES_LIST.CL_Standard,SC_SALES_LIST.CL_Unit , SC_SALES_LIST.CL_Unit_Amount,
SC_SALES_LIST.CL_Total ,SUM(convert(integer,SC_DELIVER_LIST.DL_Total)),'2'
                    
					FROM SC_SALES JOIN SC_SALES_LIST ON SC_SALES.CR_Code=SC_SALES_LIST.CL_Code LEFT OUTER JOIN SC_DELIVER_LIST
					ON SC_SALES_LIST.CL_Code=SC_DELIVER_LIST.CL_Code 
AND SC_SALES_LIST.CL_PL_Name=SC_DELIVER_LIST.DL_PL_Name where CL_Total!=DL_Total OR DL_Total IS NULL
 group by SC_SALES_LIST.CL_Sun,SC_SALES_LIST.CL_Code ,SC_SALES_LIST.CL_PL_Code, SC_SALES.CR_Date , 
 SC_SALES.CR_Customer_Name  , SC_SALES_LIST.CL_PL_Name ,SC_SALES_LIST.CL_Standard,SC_SALES_LIST.CL_Unit,
 SC_SALES_LIST.CL_Unit_Amount,SC_SALES_LIST.CL_Total"

        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            DataGridView1.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To DataGridView1.ColumnCount - 2
                    DataGridView1.Item(j, i).Value = DBT.Rows(i).Item(j)
                    'DataGridView1.Item(0, i).Value = i + 1
                Next j
                DataGridView1.Item(11, i).Value = Val(DataGridView1.Item(9, i).Value.ToString) - Val(DataGridView1.Item(10, i).Value.ToString)
            Next i
        Else
            DataGridView1.RowCount = 1
            For j = 0 To DataGridView1.ColumnCount - 1
                DataGridView1.Item(j, 0).Value = ""
            Next j
        End If

        DataGridView1_Setup = True


    End Function

    Private Sub DataGridView1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles DataGridView1.MouseDoubleClick

        a1 = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value.ToString()
        a2 = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(1).Value.ToString()
        a3 = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(2).Value.ToString()
        a4 = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(3).Value.ToString()
        a5 = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(4).Value.ToString()
        a6 = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(5).Value.ToString()
        a7 = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(6).Value.ToString()
        a8 = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(7).Value.ToString()
        a9 = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(8).Value.ToString()
        a10 = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(9).Value.ToString()
        a11 = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(10).Value.ToString()
        ' a12 = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(11).Value.ToString()

        DialogResult = DialogResult.OK
        Close()
    End Sub

    Private Sub popup_suju_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1_Setup()
    End Sub
End Class