Imports Excel = Microsoft.Office.Interop.Excel

'생산관리 > 생산소요량산출
Public Class Frm_PC_Order_Plan
    Dim Grid_Display_Ch As Boolean
    Dim Grid_Main_Row As Integer
    Dim Grid_Main_Col As Integer

    Dim Grid_Str(1000, 1000) As String
    Dim Je_ToT(1000, 1000) As String

    Dim Grid_Row As Integer
    Dim Grid_Col As Integer
    Dim Grid_Col2 As Integer

    Dim Product_Count_Cnt As Long

    Dim xlapp As Excel.Application
    Dim xlbook As Excel.Workbook
    Dim xlsheet As Excel.Worksheet
    Dim Excel_check As Boolean


    Private Sub Frm_PC_Order_Plan_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()

        Grid_Order_Setup()
        Grid_Main_Setup()

        'Grid_Main_Setup()

        '검색창
        Search_Combo.Items.Add("거래처명")
        Search_Combo.Items.Add("제품명")
        Search_Combo.Items.Add("수주일자")
        Search_Combo.Text = "거래처명"

        Com_Excel_Print.Visible = False
        Btn_Produc.Visible = False

        Search_Com.PerformClick()
        Grid_Order.ClearSelection()

        'PC_PLAN에 넣기 

    End Sub

    Public Function Grid_Order_Setup() As Boolean
        Grid_Color(Grid_Order)

        Grid_Order.AllowUserToAddRows = False
        Grid_Order.RowTemplate.Height = 20
        Grid_Order.ColumnCount = 9
        Grid_Order.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Order.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Order.RowHeadersWidth = 10
        Grid_Order.Columns(0).Width = 100
        Grid_Order.Columns(1).Width = 50
        Grid_Order.Columns(2).Width = 100
        Grid_Order.Columns(3).Width = 200
        Grid_Order.Columns(4).Width = 200
        Grid_Order.Columns(5).Width = 50
        Grid_Order.Columns(6).Width = 50
        Grid_Order.Columns(7).Width = 100
        Grid_Order.Columns(8).Width = 500

        Grid_Order.Columns(0).HeaderText = "수주코드"
        Grid_Order.Columns(1).HeaderText = "제품별코드"
        Grid_Order.Columns(2).HeaderText = "수주일자"
        Grid_Order.Columns(3).HeaderText = "거래처명"
        Grid_Order.Columns(4).HeaderText = "제품명"
        Grid_Order.Columns(5).HeaderText = "규격"
        Grid_Order.Columns(6).HeaderText = "수량"
        Grid_Order.Columns(7).HeaderText = "납품예정일"
        Grid_Order.Columns(8).HeaderText = "비고"

        Dim i As Integer
        For i = 0 To Grid_Order.ColumnCount - 1
            Grid_Order.Columns(i).ReadOnly = True
        Next i
        Grid_Order_Setup = True
    End Function

    Public Function Grid_Main_Setup() As Boolean

        Grid_Color(Grid_Main)

        Grid_Main.AllowUserToAddRows = False
        Grid_Main.RowTemplate.Height = 20
        Grid_Main.ColumnCount = 10
        Grid_Main.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Main.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.RowHeadersWidth = 10
        Grid_Main.Columns(0).Width = 200
        Grid_Main.Columns(1).Width = 100
        Grid_Main.Columns(2).Width = 100
        Grid_Main.Columns(3).Width = 100
        Grid_Main.Columns(4).Width = 100
        Grid_Main.Columns(5).Width = 100
        Grid_Main.Columns(6).Width = 100
        Grid_Main.Columns(7).Width = 100
        Grid_Main.Columns(8).Width = 100
        Grid_Main.Columns(9).Width = 1500

        Grid_Main.Columns(0).HeaderText = "제품명"
        Grid_Main.Columns(1).HeaderText = "수주량"
        Grid_Main.Columns(2).HeaderText = "완제품재고량"
        Grid_Main.Columns(3).HeaderText = "생산필요량"
        Grid_Main.Columns(4).HeaderText = "필요재료"
        Grid_Main.Columns(5).HeaderText = "생산소모갯수"
        Grid_Main.Columns(6).HeaderText = "원부재료재고량"
        Grid_Main.Columns(7).HeaderText = "산출량"
        Grid_Main.Columns(8).HeaderText = "발주필요량"
        Grid_Main.Columns(9).HeaderText = "비고"

        Grid_Main_Setup = True
    End Function

    Public Function Grid_Order_Display() As Boolean

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer

        '수주현황 보기
        StrSQL = "Select top 100 CR_Code, CL_Code, CR_Date, CR_CM_Name, CL_PL_Name, CL_Standard, CL_Qty, CR_Day, CL_Bigo FROM SC_SALES with(NOLOCK) JOIN SC_SALES_LIST ON SC_SALES.CR_CODE = SC_SALES_LIST.CL_CR_CODE ORDER BY CR_CODE DESC"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Order.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To Grid_Order.ColumnCount - 1
                    Grid_Order.Item(j, i).Value = DBT.Rows(i).Item(j)
                Next j
            Next i
        Else
            Grid_Order.RowCount = 1
            For j = 0 To Grid_Order.ColumnCount - 1
                Grid_Order.Item(j, 0).Value = ""
            Next j
        End If

        Grid_Order.MultiSelect = False
        Grid_Order.ClearSelection()

        '검색창 결과보기
        Dim DBT2 As New DataTable
        Dim k As Integer
        Dim l As Integer
        Grid_Order.RowCount = 1

        Select Case Search_Combo.Text
            Case "수주일자"
                StrSQL = "Select top 100 CR_Code, CL_Code, CR_Date, CR_CM_Name, CL_PL_Name, CL_Standard, CL_Qty, CR_Day, CL_Bigo FROM SC_SALES with(NOLOCK)
                            JOIN SC_SALES_LIST ON SC_SALES.CR_CODE = SC_SALES_LIST.CL_CR_CODE
                            WHERE CR_Date Like '%" & Search_Text.Text & "%' Order By CR_Date"
            Case "거래처명"
                StrSQL = "Select top 100 CR_Code, CL_Code, CR_Date, CR_CM_Name, CL_PL_Name, CL_Standard, CL_Qty, CR_Day, CL_Bigo FROM SC_SALES with(NOLOCK)
                            JOIN SC_SALES_LIST ON SC_SALES.CR_CODE = SC_SALES_LIST.CL_CR_CODE
                            WHERE CR_CM_Name Like '%" & Search_Text.Text & "%' Order By CR_CM_Name"
            Case "제품명"
                StrSQL = "Select top 100 CR_Code, CL_Code, CR_Date, CR_CM_Name, CL_PL_Name, CL_Standard, CL_Qty, CR_Day, CL_Bigo FROM SC_SALES with(NOLOCK)
                            JOIN SC_SALES_LIST ON SC_SALES.CR_CODE = SC_SALES_LIST.CL_CR_CODE
                            WHERE CL_PL_Name Like '%" & Search_Text.Text & "%' Order By CL_PL_Name"
        End Select
        Re_Count = DB_Select(DBT2)
        If Re_Count <> 0 Then
            Grid_Order.RowCount = Re_Count
            For k = 0 To Re_Count - 1
                For l = 0 To Grid_Order.ColumnCount - 1
                    Grid_Order.Item(l, k).Value = DBT2.Rows(k).Item(l)
                Next l
            Next k
        Else
            Grid_Order.RowCount = 1
            For l = 0 To Grid_Order.ColumnCount - 1
                Grid_Order.Item(l, 0).Value = ""
            Next l
        End If
        Grid_Order_Display = True
        Grid_Order.ClearSelection()

    End Function

    Private Sub Grid_Order_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Order.CellEnter
        'Grid_Order의 셀 선택 시 1번행의 값을 object에 심어서 Grid_Main으로 가져간다
        Grid_Main_Display(Grid_Order.Item(1, Grid_Order.CurrentCell.RowIndex).Value)
    End Sub

    Public Function Grid_Main_Display(CL_Code As String) As Boolean

        '수주테이블에서 CL_Code를 받아온다

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display_Ch = False
        Grid_Main.RowCount = 1

        ''수주정보리스트에서 제품명과 수주량 받아오기
        'StrSQL = "Select CL_PL_Name, CL_Qty, CL_PL_Code FROM SC_SALES_LIST with(NOLOCK) Where CL_Code = '" & CL_Code & "' "
        'Re_Count = DB_Select(DBT)
        'If Re_Count <> 0 Then
        '    Grid_Main.RowCount = Re_Count
        '    For i = 0 To Re_Count - 1
        '        Grid_Main.Item(0, i).Value = DBT.Rows(i).Item(0).ToString
        '        Grid_Main.Item(1, i).Value = DBT.Rows(i).Item(1).ToString
        '        Grid_Main.Item(2, i).Value = ""
        '        Grid_Main.Item(3, i).Value = ""
        '        Grid_Main.Item(4, i).Value = ""
        '        Grid_Main.Item(5, i).Value = ""
        '        Grid_Main.Item(6, i).Value = ""
        '        Grid_Main.Item(7, i).Value = DBT.Rows(i).Item(2).ToString
        '    Next i

        'Else
        '    Grid_Main.RowCount = 1
        '    For j = 0 To Grid_Main.ColumnCount - 1
        '        Grid_Main.Item(j, 0).Value = ""
        '    Next j
        'End If

        ''제품코드에서 제품의 BOM정보 가져오기
        'If Re_Count <> 0 Then
        '    For i = 0 To Grid_Main.RowCount - 1
        '        StrSQL = "Select PL_Sub_Name, PL_Qty, PL_Sub_Code FROM SI_PRODUCT_RECIPE with(NOLOCK) Where PL_Code = '" & Grid_Main.Item(7, i).Value & "' Order By Convert(Decimal, PL_Sun) "
        '        Re_Count2 = DB_Select(DBT)
        '        If Re_Count2 <> 0 Then
        '            Grid_Main.Item(2, i).Value = DBT.Rows(i).Item(0).ToString
        '            Grid_Main.Item(3, i).Value = DBT.Rows(i).Item(1).ToString
        '            Grid_Main.Item(7, i).Value = DBT.Rows(i).Item(2).ToString
        '        End If
        '    Next
        'End If

        '제품명수주량/BOM을 따로했더니 행이 1줄만 나와서(그렇겠지) 통합함
        StrSQL = "Select CL_PL_Name, CL_Qty, PL_Sub_Name, PL_Qty, PL_Sub_Code, CL_PL_Code FROM SC_SALES_LIST with(NOLOCK)
                    JOIN SI_PRODUCT_RECIPE ON SC_SALES_LIST.CL_PL_Code = SI_PRODUCT_RECIPE.PL_Code
                    Where CL_Code = '" & CL_Code & "' "
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Main.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                Grid_Main.Item(0, i).Value = DBT.Rows(i).Item(0).ToString
                Grid_Main.Item(1, i).Value = DBT.Rows(i).Item(1).ToString
                Grid_Main.Item(2, i).Value = ""
                Grid_Main.Item(3, i).Value = ""
                Grid_Main.Item(4, i).Value = DBT.Rows(i).Item(2).ToString
                Grid_Main.Item(5, i).Value = DBT.Rows(i).Item(3).ToString
                Grid_Main.Item(6, i).Value = ""
                Grid_Main.Item(7, i).Value = ""
                Grid_Main.Item(8, i).Value = DBT.Rows(i).Item(4).ToString
                Grid_Main.Item(9, i).Value = DBT.Rows(i).Item(5).ToString
            Next i

        Else
            Grid_Main.RowCount = 1
            For j = 0 To Grid_Main.ColumnCount - 1
                Grid_Main.Item(j, 0).Value = ""
            Next j

        End If

        'STOCK테이블에서 완제품의 현재 수량 정보 가져오기 / 부족량 계산
        If Re_Count <> 0 Then
            For i = 0 To Grid_Main.RowCount - 1
                StrSQL = "Select PL_QTY FROM STOCK with(NOLOCK) Where PL_CODE = '" & Grid_Main.Item(8, i).Value & "'"
                Re_Count2 = DB_Select(DBT)
                If Re_Count2 <> 0 Then
                    '완제품재고량
                    Grid_Main.Item(2, i).Value = DBT.Rows(0).Item(0).ToString
                    '완제품재고량-수주량 = 생산필요량 (+일 경우 원부재료 산출량 계산을 하지 않아야 함
                    Grid_Main.Item(3, i).Value = Val(Grid_Main.Item(2, i).Value) - Val(Grid_Main.Item(1, i).Value)
                End If
            Next
        End If

        'STOCK테이블에서 원부재료의 현재 수량 정보 가져오기 / 산출량 계산 / 부족량 계산
        If Re_Count <> 0 Then
            For i = 0 To Grid_Main.RowCount - 1
                StrSQL = "Select PL_QTY FROM STOCK with(NOLOCK) Where PL_CODE = '" & Grid_Main.Item(9, i).Value & "'"
                Re_Count2 = DB_Select(DBT)
                If Re_Count2 <> 0 Then
                    '원부재료재고량
                    Grid_Main.Item(6, i).Value = DBT.Rows(0).Item(0).ToString
                    '원부재료필요산출량 = 3*5
                    Grid_Main.Item(7, i).Value = Val(Grid_Main.Item(3, i).Value) * Val(Grid_Main.Item(5, i).Value)
                    '재고-산출량 = 발주필요량 
                    Grid_Main.Item(8, i).Value = Val(Grid_Main.Item(6, i).Value) - Val(Grid_Main.Item(7, i).Value)
                End If
            Next
        End If

        Grid_Display_Ch = True
        Grid_Main_Display = True

    End Function

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        Grid_Order_Display()
    End Sub


    Public Function Log_Setup() As Boolean
        Grid_Color(Log)

        Log.AllowUserToAddRows = False
        Log.RowTemplate.Height = 20
        Log.ColumnCount = 5
        Log.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Log.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Log.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Log.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Log.RowHeadersWidth = 10

        Log.Columns(0).HeaderText = "작업일자"
        Log.Columns(1).HeaderText = "작업화면"
        Log.Columns(2).HeaderText = "작업"

        Log.Columns(3).HeaderText = "작업자"
        Log.Columns(4).HeaderText = "작업비고"


        Dim i As Integer
        For i = 0 To Log.ColumnCount - 1
            Log.Columns(i).ReadOnly = True
        Next i
        Log_Setup = True
    End Function

    Public Function Log_Display() As Boolean

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer

        StrSQL = "Select * FROM MAN_LOG with(NOLOCK) order By ACT_DATE DESC"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Log.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To Log.ColumnCount - 1
                    Log.Item(j, i).Value = DBT.Rows(i).Item(j)
                Next j
            Next i
        Else
            Log.RowCount = 1
            For j = 0 To Log.ColumnCount - 1
                Log.Item(j, 0).Value = ""
            Next j

        End If

        Log_Display = True
        Log.ClearSelection()

    End Function

    Private Sub Insert__Order_Click(sender As Object, e As EventArgs) Handles Insert__Order.Click

        Dim form As New PCOrder_Insert
        form.Label1.Text = "insert2"
        form.contract_code.Text = Grid_Order.Item(0, Grid_Order.CurrentCell.RowIndex).Value
        form.contract_name.Text = Grid_Order.Item(1, Grid_Order.CurrentCell.RowIndex).Value
        form.contract_day.Text = Grid_Order.Item(2, Grid_Order.CurrentCell.RowIndex).Value
        form.contract_time.Text = Grid_Order.Item(3, Grid_Order.CurrentCell.RowIndex).Value
        form.customer_code.Text = Grid_Order.Item(4, Grid_Order.CurrentCell.RowIndex).Value
        form.customer_name.Text = Grid_Order.Item(5, Grid_Order.CurrentCell.RowIndex).Value
        form.CR_day.Text = Grid_Order.Item(6, Grid_Order.CurrentCell.RowIndex).Value
        form.Contract_bigo.Text = Grid_Order.Item(7, Grid_Order.CurrentCell.RowIndex).Value

        form.ShowDialog()

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If

    End Sub
End Class
