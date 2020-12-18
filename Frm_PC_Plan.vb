Imports Excel = Microsoft.Office.Interop.Excel

'생산관리 > 소요량현황
Public Class Frm_PC_Plan
    Dim Grid_Display_Ch As Boolean
    Dim Grid_Main_Row As Integer
    Dim Grid_Main_Col As Integer

    Dim Grid_Str(1000, 1000) As String
    Dim Je_ToT(1000, 1000) As String

    Dim Grid_Row As Integer
    Dim Grid_Col As Integer
    Dim Grid_Col2 As Integer
    Dim Re_Count3 As Integer
    Dim Product_Count_Cnt As Long

    Dim xlapp As Excel.Application
    Dim xlbook As Excel.Workbook
    Dim xlsheet As Excel.Worksheet
    Dim Excel_check As Boolean


    Private Sub Frm_PC_Plan_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()

        Grid_Order_Setup()
        Grid_Main_Setup()

        'Grid_Main_Setup()

        '검색창
        'Search_Combo.Items.Add("수주일자")
        'Search_Combo.Items.Add("거래처명")
        'Search_Combo.Items.Add("제품명")
        'Search_Combo.Text = "거래처명"

        'Panel_Main.Top = 136
        'Panel_Main.Left = 389
        'Panel_Excel.Top = 136
        'Panel_Excel.Left = 389

        'Panel_Main.Visible = True
        'Panel_Excel.Visible = False

        Com_Excel_Print.Visible = False
        Btn_Produc.Visible = False
        'Com_Excel.Visible = False

        'Search_Com.PerformClick()
        Grid_Order.ClearSelection()

        Grid_Order_Display()
        Grid_Main_Display()

    End Sub

    Public Function Grid_Order_Setup() As Boolean
        Grid_Color(Grid_Order)

        Grid_Order.AllowUserToAddRows = False
        Grid_Order.RowTemplate.Height = 20
        Grid_Order.ColumnCount = 6
        Grid_Order.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Order.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Order.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Order.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Order.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Order.RowHeadersWidth = 10
        Grid_Order.Columns(0).Width = 150
        Grid_Order.Columns(1).Width = 100
        Grid_Order.Columns(2).Width = 100
        Grid_Order.Columns(3).Width = 100
        Grid_Order.Columns(4).Width = 100
        Grid_Order.Columns(5).Width = 800


        Grid_Order.Columns(0).HeaderText = "수주제품명"
        Grid_Order.Columns(1).HeaderText = "총수주량"
        Grid_Order.Columns(2).HeaderText = "총재고량"
        Grid_Order.Columns(3).HeaderText = "부족분"
        Grid_Order.Columns(4).HeaderText = "테스트"
        Grid_Order.Columns(5).HeaderText = "visible"

        Dim i As Integer
        For i = 0 To Grid_Order.ColumnCount - 1
            Grid_Order.Columns(i).ReadOnly = True
        Next i
        Grid_Order_Setup = True

    End Function

    Public Function Grid_Main_Setup() As Boolean

        Grid_Color(Grid_Main)

        Grid_Main.AllowUserToAddRows = False
        Grid_Main.RowTemplate.Height = 24
        Grid_Main.ColumnCount = 8
        Grid_Main.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Main.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.RowHeadersWidth = 10
        Grid_Main.Columns(0).Width = 200
        Grid_Main.Columns(1).Width = 100
        Grid_Main.Columns(2).Width = 100
        Grid_Main.Columns(3).Width = 100
        Grid_Main.Columns(4).Width = 100
        Grid_Main.Columns(5).Width = 100
        Grid_Main.Columns(6).Width = 100
        Grid_Main.Columns(7).Width = 1500

        Grid_Main.Columns(0).HeaderText = "수주제품명"
        Grid_Main.Columns(1).HeaderText = "작업필요량"
        Grid_Main.Columns(2).HeaderText = "필요재료"
        Grid_Main.Columns(3).HeaderText = "레시피값"
        Grid_Main.Columns(4).HeaderText = "필요재료량"
        Grid_Main.Columns(5).HeaderText = "재고재료량"
        Grid_Main.Columns(6).HeaderText = "부족재료량"
        Grid_Main.Columns(7).HeaderText = "코드용(인비지블)"

        Grid_Main_Setup = True
    End Function

    Public Function Grid_Order_Display() As Boolean

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer

        '수주 전체 현황 보기
        StrSQL = "Select CL_PL_Code, CL_PL_NAME, SUM(CONVERT(integer, CL_Qty)) FROM SC_SALES_LIST with(NOLOCK) GROUP BY CL_PL_Code, CL_PL_NAME"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Order.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                Grid_Order.Item(0, i).Value = DBT.Rows(i).Item(1).ToString
                Grid_Order.Item(1, i).Value = DBT.Rows(i).Item(2).ToString
                Grid_Order.Item(2, i).Value = ""
                Grid_Order.Item(3, i).Value = ""
                Grid_Order.Item(4, i).Value = ""
                Grid_Order.Item(5, i).Value = DBT.Rows(i).Item(0).ToString
            Next i
        Else
            Grid_Order.RowCount = 1
            For j = 0 To Grid_Order.ColumnCount - 1
                Grid_Order.Item(j, 0).Value = ""
            Next j
        End If

        If Re_Count <> 0 Then
            For i = 0 To Grid_Order.RowCount - 1
                StrSQL = "Select PL_QTY FROM STOCK with(NOLOCK) Where PL_Code = '" & Grid_Order.Item(5, i).Value & "'"
                Re_Count2 = DB_Select(DBT)
                If Re_Count2 <> 0 Then
                    Grid_Order.Item(2, i).Value = DBT.Rows(i).Item(0).ToString
                    'PL_QTY*1번
                    Grid_Order.Item(3, i).Value = Val(Grid_Order.Item(2, i).Value) - Val(Grid_Order.Item(1, i).Value)
                    Grid_Order.Item(4, i).Value = ""
                End If
            Next
        End If

        Grid_Order_Display = True
        Grid_Order.ClearSelection()

    End Function

    Public Function Grid_Main_Display() As Boolean

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display_Ch = False
        Grid_Main.RowCount = 1

        'StrSQL = "Select PL_QTY FROM STOCK with(NOLOCK) Where PL_Code = '" & Grid_Order.Item(5, i).Value & "'"
        StrSQL = "Select PL_Sub_Name, PL_Qty, PL_Sub_Code FROM SC_SALES_LIST With(NOLOCK)
                    Join SI_PRODUCT_RECIPE ON SC_SALES_LIST.CL_PL_Code = SI_PRODUCT_RECIPE.PL_Code where PL_CODE = '" & Grid_Order.Item(5, i).Value & "'"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Main.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                Grid_Main.Item(0, i).Value = Grid_Order.Item(0, i).Value
                Grid_Main.Item(1, i).Value = Val(Grid_Order.Item(3, i).Value) * -1
                Grid_Main.Item(2, j).Value = DBT.Rows(j).Item(0).ToString
                Grid_Main.Item(3, j).Value = DBT.Rows(j).Item(1).ToString
                Grid_Main.Item(4, j).Value = Val(Grid_Main.Item(3, i).Value) * Val(Grid_Main.Item(1, i).Value)
                Grid_Main.Item(5, j).Value = ""
                Grid_Main.Item(6, j).Value = ""
                Grid_Main.Item(7, j).Value = DBT.Rows(j).Item(2).ToString
            Next i

            'Else
            '    Grid_Main.RowCount = 1
            '    For j = 0 To Grid_Main.ColumnCount - 1
            '        Grid_Main.Item(j, 0).Value = ""
            '    Next j

        End If

        '제품코드에서 제품의 BOM정보와 총필요량 계산하기
        'If Re_Count <> 0 Then
        '    For i = 0 To Grid_Main.RowCount - 1
        '        StrSQL = "Select PL_Sub_Name, PL_Qty, PL_Sub_Code FROM SI_PRODUCT_RECIPE with(NOLOCK) Where PL_Code = '" & Grid_Main.Item(6, i).Value & "' Order By Convert(Decimal, PL_Sun) "
        '        Re_Count2 = DB_Select(DBT)
        '        If Re_Count2 <> 0 Then
        '            Grid_Main.Item(2, i).Value = DBT.Rows(i).Item(0).ToString
        '            'PL_QTY*1번
        '            Grid_Main.Item(3, i).Value = Val(Grid_Main.Item(1, i).Value) * DBT.Rows(i).Item(1).ToString
        '            Grid_Main.Item(6, i).Value = DBT.Rows(i).Item(2).ToString
        '        End If
        '    Next
        'End If

        '재료코드에서 재료의 현재 수량 정보 가져오기 / 산출량 계산 / 부족량 계산
        'If Re_Count <> 0 Then
        '    For i = 0 To Grid_Main.RowCount - 1
        '        StrSQL = "Select PL_QTY FROM STOCK with(NOLOCK) Where PL_CODE = '" & Grid_Main.Item(7, i).Value & "'"
        '        Re_Count3 = DB_Select(DBT)
        '        If Re_Count3 <> 0 Then
        '            Grid_Main.Item(5, i).Value = DBT.Rows(0).Item(0).ToString
        '            Grid_Main.Item(6, i).Value = Val(Grid_Main.Item(5, i).Value) - Val(Grid_Main.Item(4, i).Value)
        '        End If
        '    Next
        'End If

        Grid_Display_Ch = True
        Grid_Main_Display = True

    End Function

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        Grid_Order_Display()
        Grid_Main_Display()
    End Sub

End Class
