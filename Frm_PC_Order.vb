Imports Excel = Microsoft.Office.Interop.Excel

'생산관리 > 작업지시서
Public Class Frm_PC_Order
    Dim Grid_Display_Ch As Boolean
    Dim Grid_Info_Row As Integer
    Dim Grid_Info_Col As Integer

    Dim Grid_Main_Row As Integer
    Dim Grid_Main_Col As Integer


    Dim xlapp As Excel.Application
    Dim xlbook As Excel.Workbook
    Dim xlsheet As Excel.Worksheet
    Dim Excel_check As Boolean

    Dim Info_Lab() As Button
    Dim Info_Txt() As TextBox
    Dim Info_Com() As ComboBox

    Private Sub Frm_PC_Order_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Dock = DockStyle.Fill
        Me.BringToFront()

        Grid_Order_Setup()

        '검색창
        Search_Combo.Items.Add("거래처명")
        Search_Combo.Items.Add("제품명")
        Search_Combo.Items.Add("납품예정일")
        Search_Combo.Text = "거래처명"

        Com_Excel_Print.Visible = False

        'Search_Com.PerformClick()
        Grid_Order.ClearSelection()

        'PC_PLAN에 넣기 

    End Sub

    Public Function Grid_Order_Setup() As Boolean
        Grid_Color(Grid_Order)

        Grid_Order.AllowUserToAddRows = False
        Grid_Order.RowTemplate.Height = 20
        Grid_Order.ColumnCount = 10
        Grid_Order.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Order.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Order.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Order.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Order.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Order.RowHeadersWidth = 10
        Grid_Order.Columns(0).Width = 200
        Grid_Order.Columns(1).Width = 150
        Grid_Order.Columns(2).Width = 250
        Grid_Order.Columns(3).Width = 150
        Grid_Order.Columns(4).Width = 150
        Grid_Order.Columns(5).Width = 100
        Grid_Order.Columns(6).Width = 100
        Grid_Order.Columns(7).Width = 100
        Grid_Order.Columns(8).Width = 200
        Grid_Order.Columns(9).Width = 500

        Grid_Order.Columns(0).HeaderText = "수주거래처명"
        Grid_Order.Columns(1).HeaderText = "납품예정일"
        Grid_Order.Columns(2).HeaderText = "제품명"
        Grid_Order.Columns(3).HeaderText = "등록일자"
        Grid_Order.Columns(4).HeaderText = "생산지시일자"
        Grid_Order.Columns(5).HeaderText = "작업자명"
        Grid_Order.Columns(6).HeaderText = "작업수량"
        Grid_Order.Columns(7).HeaderText = "규격"
        Grid_Order.Columns(8).HeaderText = "작업지시서코드"
        Grid_Order.Columns(9).HeaderText = "비고"

        Dim i As Integer
        For i = 0 To Grid_Order.ColumnCount - 1
            Grid_Order.Columns(i).ReadOnly = True
        Next i
        Grid_Order_Setup = True

        Grid_Order_Display()

    End Function

    Public Function Grid_Order_Display() As Boolean

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer

        '작업지시서 보기
        StrSQL = "Select CR_CM_Name, CR_Day, PO_PL_Name, PO_Date, PO_Day, PO_Name, PO_Qty, PO_Standard, PO_Code, PO_Bigo
                       FROM PC_ORDER with(NOLOCK) join SC_SALES
                       on PC_ORDER.PO_CR_CODE = SC_SALES.CR_Code ORDER BY PO_Code DESC"
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
            Case "납품예정일"
                StrSQL = "Select CR_CM_Name, CR_Day, PO_PL_Name, PO_Date, PO_Day, PO_Name, PO_Qty, PO_Standard, PO_Code, PO_Bigo
                       FROM PC_ORDER with(NOLOCK) join SC_SALES
                       on PC_ORDER.PO_CR_CODE = SC_SALES.CR_Code 
                       WHERE CR_Day Like '%" & Search_Text.Text & "%' Order By CR_Day DESC"
            Case "거래처명"
                StrSQL = "Select CR_CM_Name, CR_Day, PO_PL_Name, PO_Date, PO_Day, PO_Name, PO_Qty, PO_Standard, PO_Code, PO_Bigo
                           FROM PC_ORDER with(NOLOCK) join SC_SALES
                           ON PC_ORDER.PO_CR_CODE = SC_SALES.CR_Code 
                           WHERE CR_CM_Name Like '%" & Search_Text.Text & "%' Order By CR_CM_Name"
            Case "제품명"
                StrSQL = "Select CR_CM_Name, CR_Day, PO_PL_Name, PO_Date, PO_Day, PO_Name, PO_Qty, PO_Standard, PO_Code, PO_Bigo
                       FROM PC_ORDER with(NOLOCK) join SC_SALES
                       on PC_ORDER.PO_CR_CODE = SC_SALES.CR_Code 
                        WHERE PO_PL_Name Like '%" & Search_Text.Text & "%' Order By PO_PL_Name"
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

    Private Sub Insert_Com_Click(sender As Object, e As EventArgs) Handles Insert_Com.Click

        Dim form As New PCOrder_Insert
        form.Label1.Text = "insert2"
        form.ShowDialog()

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If
    End Sub

End Class
