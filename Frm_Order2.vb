Imports Excel = Microsoft.Office.Interop.Excel

Public Class Frm_Order2
    Dim Grid_Display_Ch As Boolean
    Dim Grid_Info_Row As Integer
    Dim Grid_Info_Col As Integer

    Dim Grid_Main_Row As Integer
    Dim Grid_Main_Col As Integer


    Dim xlapp As Excel.Application
    Dim xlbook As Excel.Workbook
    Dim xlsheet As Excel.Worksheet

    Dim Excel_check As Boolean

    Private Sub Frm_Order_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Me.Dock = DockStyle.Fill
        Me.BringToFront()

        Order.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        Log.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill


        Order_Setup()
        Log_Setup()

        Search_Com.PerformClick()
        Order.SelectionMode = DataGridViewSelectionMode.FullRowSelect

    End Sub
    Public Function Order_Setup() As Boolean
        Grid_Color(Order)

        Order.AllowUserToAddRows = False
        Order.RowTemplate.Height = 20
        Order.ColumnCount = 16
        Order.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Order.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Order.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Order.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Order.RowHeadersWidth = 10
        Order.Columns(0).Width = 100
        Order.Columns(1).Width = 70
        Order.Columns(2).Width = 80

        Order.Columns(0).HeaderText = "입고코드"
        Order.Columns(1).HeaderText = "사무실/현장"
        Order.Columns(1).Visible = False
        Order.Columns(2).HeaderText = "발주코드"

        Order.Columns(3).HeaderText = "거래처코드"
        Order.Columns(3).Visible = False
        Order.Columns(4).HeaderText = "거래처명"
        Order.Columns(5).HeaderText = "입고일"

        Order.Columns(6).HeaderText = "담당자"
        Order.Columns(7).HeaderText = "유통기한"
        Order.Columns(8).HeaderText = "품목코드"
        Order.Columns(8).Visible = False

        Order.Columns(9).HeaderText = "품목명"
        Order.Columns(10).HeaderText = "규격"
        Order.Columns(11).HeaderText = "단가"

        Order.Columns(12).HeaderText = "입고수량"
        Order.Columns(13).HeaderText = "금액"
        Order.Columns(14).HeaderText = "부가세"
        Order.Columns(15).HeaderText = "비고"


        Dim i As Integer
        For i = 0 To Order.ColumnCount - 1
            Order.Columns(i).ReadOnly = True
        Next i
        Order_Setup = True
    End Function

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

    Private Sub Insert_Com_Click(sender As Object, e As EventArgs) Handles Insert_Com.Click
        Dim form As New Order2_Insert
        form.Label1.Text = "insert"
        form.ShowDialog()

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If
    End Sub

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        Order_Display()
        Log_Display()

    End Sub

    Public Function Order_Display() As Boolean

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer

        StrSQL = "Select * FROM MC_STOCK_IN with(NOLOCK) Where MI_CM_NAME Like '%" & Search_Text.Text & "%' Order By MI_CODE"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Order.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To Order.ColumnCount - 2
                    Order.Item(j, i).Value = DBT.Rows(i).Item(j)
                Next j
            Next i
        Else
            Order.RowCount = 1
            For j = 0 To Order.ColumnCount - 2
                Order.Item(j, 0).Value = ""
            Next j

        End If

        Order_Display = True
        Order.ClearSelection()

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

    Private Sub Order_DoubleClick(sender As Object, e As EventArgs) Handles Order.DoubleClick
        If Len(Order.Item(0, Order.CurrentCell.RowIndex).Value) < 0 Then
            MsgBox("해당 행은 공백이므로 처리할 수 없습니다")
            Exit Sub
        End If

        '해당 행 내용 수정

        Dim form As New Order2_Insert
        form.Label1.Text = "update"
        form.Order_Code.Text = Order.Item(0, Order.CurrentCell.RowIndex).Value


        form.ShowDialog()

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If
    End Sub

    Private Sub Delete_Com_Click(sender As Object, e As EventArgs) Handles Delete_Com.Click
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer

        If Len(Order.Item(0, Order.CurrentCell.RowIndex).Value) < 1 Then
            MsgBox("입고코드를 선택하세요.")
            Exit Sub
        End If
        If MsgBox("입고코드  '" & Order.Item(0, Order.CurrentCell.RowIndex).Value & "'를 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "입고코드 삭제") = vbNo Then
            Exit Sub
        End If



        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO MAN_LOG Values ('" & DateTime.Now & "', '" & Me.Name & "', '입고 삭제', '" & Frm_Main.Text_Man_Name.Text & "',
                                                                     '" & Order.Item(0, Order.CurrentCell.RowIndex).Value & " 입고 삭제')"
        Re_Count = DB_Execute()

        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE MC_STOCK_IN where MI_Code = '" & Order.Item(0, Order.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        '재고 테이블 추가
        Dim stock As Integer = 0
        StrSQL = "SELECT PL_QTY FROM STOCK WHERE PL_CODE='" & Order.Item(8, Order.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Select(DBT)

        If Re_Count = 0 Then
            Exit Sub
        Else
            stock = DBT.Rows(0)("PL_QTY")
        End If

        stock = stock - Val(Order.Item(12, Order.CurrentCell.RowIndex).Value)


        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "UPDATE STOCK SET
                                          PL_QTY ='" & stock & "'
                                    WHERE PL_CODE ='" & Order.Item(8, Order.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()


        Dim SR_CODE As String

        '재고이력코드 번호자동생성
        StrSQL = "Select SR_CODE FROM STOCK_RECORD with(NOLOCK) Order By SR_CODE Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            SR_CODE = "001"
        Else
            'JP_Code = Val(DBT.Rows(0).Item(0)) + 1
            SR_CODE = Replace(DBT.Rows(0).Item(0), "SR", "")
            SR_CODE = Format(SR_CODE + 1, "00#")
        End If


        '재고이력 테이블 추가
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO STOCK_RECORD Values ('SR" & SR_CODE & "','" & Today & "','" & Order.Item(0, Order.CurrentCell.RowIndex).Value & "','입고삭제','" & Order.Item(9, Order.CurrentCell.RowIndex).Value & "','" & Order.Item(12, Order.CurrentCell.RowIndex).Value & "'
                                                                , '" & Frm_Main.Text_Man_Name.Text & "', '" & Today & "','','')"
        Re_Count = DB_Execute()



        '발주 진행중 -> 완료
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "UPDATE MC_STOCK_ORDER SET
                                          MO_ING ='진행중'
                                   WHERE MO_CODE ='" & Order.Item(2, Order.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()



        Order_Display()
        MsgBox("삭제 되었습니다.")

        Search_Com.PerformClick()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs)
        MsgBox("구현예정입니다")
    End Sub
End Class
