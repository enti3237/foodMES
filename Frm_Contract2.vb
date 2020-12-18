
Imports Excel = Microsoft.Office.Interop.Excel

Public Class Frm_Contract2
    Dim Grid_Info_Row As Integer
    Dim Grid_Info_Col As Integer
    Dim Grid_Main_Row As Integer
    Dim Grid_Main_Col As Integer
    Dim xlapp As Excel.Application
    Dim xlbook As Excel.Workbook
    Dim xlsheet As Excel.Worksheet
    Dim Excel_Check As Boolean
    Dim Grid_Display_Ch As Boolean

    Private Sub Frm_Contract2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()

        Grid_Info.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        grid_main.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        grid_info_setup()
        grid_main_setup()
        grid_info_Display()
        Grid_Info.ClearSelection()
        Log_Setup()
        Log_Display()

        'CellClick = False
    End Sub
    Public Sub grid_info_setup()
        Grid_Color(Grid_Info)

        Grid_Info.AllowUserToAddRows = False
        Grid_Info.RowTemplate.Height = 24

        Grid_Info.ColumnCount = 8
        Grid_Info.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Info.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Info.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Info.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Info.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Info.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Info.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Info.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Info.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Grid_Info.RowHeadersWidth = 40
        Grid_Info.Columns(0).Width = 35
        Grid_Info.Columns(1).Width = 80
        Grid_Info.Columns(2).Width = 150
        Grid_Info.Columns(3).Width = 200
        Grid_Info.Columns(4).Width = 200
        Grid_Info.Columns(5).Width = 200
        Grid_Info.Columns(6).Width = 200
        Grid_Info.Columns(7).Width = 200

        Grid_Info.Columns(0).HeaderText = "코드"
        Grid_Info.Columns(1).HeaderText = "담당자"
        Grid_Info.Columns(2).HeaderText = "날짜"
        Grid_Info.Columns(3).HeaderText = "시간"
        Grid_Info.Columns(4).HeaderText = "거래처코드"
        Grid_Info.Columns(5).HeaderText = "거래처명"
        Grid_Info.Columns(6).HeaderText = "납품예정일"
        Grid_Info.Columns(7).HeaderText = "비고"

        Grid_Info.Item(0, 0).Value = ""
        Grid_Info.Item(1, 0).Value = ""
        Grid_Info.Item(2, 0).Value = ""
        Grid_Info.Item(3, 0).Value = ""
        Grid_Info.Item(4, 0).Value = ""
        Grid_Info.Item(5, 0).Value = ""
        Grid_Info.Item(6, 0).Value = ""
        Grid_Info.Item(7, 0).Value = ""

        Grid_Info.Columns(0).ReadOnly = True
    End Sub

    Public Sub grid_main_setup()
        Grid_Color(Grid_Main)


        Grid_Main.AllowUserToAddRows = False
        Grid_Main.RowTemplate.Height = 20


        '열의 갯수는 하나 더 많아야 함.
        Grid_Main.ColumnCount = 9
        Grid_Main.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Main.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Main.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Grid_Main.RowHeadersWidth = 40
        Grid_Main.Columns(0).Width = 80
        Grid_Main.Columns(1).Width = 150
        Grid_Main.Columns(2).Width = 150
        Grid_Main.Columns(3).Width = 150
        Grid_Main.Columns(4).Width = 150
        Grid_Main.Columns(5).Width = 150
        Grid_Main.Columns(6).Width = 150
        Grid_Main.Columns(7).Width = 150
        Grid_Main.Columns(8).Width = 150

        Grid_Main.Columns(0).HeaderText = "수주리스트코드"
        Grid_Main.Columns(1).HeaderText = "수주코드"
        Grid_Main.Columns(2).HeaderText = "제품코드"
        Grid_Main.Columns(3).HeaderText = "제품명"
        Grid_Main.Columns(4).HeaderText = "규격"
        Grid_Main.Columns(5).HeaderText = "단가"
        Grid_Main.Columns(6).HeaderText = "수량"
        Grid_Main.Columns(7).HeaderText = "총액"
        Grid_Main.Columns(8).HeaderText = "비고"

        Grid_Main.Columns(0).ReadOnly = True
        'grid_main_setup = True
    End Sub

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        grid_info_Display()
        Log_Display()
        grid_main_Display(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value)
    End Sub

    Public Function grid_info_Display() As Boolean

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        'Grid_Display = False
        Grid_Info.RowCount = 1
        StrSQL = "Select * FROM SC_SALES with(NOLOCK) Where CR_CM_Name Like '%" & Search_Text.Text & "%' Or CR_CM_Name Is Null Order By CR_CM_Name"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Info.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To 7
                    Grid_Info.Item(j, i).Value = DBT.Rows(i).Item(j)
                Next j
            Next i
        Else
            Grid_Info.RowCount = 1
            For j = 0 To Grid_Info.ColumnCount - 1
                Grid_Info.Item(j, 0).Value = ""
            Next j

        End If
        'Grid_Info Display = True
        grid_info_Display = True
        Grid_Info.ClearSelection()
        Grid_Main.ClearSelection()

    End Function

    Public Function grid_main_Display(Grid_info_code As String) As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        'Grid_Display = False
        Grid_Main.RowCount = 1
        StrSQL = "Select CL_CODE,  CL_CR_Code, CL_PL_Code, CL_PL_Name, CL_Standard, CL_Unit_Amount, CL_Qty, CL_Gold, CL_Bigo FROM SC_Sales_list With(NOLOCK)  Where CL_CR_Code = '" & Grid_info_code & "'  "
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Main.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To 7
                    Grid_Main.Item(j, i).Value = DBT.Rows(i).Item(j)
                Next j
            Next i
        Else
            Grid_Main.RowCount = 1
            Grid_Main.Rows(0).HeaderCell.Value = ""
            For j = 0 To Grid_Main.ColumnCount - 1
                Grid_Main.Item(j, 0).Value = ""
            Next j

        End If
        'For i = 0 To Grid_Main.RowCount - 1
        '    StrSQL = "Select PD_Name  FROM SI_MACHINE with(NOLOCK)  Where PD_Code = '" & Grid_Main.Item(1, i).Value & "' "
        '    Re_Count = DB_Select(DBT)
        '    If Re_Count <> 0 Then
        '        Grid_Main.Item(2, i).Value = DBT.Rows(0).Item(0)
        '    End If
        'Next
        'Grid_Display = True
        grid_main_Display = True
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
        'Search_Com.PerformClick()
        Dim form As New Contract_Insert
        form.Label1.Text = "insert"
        form.ShowDialog()

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If
    End Sub

    Private Sub Delete_Com_Click(sender As Object, e As EventArgs) Handles Delete_Com.Click
        If Len(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value) < 1 Then
            MsgBox("수주코드를 선택하세요.")
            Exit Sub
        End If
        If MsgBox("수주코드  '" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value & "'를 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "수주코드 삭제") = vbNo Then
            Exit Sub
        End If
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO MAN_LOG Values ('" & DateTime.Now & "', '" & Me.Name & "', '수주 삭제', '" & Frm_Main.Text_Man_Name.Text & "',
                                                                     '" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value & " 수주 삭제')"
        Re_Count = DB_Execute()

        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SC_sales where CR_Code = '" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()
        'Grid_Display = False
        grid_info_Display()
        MsgBox("삭제 되었습니다.")
        Search_Com.PerformClick()
        'Grid_Display = True
    End Sub

    Private Sub Grid_Info_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Info.CellClick
        If Len(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value) > 0 Then
            'CellClick = True
            grid_main_Display(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value)
        End If
    End Sub

    Private Sub Insert__Bom_Click(sender As Object, e As EventArgs) Handles Insert__Bom.Click
        'If CellClick = False Then
        '    MsgBox("공정을 선택해주세요")
        '    Exit Sub
        'End If

        Dim form As New Contract_Insert
        form.Label1.Text = "insert2"
        form.contract_code.Text = Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value
        form.contract_name.Text = Grid_Info.Item(1, Grid_Info.CurrentCell.RowIndex).Value
        form.contract_day.Text = Grid_Info.Item(2, Grid_Info.CurrentCell.RowIndex).Value
        form.contract_time.Text = Grid_Info.Item(3, Grid_Info.CurrentCell.RowIndex).Value
        form.customer_code.Text = Grid_Info.Item(4, Grid_Info.CurrentCell.RowIndex).Value
        form.customer_name.Text = Grid_Info.Item(5, Grid_Info.CurrentCell.RowIndex).Value
        form.CR_day.Text = Grid_Info.Item(6, Grid_Info.CurrentCell.RowIndex).Value
        form.Contract_bigo.Text = Grid_Info.Item(7, Grid_Info.CurrentCell.RowIndex).Value

        form.ShowDialog()

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If
    End Sub

    Private Sub Grid_Info_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Info.CellDoubleClick
        '유효성 검사
        If Len(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value) < 0 Then
            MsgBox("해당 행은 공백이므로 처리할 수 없습니다")
            Exit Sub
        End If

        '해당 행 내용 수정
        Contract_Insert.Label1.Text = "update"
        Contract_Insert.contract_code.Text = Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value
        Dim form As New Contract_Insert
        form.Label1.Text = "update"
        form.contract_code.Text = Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value
        form.contract_name.Text = Grid_Info.Item(1, Grid_Info.CurrentCell.RowIndex).Value
        form.contract_day.Text = Grid_Info.Item(2, Grid_Info.CurrentCell.RowIndex).Value
        form.contract_time.Text = Grid_Info.Item(3, Grid_Info.CurrentCell.RowIndex).Value
        form.customer_code.Text = Grid_Info.Item(4, Grid_Info.CurrentCell.RowIndex).Value
        form.customer_name.Text = Grid_Info.Item(5, Grid_Info.CurrentCell.RowIndex).Value
        form.CR_day.Text = Grid_Info.Item(6, Grid_Info.CurrentCell.RowIndex).Value
        form.Contract_bigo.Text = Grid_Info.Item(7, Grid_Info.CurrentCell.RowIndex).Value
        form.ShowDialog()

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If
    End Sub

    Private Sub Grid_Main_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Main.CellDoubleClick
        '유효성 검사
        If Len(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value) < 0 Then
            MsgBox("해당 행은 공백이므로 처리할 수 없습니다")
            Exit Sub
        End If

        '해당 행 내용 수정
        Contract_Insert.Label1.Text = "update2"
        Contract_Insert.contract_code.Text = Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value
        Dim form As New Contract_Insert
        form.Label1.Text = "update2"
        form.TextBox5.Text = Grid_Main.Item(2, Grid_Main.CurrentCell.RowIndex).Value
        form.TextBox4.Text = Grid_Main.Item(3, Grid_Main.CurrentCell.RowIndex).Value
        form.TextBox6.Text = Grid_Main.Item(4, Grid_Main.CurrentCell.RowIndex).Value
        form.TextBox7.Text = Grid_Main.Item(5, Grid_Main.CurrentCell.RowIndex).Value
        form.TextBox8.Text = Grid_Main.Item(6, Grid_Main.CurrentCell.RowIndex).Value
        form.TextBox10.Text = Grid_Main.Item(7, Grid_Main.CurrentCell.RowIndex).Value
        form.TextBox9.Text = Grid_Main.Item(8, Grid_Main.CurrentCell.RowIndex).Value
        form.ShowDialog()

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If
    End Sub

    Private Sub Delete__Bom_Click(sender As Object, e As EventArgs) Handles Delete__Bom.Click
        '해당 칼럼 삭제
        If Len(Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        If MsgBox("수주코드 '" & Grid_Main.Item(1, Grid_Main.CurrentCell.RowIndex).Value & "'을(를) 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "삭제") = vbNo Then
            Exit Sub
        End If


        Grid_Display_Ch = False
        Dim i As Integer
        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SC_SALES_LIST Where CL_Code = '" & Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value & "' And  CL_CR_Code = '" & Grid_Main.Item(1, Grid_Main.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO MAN_LOG Values ('" & DateTime.Now & "', '" & Me.Name & "', '수주리스트 삭제', '" & Frm_Main.Text_Man_Name.Text & "',
                                                                     '" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value & " 수주리스트 삭제')"
        Re_Count = DB_Execute()
        grid_main_Display(Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True

        MsgBox("삭제되었습니다")
    End Sub
End Class
