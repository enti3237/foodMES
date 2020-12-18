

Imports Excel = Microsoft.Office.Interop.Excel

Public Class Frm_Contract
    Dim Grid_Info_Row As Integer
    Dim Grid_Info_Col As Integer
    Dim Grid_Main_Row As Integer
    Dim Grid_Main_Col As Integer
    Dim xlapp As Excel.Application
    Dim xlbook As Excel.Workbook
    Dim xlsheet As Excel.Worksheet
    Dim Excel_Check As Boolean
    Dim Grid_Display_Ch As Boolean

    Private Sub Frm_Contract_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()
        Grid_Code_Setup()
        Grid_Info_Setup()
        Grid_Main_Setup()




        Search_Combo.Items.Add("코드")
        Search_Combo.Items.Add("날짜")
        Search_Combo.Items.Add("거래처")
        Search_Combo.Text = "코드"


        Panel_Main.Top = 136
        Panel_Main.Left = 389
        Panel_Excel.Top = 136
        Panel_Excel.Left = 389

        Panel_Main.Visible = True
        Panel_Excel.Visible = False

        Com_Excel_Print.Visible = False


    End Sub
    Public Function Grid_Code_Setup() As Boolean
        Grid_Color(Grid_Code)

        Grid_Code.AllowUserToAddRows = False
        Grid_Code.RowTemplate.Height = 20
        Grid_Code.ColumnCount = 3
        Grid_Code.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Code.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Code.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Code.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Code.RowHeadersWidth = 10
        Grid_Code.Columns(0).Width = 100
        Grid_Code.Columns(1).Width = 70
        Grid_Code.Columns(2).Width = 80

        Grid_Code.Columns(0).HeaderText = "코드"
        Grid_Code.Columns(1).HeaderText = "날짜"
        Grid_Code.Columns(2).HeaderText = "거래처"

        Dim i As Integer
        For i = 0 To Grid_Code.ColumnCount - 1
            Grid_Code.Columns(i).ReadOnly = True
        Next i
        Grid_Code_Setup = True
    End Function

    Public Function Grid_Info_Setup() As Boolean

        Grid_Color(Grid_Info)


        Grid_Info.AllowUserToAddRows = False
        Grid_Info.RowTemplate.Height = 24
        Grid_Info.ColumnCount = 2
        Grid_Info.RowCount = 14





        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Info.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Info.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Info.RowHeadersWidth = 10
        Grid_Info.Columns(0).Width = 100
        Grid_Info.Columns(1).Width = 240
        Grid_Info.Columns(0).HeaderText = "구분"
        Grid_Info.Columns(1).HeaderText = "내용"

        Grid_Info.Item(0, 0).Value = "코드"
        Grid_Info.Item(0, 1).Value = "담당자"
        Grid_Info.Item(0, 2).Value = "날짜"
        Grid_Info.Item(0, 3).Value = "시간"
        Grid_Info.Item(0, 4).Value = "거래처코드"
        Grid_Info.Item(0, 5).Value = "거래처명"
        Grid_Info.Item(0, 6).Value = "담당자"
        Grid_Info.Item(0, 7).Value = "연락처"
        Grid_Info.Item(0, 8).Value = "지불조건"

        Grid_Info.Item(0, 9).Value = "납품예정일1"
        Grid_Info.Item(0, 10).Value = "납품예정일1"
        Grid_Info.Item(0, 11).Value = "납품예정일3"
        Grid_Info.Item(0, 12).Value = "비고"
        Grid_Info.Item(0, 13).Value = "관리"


        'Grid_Info.Columns(0).ReadOnly = True
        'Grid_Info.Columns(1).ReadOnly = True

        Grid_Info.Rows(0).ReadOnly = True
        Grid_Info.Rows(1).ReadOnly = True
        Grid_Info.Rows(2).ReadOnly = True
        Grid_Info.Rows(3).ReadOnly = True
        Grid_Info.Rows(4).ReadOnly = False

        Grid_Info_Setup = True
    End Function
    Public Function Grid_Main_Setup() As Boolean
        Grid_Color(Grid_Main)

        Grid_Main.AllowUserToAddRows = False
        Grid_Main.RowTemplate.Height = 24
        Grid_Main.ColumnCount = 12
        Grid_Main.RowCount = 0

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Main.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Grid_Main.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Grid_Main.RowHeadersWidth = 10


        Grid_Main.Columns(0).HeaderText = "순번"
        Grid_Main.Columns(1).HeaderText = "제품코드"
        Grid_Main.Columns(2).HeaderText = "제품명"
        Grid_Main.Columns(3).HeaderText = "규격"
        Grid_Main.Columns(4).HeaderText = "단위"
        Grid_Main.Columns(5).HeaderText = "단가"
        Grid_Main.Columns(6).HeaderText = "Day1"
        Grid_Main.Columns(7).HeaderText = "Day2"
        Grid_Main.Columns(8).HeaderText = "Day3"
        Grid_Main.Columns(9).HeaderText = "Total"
        Grid_Main.Columns(10).HeaderText = "금액"
        Grid_Main.Columns(11).HeaderText = "비고"

        Dim i As Integer

        Grid_Main.Columns(0).Width = 60
        Grid_Main.Columns(1).Width = 150
        Grid_Main.Columns(2).Width = 150

        For i = 3 To Grid_Main.ColumnCount - 1
            Grid_Main.Columns(i).Width = 80
        Next i

        Grid_Main_Setup = True
    End Function

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        '검색
        Grid_Code_Display()
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
            Panel_Main.Visible = True
            Panel_Excel.Visible = False
            Com_Excel_Print.Visible = False
        End If

        Grid_Code.MultiSelect = False
        Grid_Code.ClearSelection()

    End Sub
    Public Function Grid_Code_Display() As Boolean

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Code.RowCount = 1
        Select Case Search_Combo.Text
            Case "코드"
                StrSQL = "Select CR_Code, CR_Date, CR_Customer_Name FROM SC_SALES with(NOLOCK) Where CR_Code Like '" & Search_Text.Text & "%'  Order  By CR_Code DESC"
            Case "날짜"
                StrSQL = "Select CR_Code, CR_Date, CR_Customer_Name FROM SC_SALES with(NOLOCK) Where CR_Date Like '" & Search_Text.Text & "%' Or CR_Date Is Null  Order By CR_Date"
            Case "거래처"
                StrSQL = "Select CR_Code, CR_Date, CR_Customer_Name FROM SC_SALES with(NOLOCK) Where CR_Customer_Name Like '" & Search_Text.Text & "%' Or CR_Customer_Name Is Null  Order By CR_Customer_Name"
        End Select

        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Code.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To Grid_Code.ColumnCount - 1
                    Grid_Code.Item(j, i).Value = DBT.Rows(i).Item(j)
                Next j
            Next i
        Else
            Grid_Code.RowCount = 1
            For j = 0 To Grid_Code.ColumnCount - 1
                Grid_Code.Item(j, 0).Value = ""
            Next j
        End If
        Grid_Code_Display = True

        Grid_Code.MultiSelect = False
        Grid_Code.ClearSelection()

    End Function

    Private Sub Insert_Com_Click(sender As Object, e As EventArgs) Handles Insert_Com.Click
        '새로운 코드 추가
        Dim DBT As New DataTable
        Dim JP_Code As Long
        Dim My_Date As String
        Dim My_Time As String

        StrSQL = "Select GETDATE() "
        Re_Count = DB_Select(DBT)


        If Re_Count = 0 Then
            Exit Sub
        Else
            My_Date = DBT.Rows(0).Item(0)
            JP_Code = Mid(My_Date, 1, 4) & Mid(My_Date, 6, 2) & Mid(My_Date, 9, 2)
            ' My_Time = Mid(My_Date, 12, 8)
            My_Time = DateTime.Now.ToString("HH:mm:ss")
            My_Date = Mid(My_Date, 1, 4) & "-" & Mid(My_Date, 6, 2) & "-" & Mid(My_Date, 9, 2)
        End If


        StrSQL = "Select CR_Code FROM SC_SALES with(NOLOCK) Where CR_Date Like  '" & My_Date & "%' Order By CR_Code Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            JP_Code = JP_Code & "001"
        Else
            JP_Code = JP_Code & Format(Val(Mid(DBT.Rows(0).Item(0), 11, 3)) + 1, "00#")
        End If

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO SC_SALES (CR_Code, CR_Date, CR_Time, CR_Name, CR_Day1, CR_Sel)  Values ('CC" & JP_Code & "', '" & My_Date & "','" & My_Time & "','" & Frm_Main.Text_Man_Name.Text & "', '" & My_Date & "','진행중' )"
        Re_Count = DB_Execute()
        Grid_Code_Display()
        '위치로 이동한다.


        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
            Panel_Main.Visible = True
            Grid_Info_Display(Grid_Info, "Select * FROM SC_SALES With(NOLOCK) Where CR_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "'")
            Label_Color(Com_Contract, Color_Label, Di_Panel2)
            Grid_Main_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
            Panel_Main.Visible = True
            Panel_Excel.Visible = False
            Com_Excel_Print.Visible = False
        End If

    End Sub

    Private Sub Delete_Com_Click(sender As Object, e As EventArgs) Handles Delete_Com.Click

        '해당 칼럼 삭제
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        If MsgBox("수주전표  '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "'를 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "수주전표 삭제") = vbNo Then
            Exit Sub
        End If
        '납품전표에 처리된 내용이 있으면 삭제 할 수 없다.
        Dim DBT As New DataTable

        StrSQL = "Select DL_Code FROM SC_DELIVER with(NOLOCK) Where DL_CR_Code Like '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Select(DBT)
        If Re_Count > 0 Then
            MsgBox("납품 전표에 처리된 내역이 있어 삭제 할 수 없습니다.")
            Exit Sub
        End If




        Grid_Display_Ch = False
        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SC_SALES Where CR_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' "
        Re_Count = DB_Execute()

        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SC_SALES_LIST Where CL_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' "
        Re_Count = DB_Execute()

        Grid_Code_Display()
        Grid_Info_Display(Grid_Info, "Select * FROM SC_SALES With(NOLOCK) Where CR_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "'")
        'Grid_Info_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True

    End Sub

    Private Sub Grid_Code_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Code.CellContentClick

    End Sub

    Private Sub Grid_Code_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Code.CellEnter
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
            Grid_Display_Ch = False
            Panel_Main.Visible = True
            Grid_Info_Display(Grid_Info, "Select * FROM SC_SALES With(NOLOCK) Where CR_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "'")
            Main_Grid_day()
            Label_Color(Com_Contract, Color_Label, Di_Panel2)
            Grid_Main_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
            Panel_Main.Visible = True
            Panel_Excel.Visible = False
            Com_Excel_Print.Visible = False
            Grid_Display_Ch = True
        End If

    End Sub
    'Public Function Grid_Info_Display(Code As String) As Boolean
    '    Grid_Display_Ch = False
    '    Grid_Info_Display = True
    '    Dim DBT As New DataTable
    '    Dim i As Integer
    '    StrSQL = "Select * FROM SC_SALES With(NOLOCK) Where CR_Code = '" & Code & "'"
    '    Re_Count = DB_Select(DBT)
    '    If Re_Count > 0 Then
    '        For i = 0 To DBT.Columns.Count - 1
    '            Grid_Info.Item(1, i).Value = DBT.Rows(0).Item(i).ToString
    '        Next i
    '    Else
    '        For i = 0 To DBT.Columns.Count - 1
    '            Grid_Info.Item(1, i).Value = ""
    '        Next i
    '    End If
    '    Main_Grid_day()
    '    Grid_Display_Ch = True

    'End Function

    Private Sub Grid_Info_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Info.CellContentClick

    End Sub

    Private Sub Grid_Info_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Info.CellEnter
        '선택되었을때 처리
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) < 1 Then
            Exit Sub
        End If
        'Car_Code 선택 처리
        Grid_Info_Row = Grid_Info.CurrentCell.RowIndex
        Grid_Info_Col = Grid_Info.CurrentCell.ColumnIndex
        Grid_Info_Combo.Visible = False
        If Grid_Info_Row < 0 Then
            Exit Sub
        End If
        If Grid_Info_Col < 0 Then
            Exit Sub
        End If
        Dim i As Integer
        Dim WW As Long
        Dim DBT As New DataTable
        If Grid_Info_Col = 1 Then
            Select Case Grid_Info_Row
                Case 4
                    StrSQL = "Select CM_Code  FROM SI_CUSTOMER With(NOLOCK) Order By CM_Name"
                    Re_Count = DB_Select(DBT)
                    Grid_Info_Combo.Items.Clear()
                    If Re_Count <> 0 Then
                        For i = 0 To Re_Count - 1
                            Grid_Info_Combo.Items.Add(DBT.Rows(i)("CM_Code"))
                        Next i
                    End If
                    Grid_Info_Combo.Top = Grid_Info.Top + Grid_Info.ColumnHeadersHeight + (Grid_Info_Row) * 24
                    WW = 0
                    For i = 0 To Grid_Info_Col - 1
                        WW = WW + Grid_Info.Columns(i).Width
                    Next
                    Grid_Info_Combo.Left = Grid_Info.Left + Grid_Info.RowHeadersWidth + WW
                    Grid_Info_Combo.Width = Grid_Info.Columns(Grid_Info_Col).Width
                    Grid_Info_Combo.Text = Grid_Info.CurrentCell.Value.ToString
                    Grid_Info_Combo.BackColor = Grid_Info.RowsDefaultCellStyle.BackColor
                    Grid_Info_Combo.Visible = True
                    Grid_Info_Combo.Focus.ToString()
                Case 5
                    StrSQL = "Select CM_Name  FROM SI_CUSTOMER With(NOLOCK) Order By CM_Name"
                    Re_Count = DB_Select(DBT)
                    Grid_Info_Combo.Items.Clear()
                    If Re_Count <> 0 Then
                        For i = 0 To Re_Count - 1
                            Grid_Info_Combo.Items.Add(DBT.Rows(i)("CM_Name"))
                        Next i
                    End If
                    Grid_Info_Combo.Top = Grid_Info.Top + Grid_Info.ColumnHeadersHeight + (Grid_Info_Row) * 24
                    WW = 0
                    For i = 0 To Grid_Info_Col - 1
                        WW = WW + Grid_Info.Columns(i).Width
                    Next
                    Grid_Info_Combo.Left = Grid_Info.Left + Grid_Info.RowHeadersWidth + WW
                    Grid_Info_Combo.Width = Grid_Info.Columns(Grid_Info_Col).Width
                    Grid_Info_Combo.Text = Grid_Info.CurrentCell.Value.ToString
                    Grid_Info_Combo.BackColor = Grid_Info.RowsDefaultCellStyle.BackColor
                    Grid_Info_Combo.Visible = True
                    Grid_Info_Combo.Focus.ToString()
                Case 8
                    StrSQL = "Select Base_Name  FROM SI_BASE_CODE With(NOLOCK) Where Base_Sel = '지불조건' Order By Base_Name"
                    Re_Count = DB_Select(DBT)
                    Grid_Info_Combo.Items.Clear()
                    If Re_Count <> 0 Then
                        For i = 0 To Re_Count - 1
                            Grid_Info_Combo.Items.Add(DBT.Rows(i)("Base_Name"))
                        Next i
                    End If
                    Grid_Info_Combo.Top = Grid_Info.Top + Grid_Info.ColumnHeadersHeight + (Grid_Info_Row) * 24
                    WW = 0
                    For i = 0 To Grid_Info_Col - 1
                        WW = WW + Grid_Info.Columns(i).Width
                    Next
                    Grid_Info_Combo.Left = Grid_Info.Left + Grid_Info.RowHeadersWidth + WW
                    Grid_Info_Combo.Width = Grid_Info.Columns(Grid_Info_Col).Width
                    Grid_Info_Combo.Text = Grid_Info.CurrentCell.Value.ToString
                    Grid_Info_Combo.BackColor = Grid_Info.RowsDefaultCellStyle.BackColor
                    Grid_Info_Combo.Visible = True
                    Grid_Info_Combo.Focus.ToString()
                Case 13
                    Grid_Info_Combo.Items.Clear()
                    Grid_Info_Combo.Items.Add("완료")
                    Grid_Info_Combo.Items.Add("진행중")

                    Grid_Info_Combo.Top = Grid_Info.Top + Grid_Info.ColumnHeadersHeight + (Grid_Info_Row) * 24
                    WW = 0
                    For i = 0 To Grid_Info_Col - 1
                        WW = WW + Grid_Info.Columns(i).Width
                    Next
                    Grid_Info_Combo.Left = Grid_Info.Left + Grid_Info.RowHeadersWidth + WW
                    Grid_Info_Combo.Width = Grid_Info.Columns(Grid_Info_Col).Width
                    Grid_Info_Combo.Text = Grid_Info.CurrentCell.Value
                    Grid_Info_Combo.BackColor = Grid_Info.RowsDefaultCellStyle.BackColor
                    Grid_Info_Combo.Visible = True
                    Grid_Info_Combo.Focus.ToString()
                Case Else
                    Exit Sub
            End Select
        End If
    End Sub

    Private Sub Grid_Info_Combo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Grid_Info_Combo.SelectedIndexChanged

    End Sub

    Private Sub Grid_Info_Combo_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Grid_Info_Combo.SelectionChangeCommitted
        Dim DBT As New DataTable
        Grid_Info.Item(Grid_Info_Col, Grid_Info_Row).Value = Grid_Info_Combo.SelectedItem.ToString()
        'Grid_Info.Item(0, 4).Value = "거래처코드"
        'Grid_Info.Item(0, 5).Value = "거래처명"
        'Grid_Info.Item(0, 6).Value = "담당자"
        'Grid_Info.Item(0, 7).Value = "연락처"
        'Grid_Info.Item(0, 8).Value = "지불조건"
        If Grid_Info_Row = 4 Then
            StrSQL = "Select CM_Code, CM_Name, CM_Man_Name, CM_Man_Tel  FROM SI_CUSTOMER With(NOLOCK) Where CM_Code = '" & Grid_Info_Combo.SelectedItem.ToString() & "'"
            Re_Count = DB_Select(DBT)
            Grid_Info_Combo.Items.Clear()
            If Re_Count <> 0 Then
                Grid_Info.Item(Grid_Info_Col, 5).Value = DBT.Rows(0)("CM_Name")
                Grid_Info.Item(Grid_Info_Col, 6).Value = DBT.Rows(0)("CM_Man_Name")
                Grid_Info.Item(Grid_Info_Col, 7).Value = DBT.Rows(0)("CM_Man_Tel")
                Grid_Info.Rows(5).HeaderCell.Value = "*"
                Grid_Info.Rows(6).HeaderCell.Value = "*"
                Grid_Info.Rows(7).HeaderCell.Value = "*"
            End If
        End If
        If Grid_Info_Row = 5 Then
            StrSQL = "Select CM_Code, CM_Name, CM_Man_Name, CM_Man_Tel  FROM SI_CUSTOMER With(NOLOCK) Where CM_Name = '" & Grid_Info_Combo.SelectedItem.ToString() & "'"
            Re_Count = DB_Select(DBT)
            Grid_Info_Combo.Items.Clear()
            If Re_Count <> 0 Then
                Grid_Info.Item(Grid_Info_Col, 4).Value = DBT.Rows(0)("CM_Code")
                Grid_Info.Item(Grid_Info_Col, 6).Value = DBT.Rows(0)("CM_Man_Name")
                Grid_Info.Item(Grid_Info_Col, 7).Value = DBT.Rows(0)("CM_Man_Tel")
                Grid_Info.Rows(4).HeaderCell.Value = "*"
                Grid_Info.Rows(6).HeaderCell.Value = "*"
                Grid_Info.Rows(7).HeaderCell.Value = "*"
            End If
        End If
        Grid_Info_Combo.Visible = False

    End Sub

    Private Sub Save_Com_Click(sender As Object, e As EventArgs) Handles Save_Com.Click
        Grid_Display_Ch = False
        Grid_Info_Save(Grid_Info, "Contract")
        Grid_Display_Ch = True

    End Sub

    Private Sub Grid_Info_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Info.CellValueChanged
        If Grid_Display_Ch = True Then
            Grid_Info.CurrentRow.HeaderCell.Value = "*"
            Main_Grid_day()
        End If

    End Sub


    Public Function Main_Grid_day() As Boolean
        Grid_Main.Columns(6).HeaderText = Grid_Info.Item(1, 9).Value.ToString
        Grid_Main.Columns(7).HeaderText = Grid_Info.Item(1, 10).Value.ToString
        Grid_Main.Columns(8).HeaderText = Grid_Info.Item(1, 11).Value.ToString
        Main_Grid_day = True

    End Function

    Private Sub Insert__Main_Click(sender As Object, e As EventArgs) Handles Insert__Main.Click
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        Grid_Display_Ch = False
        Dim DBT As New DataTable
        Dim Db_Sun As Long
        StrSQL = "Select CL_Sun FROM SC_SALES_LIST with(NOLOCK) Where CL_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' Order By Convert(Decimal,CL_Sun) Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            Db_Sun = 1
        Else
            Db_Sun = DBT.Rows(0).Item(0) + 1
        End If

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO SC_SALES_LIST (CL_Code, CL_Sun)  Values ('" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "', '" & Db_Sun & "')"
        Re_Count = DB_Execute()
        Grid_Main_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True

    End Sub
    Public Function Grid_Main_Display(PL_Code As String) As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display_Ch = False
        Grid_Main.RowCount = 0
        StrSQL = "Select CL_Sun, CL_PL_Code, CL_PL_Name, CL_Standard, CL_Unit, CL_Unit_Amount, CL_day01, CL_day02, CL_day03, CL_Total, CL_Gold, CL_Bigo  FROM SC_SALES_LIST with(NOLOCK)  Where CL_Code = '" & PL_Code & "' Order By Convert(Decimal,CL_Sun)"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Main.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To Grid_Main.ColumnCount - 1
                    Grid_Main.Item(j, i).Value = DBT.Rows(i).Item(j).ToString
                Next j
            Next i
        Else
            Grid_Main.RowCount = 1
            For j = 0 To Grid_Main.ColumnCount - 1
                Grid_Main.Item(j, 0).Value = ""
            Next j
        End If
        Main_Grid_day()

        Grid_Display_Ch = True
        Grid_Main_Display = True

    End Function

    Private Sub Grid_Main_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Main.CellValueChanged
        If Grid_Display_Ch = True Then
            Grid_Main.CurrentRow.HeaderCell.Value = "*"
            Grid_Main.Item(9, Grid_Main.CurrentCell.RowIndex).Value = Val(Grid_Main.Item(6, Grid_Main.CurrentCell.RowIndex).Value) + Val(Grid_Main.Item(7, Grid_Main.CurrentCell.RowIndex).Value) + Val(Grid_Main.Item(8, Grid_Main.CurrentCell.RowIndex).Value)
            Grid_Main.Item(10, Grid_Main.CurrentCell.RowIndex).Value = Val(Grid_Main.Item(9, Grid_Main.CurrentCell.RowIndex).Value) * Val(Grid_Main.Item(5, Grid_Main.CurrentCell.RowIndex).Value.ToString)
        End If

    End Sub
    Private Sub Grid_Main_MouseClick(sender As Object, e As MouseEventArgs) Handles Grid_Main.MouseClick


        Grid_Main_Row = Grid_Main.CurrentCell.RowIndex
        Grid_Main_Col = Grid_Main.CurrentCell.ColumnIndex
        Grid_Main_Combo.Visible = False
        If Grid_Main_Row < 0 Then
            Exit Sub
        End If
        If Grid_Main_Col < 0 Then
            Exit Sub
        End If
        Dim DBT As New DataTable

        Select Case Grid_Main_Col
            Case 1
                StrSQL = "Select SI_Product.PL_Code, SI_Product.PL_Name  FROM SI_PRODUCT with(NOLOCK), SI_PRODUCT_CUSTOMER with(NOLOCK) Where SI_PRODUCT_CUSTOMER.PL_CM_Code = '" & Grid_Info.Item(1, 4).Value & "' And SI_PRODUCT_CUSTOMER.PL_Code = SI_Product.PL_Code And SI_PRODUCT_CUSTOMER.PL_Sel = '매출' Order By PL_Code"
                Re_Count = DB_Select(DBT)
                Grid_Main_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        Grid_Main_Combo.Items.Add(DBT.Rows(i).Item(0))
                    Next i
                End If
                Grid_Main_Combo.Top = Grid_Main.Top + Grid_Main.ColumnHeadersHeight + (Grid_Main_Row) * 24
                Grid_Main_Combo.Left = Grid_Main.Left + Grid_Main.RowHeadersWidth + Grid_Main.Columns(0).Width
                Grid_Main_Combo.Width = Grid_Main.Columns(Grid_Main_Col).Width
                Grid_Main_Combo.Text = Grid_Main.CurrentCell.Value.ToString
                Grid_Main_Combo.BackColor = Grid_Main.RowsDefaultCellStyle.BackColor
                Grid_Main_Combo.Visible = True
                Grid_Main_Combo.Focus.ToString()
            Case 2
                StrSQL = "Select SI_Product.PL_Code, SI_Product.PL_Name  FROM SI_PRODUCT with(NOLOCK), SI_PRODUCT_CUSTOMER with(NOLOCK) Where SI_PRODUCT_CUSTOMER.PL_CM_Code = '" & Grid_Info.Item(1, 4).Value & "' And SI_PRODUCT_CUSTOMER.PL_Code = SI_Product.PL_Code And SI_PRODUCT_CUSTOMER.PL_Sel = '매출' Order By PL_Name"
                Re_Count = DB_Select(DBT)
                Grid_Main_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        Grid_Main_Combo.Items.Add(DBT.Rows(i).Item(1))
                    Next i
                End If
                Grid_Main_Combo.Top = Grid_Main.Top + Grid_Main.ColumnHeadersHeight + (Grid_Main_Row) * 24
                Grid_Main_Combo.Left = Grid_Main.Left + Grid_Main.RowHeadersWidth + Grid_Main.Columns(0).Width + Grid_Main.Columns(1).Width
                Grid_Main_Combo.Width = Grid_Main.Columns(Grid_Main_Col).Width
                Grid_Main_Combo.Text = Grid_Main.CurrentCell.Value.ToString
                Grid_Main_Combo.BackColor = Grid_Main.RowsDefaultCellStyle.BackColor
                Grid_Main_Combo.Visible = True
                Grid_Main_Combo.Focus.ToString()
        End Select


    End Sub
    Private Sub Grid_Main_Scroll(sender As Object, e As ScrollEventArgs) Handles Grid_Main.Scroll
        Grid_Main_Combo.Visible = False
    End Sub

    Private Sub Grid_Main_Combo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Grid_Main_Combo.SelectedIndexChanged

    End Sub

    Private Sub Grid_Main_Combo_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Grid_Main_Combo.SelectionChangeCommitted
        Dim DBT As New DataTable
        Grid_Main.Item(Grid_Main_Col, Grid_Main_Row).Value = Grid_Main_Combo.SelectedItem.ToString()
        Select Case Grid_Main_Col
            Case 1
                StrSQL = "Select PL_Code, PL_Name, PL_Standard ,PL_Unit, PL_Unit_Gold FROM SI_PRODUCT with(NOLOCK) Where PL_Code = '" & Grid_Main_Combo.SelectedItem.ToString() & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Main.Item(2, Grid_Main_Row).Value = DBT.Rows(0).Item(1)
                    Grid_Main.Item(3, Grid_Main_Row).Value = DBT.Rows(0).Item(2)
                    Grid_Main.Item(4, Grid_Main_Row).Value = DBT.Rows(0).Item(3)
                    Grid_Main.Item(5, Grid_Main_Row).Value = DBT.Rows(0).Item(4)

                End If
            Case 2
                StrSQL = "Select PL_Code, PL_Name, PL_Standard ,PL_Unit, PL_Unit_Gold  FROM SI_PRODUCT with(NOLOCK) Where PL_Name= '" & Replace(Grid_Main_Combo.SelectedItem.ToString(), "'", "''") & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Main.Item(1, Grid_Main_Row).Value = DBT.Rows(0).Item(0)
                    Grid_Main.Item(3, Grid_Main_Row).Value = DBT.Rows(0).Item(2)
                    Grid_Main.Item(4, Grid_Main_Row).Value = DBT.Rows(0).Item(3)
                    Grid_Main.Item(5, Grid_Main_Row).Value = DBT.Rows(0).Item(4)
                End If
        End Select

        Grid_Main_Combo.Visible = False
        'Grid_Main.Columns(1).HeaderText = "제품코드"
        'Grid_Main.Columns(2).HeaderText = "제품명"
        'Grid_Main.Columns(3).HeaderText = "규격"
        'Grid_Main.Columns(4).HeaderText = "단위"
        'Grid_Main.Columns(5).HeaderText = "단가"
    End Sub

    Private Sub Grid_Main_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Main.CellContentClick

    End Sub

    Private Sub Save_Main_Click(sender As Object, e As EventArgs) Handles Save_Main.Click
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        Dim i As Integer

        Grid_Display_Ch = False
        For i = 0 To Grid_Main.Rows.Count - 1
            If Grid_Main.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SC_SALES_LIST Set CL_PL_Code = '" & Grid_Main.Item(1, i).Value & "', CL_PL_Name = '" & Grid_Main.Item(2, i).Value & "', CL_Standard = '" & Grid_Main.Item(3, i).Value & "', CL_Unit = '" & Grid_Main.Item(4, i).Value & "', CL_Unit_Amount = '" & Grid_Main.Item(5, i).Value & "', CL_day01 = '" & Grid_Main.Item(6, i).Value & "', CL_day02 = '" & Grid_Main.Item(7, i).Value & "', CL_day03 = '" & Grid_Main.Item(8, i).Value & "', CL_Total = '" & Grid_Main.Item(9, i).Value & "',CL_Gold = '" & Grid_Main.Item(10, i).Value & "',CL_Bigo = '" & Grid_Main.Item(11, i).Value & "' where CL_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And CL_Sun = '" & Grid_Main.Item(0, i).Value & "'"
                Re_Count = DB_Execute()
                Grid_Main.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display_Ch = True

    End Sub
    Private Sub Grid_Main_Combo_LostFocus(sender As Object, e As EventArgs) Handles Grid_Main_Combo.LostFocus
        Grid_Main_Combo.Visible = False
    End Sub

    Private Sub Delete__Main_Click(sender As Object, e As EventArgs) Handles Delete__Main.Click
        '해당 칼럼 삭제
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        Grid_Display_Ch = False
        Dim i As Integer
        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SC_SALES_LIST Where CL_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And  CL_Sun = '" & Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        For i = Grid_Main.CurrentCell.RowIndex + 1 To Grid_Main.RowCount - 1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "Update  Contract_List Set CL_Sun = '" & i & "' Where CL_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And  CL_Sun = '" & i + 1 & "'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        Grid_Main_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True

    End Sub

    Private Sub Com_Excel_Click(sender As Object, e As EventArgs) Handles Com_Excel.Click
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        Label_Color(sender, Color_Label, Di_Panel2)
        '판넬 보이기
        Panel_Main.Visible = False
        Panel_Excel.Visible = True
        Com_Excel_Print.Visible = True
        Cursor = Cursors.AppStarting
        If Excel_Check = True Then
            xlbook.Close()
            xlapp.Quit()
            xlsheet = Nothing
            xlbook = Nothing
            xlapp = Nothing
            releaseObject(xlsheet)
            releaseObject(xlbook)
            releaseObject(xlapp)
            Excel_Check = False
        End If
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Long


        xlapp = New Excel.Application
        xlbook = xlapp.Workbooks.Open(Application.StartupPath & "\excel\수주.xls")
        xlsheet = xlbook.ActiveSheet
        xlapp.DisplayAlerts = False
        xlapp.WindowState = Excel.XlWindowState.xlMinimized
        xlapp.Visible = False
        xlapp.DisplayFormulaBar = False
        xlapp.DisplayStatusBar = False
        xlapp.ActiveWindow.DisplayWorkbookTabs = False
        xlapp.ActiveWindow.DisplayHeadings = False
        SetParent(xlapp.Hwnd, Panel_Excel.Handle)
        SendMessage(xlapp.Hwnd, WM_SYSCOMMAND, SC_MAXIMIZE, 0)
        '지불조건
        xlsheet.Cells(5, 4).Value = Grid_Info.Item(1, 8).Value
        '담당자 
        xlsheet.Cells(6, 4).Value = Grid_Info.Item(1, 6).Value
        '발주일자
        xlsheet.Cells(4, 14).Value = Grid_Info.Item(1, 2).Value
        '발주코드
        xlsheet.Cells(5, 14).Value = Grid_Info.Item(1, 0).Value
        '담당자 전화 번로
        xlsheet.Cells(6, 14).Value = Grid_Info.Item(1, 7).Value
        '일자
        xlsheet.Cells(10, 9).Value = Mid(Grid_Info.Item(1, 9).Value, 1, 4) & " " & Mid(Grid_Info.Item(1, 9).Value, 6, 5)
        xlsheet.Cells(10, 10).Value = Mid(Grid_Info.Item(1, 10).Value, 1, 4) & " " & Mid(Grid_Info.Item(1, 10).Value, 6, 5)
        xlsheet.Cells(10, 11).Value = Mid(Grid_Info.Item(1, 11).Value, 1, 4) & " " & Mid(Grid_Info.Item(1, 11).Value, 6, 5)
        j = 0
        For i = 0 To Grid_Main.RowCount - 1
            xlsheet.Cells(12 + i, 2).Value = Grid_Main.Item(0, i).Value.ToString
            xlsheet.Cells(12 + i, 3).Value = Grid_Main.Item(1, i).Value.ToString
            xlsheet.Cells(12 + i, 4).Value = Grid_Main.Item(2, i).Value.ToString
            xlsheet.Cells(12 + i, 5).Value = Grid_Main.Item(3, i).Value.ToString
            xlsheet.Cells(12 + i, 7).Value = Grid_Main.Item(4, i).Value.ToString
            xlsheet.Cells(12 + i, 8).Value = Grid_Main.Item(5, i).Value.ToString
            xlsheet.Cells(12 + i, 9).Value = Grid_Main.Item(6, i).Value.ToString
            xlsheet.Cells(12 + i, 10).Value = Grid_Main.Item(7, i).Value.ToString
            xlsheet.Cells(12 + i, 11).Value = Grid_Main.Item(8, i).Value.ToString
            xlsheet.Cells(12 + i, 12).Value = Grid_Main.Item(9, i).Value.ToString
            xlsheet.Cells(12 + i, 13).Value = Grid_Main.Item(10, i).Value.ToString
            xlsheet.Cells(12 + i, 15).Value = Grid_Main.Item(11, i).Value.ToString
            j = j + Val(Grid_Main.Item(9, i).Value.ToString)
        Next i
        '합계
        xlsheet.Cells(27, 6).Value = j
        '거래처명
        xlsheet.Cells(29, 4).Value = Grid_Info.Item(1, 5).Value
        '거래처 주소
        StrSQL = "Select CM_Add1, CM_Add2  FROM SI_CUSTOMER with(NOLOCK) Where CM_Code= '" & Grid_Info.Item(1, 4).Value & "'"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            xlsheet.Cells(41, 5).Value = DBT.Rows(0).Item(0) & " " & DBT.Rows(0).Item(1)
        End If
        xlbook.SaveAs(Application.StartupPath & "\Excel\수주\" & Grid_Info.Item(1, 0).Value & Grid_Info.Item(1, 5).Value & ".xlsx")
        Cursor = Cursors.Default
        Excel_Check = True
    End Sub

    Private Sub Com_Contract_Click(sender As Object, e As EventArgs) Handles Com_Contract.Click
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        Label_Color(sender, Color_Label, Di_Panel2)
        '판넬 보이기
        Panel_Main.Visible = True
        Panel_Excel.Visible = False
        Com_Excel_Print.Visible = False


        'Print_Com.Visible = False
        'Excel_Save_Com.Visible = False

    End Sub

    Private Sub Com_Excel_Print_Click(sender As Object, e As EventArgs) Handles Com_Excel_Print.Click
        If Excel_Check = True Then
            xlsheet.PrintOut(From:=1, To:=1, Copies:=1, Collate:=True)
        End If

    End Sub
    Private Sub Grid_Info_Combo_LostFocus(sender As Object, e As EventArgs) Handles Grid_Info_Combo.LostFocus
        Grid_Info_Combo.Visible = False
    End Sub
    Private Sub Com_Excel_Print_VisibleChanged(sender As Object, e As EventArgs) Handles Com_Excel_Print.VisibleChanged
        If Excel_Check = True Then
            xlbook.Close()
            xlapp.Quit()
            xlsheet = Nothing
            xlbook = Nothing
            xlapp = Nothing
            releaseObject(xlsheet)
            releaseObject(xlbook)
            releaseObject(xlapp)
            Excel_Check = False
        End If
    End Sub

End Class
