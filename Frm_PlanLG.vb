Imports System.Text
Imports Excel = Microsoft.Office.Interop.Excel

Public Class Frm_PlanLG
    Dim Grid_Display_Ch As Boolean
    Dim Grid_Info_Row As Integer
    Dim Grid_Info_Col As Integer
    Dim xlapp As Excel.Application
    Dim xlbook As Excel.Workbook
    Dim xlsheet As Excel.Worksheet
    Dim Excel_Check As Boolean

    Private Sub Frm_PlanLG_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()
        Grid_Code_Setup()
        Grid_Info_Setup()
        Grid_LGPlan_Setup()

        'LGPlan_2
        Search_Combo.Items.Add("코드")
        Search_Combo.Items.Add("날짜")
        Search_Combo.Text = "코드"

        LGPlan_Panel.Top = 136
        LGPlan_Panel.Left = 389

        LGPlan_Panel.Visible = True

    End Sub
    Public Function Grid_Code_Setup() As Boolean
        Grid_Color(Grid_Code)
        'Grid_Code.GridColor = Color.White
        'Grid_Code.BackgroundColor = Color.White
        'Grid_Code.EnableHeadersVisualStyles = False
        'Grid_Code.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(91, 155, 213)
        'Grid_Code.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        'Grid_Code.RowsDefaultCellStyle.BackColor = Color.FromArgb(210, 222, 239)
        'Grid_Code.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(234, 239, 247)

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
        Grid_Code.Columns(2).HeaderText = "담당자"

        Dim i As Integer
        For i = 0 To Grid_Code.ColumnCount - 1
            Grid_Code.Columns(i).ReadOnly = True
        Next i
        Grid_Code_Setup = True
    End Function

    Public Function Grid_Info_Setup() As Boolean
        Grid_Color(Grid_Info)
        'Grid_Info.GridColor = Color.White
        'Grid_Info.BackgroundColor = Color.White
        'Grid_Info.EnableHeadersVisualStyles = False
        'Grid_Info.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(91, 155, 213)
        'Grid_Code.ColumnHeadersDefaultCellStyle.ForeColor = Color.White

        'Grid_Info.RowsDefaultCellStyle.BackColor = Color.FromArgb(210, 222, 239)
        'Grid_Info.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(234, 239, 247)


        Grid_Info.AllowUserToAddRows = False
        Grid_Info.RowTemplate.Height = 24
        Grid_Info.ColumnCount = 2
        Grid_Info.RowCount = 7





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
        Grid_Info.Item(0, 4).Value = "업체"
        Grid_Info.Item(0, 5).Value = "주문날짜"
        Grid_Info.Item(0, 6).Value = "비고"

        'Grid_Info.Columns(0).ReadOnly = True
        'Grid_Info.Columns(1).ReadOnly = True

        Grid_Info.Rows(0).ReadOnly = True
        Grid_Info.Rows(1).ReadOnly = True
        Grid_Info.Rows(2).ReadOnly = True
        Grid_Info.Rows(3).ReadOnly = True
        Grid_Info.Rows(4).ReadOnly = False

        Grid_Info_Setup = True
    End Function

    Public Function Grid_LGPlan_Setup() As Boolean
        Grid_Color(Grid_LGPlan)
        'Grid_LGPlan.GridColor = Color.White
        'Grid_LGPlan.BackgroundColor = Color.White
        'Grid_LGPlan.EnableHeadersVisualStyles = False
        'Grid_LGPlan.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(91, 155, 213)
        'Grid_LGPlan.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        'Grid_LGPlan.RowsDefaultCellStyle.BackColor = Color.FromArgb(210, 222, 239)
        'Grid_LGPlan.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(234, 239, 247)

        Grid_LGPlan.AllowUserToAddRows = False
        Grid_LGPlan.RowTemplate.Height = 20
        Grid_LGPlan.ColumnCount = 43
        'Grid_LGPlan.RowCount = 2500
        Grid_LGPlan.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_LGPlan.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_LGPlan.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_LGPlan.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_LGPlan.RowHeadersWidth = 10
        Grid_LGPlan.Columns(0).Width = 60


        Grid_LGPlan.Columns(0).HeaderText = "순번"
        Grid_LGPlan.Columns(1).HeaderText = "Line"
        Grid_LGPlan.Columns(2).HeaderText = "그룹"
        Grid_LGPlan.Columns(3).HeaderText = "주문순번"
        Grid_LGPlan.Columns(4).HeaderText = "PartNo"
        Grid_LGPlan.Columns(5).HeaderText = "BuyerModel"
        Grid_LGPlan.Columns(6).HeaderText = "Total"
        Grid_LGPlan.Columns(7).HeaderText = "남은 Qty"
        Grid_LGPlan.Columns(8).HeaderText = "day01"
        Grid_LGPlan.Columns(9).HeaderText = "day02"
        Grid_LGPlan.Columns(10).HeaderText = "day03"
        Grid_LGPlan.Columns(11).HeaderText = "day04"
        Grid_LGPlan.Columns(12).HeaderText = "day05"
        Grid_LGPlan.Columns(13).HeaderText = "day06"
        Grid_LGPlan.Columns(14).HeaderText = "day07"
        Grid_LGPlan.Columns(15).HeaderText = "day08"
        Grid_LGPlan.Columns(16).HeaderText = "day09"
        Grid_LGPlan.Columns(17).HeaderText = "day10"
        Grid_LGPlan.Columns(18).HeaderText = "day11"
        Grid_LGPlan.Columns(19).HeaderText = "day12"
        Grid_LGPlan.Columns(20).HeaderText = "day13"
        Grid_LGPlan.Columns(21).HeaderText = "day14"
        Grid_LGPlan.Columns(22).HeaderText = "day15"
        Grid_LGPlan.Columns(23).HeaderText = "day16"
        Grid_LGPlan.Columns(24).HeaderText = "day17"
        Grid_LGPlan.Columns(25).HeaderText = "day18"
        Grid_LGPlan.Columns(26).HeaderText = "day19"
        Grid_LGPlan.Columns(27).HeaderText = "day20"
        Grid_LGPlan.Columns(28).HeaderText = "day21"
        Grid_LGPlan.Columns(29).HeaderText = "day22"
        Grid_LGPlan.Columns(30).HeaderText = "P/Period"
        Grid_LGPlan.Columns(31).HeaderText = "우선순위"
        Grid_LGPlan.Columns(32).HeaderText = "예정일"
        Grid_LGPlan.Columns(33).HeaderText = "MKT"
        Grid_LGPlan.Columns(34).HeaderText = "Buyer"
        Grid_LGPlan.Columns(35).HeaderText = "MixSeqNo"
        Grid_LGPlan.Columns(36).HeaderText = "Mixinfo"
        Grid_LGPlan.Columns(37).HeaderText = "MixRate"
        Grid_LGPlan.Columns(38).HeaderText = "sqlitWO"
        Grid_LGPlan.Columns(39).HeaderText = "시작시간"
        Grid_LGPlan.Columns(40).HeaderText = "지시시작번호"
        Grid_LGPlan.Columns(41).HeaderText = "지시종료"
        Grid_LGPlan.Columns(42).HeaderText = "Tool"


        Dim i As Integer
        For i = 1 To Grid_LGPlan.ColumnCount - 1
            Grid_LGPlan.Columns(i).Width = 80
            'Grid_LGPlan.Columns(i).ReadOnly = True
        Next i


        '일자 표시
        Dim My_Date_Date As Date
        Dim My_Date As String
        Dim DBT As New DataTable
        StrSQL = "Select GETDATE() "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            Exit Function
        Else
            My_Date = DBT.Rows(0).Item(0)
            My_Date_Date = My_Date
            My_Date = Mid(My_Date, 1, 4) & "-" & Mid(My_Date, 6, 2) & "-" & Mid(My_Date, 9, 2)
        End If
        Dim j As Integer
        j = 0
        'For i = 0 To 18
        '    My_Date_Date = DateAdd(DateInterval.Day, 1, My_Date_Date)
        '    My_Date = Format(My_Date_Date, "yyyy-MM-dd")
        '    Grid_LGPlan.Columns(i + 11).HeaderText = My_Date

        'Next i
        For i = 0 To 21
            My_Date_Date = DateAdd(DateInterval.Day, 1, My_Date_Date)
            My_Date = Format(My_Date_Date, "yyyy-MM-dd")
            Grid_LGPlan.Columns(i + 8).HeaderText = My_Date

        Next i

        Grid_LGPlan_Setup = True



    End Function


    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        '검색
        Grid_Code_Display()
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
            LGPlan_Panel.Visible = True
            Grid_Info_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
            Grid_LGPlan_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)


        End If

    End Sub
    Public Function Grid_Code_Display() As Boolean

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Code.RowCount = 1
        Select Case Search_Combo.Text
            Case "코드"
                StrSQL = "Select PL_Code, PL_Date, PL_Name From Plan_LG_List with(NOLOCK) Where PL_Code Like '%" & Search_Text.Text & "%'  And PL_Customer = 'LG' Order  By PL_Code DESC"
            Case "날짜"
                StrSQL = "Select PL_Code, PL_Date, PL_Name From Plan_LG_List with(NOLOCK) Where PL_Code Like '%" & Search_Text.Text & "%' Or PL_Date Is not Null  And PL_Customer = 'LG' Order By PL_Date"
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
            My_Time = DateTime.Now.ToString("HH:mm:ss")
            My_Date = Mid(My_Date, 1, 4) & "-" & Mid(My_Date, 6, 2) & "-" & Mid(My_Date, 9, 2)
        End If


        StrSQL = "Select PL_Code From Plan_LG_List with(NOLOCK) Where PL_Date Like  '%" & My_Date & "%' Order By PL_Code Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            JP_Code = JP_Code & "001"
        Else
            JP_Code = JP_Code & Format(Val(Mid(DBT.Rows(0).Item(0), 11, 3)) + 1, "00#")
        End If

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert into Plan_LG_List (PL_Code, PL_Date, PL_Time, PL_Name, PL_Customer)  Values ('LG" & JP_Code & "', '" & My_Date & "','" & My_Time & "','" & Frm_Main.Text_Man_Name.Text & "','LG' )"
        Re_Count = DB_Execute()
        Grid_Code_Display()
        LGPlan_Panel.Visible = True
        Grid_Info_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
        Grid_LGPlan_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)


    End Sub

    Private Sub Grid_Code_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Code.CellEnter
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
            LGPlan_Panel.Visible = True
            Grid_Info_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
            Grid_LGPlan_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)


        End If

    End Sub
    Public Function Grid_Info_Display(Code As String) As Boolean
        Grid_Display_Ch = False

        Grid_Info_Display = True
        Dim DBT As New DataTable
        Dim i As Integer
        StrSQL = "Select * From Plan_LG_List With(NOLOCK) Where PL_Code = '" & Code & "'"
        Re_Count = DB_Select(DBT)
        If Re_Count > 0 Then
            For i = 0 To DBT.Columns.Count - 1
                Grid_Info.Item(1, i).Value = DBT.Rows(0).Item(i).ToString
            Next i
        Else
            For i = 0 To DBT.Columns.Count - 1
                Grid_Info.Item(1, i).Value = ""
            Next i
        End If

        Grid_Info.Item(1, 1).ReadOnly = True
        Grid_Info.Item(1, 2).ReadOnly = True
        Grid_Info.Item(1, 3).ReadOnly = True
        Grid_Info.Item(1, 4).ReadOnly = True

        Grid_Display_Ch = True

    End Function




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
                Case Else
                    Exit Sub
            End Select
        End If

    End Sub

    Private Sub Grid_Info_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Info.CellValueChanged
        If Grid_Display_Ch = True Then
            Grid_Info.CurrentRow.HeaderCell.Value = "*"
        End If

    End Sub

    Private Sub Grid_Info_SelectionChanged(sender As Object, e As EventArgs) Handles Grid_Info.SelectionChanged
        Grid_Info_Combo.Visible = False

    End Sub


    Private Sub Grid_Info_Combo_LostFocus(sender As Object, e As EventArgs) Handles Grid_Info_Combo.LostFocus
        Grid_Info_Combo.Visible = False

    End Sub

    Private Sub Grid_Info_Combo_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Grid_Info_Combo.SelectionChangeCommitted
        Dim DBT As New DataTable
        Grid_Info.Item(Grid_Info_Col, Grid_Info_Row).Value = Grid_Info_Combo.SelectedItem.ToString()

    End Sub

    Private Sub Save_Com_Click(sender As Object, e As EventArgs) Handles Save_Com.Click
        'Grid_Info 를 저장 한다.
        'DataBase에서 필드 이름 가지고 오기

        Dim Table_Name As String
        Dim j As Integer
        Dim DBT As New DataTable
        Dim Field_Name(500) As String
        Dim i As Integer
        j = 0

        If Grid_Code.Item(0, 0).Value = "" Then
            MsgBox("공백은 저장할 수 없습니다")
            Exit Sub
        End If


        For i = 1 To Grid_Info.RowCount - 1
            If Grid_Info.Rows(i).HeaderCell.Value = "*" Then
                j = 1
            End If
        Next i
        If j = 0 Then
        Else
            Table_Name = "Plan_LG_List"
            j = 0
            StrSQL = "Select sys.Columns.Name From sys.Tables With(NOLOCK) , sys.Columns With(NOLOCK) Where sys.Tables.name = '" & Table_Name & "' And sys.Tables.Object_id = sys.Columns.Object_id Order By sys.Columns.column_id"
            Re_Count = DB_Select(DBT)
            If Re_Count <> 0 Then
                For i = 0 To DBT.Rows.Count - 1
                    Field_Name(j) = DBT.Rows(i)("Name").ToString
                    j = j + 1
                Next i
            End If
            j = j - 1
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "UPDATE " & Table_Name & " SET "
            Field_Name(500) = "Ok"
            For i = 1 To j
                If Grid_Info.Rows(i).HeaderCell.Value = "*" Then
                    If Field_Name(500) = "" Then
                        StrSQL = StrSQL & ","
                    End If
                    StrSQL = StrSQL & " " & Table_Name & "." & Field_Name(i) & " = '" & Replace(Grid_Info.Item(1, i).Value, "'", "''") & "'"
                    If Field_Name(500) = "Ok" Then
                        Field_Name(500) = ""
                    End If
                End If
            Next i
            StrSQL = StrSQL & " WHERE " & Table_Name & "." & Field_Name(0) & " = '" & Grid_Info.Item(1, 0).Value & "'"
            Re_Count = DB_Execute()

        End If




        'Grid_LGPlan 저장
        '이미 Data 가 있으면 저장 필요가 없다
        StrSQL = "Select *  From Plan_LG with(NOLOCK)  Where PL_Code = '" & Grid_Info.Item(1, 0).Value & "' Order By Convert(Decimal,PL_Sun)"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            ' MsgBox("이미 저장 처리가 되었습니다.")

            For i = 0 To Grid_LGPlan.RowCount - 1
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE Plan_LG SET "

                StrSQL = StrSQL & "PL_ScheduleGroup = '" & Grid_LGPlan.Item(2, i).Value & "'"
                StrSQL = StrSQL & "PL_WorkOrder = '" & Grid_LGPlan.Item(3, i).Value & "'"
                StrSQL = StrSQL & "PL_BuyerModel = '" & Grid_LGPlan.Item(5, i).Value & "'"
                StrSQL = StrSQL & "PL_Total = '" & Grid_LGPlan.Item(6, i).Value & "'"
                StrSQL = StrSQL & "PL_RemainedQty = '" & Grid_LGPlan.Item(7, i).Value & "'"
                StrSQL = StrSQL & "PL_day01 = '" & Grid_LGPlan.Item(8, i).Value & "'"
                StrSQL = StrSQL & "PL_day02 = '" & Grid_LGPlan.Item(9, i).Value & "'"
                StrSQL = StrSQL & "PL_day03 = '" & Grid_LGPlan.Item(10, i).Value & "'"

                StrSQL = StrSQL & "PL_day04 = '" & Grid_LGPlan.Item(11, i).Value & "'"
                StrSQL = StrSQL & "PL_day05 = '" & Grid_LGPlan.Item(12, i).Value & "'"
                StrSQL = StrSQL & "PL_day06 = '" & Grid_LGPlan.Item(13, i).Value & "'"
                StrSQL = StrSQL & "PL_day07 = '" & Grid_LGPlan.Item(14, i).Value & "'"
                StrSQL = StrSQL & "PL_day08 = '" & Grid_LGPlan.Item(15, i).Value & "'"
                StrSQL = StrSQL & "PL_day09 = '" & Grid_LGPlan.Item(16, i).Value & "'"
                StrSQL = StrSQL & "PL_day10 = '" & Grid_LGPlan.Item(17, i).Value & "'"
                StrSQL = StrSQL & "PL_day11 = '" & Grid_LGPlan.Item(18, i).Value & "'"
                StrSQL = StrSQL & "PL_day12 = '" & Grid_LGPlan.Item(19, i).Value & "'"
                StrSQL = StrSQL & "PL_day13 = '" & Grid_LGPlan.Item(20, i).Value & "'"

                StrSQL = StrSQL & "PL_day14 = '" & Grid_LGPlan.Item(21, i).Value & "'"
                    StrSQL = StrSQL & "PL_day15 = '" & Grid_LGPlan.Item(22, i).Value & "'"
                    StrSQL = StrSQL & "PL_day16 = '" & Grid_LGPlan.Item(23, i).Value & "'"
                    StrSQL = StrSQL & "PL_day17 = '" & Grid_LGPlan.Item(24, i).Value & "'"
                    StrSQL = StrSQL & "PL_day18 = '" & Grid_LGPlan.Item(25, i).Value & "'"
                    StrSQL = StrSQL & "PL_day19 = '" & Grid_LGPlan.Item(26, i).Value & "'"
                    StrSQL = StrSQL & "PL_day20 = '" & Grid_LGPlan.Item(27, i).Value & "'"

                    StrSQL = StrSQL & "PL_day21 = '" & Grid_LGPlan.Item(28, i).Value & "'"
                    StrSQL = StrSQL & "PL_day22 = '" & Grid_LGPlan.Item(29, i).Value & "'"
                    StrSQL = StrSQL & "PL_Prod_Period = '" & Grid_LGPlan.Item(30, i).Value & "'"
                    StrSQL = StrSQL & "PL_Priority = '" & Grid_LGPlan.Item(31, i).Value & "'"
                    StrSQL = StrSQL & "PL_NeedByDate = '" & Grid_LGPlan.Item(32, i).Value & "'"
                    StrSQL = StrSQL & "PL_MKT = '" & Grid_LGPlan.Item(33, i).Value & "'"
                    StrSQL = StrSQL & "PL_Buyer = '" & Grid_LGPlan.Item(34, i).Value & "'"
                    StrSQL = StrSQL & "PL_MixSeqNo = '" & Grid_LGPlan.Item(35, i).Value & "'"
                    StrSQL = StrSQL & "PL_MixInfo = '" & Grid_LGPlan.Item(36, i).Value & "'"
                    StrSQL = StrSQL & "PL_MixRate = '" & Grid_LGPlan.Item(37, i).Value & "'"
                    StrSQL = StrSQL & "PL_SqlitWO = '" & Grid_LGPlan.Item(38, i).Value & "'"
                    StrSQL = StrSQL & "PL_StartTime = '" & Grid_LGPlan.Item(39, i).Value & "'"
                    StrSQL = StrSQL & "PL_ProductionStartNo = '" & Grid_LGPlan.Item(40, i).Value & "'"
                    StrSQL = StrSQL & "PL_ProductionEnd = '" & Grid_LGPlan.Item(41, i).Value & "'"
                    StrSQL = StrSQL & "PL_Tool = '" & Grid_LGPlan.Item(42, i).Value & "'"

                    StrSQL = StrSQL & "WHERE PL_Code = '" & Grid_Info.Item(1, 0).Value & "' AND PL_Sun ='" & Grid_LGPlan.Item(0, i).Value & "'"
                    StrSQL = StrSQL & "AND PL_Line = '" & Grid_LGPlan.Item(1, i).Value & "' AND PL_PartNo ='" & Grid_LGPlan.Item(4, i).Value & "'"

                StrSQL = StrSQL & ")"
                Re_Count = DB_Execute()
            Next i
            MessageBox.Show("기존 내용이 수정되었습니다")
            Exit Sub

        End If
        '저장(추가)한다.
        If Len(Grid_LGPlan.Item(0, 0).Value) > 0 Then
        Else
            Exit Sub
        End If
        For i = 0 To Grid_LGPlan.RowCount - 1
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "Insert Into Plan_LG Values ('" & Grid_Info.Item(1, 0).Value & "'"
            For j = 0 To Grid_LGPlan.ColumnCount - 1
                StrSQL = StrSQL & ",'" & Replace(Grid_LGPlan.Item(j, i).Value, "'", "''") & "'"
            Next j
            StrSQL = StrSQL & ")"
            Re_Count = DB_Execute()
        Next i

        MessageBox.Show("저장되었습니다")


    End Sub

    Private Sub Load_Com_Click(sender As Object, e As EventArgs) Handles Load_Com.Click
        Dim DBT As New DataTable
        Dim PL_Code As String
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            MsgBox("코드번호를 선택해주세요.")
            Exit Sub
        End If
        PL_Code = Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value
        '기존에 Data 있는가?
        StrSQL = "Select *  From Plan_LG with(NOLOCK)  Where PL_Code = '" & PL_Code & "' Order By Convert(Decimal,PL_Sun)"
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
        Else
            MsgBox("이미 저장 처리 되었습니다.")
            Exit Sub
        End If


        'tsv파일(LG일정계획 다운로드)을 불러온다.
        OpenFileDialog1.InitialDirectory = Application.StartupPath
        OpenFileDialog1.Filter = "xlsx file(*.xlsx)|*.xlsx|All files(*,*)|*.*"
        OpenFileDialog1.FilterIndex = 0
        OpenFileDialog1.RestoreDirectory = True
        Dim i As Integer
        Dim j As Integer
        Dim Ex_Count As Integer

        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            Try
                Dim xlapp As Excel.Application
                Dim xlbook As Excel.Workbook
                Dim xlsheet As Excel.Worksheet
                xlapp = New Excel.Application
                xlbook = xlapp.Workbooks.Open(OpenFileDialog1.FileName)
                'xlsheet = xlbook.ActiveSheet
                'xlsheet = xlbook.Sheets(2)
                xlsheet = xlbook.Sheets(1)
                xlapp.DisplayAlerts = False
                xlapp.Visible = False
                Ex_Count = -1
                'Grid_LGPlan.RowCount = 2500
                Grid_LGPlan.RowCount = xlsheet.UsedRange.Rows.Count
                'Grid_LGPlan.RowCount = 10
                For i = 2 To xlsheet.UsedRange.Rows.Count - 1

                    If Len(xlsheet.Cells(i, 1).Value) = 0 Then
                        Grid_LGPlan.RowCount = i - 1
                        Exit For
                    End If
                    If Strings.Right(xlsheet.Cells(i, 1).Value, 5) = "Total" Then
                    Else
                        Ex_Count = Ex_Count + 1
                        Grid_LGPlan.Item(0, Ex_Count).Value = Ex_Count + 1
                        For j = 1 To 42
                            Grid_LGPlan.Item(j, Ex_Count).Value = xlsheet.Cells(i, j + 1).Value
                            'If j <= 8 Then
                            '    If j = 7 Then
                            '        Grid_LGPlan.Item(j, Ex_Count).Value = ""
                            '    Else
                            '        Grid_LGPlan.Item(j, Ex_Count).Value = xlsheet.Cells(i, j + 1).Value
                            '    End If
                            'Else
                            '    '   Grid_LGPlan.Item(j + 2, Ex_Count).Value = xlsheet.Cells(i, j).Value
                            '    Grid_LGPlan.Item(j, Ex_Count).Value = xlsheet.Cells(i, j + 1).Value
                            'End If
                        Next j
                        ''상위 코드 검색
                        ''   StrSQL = "Select Product_List.PL_Code, Product_List.PL_Name FROM SI_PRODUCT_List With(NOLOCK), Product_LG With(NOLOCK) Where Product_LG.PL_Code = Product_List.PL_Code And Product_LG.PL_LG_Code  = '" & Grid_LGPlan.Item(2, Ex_Count).Value & "'"
                        'StrSQL = "Select SI_Product.PL_Code FROM SI_PRODUCT With(NOLOCK), Product_LG With(NOLOCK) Where Product_LG.PL_Code = SI_Product.PL_Code And Product_LG.PL_LG_Code  = '" & Grid_LGPlan.Item(4, Ex_Count).Value & "'"
                        ''MsgBox(Grid_LGPlan.Item(4, Ex_Count).Value)

                        'Re_Count = DB_Select(DBT)
                        'If Re_Count <> 0 Then
                        '    'MsgBox("일치")
                        '    ' Grid_LGPlan.Item(4, Ex_Count).Value = DBT.Rows(0).Item(0)
                        '    ' Grid_LGPlan.Item(10, Ex_Count).Value = DBT.Rows(0).Item(1)
                        'Else
                        '    '제품코드가 안들어있을때
                        '    'Grid_LGPlan.Item(4, Ex_Count).Value = ""
                        '    Grid_LGPlan.Rows.RemoveAt(Ex_Count)
                        '    Ex_Count = Ex_Count - 1
                        '    ' Grid_LGPlan.Item(10, Ex_Count).Value = ""
                        'End If


                    End If
                Next i

                ' MsgBox(Grid_LGPlan.Rows.Count)
                Dim count As Integer
                count = Grid_LGPlan.Rows.Count - 1
                For k = 0 To Grid_LGPlan.Rows.Count - 1

                    ''상위 코드 검색
                    '   StrSQL = "Select Product_List.PL_Code, Product_List.PL_Name FROM SI_PRODUCT_List With(NOLOCK), Product_LG With(NOLOCK) Where Product_LG.PL_Code = Product_List.PL_Code And Product_LG.PL_LG_Code  = '" & Grid_LGPlan.Item(2, Ex_Count).Value & "'"
                    'MsgBox(Grid_LGPlan.Item(4, i).Value)
                    If Grid_LGPlan.Item(4, count).Value = "" Then
                        k = k + 1
                    End If
                    StrSQL = "Select PL_LG_Code FROM SI_PRODUCT_LG With(NOLOCK) Where PL_LG_Code = '" & Grid_LGPlan.Item(4, count).Value & "'"
                    'MsgBox(Grid_LGPlan.Item(4, Ex_Count).Value)

                    Re_Count = DB_Select(DBT)

                    If Re_Count <> 0 Then

                    Else
                        ' 없음
                        '  Grid_LGPlan.Rows.Remove(Grid_LGPlan.Rows(count))
                        Grid_LGPlan.Rows.RemoveAt(count)
                    End If

                    count = count - 1

                Next k




                ' Grid_LGPlan.RowCount = Ex_Count + 1
                xlsheet = Nothing
                    xlbook.Close()
                xlbook = Nothing
                xlapp.Quit()
                xlapp = Nothing
                Dim xlp() As Process = Process.GetProcessesByName("EXCEL")
                For Each Process As Process In xlp
                    Process.Kill()
                Next
                releaseObject(xlsheet)
                releaseObject(xlbook)
                releaseObject(xlapp)
                MsgBox("로드 완료되었습니다.")



            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        Else
            MsgBox("취소하셨습니다.")
            Exit Sub
        End If
        '등록된 코드 만 처리 한다.
        Excel_Com.Visible = True

        Exit Sub

    End Sub


    Public Function Grid_LGPlan_Display(Code As String) As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer

        Grid_Display_Ch = False
        Grid_LGPlan.RowCount = 1
        StrSQL = "Select *  From Plan_LG with(NOLOCK)  Where PL_Code = '" & Code & "' Order By Convert(Decimal,PL_Sun)"
        Re_Count = DB_Select(DBT)

        If Re_Count <> 0 Then
            Grid_LGPlan.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 1 To Grid_LGPlan.ColumnCount - 1
                    Grid_LGPlan.Item(j - 1, i).Value = DBT.Rows(i).Item(j)
                Next j
            Next i
        Else
            Grid_LGPlan.RowCount = 1
            For j = 0 To Grid_LGPlan.ColumnCount - 1
                Grid_LGPlan.Item(j, 0).Value = ""
            Next j
        End If
        Grid_Display_Ch = True
        Grid_LGPlan_Display = True

    End Function


    Private Sub Grid_Code_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Code.CellContentClick

    End Sub

    Private Sub Excel_Com_Click(sender As Object, e As EventArgs) Handles Excel_Com.Click
        Label_Color(sender, Color_Label, Di_Panel2)
        '판넬 보이기

        Cursor = Cursors.AppStarting


        '여기서 처리

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

        Dim i As Integer
        Dim j As Integer


        xlapp = New Excel.Application
        xlbook = xlapp.Workbooks.Open(Application.StartupPath & "\excel\null.xls")
        xlsheet = xlbook.ActiveSheet
        xlapp.DisplayAlerts = False
        xlapp.WindowState = Excel.XlWindowState.xlMinimized
        xlapp.Visible = False
        xlapp.DisplayFormulaBar = False
        xlapp.DisplayStatusBar = False




        'PInvoke 에러시 프로젝트 속성에서 컴파일을 x64로 하거나 Any CPU일 경우 32비트 기본 사용 체크를 해제한다.
        'SetWindowLong(xlapp.Hwnd, GWL_STYLE, GetWindowLong(xlapp.Hwnd, GWL_STYLE) And Not WS_CAPTION)

        xlapp.ActiveWindow.DisplayWorkbookTabs = False
        xlapp.ActiveWindow.DisplayHeadings = False

        'xlapp.ActiveWindow.DisplayWorkbookTabs = True
        'xlapp.ActiveWindow.DisplayHeadings = True


        ' SetParent(xlapp.Hwnd, Panel_Excel.Handle)
        SendMessage(xlapp.Hwnd, WM_SYSCOMMAND, SC_MAXIMIZE, 0)



        '제목

        For j = 0 To Grid_LGPlan.ColumnCount - 1
            xlsheet.Cells(1, j + 1).Value = Grid_LGPlan.Columns(j).HeaderText
            xlsheet.Columns(j + 1).ColumnWidth = Grid_LGPlan.Columns(j).Width / 5

            'xlsheet.Cells(1, 1).Value = Grid_LGPlan.Columns(1).HeaderText
            'xlsheet.Cells(1, 2).Value = Grid_LGPlan.Columns(3).HeaderText
            'xlsheet.Cells(1, 3).Value = Grid_LGPlan.Columns(4).HeaderText
            'xlsheet.Cells(1, 4).Value = "P/N"
            'xlsheet.Cells(1, 5).Value = "MOTOR"
            'xlsheet.Cells(1, 6).Value = "HEARTER"
        Next j

        'For j = 6 To 29 'Grid_LGPlan.ColumnCount - 1
        '    'xlsheet.Cells(1, j + 1).Value = Grid_LGPlan.Columns(j).HeaderText
        '    'xlsheet.Columns(j + 1).ColumnWidth = Grid_LGPlan.Columns(j).Width / 5
        '    xlsheet.Cells(1, j + 1).value = Grid_LGPlan.Columns(j).HeaderText
        '    xlsheet.Columns(j + 1).ColumnWidth = Grid_LGPlan.Columns(j).Width / 7
        'Next j

        'xlsheet.Cells(1, 31).Value = Grid_LGPlan.Columns(39).HeaderText
        'xlsheet.Cells(1, 32).Value = Grid_LGPlan.Columns(42).HeaderText

        'For i = 0 To Grid_LGPlan.RowCount - 1
        '    For j = 0 To Grid_LGPlan.ColumnCount - 1
        '        xlsheet.Cells(i + 2, j + 1).Value = Grid_LGPlan.Item(j, i).Value.ToString
        '    Next j
        'Next i


        'For j = 0 To Grid_LGPlan.ColumnCount - 1

        '    'MsgBox("j : " & j & "----" & Grid_LGPlan.Item(j, 0).Value.ToString)
        '    If Grid_LGPlan.Item(j, 0).Value = Nothing Then
        '        Grid_LGPlan.Item(j, 0).Value = ""
        '    End If
        '    xlsheet.Cells(2, j + 1).Value = Grid_LGPlan.Item(j, 0).Value.ToString
        'Next j



        For i = 0 To Grid_LGPlan.RowCount - 1
            For j = 0 To Grid_LGPlan.ColumnCount - 1
                If Grid_LGPlan.Item(j, i).Value = Nothing Then
                    Grid_LGPlan.Item(j, i).Value = ""
                End If
                xlsheet.Cells(i + 2, j + 1).Value = "'" & Grid_LGPlan.Item(j, i).Value.ToString
            Next j
        Next i

        ''합계

        'xlsheet.Cells(27, 6).Value = j

        ''거래처명
        'xlsheet.Cells(29, 4).Value = Grid_Info.Item(1, 5).Value
        ''거래처 주소
        'StrSQL = "Select CM_Add1, CM_Add2  FROM SI_CUSTOMER with(NOLOCK) Where CM_Code= '" & Grid_Info.Item(1, 4).Value & "'"
        'Re_Count = DB_Select(DBT)
        'If Re_Count <> 0 Then
        '    xlsheet.Cells(41, 5).Value = DBT.Rows(0).Item(0) & " " & DBT.Rows(0).Item(1)
        'End If

        xlbook.SaveAs(Application.StartupPath & "\Excel\LG제품\" & Format(Now, "MMdd") & "일.xls")
        xlapp.DisplayAlerts = False
        Cursor = Cursors.Default
        Excel_Check = True

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

    Private Sub Delete_Com_Click(sender As Object, e As EventArgs) Handles Delete_Com.Click
        Dim DBT As New DataTable
        Dim JP_Code As Long
        Dim My_Date As String
        Dim My_Time As String

        If Len(Grid_Info.Item(1, 0).Value) < 1 Then
            Exit Sub
        End If


        If MsgBox("코드 '" & Grid_Info.Item(1, 0).Value & " 를 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "코드 삭제") = vbNo Then
            Exit Sub
        End If
        Grid_Display_Ch = False

        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_PRODUCT Where PL_Code = '" & Grid_Info.Item(1, 0).Value & "' "
        Re_Count = DB_Execute()


    End Sub
End Class
