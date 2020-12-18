﻿
Imports Excel = Microsoft.Office.Interop.Excel


Public Class Input
    Dim Grid_Info_Row As Integer
    Dim Grid_Info_Col As Integer
    Dim Grid_Main_Row As Integer
    Dim Grid_Main_Col As Integer
    Dim xlapp As Excel.Application
    Dim xlbook As Excel.Workbook
    Dim xlsheet As Excel.Worksheet
    Dim Excel_Check As Boolean
    Dim Grid_Display_Ch As Boolean
    Dim Info_Lab() As Button
    Dim Info_Txt() As TextBox
    Dim Info_Com() As ComboBox






    Private Sub Frm_Contract_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Info_Lab = New Button() {Info_La0, Info_La1, Info_La2, Info_La3, Info_La4, Info_La5, Info_La6, Info_La7, Info_La8, Info_La9, Info_La10, Info_La11, Info_La12, Info_La13}
        Info_Txt = New TextBox() {Info_Tx0, Info_Tx1, Info_Tx2, Info_Tx3, Info_Tx4, Info_Tx5, Info_Tx6, Info_Tx7, Info_Tx8, Info_Tx9, Info_Tx10, Info_Tx11, Info_Tx12, Info_Tx13}
        Info_Com = New ComboBox() {Info_Co0, Info_Co1, Info_Co2, Info_Co3, Info_Co4, Info_Co5, Info_Co6, Info_Co7, Info_Co8, Info_Co9, Info_Co10, Info_Co11, Info_Co12, Info_Co13}


        Me.Dock = DockStyle.Fill
        Me.BringToFront()
        Grid_Code_Setup()
        Grid_Info_Setup()
        Grid_Main_Setup()




        Search_Combo.Items.Add("코드")
        Search_Combo.Items.Add("날짜")
        '  Search_Combo.Items.Add("거래처")
        Search_Combo.Text = "코드"


        Panel_Main.Top = 136
        Panel_Main.Left = 389
        Panel_Excel.Top = 136
        Panel_Excel.Left = 389

        Panel_Main.Visible = True
        Panel_Excel.Visible = False

        Com_Excel_Print.Visible = False
        Search_Com.PerformClick()

    End Sub
    Public Function Grid_Code_Setup() As Boolean
        Grid_Color(Grid_Code)

        Grid_Code.AllowUserToAddRows = False
        Grid_Code.RowTemplate.Height = 20
        Grid_Code.ColumnCount = 2
        Grid_Code.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Code.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Code.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        '    Grid_Code.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Code.RowHeadersWidth = 10
        Grid_Code.Columns(0).Width = 100
        Grid_Code.Columns(1).Width = 70
        '       Grid_Code.Columns(2).Width = 80

        Grid_Code.Columns(0).HeaderText = "코드"
        Grid_Code.Columns(1).HeaderText = "날짜"
        '  Grid_Code.Columns(2).HeaderText = "거래처"

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
        Grid_Info.RowCount = 4





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
        'Grid_Info.Item(0, 4).Value = "거래처코드"
        'Grid_Info.Item(0, 5).Value = "거래처명"
        'Grid_Info.Item(0, 6).Value = "담당자"
        'Grid_Info.Item(0, 7).Value = "연락처"
        'Grid_Info.Item(0, 8).Value = "지불조건"

        'Grid_Info.Item(0, 9).Value = "납품예정일1"
        'Grid_Info.Item(0, 10).Value = "납품예정일1"
        'Grid_Info.Item(0, 11).Value = "납품예정일3"
        'Grid_Info.Item(0, 12).Value = "비고"
        'Grid_Info.Item(0, 13).Value = "관리"

        Dim i As Integer
        For i = 0 To Grid_Info.RowCount - 1
            Info_Lab(i).Text = Grid_Info.Item(0, i).Value
        Next i
        'Grid_Info.Columns(0).ReadOnly = True
        'Grid_Info.Columns(1).ReadOnly = True

        Grid_Info.Rows(0).ReadOnly = True
        Grid_Info.Rows(1).ReadOnly = True
        Grid_Info.Rows(2).ReadOnly = True
        Grid_Info.Rows(3).ReadOnly = True
        '  Grid_Info.Rows(4).ReadOnly = False

        'Info_Co Setup

        For i = 0 To Grid_Info.RowCount - 1
            Info_Com(i).Visible = False
        Next i

        Info_Com(4).Visible = True
        Info_Com(5).Visible = True
        Info_Com(8).Visible = True
        Info_Com(13).Visible = True

        Info_Txt(4).Visible = False
        Info_Txt(5).Visible = False
        Info_Txt(8).Visible = False

        Grid_Info_Setup = True
    End Function
    Public Function Grid_Main_Setup() As Boolean
        Grid_Color(Grid_Main)

        Grid_Main.AllowUserToAddRows = False
        Grid_Main.RowTemplate.Height = 24
        Grid_Main.ColumnCount = 8
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
        Grid_Main.Columns(6).HeaderText = "수량"
        Grid_Main.Columns(7).HeaderText = "비고"


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

    End Sub
    Public Function Grid_Code_Display() As Boolean

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Code.RowCount = 1
        Select Case Search_Combo.Text
            Case "코드"
                StrSQL = "Select MS_Code, MS_Date FROM MC_STOCK_OUT with(NOLOCK) Where MS_Code Like '%" & Search_Text.Text & "%'  Order  By MS_Code DESC"
            Case "날짜"
                StrSQL = "Select MS_Code, MS_Date FROM MC_STOCK_OUT with(NOLOCK) Where MS_Date Like '%" & Search_Text.Text & "%' Or MS_Date Is Null  Order By MS_Date"
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
            My_Time = DateTime.Now.ToString("HH:mm:ss")
            My_Date = Mid(My_Date, 1, 4) & "-" & Mid(My_Date, 6, 2) & "-" & Mid(My_Date, 9, 2)
        End If


        StrSQL = "Select MS_Code FROM MC_STOCK_OUT with(NOLOCK) Where MS_Date Like  '" & My_Date & "%' Order By MS_Code Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            JP_Code = JP_Code & "001"
        Else
            JP_Code = JP_Code & Format(Val(Mid(DBT.Rows(0).Item(0), 11, 3)) + 1, "00#")
        End If

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO MC_STOCK_OUT (MS_Code, MS_NAME, MS_Date, MS_Time)  Values ('CC" & JP_Code & "','" & Frm_Main.Text_Man_Name.Text & "', '" & My_Date & "','" & My_Time & "')"
        Re_Count = DB_Execute()
        Grid_Code_Display()
        '위치로 이동한다.


        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
            Panel_Main.Visible = True
            Grid_Info_Display(Grid_Info, "Select * FROM MC_STOCK_OUT With(NOLOCK) Where MS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "'")
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

        StrSQL = "Select DL_Code FROM SC_DELIVER with(NOLOCK) Where DL_MS_Code Like '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Select(DBT)
        If Re_Count > 0 Then
            MsgBox("납품 전표에 처리된 내역이 있어 삭제 할 수 없습니다.")
            Exit Sub
        End If




        Grid_Display_Ch = False
        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE MC_STOCK_OUT Where MS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' "
        Re_Count = DB_Execute()

        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE MC_STOCK_OUT_LIST Where MS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' "
        Re_Count = DB_Execute()

        Grid_Code_Display()
        Grid_Info_Display(Grid_Info, "Select * FROM MC_STOCK_OUT With(NOLOCK) Where MS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "'")
        'Grid_Info_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True
        Grid_Main.RowCount = 0

        MsgBox("삭제되었습니다")
    End Sub


    Private Sub Grid_Code_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Code.CellEnter
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
            Grid_Display_Ch = False
            Panel_Main.Visible = True
            Grid_Info_Display2(Grid_Info, "Select * FROM MC_STOCK_OUT With(NOLOCK) Where MS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "'", Info_Txt, Info_Com)
            '  Main_Grid_day()
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
    '    StrSQL = "Select * FROM MC_STOCK_OUT With(NOLOCK) Where MS_Code = '" & Code & "'"
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
                'Case 4
                '    StrSQL = "Select CM_Code  FROM SI_CUSTOMER With(NOLOCK) Order By CM_Name"
                '    Re_Count = DB_Select(DBT)
                '    Grid_Info_Combo.Items.Clear()
                '    If Re_Count <> 0 Then
                '        For i = 0 To Re_Count - 1
                '            Grid_Info_Combo.Items.Add(DBT.Rows(i)("CM_Code"))
                '        Next i
                '    End If
                '    Grid_Info_Combo.Top = Grid_Info.Top + Grid_Info.ColumnHeadersHeight + (Grid_Info_Row) * 24
                '    WW = 0
                '    For i = 0 To Grid_Info_Col - 1
                '        WW = WW + Grid_Info.Columns(i).Width
                '    Next
                '    Grid_Info_Combo.Left = Grid_Info.Left + Grid_Info.RowHeadersWidth + WW
                '    Grid_Info_Combo.Width = Grid_Info.Columns(Grid_Info_Col).Width
                '    Grid_Info_Combo.Text = Grid_Info.CurrentCell.Value.ToString
                '    Grid_Info_Combo.BackColor = Grid_Info.RowsDefaultCellStyle.BackColor
                '    Grid_Info_Combo.Visible = True
                '    Grid_Info_Combo.Focus.ToString()
                'Case 5
                '    StrSQL = "Select CM_Name  FROM SI_CUSTOMER With(NOLOCK) Order By CM_Name"
                '    Re_Count = DB_Select(DBT)
                '    Grid_Info_Combo.Items.Clear()
                '    If Re_Count <> 0 Then
                '        For i = 0 To Re_Count - 1
                '            Grid_Info_Combo.Items.Add(DBT.Rows(i)("CM_Name"))
                '        Next i
                '    End If
                '    Grid_Info_Combo.Top = Grid_Info.Top + Grid_Info.ColumnHeadersHeight + (Grid_Info_Row) * 24
                '    WW = 0
                '    For i = 0 To Grid_Info_Col - 1
                '        WW = WW + Grid_Info.Columns(i).Width
                '    Next
                '    Grid_Info_Combo.Left = Grid_Info.Left + Grid_Info.RowHeadersWidth + WW
                '    Grid_Info_Combo.Width = Grid_Info.Columns(Grid_Info_Col).Width
                '    Grid_Info_Combo.Text = Grid_Info.CurrentCell.Value.ToString
                '    Grid_Info_Combo.BackColor = Grid_Info.RowsDefaultCellStyle.BackColor
                '    Grid_Info_Combo.Visible = True
                '    Grid_Info_Combo.Focus.ToString()

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

        '순번추가
        Grid_Display_Ch = False
        'Dim DBT As New DataTable
        Dim Db_Sun As Long
        StrSQL = "Select MS_Sun FROM MC_STOCK_OUT_LIST with(NOLOCK) Where MS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' Order By Convert(Decimal,MS_Sun) Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            Db_Sun = 1
        Else
            Exit Sub
        End If

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO MC_STOCK_OUT_LIST (MS_Code, MS_Sun)  Values ('" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "', '" & Db_Sun & "')"
        Re_Count = DB_Execute()
        Grid_Main_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True

    End Sub

    Private Sub Save_Com_Click(sender As Object, e As EventArgs) Handles Save_Com.Click

        If Grid_Code.Item(0, 0).Value = "" Then
            MsgBox("공백은 저장할 수 없습니다")
            Exit Sub
        End If

        Grid_Display_Ch = False
        Grid_Info_Save(Grid_Info, "MC_STOCK_OUT")
        Grid_Display_Ch = True
        Search_Com.PerformClick()
        MsgBox("저장되었습니다")
    End Sub

    Private Sub Grid_Info_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Info.CellValueChanged
        If Grid_Display_Ch = True Then
            Grid_Info.CurrentRow.HeaderCell.Value = "*"
            '   Main_Grid_day()
        End If

    End Sub


    'Public Function Main_Grid_day() As Boolean
    '    Grid_Main.Columns(6).HeaderText = Grid_Info.Item(1, 9).Value.ToString
    '    Grid_Main.Columns(7).HeaderText = Grid_Info.Item(1, 10).Value.ToString
    '    Grid_Main.Columns(8).HeaderText = Grid_Info.Item(1, 11).Value.ToString
    '    Main_Grid_day = True

    'End Function

    Private Sub Insert__Main_Click(sender As Object, e As EventArgs) Handles Insert__Main.Click
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        Grid_Display_Ch = False
        Dim DBT As New DataTable
        Dim Db_Sun As Long
        StrSQL = "Select MS_Sun FROM MC_STOCK_OUT_LIST with(NOLOCK) Where MS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' Order By Convert(Decimal,MS_Sun) Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            Db_Sun = 1
        Else
            Db_Sun = DBT.Rows(0).Item(0) + 1
        End If

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO MC_STOCK_OUT_LIST (MS_Code, MS_Sun)  Values ('" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "', '" & Db_Sun & "')"
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
        StrSQL = "Select MS_Sun, MS_PL_Code, MS_PL_Name, MS_Standard, MS_Unit, MS_Unit_Amount, MS_Num, MS_Bigo  FROM MC_STOCK_OUT_LIST with(NOLOCK)  Where MS_Code = '" & PL_Code & "' Order By Convert(Decimal,MS_Sun)"
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
        '   Main_Grid_day()

        Grid_Display_Ch = True
        Grid_Main_Display = True

    End Function

    Private Sub Grid_Main_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Main.CellValueChanged

        If IsNothing(Grid_Main.CurrentRow) Then
            Exit Sub
        End If
        If Grid_Display_Ch = True Then
            Grid_Main.CurrentRow.HeaderCell.Value = "*"
            '    Grid_Main.Item(9, Grid_Main.CurrentCell.RowIndex).Value = Val(Grid_Main.Item(6, Grid_Main.CurrentCell.RowIndex).Value) + Val(Grid_Main.Item(7, Grid_Main.CurrentCell.RowIndex).Value) + Val(Grid_Main.Item(8, Grid_Main.CurrentCell.RowIndex).Value)
            '    Grid_Main.Item(10, Grid_Main.CurrentCell.RowIndex).Value = Val(Grid_Main.Item(9, Grid_Main.CurrentCell.RowIndex).Value) * Val(Grid_Main.Item(5, Grid_Main.CurrentCell.RowIndex).Value.ToString)
        End If

    End Sub
    Private Sub Grid_Main_MouseClick(sender As Object, e As MouseEventArgs) Handles Grid_Main.MouseClick

        If String.IsNullOrEmpty(Grid_Main.CurrentCell.RowIndex) Then
            Exit Sub
        End If

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
            Case 2
                
                '팝업창 띄우기
                Dim popup_product As New popup_product
                popup_product.ShowDialog()

                '결과값이 ok일때 그리드뷰에 값넣어주기
                If popup_product.DialogResult = DialogResult.OK Then
                    Grid_Main.Rows(Grid_Main.CurrentRow.Index).Cells(2).Value = popup_product.Product_Name
                    Grid_Main.Rows(Grid_Main.CurrentRow.Index).Cells(1).Value = popup_product.Product_Code
                    Grid_Main.Rows(Grid_Main.CurrentRow.Index).Cells(3).Value = popup_product.Product_Standard
                    Grid_Main.Rows(Grid_Main.CurrentRow.Index).Cells(4).Value = popup_product.Product_Unit
                    Grid_Main.Rows(Grid_Main.CurrentRow.Index).Cells(5).Value = popup_product.Product_Gold

                End If

                '다 넣고 커서 포커스 변경
                Grid_Main.CurrentCell = Grid_Main.Rows(Grid_Main.CurrentCell.RowIndex).Cells(6)

            Case 1
                
                '팝업창 띄우기
                Dim popup_product As New popup_product
                popup_product.ShowDialog()

                '결과값이 ok일때 그리드뷰에 값넣어주기
                If popup_product.DialogResult = DialogResult.OK Then
                    Grid_Main.Rows(Grid_Main.CurrentRow.Index).Cells(2).Value = popup_product.Product_Name
                    Grid_Main.Rows(Grid_Main.CurrentRow.Index).Cells(1).Value = popup_product.Product_Code
                    Grid_Main.Rows(Grid_Main.CurrentRow.Index).Cells(3).Value = popup_product.Product_Standard
                    Grid_Main.Rows(Grid_Main.CurrentRow.Index).Cells(4).Value = popup_product.Product_Unit
                    Grid_Main.Rows(Grid_Main.CurrentRow.Index).Cells(5).Value = popup_product.Product_Gold

                End If

                '다 넣고 커서 포커스 변경
                Grid_Main.CurrentCell = Grid_Main.Rows(Grid_Main.CurrentCell.RowIndex).Cells(6)

        End Select

    End Sub
    Private Sub Grid_Main_Scroll(sender As Object, e As ScrollEventArgs) Handles Grid_Main.Scroll
        Grid_Main_Combo.Visible = False
    End Sub

    Private Sub Grid_Main_Combo_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub


    Private Sub Grid_Main_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Main.CellContentClick

    End Sub

    Private Sub Save_Main_Click(sender As Object, e As EventArgs) Handles Save_Main.Click
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        Dim i As Integer


        'LIST 저장

        Grid_Display_Ch = False
        For i = 0 To Grid_Main.Rows.Count - 1
            If Grid_Main.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE MC_STOCK_OUT_LIST Set MS_PL_Code = '" & Grid_Main.Item(1, i).Value & "', 
                    MS_PL_Name = '" & Grid_Main.Item(2, i).Value & "', MS_Standard = '" & Grid_Main.Item(3, i).Value & "', 
                    MS_Unit = '" & Grid_Main.Item(4, i).Value & "', MS_Unit_Amount = '" & Grid_Main.Item(5, i).Value & "', 
                    MS_Num = '" & Grid_Main.Item(6, i).Value & "', MS_Bigo = '" & Grid_Main.Item(7, i).Value & "' 
                    where MS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' 
                    And MS_Sun = '" & Grid_Main.Item(0, i).Value & "'"
                Re_Count = DB_Execute()
                Grid_Main.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display_Ch = True


        'MASTER 저장

        Grid_Display_Ch = False
        Grid_Info_Save(Grid_Info, "MC_STOCK_OUT")
        Grid_Display_Ch = True
        Search_Com.PerformClick()


        MsgBox("저장되었습니다")
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

        If MsgBox("순번 '" & Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value & "'을(를) 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "삭제") = vbNo Then
            Exit Sub
        End If


        Grid_Display_Ch = False
        Dim i As Integer
        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE MC_STOCK_OUT_LIST Where MS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And  MS_Sun = '" & Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        For i = Grid_Main.CurrentCell.RowIndex + 1 To Grid_Main.RowCount - 1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "Update MC_STOCK_OUT_LIST Set MS_Sun = '" & i & "' Where MS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And  MS_Sun = '" & i + 1 & "'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        Grid_Main_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True

        MsgBox("삭제되었습니다")
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
        '엑셀파일 수정(20200324)
        xlbook = xlapp.Workbooks.Open(Application.StartupPath & "\Excel\수주.xlsx")
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
        '공급업체명
        StrSQL = "Select CM_Name From Company with(NOLOCK) Where CM_Code= '10000'"
        Re_Count = DB_Select(DBT)
        xlsheet.Cells(4, 4).Value = DBT.Rows(0).Item(0)

        '지불조건
        xlsheet.Cells(5, 4).Value = Grid_Info.Item(1, 8).Value
        '담당자 
        xlsheet.Cells(6, 4).Value = Grid_Info.Item(1, 6).Value
        '발주일자
        xlsheet.Cells(4, 14).Value = Grid_Info.Item(1, 2).Value
        '발주코드
        xlsheet.Cells(5, 14).Value = Grid_Info.Item(1, 0).Value
        '담당자 전화 번호
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
        xlsheet.Cells(33, 6).Value = j
        '거래처명
        xlsheet.Cells(35, 4).Value = Grid_Info.Item(1, 5).Value
        '거래처 주소
        StrSQL = "Select CM_Add1, CM_Add2  FROM SI_CUSTOMER with(NOLOCK) Where CM_Code= '" & Grid_Info.Item(1, 4).Value & "'"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            xlsheet.Cells(41, 5).Value = DBT.Rows(0).Item(0) & " " & DBT.Rows(0).Item(1)
        End If
        'xlbook.SaveAs(Application.StartupPath & "\Excel\수주\" & Grid_Info.Item(1, 0).Value & Grid_Info.Item(1, 5).Value & ".xls")
        Cursor = Cursors.Default
        Excel_Check = True
    End Sub

    Private Sub Com_Contract_Click(sender As Object, e As EventArgs) Handles Com_Contract.Click

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

    Private Sub Info_Tx0_TextChanged(sender As Object, e As EventArgs) Handles Info_Tx0.TextChanged, Info_Tx1.TextChanged, Info_Tx2.TextChanged, Info_Tx3.TextChanged, Info_Tx4.TextChanged, Info_Tx5.TextChanged, Info_Tx6.TextChanged, Info_Tx7.TextChanged, Info_Tx8.TextChanged, Info_Tx9.TextChanged, Info_Tx10.TextChanged, Info_Tx11.TextChanged, Info_Tx12.TextChanged, Info_Tx13.TextChanged
        Dim index As Integer

        If Grid_Display_Ch = False Then
            Exit Sub

        Else
            'MsgBox(Mid(sender.name.ToString, 8, 2))
        End If

        index = Val(Mid(sender.name.ToString, 8, 2))
        Grid_Info.Item(1, index).Value = sender.text
        Grid_Info.Rows(index).HeaderCell.Value = "*"
        Info_Com(index).Text = sender.text



    End Sub

    Private Sub Info_Co0_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Info_Co0.SelectedIndexChanged

    End Sub

    Private Sub Info_Co0_MouseClick(sender As Object, e As MouseEventArgs) Handles Info_Co4.MouseClick, Info_Co5.MouseClick, Info_Co8.MouseClick, Info_Co13.MouseClick

        Dim index As Integer

        If Grid_Display_Ch = False Then
            Exit Sub
        Else
            'MsgBox(Mid(sender.name.ToString, 8, 2))
        End If

        index = Val(Mid(sender.name.ToString, 8, 2))
        Select Case index
            Case 4
                StrSQL = "Select CM_Code  FROM SI_CUSTOMER With(NOLOCK) Order By CM_Name"
            Case 5
                StrSQL = "Select CM_Name  FROM SI_CUSTOMER With(NOLOCK) Order By CM_Name"
            Case 8
                StrSQL = "Select Base_Name  FROM SI_BASE_CODE With(NOLOCK) Where Base_Sel = '지불조건' Order By Base_Name"
            Case 13
                sender.Items.Clear()
                sender.Items.Add("완료")
                sender.Items.Add("진행중")
                sender.text = Info_Tx13.Text
                Exit Sub
            Case Else
                Exit Sub

        End Select
        Dim DBT As New DataTable

        Re_Count = DB_Select(DBT)
        sender.Items.Clear()
        If Re_Count <> 0 Then
            For i = 0 To Re_Count - 1
                sender.Items.Add(DBT.Rows(i).Item(0).ToString)
            Next i
        End If
        sender.text = Info_Txt(index).Text

    End Sub

    Private Sub Info_Co0_SelectedValueChanged(sender As Object, e As EventArgs) Handles Info_Co4.SelectedValueChanged, Info_Co5.SelectedValueChanged, Info_Co8.SelectedValueChanged, Info_Co13.SelectedValueChanged
        Dim index As Integer

        If Grid_Display_Ch = False Then
            Exit Sub
        End If
        index = Val(Mid(sender.name.ToString, 8, 2))

        Info_Txt(index).Text = Info_Com(index).Text
        Grid_Info.Rows(index).HeaderCell.Value = "*"

        Dim DBT As New DataTable
        Select Case index
            Case 4
                StrSQL = "Select CM_Code, CM_Name, CM_Man_Name, CM_Man_Tel  FROM SI_CUSTOMER With(NOLOCK) Where CM_Code = '" & sender.text.ToString() & "'"
                Re_Count = DB_Select(DBT)
                Grid_Info_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    Info_Txt(5).Text = DBT.Rows(0)("CM_Name")
                    Info_Txt(6).Text = DBT.Rows(0)("CM_Man_Name")
                    Info_Txt(7).Text = DBT.Rows(0)("CM_Man_Tel")
                    Grid_Info.Rows(4).HeaderCell.Value = "*"
                    Grid_Info.Rows(5).HeaderCell.Value = "*"
                    Grid_Info.Rows(6).HeaderCell.Value = "*"
                    Grid_Info.Rows(7).HeaderCell.Value = "*"
                End If
            Case 5
                StrSQL = "Select CM_Code, CM_Name, CM_Man_Name, CM_Man_Tel  FROM SI_CUSTOMER With(NOLOCK) Where CM_Name = '" & sender.text.ToString() & "'"
                Re_Count = DB_Select(DBT)
                Grid_Info_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    Info_Txt(4).Text = DBT.Rows(0)("CM_Code")
                    Info_Txt(6).Text = DBT.Rows(0)("CM_Man_Name")
                    Info_Txt(7).Text = DBT.Rows(0)("CM_Man_Tel")
                    Grid_Info.Rows(4).HeaderCell.Value = "*"
                    Grid_Info.Rows(5).HeaderCell.Value = "*"
                    Grid_Info.Rows(6).HeaderCell.Value = "*"
                    Grid_Info.Rows(7).HeaderCell.Value = "*"
                End If

        End Select

    End Sub

    Private Sub Info_Co4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Info_Co4.SelectedIndexChanged

    End Sub

    Private Sub Info_La0_Click(sender As Object, e As EventArgs) Handles Info_La0.Click

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

    Private Sub Grid_Info_MouseClick(sender As Object, e As MouseEventArgs) Handles Grid_Info.MouseClick
        If String.IsNullOrEmpty(Grid_Info.CurrentCell.RowIndex) Then
            Exit Sub
        End If

        Grid_Info_Row = Grid_Info.CurrentCell.RowIndex
        Grid_Info_Col = Grid_Info.CurrentCell.ColumnIndex
        Grid_Info_Combo.Visible = False
        If Grid_Info_Row < 0 Then
            Exit Sub
        End If
        If Grid_Info_Col < 0 Then
            Exit Sub
        End If
        Dim DBT As New DataTable

        Select Case Grid_Info_Row
            Case 4
                
                '팝업창 띄우기
                Dim popup_account As New popup_account
                popup_account.ShowDialog()

                '결과값이 ok일때 그리드뷰에 값넣어주기
                If popup_account.DialogResult = DialogResult.OK Then

                    Grid_Info.Item(1, 4).Value = popup_account.Customer_Code
                    Grid_Info.Item(1, 5).Value = popup_account.Customer_Name
                    Grid_Info.Item(1, 6).Value = popup_account.Customer_Man_Name
                    Grid_Info.Item(1, 7).Value = popup_account.Customer_Man_Tel
                    Grid_Info.Rows(4).HeaderCell.Value = "*"
                    Grid_Info.Rows(5).HeaderCell.Value = "*"
                    Grid_Info.Rows(6).HeaderCell.Value = "*"
                    Grid_Info.Rows(7).HeaderCell.Value = "*"
                    Grid_Display_Ch = False
                    'Dim DBT As New DataTable
                    Dim Db_Sun As Long
                    StrSQL = "Select MS_Sun FROM MC_STOCK_OUT_LIST with(NOLOCK) Where MS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' Order By Convert(Decimal,MS_Sun) Desc "
                    Re_Count = DB_Select(DBT)
                    If Re_Count = 0 Then
                        Db_Sun = 1
                    Else
                        Exit Sub
                    End If

                    '추가한다
                    StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                    StrSQL = StrSQL & "Insert INTO MC_STOCK_OUT_LIST (MS_Code, MS_Sun)  Values ('" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "', '" & Db_Sun & "')"
                    Re_Count = DB_Execute()
                    Grid_Main_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
                    Grid_Display_Ch = True

                End If

                '다 넣고 커서 포커스 변경
                Grid_Info.CurrentCell = Grid_Info.Rows(Grid_Info.CurrentCell.RowIndex).Cells(1)

            Case 5
                
                '팝업창 띄우기
                Dim popup_account As New popup_account
                popup_account.ShowDialog()

                '결과값이 ok일때 그리드뷰에 값넣어주기
                If popup_account.DialogResult = DialogResult.OK Then
                    Grid_Info.Item(1, 4).Value = popup_account.Customer_Code
                    Grid_Info.Item(1, 5).Value = popup_account.Customer_Name
                    Grid_Info.Item(1, 6).Value = popup_account.Customer_Man_Name
                    Grid_Info.Item(1, 7).Value = popup_account.Customer_Man_Tel
                    Grid_Info.Rows(4).HeaderCell.Value = "*"
                    Grid_Info.Rows(5).HeaderCell.Value = "*"
                    Grid_Info.Rows(6).HeaderCell.Value = "*"
                    Grid_Info.Rows(7).HeaderCell.Value = "*"
                    Grid_Display_Ch = False
                    'Dim DBT As New DataTable
                    Dim Db_Sun As Long
                    StrSQL = "Select MS_Sun FROM MC_STOCK_OUT_LIST with(NOLOCK) Where MS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' Order By Convert(Decimal,MS_Sun) Desc "
                    Re_Count = DB_Select(DBT)
                    If Re_Count = 0 Then
                        Db_Sun = 1
                    Else
                        Exit Sub
                    End If

                    '추가한다
                    StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                    StrSQL = StrSQL & "Insert INTO MC_STOCK_OUT_LIST (MS_Code, MS_Sun)  Values ('" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "', '" & Db_Sun & "')"
                    Re_Count = DB_Execute()
                    Grid_Main_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
                    Grid_Display_Ch = True
                End If

                '다 넣고 커서 포커스 변경
                Grid_Info.CurrentCell = Grid_Info.Rows(Grid_Info.CurrentCell.RowIndex).Cells(1)


        End Select
    End Sub
End Class
