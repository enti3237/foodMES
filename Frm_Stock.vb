﻿Imports Excel = Microsoft.Office.Interop.Excel
Public Class Frm_Stock
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

    Private Sub Frm_Stock_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Info_Lab = New Button() {Info_La0, Info_La1, Info_La2, Info_La3, Info_La4, Info_La5, Info_La6, Info_La7, Info_La8, Info_La9, Info_La10, Info_La11}
        Info_Txt = New TextBox() {Info_Tx0, Info_Tx1, Info_Tx2, Info_Tx3, Info_Tx4, Info_Tx5, Info_Tx6, Info_Tx7, Info_Tx8, Info_Tx9, Info_Tx10, Info_Tx11}
        Info_Com = New ComboBox() {Info_Co0, Info_Co1, Info_Co2, Info_Co3, Info_Co4, Info_Co5, Info_Co6, Info_Co7, Info_Co8, Info_Co9, Info_Co10, Info_Co11}

        Me.Dock = DockStyle.Fill
        Me.BringToFront()
        Grid_Code_Setup()
        Grid_Info_Setup()
        Grid_Main_Setup()
        DataGridView1_Setup()


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

        Com_Excel.Visible = True


        Search_Com.PerformClick()
        Grid_Code.ClearSelection()


        Panel1.Visible = False
        Info_Co11.Enabled = False

    End Sub

    Public Function DataGridView1_Setup() As Boolean
        Grid_Color(DataGridView1)

        DataGridView1.AllowUserToAddRows = False
        DataGridView1.RowTemplate.Height = 24
        DataGridView1.ColumnCount = 5
        DataGridView1.RowCount = 0

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        DataGridView1.RowHeadersWidth = 10
        DataGridView1.Columns(0).Width = 60



        DataGridView1.Columns(0).HeaderText = "코드"
        DataGridView1.Columns(1).HeaderText = "일자"
        DataGridView1.Columns(2).HeaderText = "거래처"
        DataGridView1.Columns(3).HeaderText = "제품명"
        DataGridView1.Columns(4).HeaderText = "수량"




        DataGridView1.Columns(0).Width = 60
        DataGridView1.Columns(1).Width = 40
        DataGridView1.Columns(2).Width = 60
        DataGridView1.Columns(3).Width = 60
        DataGridView1.Columns(4).Width = 40

        DataGridView1_Setup = True
    End Function
    Public Function DataGridView1_Display() As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        ' Grid_Code.RowCount = 1


        StrSQL = "select MC_STOCK_ORDER.CO_Code, MC_STOCK_ORDER.CO_Date, MC_STOCK_ORDER.CO_Customer_Name,MC_STOCK_ORDER_LIST.CO_PL_Name,
MC_STOCK_ORDER_LIST.CO_Total from MC_STOCK_ORDER JOIN MC_STOCK_ORDER_LIST ON MC_STOCK_ORDER.CO_code=MC_STOCK_ORDER_LIST.CO_Code  where CO_Customer_Name='" + Grid_Info.Item(1, 5).Value + "'"

        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            DataGridView1.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To DataGridView1.ColumnCount - 1
                    DataGridView1.Item(j, i).Value = DBT.Rows(i).Item(j)
                Next j
            Next i
        Else
            DataGridView1.RowCount = 1
            For j = 0 To DataGridView1.ColumnCount - 1
                DataGridView1.Item(j, 0).Value = ""
            Next j
        End If
        Panel1.Visible = True
    End Function


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
        Grid_Info.RowCount = 12


        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Info.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Info.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Info.RowHeadersWidth = 10
        Grid_Info.Columns(0).Width = 100
        Grid_Info.Columns(1).Width = 240
        Grid_Info.Columns(0).HeaderText = "구분"
        Grid_Info.Columns(1).HeaderText = "내용"

        Grid_Info.Rows(11).HeaderCell.Value = "*"

        Grid_Info.Item(0, 0).Value = "코드"
        Grid_Info.Item(0, 1).Value = "담당자"
        Grid_Info.Item(0, 2).Value = "날짜"
        Grid_Info.Item(0, 3).Value = "시간"
        Grid_Info.Item(0, 4).Value = "거래처코드"
        Grid_Info.Item(0, 5).Value = "거래처명"
        Grid_Info.Item(0, 6).Value = "담당자"
        Grid_Info.Item(0, 7).Value = "연락처"
        Grid_Info.Item(0, 8).Value = "입고일자"
        Grid_Info.Item(0, 9).Value = "수량"
        Grid_Info.Item(0, 10).Value = "비고"

        Grid_Info.Item(0, 11).Value = "발주전표"


        Dim i As Integer
        For i = 0 To Grid_Info.RowCount - 1
            Info_Lab(i).Text = Grid_Info.Item(0, i).Value
        Next i

        For i = 0 To Grid_Info.RowCount - 1
            Info_Com(i).Visible = False
        Next i
        
        Info_Com(4).Visible = True
        Info_Com(5).Visible = True
        Info_Com(11).Visible = True

        Info_Tx9.ReadOnly = True

        'Grid_Info.Columns(0).ReadOnly = True
        'Grid_Info.Columns(1).ReadOnly = True

        'Grid_Info.Rows(9).Visible = False

        Grid_Info.Rows(0).ReadOnly = True
        Grid_Info.Rows(1).ReadOnly = True
        Grid_Info.Rows(2).ReadOnly = True
        Grid_Info.Rows(3).ReadOnly = True
        Grid_Info.Rows(9).ReadOnly = True
        Grid_Info.Rows(11).ReadOnly = True
        Grid_Info_Setup = True

        Info_Txt(4).Visible = False
        Info_Txt(5).Visible = False
        Info_Txt(11).Visible = False

    End Function
    Public Function Grid_Main_Setup() As Boolean
        Grid_Color(Grid_Main)

        Grid_Main.AllowUserToAddRows = False
        Grid_Main.RowTemplate.Height = 24
        Grid_Main.ColumnCount = 9
        Grid_Main.RowCount = 1

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
        Grid_Main.Columns(7).HeaderText = "금액"
        Grid_Main.Columns(8).HeaderText = "비고"
        'Grid_Main.Columns(9).HeaderText = "보관방법"
        'Grid_Main.Columns(10).HeaderText = "이물"
        'Grid_Main.Columns(11).HeaderText = "이취"
        'Grid_Main.Columns(12).HeaderText = "성상"
        'Grid_Main.Columns(13).HeaderText = "포장상태"
        'Grid_Main.Columns(14).HeaderText = "검사성적서"







        Dim i As Integer

        Grid_Main.Columns(0).Width = 30
        Grid_Main.Columns(1).Width = 100
        Grid_Main.Columns(2).Width = 200

        For i = 3 To Grid_Main.ColumnCount - 1
            Grid_Main.Columns(i).Width = 80
        Next i

        Grid_Main_Setup = True
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


        StrSQL = "Select CS_Code From MC_STOCK_IN with(NOLOCK) Where CS_Date Like  '" & My_Date & "%' Order By CS_Code Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            JP_Code = JP_Code & "001"
        Else
            JP_Code = JP_Code & Format(Val(Mid(DBT.Rows(0).Item(0), 11, 3)) + 1, "00#")
        End If

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert into MC_STOCK_IN (CS_Code, CS_Date, CS_Time, CS_Name, CS_Day, CS_Check)  Values ('CS" & JP_Code & "', '" & My_Date & "','" & My_Time & "','" & Frm_Main.Text_Man_Name.Text & "','" & My_Date & "','체크')"
        Re_Count = DB_Execute()
        Grid_Code_Display()
    End Sub
    Public Function Grid_Code_Display() As Boolean

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Code.RowCount = 1

        Select Case Search_Combo.Text
            Case "코드"
                StrSQL = "Select CS_Code, CS_Date, CS_CUSTOMER_NAME From MC_STOCK_IN with(NOLOCK) Where CS_Code Like '%" & Search_Text.Text & "%' Order  By CS_Code DESC"
            Case "날짜"
                StrSQL = "Select CS_Code, CS_Date, CS_CUSTOMER_NAME From MC_STOCK_IN with(NOLOCK) Where CS_Date Like '%" & Search_Text.Text & "%'  Or CS_Date Is Null  Order By CS_Date"
            Case "거래처"
                StrSQL = "Select CS_Code, CS_Date, CS_CUSTOMER_NAME From MC_STOCK_IN with(NOLOCK) Where CS_CUSTOMER_NAME Like '%" & Search_Text.Text & "%' Or CS_CUSTOMER_NAME Is Null  Order By CS_CUSTOMER_NAME"
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
    Private Sub Grid_Code_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Code.CellEnter

        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
            Panel_Main.Visible = True
            'Grid_Info_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
            'Grid_Info_Display2(Grid_Info, "Select * From MC_STOCK_IN With(NOLOCK) Where CS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "'", Info_Txt, Info_Com)

            Dim DBT As New DataTable
            Dim i As Integer
            StrSQL = "SELECT * From MC_STOCK_IN With(NOLOCK) Where CS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "'"
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

            For i = 0 To Grid_Info.Rows.Count - 1
                Info_Txt(i).Text = Grid_Info.Item(1, i).Value
                Info_Com(i).Text = Grid_Info.Item(1, i).Value
            Next i

            '  Grid_Info_Display2 = True

            Label_Color(Com_Main, Color_Label, Di_Panel2)
            Grid_Main_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
            Panel_Main.Visible = True
            Panel_Excel.Visible = False
            Com_Excel_Print.Visible = False
        End If



    End Sub
    Public Function Grid_Info_Display(Code As String) As Boolean
        Grid_Display_Ch = False
        Grid_Info_Display = True
        Dim DBT As New DataTable
        Dim i As Integer
        StrSQL = "Select * From MC_STOCK_IN With(NOLOCK) Where CS_Code = '" & Code & "'"
        Re_Count = DB_Select(DBT)
        If Re_Count > 0 Then
            For i = 0 To DBT.Columns.Count - 4
                Grid_Info.Item(1, i).Value = DBT.Rows(0).Item(i).ToString
            Next i
        Else
            For i = 0 To DBT.Columns.Count - 4
                Grid_Info.Item(1, i).Value = ""
            Next i
        End If
        Grid_Display_Ch = True
    End Function
    Public Function Grid_Main_Display(PL_Code As String) As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display_Ch = False
        Grid_Main.RowCount = 0
        StrSQL = "Select CS_Sun, CS_PL_Code, CS_PL_Name, CS_Standard, CS_Unit, CS_Unit_Amount, CS_Total, CS_Gold, CS_Bigo
                From MC_STOCK_IN_List  Where CS_Code = '" & PL_Code & "' Order By Convert(Decimal,CS_Sun)"
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

        Grid_Info_Combo.Visible = False
        Grid_Display_Ch = True
        Grid_Main_Display = True

    End Function

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        Grid_Code_Display()
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
                    '팝업창 띄우기
                    Dim popup_account As New popup_account
                    popup_account.ShowDialog()

                    '결과값이 ok일때 그리드뷰에 값넣어주기
                    If popup_account.DialogResult = DialogResult.OK Then
                        Grid_Info.Item(Grid_Info_Col, 5).Value = popup_account.Customer_Name
                        Grid_Info.Item(Grid_Info_Col, 6).Value = popup_account.Customer_Man_Name
                        Grid_Info.Item(Grid_Info_Col, 7).Value = popup_account.Customer_Man_Tel
                        Grid_Info.Item(Grid_Info_Col, 4).Value = popup_account.Customer_Code
                    End If
                    'StrSQL = "Select CM_Code  From SI_CUSTOMER With(NOLOCK) Order By CM_Code"
                    'Re_Count = DB_Select(DBT)
                    'Grid_Info_Combo.Items.Clear()
                    'If Re_Count <> 0 Then
                    '    For i = 0 To Re_Count - 1
                    '        Grid_Info_Combo.Items.Add(DBT.Rows(i)("CM_Code"))
                    '    Next i
                    'End If
                    'Grid_Info_Combo.Top = Grid_Info.Top + Grid_Info.ColumnHeadersHeight + (Grid_Info_Row) * 24
                    'WW = 0
                    'For i = 0 To Grid_Info_Col - 1
                    '    WW = WW + Grid_Info.Columns(i).Width
                    'Next
                    'Grid_Info_Combo.Left = Grid_Info.Left + Grid_Info.RowHeadersWidth + WW
                    'Grid_Info_Combo.Width = Grid_Info.Columns(Grid_Info_Col).Width
                    'Grid_Info_Combo.Text = Grid_Info.CurrentCell.Value.ToString
                    'Grid_Info_Combo.BackColor = Grid_Info.RowsDefaultCellStyle.BackColor
                    'Grid_Info_Combo.Visible = True
                    'Grid_Info_Combo.Focus.ToString()
                Case 5
                    '팝업창 띄우기
                    Dim popup_account As New popup_account
                    popup_account.ShowDialog()

                    '결과값이 ok일때 그리드뷰에 값넣어주기
                    If popup_account.DialogResult = DialogResult.OK Then
                        Grid_Info.Item(Grid_Info_Col, 5).Value = popup_account.Customer_Name
                        Grid_Info.Item(Grid_Info_Col, 6).Value = popup_account.Customer_Man_Name
                        Grid_Info.Item(Grid_Info_Col, 7).Value = popup_account.Customer_Man_Tel
                        Grid_Info.Item(Grid_Info_Col, 4).Value = popup_account.Customer_Code
                    End If
                    'StrSQL = "Select CM_Name  From SI_CUSTOMER With(NOLOCK) Order By CM_Name"
                    'Re_Count = DB_Select(DBT)
                    'Grid_Info_Combo.Items.Clear()
                    'If Re_Count <> 0 Then
                    '    For i = 0 To Re_Count - 1
                    '        Grid_Info_Combo.Items.Add(DBT.Rows(i)("CM_Name"))
                    '    Next i
                    'End If
                    'Grid_Info_Combo.Top = Grid_Info.Top + Grid_Info.ColumnHeadersHeight + (Grid_Info_Row) * 24
                    'WW = 0
                    'For i = 0 To Grid_Info_Col - 1
                    '    WW = WW + Grid_Info.Columns(i).Width
                    'Next
                    'Grid_Info_Combo.Left = Grid_Info.Left + Grid_Info.RowHeadersWidth + WW
                    'Grid_Info_Combo.Width = Grid_Info.Columns(Grid_Info_Col).Width
                    'Grid_Info_Combo.Text = Grid_Info.CurrentCell.Value.ToString
                    'Grid_Info_Combo.BackColor = Grid_Info.RowsDefaultCellStyle.BackColor
                    'Grid_Info_Combo.Visible = True
                    'Grid_Info_Combo.Focus.ToString()
                Case Else
                    Exit Sub
            End Select
        End If
    End Sub
    Private Sub Grid_Info_Combo_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Grid_Info_Combo.SelectionChangeCommitted
        Dim DBT As New DataTable
        Grid_Info.Item(Grid_Info_Col, Grid_Info_Row).Value = Grid_Info_Combo.SelectedItem.ToString()

        If Grid_Info_Row = 4 Then
            StrSQL = "Select CM_Code, CM_Name, CM_Man_Name, CM_Man_Tel  From SI_CUSTOMER With(NOLOCK) Where CM_Code = '" & Grid_Info_Combo.SelectedItem.ToString() & "'"
            Re_Count = DB_Select(DBT)
            Grid_Info_Combo.Items.Clear()
            If Re_Count <> 0 Then
                Grid_Info.Item(Grid_Info_Col, 5).Value = DBT.Rows(0)("CM_Name")
                Grid_Info.Item(Grid_Info_Col, 6).Value = DBT.Rows(0)("CM_Man_Name")
                Grid_Info.Item(Grid_Info_Col, 7).Value = DBT.Rows(0)("CM_Man_Tel")
                Grid_Info.Rows(4).HeaderCell.Value = "*"
                Grid_Info.Rows(5).HeaderCell.Value = "*"
                Grid_Info.Rows(6).HeaderCell.Value = "*"
                Grid_Info.Rows(7).HeaderCell.Value = "*"
            End If
        End If
        If Grid_Info_Row = 5 Then
            StrSQL = "Select CM_Code, CM_Name, CM_Man_Name, CM_Man_Tel  From SI_CUSTOMER With(NOLOCK) Where CM_Name = '" & Grid_Info_Combo.SelectedItem.ToString() & "'"
            Re_Count = DB_Select(DBT)
            Grid_Info_Combo.Items.Clear()
            If Re_Count <> 0 Then
                Grid_Info.Item(Grid_Info_Col, 4).Value = DBT.Rows(0)("CM_Code")
                Grid_Info.Item(Grid_Info_Col, 6).Value = DBT.Rows(0)("CM_Man_Name")
                Grid_Info.Item(Grid_Info_Col, 7).Value = DBT.Rows(0)("CM_Man_Tel")
                Grid_Info.Rows(4).HeaderCell.Value = "*"
                Grid_Info.Rows(5).HeaderCell.Value = "*"
                Grid_Info.Rows(6).HeaderCell.Value = "*"
                Grid_Info.Rows(7).HeaderCell.Value = "*"
            End If
        End If
        Grid_Info.Rows(Grid_Info_Row).HeaderCell.Value = "*"
        Grid_Info_Combo.Visible = False


    End Sub

    Private Sub Grid_Info_Combo_LostFocus(sender As Object, e As EventArgs) Handles Grid_Info_Combo.LostFocus
        Grid_Info_Combo.Visible = False

    End Sub

    Private Sub Save_Com_Click(sender As Object, e As EventArgs) Handles Save_Com.Click
        'Grid_Info 를 저장 한다.
        'DataBase에서 필드 이름 가지고 오기

        If Grid_Code.Item(0, 0).Value = "" Then
            MsgBox("공백은 저장할 수 없습니다")
            Exit Sub
        End If

        Dim Table_Name As String
        Dim j As Long
        Dim DBT As DataTable = Nothing
        Dim Field_Name(500) As String
        Dim i As Integer
        j = 0
        For i = 1 To Grid_Info.RowCount - 1
            If Grid_Info.Rows(i).HeaderCell.Value = "*" Then
                j = 1
            End If
        Next i
        If j = 0 Then
        Else
            Table_Name = "MC_STOCK_IN"
            j = 0
            StrSQL = "Select sys.Columns.Name From sys.Tables with(NOLOCK) , sys.Columns with(NOLOCK) Where sys.Tables.name = '" & Table_Name & "' And sys.Tables.Object_id = sys.Columns.Object_id Order By sys.Columns.column_id"
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

        If Grid_Main.Item(0, 0).Value = "1" Then
        Else
            Exit Sub

        End If




        MsgBox("저장되었습니다")

        ''Grid_Info 를 저장 한다.
        ''DataBase에서 필드 이름 가지고 오기


        'Dim Table_Name As String
        'Dim j As Long
        'Dim DBT As DataTable = Nothing
        'Dim Field_Name(500) As String
        'Dim i As Integer
        'j = 0

        'If Grid_Code.Item(0, 0).Value = "" Then
        '    MsgBox("공백은 저장할 수 없습니다")
        '    Exit Sub
        'End If


        'For i = 1 To Grid_Info.RowCount - 1
        '    If Grid_Info.Rows(i).HeaderCell.Value = "*" Then
        '        j = 1
        '    End If
        'Next i
        'If j = 0 Then
        'Else
        '    Table_Name = "MC_STOCK_IN"
        '    j = 0
        '    StrSQL = "Select sys.Columns.Name From sys.Tables with(NOLOCK) , sys.Columns with(NOLOCK) Where sys.Tables.name = '" & Table_Name & "' And sys.Tables.Object_id = sys.Columns.Object_id Order By sys.Columns.column_id"
        '    Re_Count = DB_Select(DBT)
        '    If Re_Count <> 0 Then
        '        For i = 0 To DBT.Rows.Count - 1
        '            Field_Name(j) = DBT.Rows(i)("Name").ToString
        '            j = j + 1
        '        Next i
        '    End If
        '    j = j - 1
        '    StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        '    StrSQL = StrSQL & "UPDATE " & Table_Name & " SET "
        '    Field_Name(500) = "Ok"
        '    For i = 1 To j
        '        If Grid_Info.Rows(i).HeaderCell.Value = "*" Then
        '            If Field_Name(500) = "" Then
        '                StrSQL = StrSQL & ","
        '            End If
        '            StrSQL = StrSQL & " " & Table_Name & "." & Field_Name(i) & " = '" & Replace(Grid_Info.Item(1, i).Value, "'", "''") & "'"
        '            If Field_Name(500) = "Ok" Then
        '                Field_Name(500) = ""
        '            End If
        '        End If
        '    Next i
        '    StrSQL = StrSQL & " WHERE " & Table_Name & "." & Field_Name(0) & " = '" & Grid_Info.Item(1, 0).Value & "'"
        '    Re_Count = DB_Execute()
        'End If

        'If Grid_Main.Item(0, 0).Value = "1" Then
        'Else
        '    Exit Sub

        'End If


        ''수량 변경
        ''수량이 변경 되었는지 확인 한다.
        'Dim Val_Check As String
        'Dim PP_G_Sun As String
        'Dim PP_G_Amount As String

        'Val_Check = ""
        'StrSQL = "Select CS_Check From MC_STOCK_IN With(NOLOCK) Where CS_Code = '" & Grid_Info.Item(1, 0).Value & "'"
        'Re_Count = DB_Select(DBT)
        'Grid_Info_Combo.Items.Clear()
        'If Re_Count <> 0 Then
        '    If DBT.Rows(0)("CS_Check") = "처리" Then
        '        Val_Check = "처리"
        '        MsgBox("이미 처리 되었습니다. 삭제 후 다시 처리 하세요")
        '        Exit Sub
        '    End If
        'End If


        ''수량을 변경한다.
        'For i = 0 To Grid_Main.RowCount - 1
        '    PP_G_Sun = "0"
        '    PP_G_Amount = "0"
        '    StrSQL = "Select PP_Sun, PP_Amount From SI_PROCESS_LIST With(NOLOCK) Where PP_Code = '" & Grid_Main.Item(1, i).Value & "' Order By Convert(Decimal,PP_Sun)"
        '    Re_Count = DB_Select(DBT)
        '    If Re_Count > 0 Then
        '        PP_G_Sun = DBT.Rows(0)("PP_Sun")
        '        PP_G_Amount = DBT.Rows(0)("PP_Amount")
        '    End If
        '    If PP_G_Sun = "0" Then
        '    Else
        '        '변경
        '        PP_G_Amount = Val(PP_G_Amount) + Val(Grid_Main.Item(6, i).Value)
        '        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        '        StrSQL = StrSQL & "Update SI_PROCESS_LIST Set PP_Amount = '" & PP_G_Amount & "' Where PP_Code = '" & Grid_Main.Item(1, i).Value & "' And PP_Sun = '" & PP_G_Sun & "'"
        '        Re_Count = DB_Execute()
        '    End If
        'Next i

        'StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        'StrSQL = StrSQL & "Update MC_STOCK_IN Set CS_Check = '처리' Where CS_Code = '" & Grid_Info.Item(1, 0).Value & "'"
        'Re_Count = DB_Execute()

        'Grid_Info.Item(1, 9).Value = "처리"
        'MsgBox("저장되었습니다")

    End Sub

    Private Sub Insert__Main_Click(sender As Object, e As EventArgs) Handles Insert__Main.Click
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        If Grid_Info.Item(1, 9).Value = "처리" Then
            MsgBox("수량 처리가 되어서 수정 할 수 없습니다.")
            Exit Sub
        End If

        Grid_Display_Ch = False

        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        Grid_Display_Ch = False
        Dim DBT As New DataTable
        Dim Db_Sun As Long
        StrSQL = "Select CS_Sun From MC_STOCK_IN_List with(NOLOCK) Where CS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' Order By Convert(Decimal,CS_Sun) Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            Db_Sun = 1
        Else
            Db_Sun = DBT.Rows(0).Item(0) + 1
        End If


        ''새로운 코드 추가
        'Dim DBT As New DataTable
        'Dim JP_Code As Long
        'Dim My_Date As String
        'Dim My_Time As String

        'StrSQL = "Select GETDATE() "
        'Re_Count = DB_Select(DBT)


        'If Re_Count = 0 Then
        '    Exit Sub
        'Else
        '    My_Date = DBT.Rows(0).Item(0)
        '    JP_Code = Mid(My_Date, 1, 4) & Mid(My_Date, 6, 2) & Mid(My_Date, 9, 2)
        '    My_Time = DateTime.Now.ToString("HH:mm:ss")
        '    My_Date = Mid(My_Date, 1, 4) & "-" & Mid(My_Date, 6, 2) & "-" & Mid(My_Date, 9, 2)
        'End If


        ''Dim DBT As New DataTable
        'Dim Db_Sun As Long
        'StrSQL = "Select CS_Sun From MC_STOCK_IN_List with(NOLOCK) Where CS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' Order By Convert(Decimal,CS_Sun) Desc "
        'Re_Count = DB_Select(DBT)
        'If Re_Count = 0 Then
        '    Db_Sun = 1
        'Else
        '    Db_Sun = DBT.Rows(0).Item(0) + 1
        'End If

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert into MC_STOCK_IN_List (CS_Code, CS_Sun)  Values ('" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "', '" & Db_Sun & "')"
        Re_Count = DB_Execute()
        Grid_Main_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True
    End Sub

    Private Sub Grid_Main_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Main.CellValueChanged
        If Grid_Display_Ch = True Then
            Grid_Main.CurrentRow.HeaderCell.Value = "*"
            '콤마 넣으면 값 계산이 이상하므로 콤바를 넣어도 계산되게
            If Len(Grid_Main.Item(5, Grid_Main.CurrentCell.RowIndex).Value) > 1 Then
                Grid_Main.Item(5, Grid_Main.CurrentCell.RowIndex).Value = Replace(Grid_Main.Item(5, Grid_Main.CurrentCell.RowIndex).Value.ToString, ",", "")
            End If

            Grid_Main.Item(7, Grid_Main.CurrentCell.RowIndex).Value = Val(Grid_Main.Item(5, Grid_Main.CurrentCell.RowIndex).Value.ToString) * Val(Grid_Main.Item(6, Grid_Main.CurrentCell.RowIndex).Value.ToString)

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
                'StrSQL = "Select SI_PRODUCT.PL_Code, SI_PRODUCT.PL_Name  From SI_PRODUCT with(NOLOCK), SI_PRODUCT_CUSTOMER with(NOLOCK) Where SI_PRODUCT_CUSTOMER.PL_CM_Code = '" & Grid_Info.Item(1, 4).Value & "' And SI_PRODUCT_CUSTOMER.PL_Code = SI_PRODUCT.PL_Code And SI_PRODUCT_CUSTOMER.PL_Sel = '매입' Order By PL_Code"
                'Re_Count = DB_Select(DBT)
                'Grid_Main_Combo.Items.Clear()
                'If Re_Count <> 0 Then
                '    For i = 0 To Re_Count - 1
                '        Grid_Main_Combo.Items.Add(DBT.Rows(i).Item(0))
                '    Next i
                'End If
                'Grid_Main_Combo.Top = Grid_Main.Top + Grid_Main.ColumnHeadersHeight + (Grid_Main_Row) * 24
                'Grid_Main_Combo.Left = Grid_Main.Left + Grid_Main.RowHeadersWidth + Grid_Main.Columns(0).Width
                'Grid_Main_Combo.Width = Grid_Main.Columns(Grid_Main_Col).Width
                'Grid_Main_Combo.Text = Grid_Main.CurrentCell.Value.ToString
                'Grid_Main_Combo.BackColor = Grid_Main.RowsDefaultCellStyle.BackColor
                'Grid_Main_Combo.Visible = True
                'Grid_Main_Combo.Focus.ToString()
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
                'StrSQL = "Select SI_PRODUCT.PL_Code, SI_PRODUCT.PL_Name  From SI_PRODUCT with(NOLOCK), SI_PRODUCT_CUSTOMER with(NOLOCK) Where SI_PRODUCT_CUSTOMER.PL_CM_Code = '" & Grid_Info.Item(1, 4).Value & "' And SI_PRODUCT_CUSTOMER.PL_Code = SI_PRODUCT.PL_Code And SI_PRODUCT_CUSTOMER.PL_Sel = '매입' Order By PL_Name"
                'Re_Count = DB_Select(DBT)
                'Grid_Main_Combo.Items.Clear()
                'If Re_Count <> 0 Then
                '    For i = 0 To Re_Count - 1
                '        Grid_Main_Combo.Items.Add(DBT.Rows(i).Item(1))
                '    Next i
                'End If
                'Grid_Main_Combo.Top = Grid_Main.Top + Grid_Main.ColumnHeadersHeight + (Grid_Main_Row) * 24
                'Grid_Main_Combo.Left = Grid_Main.Left + Grid_Main.RowHeadersWidth + Grid_Main.Columns(0).Width + Grid_Main.Columns(1).Width
                'Grid_Main_Combo.Width = Grid_Main.Columns(Grid_Main_Col).Width
                'Grid_Main_Combo.Text = Grid_Main.CurrentCell.Value.ToString
                'Grid_Main_Combo.BackColor = Grid_Main.RowsDefaultCellStyle.BackColor
                'Grid_Main_Combo.Visible = True
                'Grid_Main_Combo.Focus.ToString()
        End Select


    End Sub
    Private Sub Grid_Main_Scroll(sender As Object, e As ScrollEventArgs) Handles Grid_Main.Scroll
        Grid_Main_Combo.Visible = False
    End Sub
    Private Sub Grid_Main_Combo_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Grid_Main_Combo.SelectionChangeCommitted
        Dim DBT As New DataTable
        Grid_Main.Item(Grid_Main_Col, Grid_Main_Row).Value = Grid_Main_Combo.SelectedItem.ToString()
        Select Case Grid_Main_Col
            Case 1
                StrSQL = "Select PL_Code, PL_Name, PL_Standard ,PL_Unit, PL_Unit_Gold From SI_PRODUCT with(NOLOCK) Where PL_Code = '" & Grid_Main_Combo.SelectedItem.ToString() & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Main.Item(2, Grid_Main_Row).Value = DBT.Rows(0).Item(1)
                    Grid_Main.Item(3, Grid_Main_Row).Value = DBT.Rows(0).Item(2)
                    Grid_Main.Item(4, Grid_Main_Row).Value = DBT.Rows(0).Item(3)
                    Grid_Main.Item(5, Grid_Main_Row).Value = DBT.Rows(0).Item(4)

                End If
            Case 2
                StrSQL = "Select PL_Code, PL_Name, PL_Standard ,PL_Unit, PL_Unit_Gold  From SI_PRODUCT with(NOLOCK) Where PL_Name= '" & Replace(Grid_Main_Combo.SelectedItem.ToString(), "'", "''") & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Main.Item(1, Grid_Main_Row).Value = DBT.Rows(0).Item(0)
                    Grid_Main.Item(3, Grid_Main_Row).Value = DBT.Rows(0).Item(2)
                    Grid_Main.Item(4, Grid_Main_Row).Value = DBT.Rows(0).Item(3)
                    Grid_Main.Item(5, Grid_Main_Row).Value = DBT.Rows(0).Item(4)
                End If
        End Select
        Grid_Main_Combo.Visible = False
    End Sub

    Private Sub Grid_Main_Combo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Grid_Main_Combo.SelectedIndexChanged

    End Sub

    Private Sub Grid_Main_Combo_LostFocus(sender As Object, e As EventArgs) Handles Grid_Main_Combo.LostFocus
        Grid_Main_Combo.Visible = False

    End Sub

    Private Sub Save_Main_Click(sender As Object, e As EventArgs) Handles Save_Main.Click
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        If Grid_Info.Item(1, 9).Value = "처리" Then
            MsgBox("수량 처리가 되어서 수정 할 수 없습니다.")
            Exit Sub
        End If

        Dim i As Integer
        Dim UD As Integer
        Grid_Display_Ch = False
        For i = 0 To Grid_Main.Rows.Count - 1
            If Grid_Main.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "update MC_STOCK_IN_List Set CS_PL_Code = '" & Grid_Main.Item(1, i).Value & "', 
                                        CS_PL_Name = '" & Grid_Main.Item(2, i).Value & "', CS_Standard = '" & Grid_Main.Item(3, i).Value & "', 
                                        CS_Unit = '" & Grid_Main.Item(4, i).Value & "', CS_Unit_Amount = '" & Grid_Main.Item(5, i).Value & "', 
                                        CS_Total = '" & Grid_Main.Item(6, i).Value & "', CS_Gold = '" & Grid_Main.Item(7, i).Value & "', 
                                        CS_Bigo = '" & Grid_Main.Item(8, i).Value & "'

                            where CS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And CS_Sun = '" & Grid_Main.Item(0, i).Value & "'"
                Re_Count = DB_Execute()
                Grid_Main.Rows(i).HeaderCell.Value = ""
                UD = 1
            End If
        Next i
        Grid_Display_Ch = True

        If UD <> 1 Then
            For i = 0 To Grid_Main.Rows.Count - 1
                StrSQL = "INSERT INTO MC_STOCK_IN_LIST (CS_Code , CS_Sun , CS_PL_Code , CS_PL_Name,
CS_Standard , CS_Unit , CS_Unit_Amount , CS_Total, CS_Gold, CS_Bigo) VALUES ('" & Grid_Info.Item(1, 0).Value & "','" & Grid_Main.Item(0, i).Value & "','" & Grid_Main.Item(1, i).Value & "','" & Grid_Main.Item(2, i).Value & "','" & Grid_Main.Item(3, i).Value & "','" & Grid_Main.Item(4, i).Value & "','" & Grid_Main.Item(5, i).Value & "','" & Grid_Main.Item(6, i).Value & "','" & Grid_Main.Item(7, i).Value & "','" & Grid_Main.Item(8, i).Value & "')"
                Re_Count = DB_Execute()
            Next i
        End If

        MsgBox("저장되었습니다")
    End Sub

    Private Sub Grid_Main_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Main.CellContentClick

    End Sub

    Private Sub Delete__Main_Click(sender As Object, e As EventArgs) Handles Delete__Main.Click
        '해당 칼럼 삭제
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        If Grid_Info.Item(1, 9).Value = "처리" Then
            MsgBox("수량 처리가 되어서 수정 할 수 없습니다.")
            Exit Sub
        End If

        Grid_Display_Ch = False
        Dim i As Integer
        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Delete MC_STOCK_IN_List Where CS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And  CS_Sun = '" & Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        For i = Grid_Main.CurrentCell.RowIndex + 1 To Grid_Main.RowCount - 1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "Update  MC_STOCK_IN_List Set CS_Sun = '" & i & "' Where CS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And  CS_Sun = '" & i + 1 & "'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        Grid_Main_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True
    End Sub

    Private Sub Delete_Com_Click(sender As Object, e As EventArgs) Handles Delete_Com.Click

        '해당 칼럼 삭제
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        If MsgBox("입고전표  '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "'를 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "입고전표 삭제") = vbNo Then
            Exit Sub
        End If


        '수량 변경
        '수량이 변경 되었는지 확인 한다.
        Dim Val_Check As String
        Dim PP_G_Sun As String
        Dim PP_G_Amount As String
        Dim DBT As DataTable = Nothing

        Val_Check = ""
        StrSQL = "Select CS_Check From MC_STOCK_IN With(NOLOCK) Where CS_Code = '" & Grid_Info.Item(1, 0).Value & "'"
        Re_Count = DB_Select(DBT)
        Grid_Info_Combo.Items.Clear()
        If Re_Count <> 0 Then
            If DBT.Rows(0)("CS_Check") = "처리" Then
                Val_Check = "처리"
            End If
        End If


        If Val_Check = "처리" Then
            '기존 Data를 삭제 한다.
            'Grid의 수량을 가지고 온다.
            For i = 0 To Grid_Main.RowCount - 1
                PP_G_Sun = "0"
                PP_G_Amount = "0"
                StrSQL = "Select PP_Sun, PP_Amount From SI_PROCESS_LIST With(NOLOCK) Where PP_Code = '" & Grid_Main.Item(1, i).Value & "' Order By Convert(Decimal,PP_Sun) "
                Re_Count = DB_Select(DBT)
                If Re_Count > 0 Then
                    PP_G_Sun = DBT.Rows(0)("PP_Sun")
                    PP_G_Amount = DBT.Rows(0)("PP_Amount")
                End If
                If PP_G_Sun = "0" Then
                Else
                    '변경
                    PP_G_Amount = Val(PP_G_Amount) - Val(Grid_Main.Item(6, i).Value)
                    StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                    StrSQL = StrSQL & "Update SI_PROCESS_LIST Set PP_Amount = '" & PP_G_Amount & "' Where PP_Code = '" & Grid_Main.Item(1, i).Value & "' And PP_Sun = '" & PP_G_Sun & "'"
                    Re_Count = DB_Execute()
                End If
            Next i

        End If





        Grid_Display_Ch = False
        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Delete MC_STOCK_IN Where CS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' "
        Re_Count = DB_Execute()

        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Delete MC_STOCK_IN_List Where CS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' "
        Re_Count = DB_Execute()

        Grid_Code_Display()
        Grid_Info_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True
    End Sub
    Private Sub Grid_Info_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Info.CellValueChanged
        If Grid_Display_Ch = True Then
            For i = 0 To Grid_Info.RowCount - 1
                Grid_Info.Rows(i).HeaderCell.Value = "*"
            Next i
            'Main_Grid_day()
        End If
    End Sub

    Private Sub Panel_Menu_Paint(sender As Object, e As PaintEventArgs) Handles Panel_Menu.Paint

    End Sub

    Private Sub Info_Co3_MouseClick(sender As Object, e As MouseEventArgs) Handles Info_Co4.MouseClick, Info_Co5.MouseClick
        Dim index As Integer

        If Grid_Display_Ch = False Then
            Exit Sub
        Else
            'MsgBox(Mid(sender.name.ToString, 8, 2))
        End If

        index = Val(Mid(sender.name.ToString, 8, 2))
        Select Case index
            Case 4
                StrSQL = "Select CM_Code  From SI_CUSTOMER With(NOLOCK) Order By CM_Code"
            Case 5
                StrSQL = "Select CM_Name  From SI_CUSTOMER With(NOLOCK) Order By CM_Name"
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

    Private Sub Info_Co0_SelectedValueChanged(sender As Object, e As EventArgs) Handles Info_Co4.SelectedValueChanged, Info_Co5.SelectedValueChanged
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
                StrSQL = "Select CM_Code, CM_Name, CM_Man_Name, CM_Man_Tel  From SI_CUSTOMER With(NOLOCK) Where CM_Code = '" & sender.text.ToString() & "'"
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
                StrSQL = "Select CM_Code, CM_Name, CM_Man_Name, CM_Man_Tel  From SI_CUSTOMER With(NOLOCK) Where CM_Name = '" & sender.text.ToString() & "'"
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

    Private Sub Info_Tx0_TextChanged(sender As Object, e As EventArgs) Handles Info_Tx0.TextChanged, Info_Tx1.TextChanged, Info_Tx2.TextChanged, Info_Tx3.TextChanged, Info_Tx4.TextChanged, Info_Tx5.TextChanged, Info_Tx6.TextChanged, Info_Tx7.TextChanged, Info_Tx8.TextChanged, Info_Tx9.TextChanged, Info_Tx10.TextChanged
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

        '여기서 처리
        If Excel_check = False Then
        Else
            xlbook.Close()
            xlapp.Quit()

            xlsheet = Nothing
            xlbook = Nothing
            xlapp = Nothing
            releaseObject(xlsheet)
            releaseObject(xlbook)
            releaseObject(xlapp)
            Excel_check = True
        End If


        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer





        xlapp = New Excel.Application
        xlbook = xlapp.Workbooks.Open(Application.StartupPath & "\excel\부재료 입고검사.xlsx")
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

        'PInvoke 에러시 프로젝트 속성에서 컴파일을 x64로 하거나 Any CPU일 경우 32비트 기본 사용 체크를 해제한다.
        'SetWindowLong(xlapp.Hwnd, GWL_STYLE, GetWindowLong(xlapp.Hwnd, GWL_STYLE) And Not WS_CAPTION)

        '거래처명
        xlsheet.Cells(4, 3).Value = Grid_Info.Item(1, 5).Value
        ''지불조건
        'xlsheet.Cells(5, 4).Value = Grid_Info.Item(1, 8).Value
        '담당자 
        ' xlsheet.Cells(6, 4).Value = Grid_Info.Item(1, 6).Value


        '입고일자
        xlsheet.Cells(3, 1).Value = Grid_Info.Item(1, 8).Value
        '입고코드
        '   xlsheet.Cells(5, 14).Value = Grid_Info.Item(1, 0).Value
        '담당자 전화 번호
        '   xlsheet.Cells(6, 14).Value = Grid_Info.Item(1, 7).Value

        ''납품장소
        'StrSQL = "Select CM_Name  From Company with(NOLOCK) Where CM_Code= '10000'"
        'Re_Count = DB_Select(DBT)
        'If Re_Count <> 0 Then
        '    xlsheet.Cells(30, 5).Value = DBT.Rows(0).Item(0)
        'End If


        ''일자
        'xlsheet.Cells(10, 9).Value = Mid(Grid_Info.Item(1, 9).Value, 1, 4) & " " & Mid(Grid_Info.Item(1, 9).Value, 6, 5)
        'xlsheet.Cells(10, 10).Value = Mid(Grid_Info.Item(1, 10).Value, 1, 4) & " " & Mid(Grid_Info.Item(1, 10).Value, 6, 5)
        'xlsheet.Cells(10, 11).Value = Mid(Grid_Info.Item(1, 11).Value, 1, 4) & " " & Mid(Grid_Info.Item(1, 11).Value, 6, 5)

        j = 0
        '    For i = 0 To Grid_Main.RowCount - 1
        xlsheet.Cells(5, 3).Value = Grid_Main.Item(2, 0).Value.ToString '제품명
        xlsheet.Cells(6, 3).Value = Grid_Main.Item(6, 0).Value.ToString '수량
        '  xlsheet.Cells(7, 3).Value = Grid_Info.Item(1, 11).Value.ToString '도축일자
        ' xlsheet.Cells(8, 3).Value = Grid_Info.Item(1, 12).Value.ToString '유통기한

        xlsheet.Cells(10, 1).Value = Grid_Main.Item(9, 0).Value.ToString '보관방법
        xlsheet.Cells(10, 2).Value = Grid_Main.Item(10, 0).Value.ToString '이물
        xlsheet.Cells(10, 3).Value = Grid_Main.Item(11, 0).Value.ToString '이취
        xlsheet.Cells(10, 4).Value = Grid_Main.Item(12, 0).Value.ToString '성상

        xlsheet.Cells(10, 6).Value = Grid_Main.Item(14, 0).Value.ToString '포장상태
        xlsheet.Cells(10, 7).Value = Grid_Main.Item(13, 0).Value.ToString '검사성적서
        ' xlsheet.Cells(15, 1).Value = Grid_Main.Item(15, 0).Value.ToString '차량온도기록지


        '  Next i


        '합계

        '    xlsheet.Cells(27, 6).Value = j

        ''거래 일자
        'j = 0
        'For j = 0 To Grid_Main.ColumnCount - 1
        '    xlsheet.Cells(1, j + 1).Value = "'" & Grid_Main.Columns(j).HeaderText
        'Next j
        'For i = 0 To Grid_Main.RowCount - 1
        '    '제목표시
        '    For j = 0 To Grid_Main.ColumnCount - 1
        '        xlsheet.Cells(i + 2, j + 1).Value = Grid_Main.Item(j, i).Value.ToString
        '    Next j
        'Next i
        xlbook.SaveAs(Application.StartupPath & "\Excel\입고검사\" & Grid_Info.Item(1, 0).Value & ".xlsx")
        Cursor = Cursors.Default
        Excel_check = True


        ''거래처명
        'xlsheet.Cells(4, 4).Value = Grid_Info.Item(1, 5).Value
        '''지불조건
        ''xlsheet.Cells(5, 4).Value = Grid_Info.Item(1, 8).Value
        ''담당자 
        'xlsheet.Cells(6, 4).Value = Grid_Info.Item(1, 6).Value


        ''입고일자
        'xlsheet.Cells(4, 14).Value = Grid_Info.Item(1, 8).Value
        ''입고코드
        'xlsheet.Cells(5, 14).Value = Grid_Info.Item(1, 0).Value
        ''담당자 전화 번호
        'xlsheet.Cells(6, 14).Value = Grid_Info.Item(1, 7).Value

        ''납품장소
        'StrSQL = "Select CM_Name  From Company with(NOLOCK) Where CM_Code= '10000'"
        'Re_Count = DB_Select(DBT)
        'If Re_Count <> 0 Then
        '    xlsheet.Cells(30, 5).Value = DBT.Rows(0).Item(0)
        'End If


        '''일자
        ''xlsheet.Cells(10, 9).Value = Mid(Grid_Info.Item(1, 9).Value, 1, 4) & " " & Mid(Grid_Info.Item(1, 9).Value, 6, 5)
        ''xlsheet.Cells(10, 10).Value = Mid(Grid_Info.Item(1, 10).Value, 1, 4) & " " & Mid(Grid_Info.Item(1, 10).Value, 6, 5)
        ''xlsheet.Cells(10, 11).Value = Mid(Grid_Info.Item(1, 11).Value, 1, 4) & " " & Mid(Grid_Info.Item(1, 11).Value, 6, 5)

        'j = 0
        'For i = 0 To Grid_Main.RowCount - 1
        '    xlsheet.Cells(12 + i, 2).Value = Grid_Main.Item(0, i).Value.ToString '순번
        '    xlsheet.Cells(12 + i, 3).Value = Grid_Main.Item(1, i).Value.ToString '제품코드
        '    xlsheet.Cells(12 + i, 5).Value = Grid_Main.Item(2, i).Value.ToString '제품명
        '    xlsheet.Cells(12 + i, 7).Value = Grid_Main.Item(3, i).Value.ToString '규격
        '    xlsheet.Cells(12 + i, 8).Value = Grid_Main.Item(4, i).Value.ToString '단위
        '    xlsheet.Cells(12 + i, 9).Value = Grid_Main.Item(5, i).Value.ToString '단가
        '    xlsheet.Cells(12 + i, 10).Value = Grid_Main.Item(6, i).Value.ToString '수량
        '    xlsheet.Cells(12 + i, 11).Value = Grid_Main.Item(7, i).Value.ToString '금액
        '    xlsheet.Cells(12 + i, 13).Value = Grid_Main.Item(8, i).Value.ToString '비고
        '    'xlsheet.Cells(12 + i, 12).Value = Grid_Main.Item(9, i).Value.ToString
        '    'xlsheet.Cells(12 + i, 13).Value = Grid_Main.Item(10, i).Value.ToString
        '    'xlsheet.Cells(12 + i, 15).Value = Grid_Main.Item(11, i).Value.ToString
        '    j = j + Val(Grid_Main.Item(8, i).Value.ToString)
        'Next i


        ''합계

        'xlsheet.Cells(27, 6).Value = j

        '''거래 일자
        ''j = 0
        ''For j = 0 To Grid_Main.ColumnCount - 1
        ''    xlsheet.Cells(1, j + 1).Value = "'" & Grid_Main.Columns(j).HeaderText
        ''Next j
        ''For i = 0 To Grid_Main.RowCount - 1
        ''    '제목표시
        ''    For j = 0 To Grid_Main.ColumnCount - 1
        ''        xlsheet.Cells(i + 2, j + 1).Value = Grid_Main.Item(j, i).Value.ToString
        ''    Next j
        ''Next i
        'xlbook.SaveAs(Application.StartupPath & "\Excel\입고검사\" & Grid_Info.Item(1, 0).Value & Grid_Info.Item(1, 5).Value & ".xlsx")
        'Cursor = Cursors.Default
        'Excel_check = True



    End Sub

    Private Sub Com_Main_Click(sender As Object, e As EventArgs) Handles Com_Main.Click
        If Excel_check = True Then

            xlbook.Close()
            xlapp.Quit()

            xlsheet = Nothing
            xlbook = Nothing
            xlapp = Nothing

            releaseObject(xlsheet)
            releaseObject(xlbook)
            releaseObject(xlapp)
            Excel_check = False
        End If

        Panel_Main.Visible = True
        Panel_Excel.Visible = False
        Com_Excel_Print.Visible = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If (Grid_Info.Item(1, 5).Value = "") Then
            MsgBox("거래처 명이 비워져있습니다.")
        Else
            Dim Dbt As New DataTable
            Dim checkdata As String
            StrSQL = "select count(CO_Name) from MC_STOCK_ORDER_LIST JOIN MC_STOCK_ORDER ON MC_STOCK_ORDER.CO_Code=MC_STOCK_ORDER_LIST.CO_Code where CO_Customer_Name='" + Grid_Info.Item(1, 5).Value + "';"
            Re_Count = DB_Select(Dbt)

            checkdata = Dbt.Rows(0)(0).ToString()
            If (checkdata = "0") Then
                MsgBox("발주가 없습니다")
            Else
                DataGridView1_Display()
            End If
        End If


    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Panel1.Visible = False
    End Sub

    Private Sub DataGridView1_CellMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDoubleClick


        Info_Co11.Text = DataGridView1.Item(0, DataGridView1.CurrentCell.RowIndex).Value
        Grid_Info.Item(1, 11).Value = DataGridView1.Item(0, DataGridView1.CurrentCell.RowIndex).Value

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display_Ch = False

        Grid_Main.RowCount = 0

        'MsgBox(Today)
        'MsgBox(DateTime.Now.ToString)
        'MsgBox(DateTime.Today)


        StrSQL = "select CO_Sun, CO_PL_Code, CO_PL_Name, CO_Standard, CO_Unit, CO_Unit_Amount, CO_Total, CO_Gold, CO_Bigo 
            from MC_STOCK_ORDER_LIST where CO_Code= '" & DataGridView1.Item(0, DataGridView1.CurrentCell.RowIndex).Value & "' Order By Convert(Decimal,CO_Sun)"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Main.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To Grid_Main.ColumnCount - 3
                    Grid_Main.Item(j, i).Value = DBT.Rows(i).Item(j).ToString

                Next j
            Next i
        Else
            Grid_Main.RowCount = 1
            For j = 0 To Grid_Main.ColumnCount - 1
                Grid_Main.Item(j, 0).Value = ""
            Next j
        End If

        'Grid_Main.Item(9, Grid_Main.CurrentCell.RowIndex).Value = Int(Val(Grid_Main.Item(8, Grid_Main.CurrentCell.RowIndex).Value) * 0.1)


        Grid_Info_Combo.Visible = False
        Grid_Display_Ch = True


        ' Grid_Main_Display(DataGridView1.Item(0, DataGridView1.CurrentCell.RowIndex).Value)

        Panel1.Visible = False
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer


        StrSQL = "select MC_STOCK_ORDER.CO_Code, MC_STOCK_ORDER.CO_Date, MC_STOCK_ORDER.CO_Customer_Name,MC_STOCK_ORDER_LIST.CO_PL_Name,
MC_STOCK_ORDER_LIST.CO_Total from MC_STOCK_ORDER JOIN MC_STOCK_ORDER_LIST ON MC_STOCK_ORDER.CO_code=MC_STOCK_ORDER_LIST.CO_Code  where CO_PL_Name Like'%" + TextBox1.Text + "%'"

        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            DataGridView1.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To DataGridView1.ColumnCount - 1
                    DataGridView1.Item(j, i).Value = DBT.Rows(i).Item(j)
                Next j
            Next i
        Else
            DataGridView1.RowCount = 1
            For j = 0 To DataGridView1.ColumnCount - 1
                DataGridView1.Item(j, 0).Value = ""
            Next j
        End If
        Panel1.Visible = True
        TextBox1.Text = ""
    End Sub

    'Private Sub Grid_Info_Combo_TextChanged(sender As Object, e As EventArgs) Handles Grid_Info_Combo.TextChanged
    '    Dim P_Code As String
    '    Dim P_Su As String
    '    Dim i As Integer
    '    Dim DBT As New DataTable



    '    Dim JP_Code As Long
    '    Dim My_Date As String
    '    Dim My_Time As String
    '    If Len(TextBox1.Text) < 1 Then
    '    Else

    '        '추가 한다.
    '        StrSQL = "Select GETDATE() "
    '        Re_Count = DB_Select(DBT)

    '        Dim Db_Sun As Long
    '            StrSQL = "Select DL_Sun FROM SC_DELIVER_LIST with(NOLOCK) Where DL_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' Order By Convert(Decimal,DL_Sun) Desc "
    '            Re_Count = DB_Select(DBT)
    '            If Re_Count = 0 Then
    '                Db_Sun = 1
    '            Else
    '                Db_Sun = DBT.Rows(0).Item(0) + 1
    '            End If

    '        '추가한다
    '        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
    '        StrSQL = StrSQL & "Insert INTO MC_STOCK_IN_LIST ( CS_Sun)  Values ( '" & Db_Sun & "')"
    '        'Re_Count = DB_Execute()
    '    End If
    'End Sub
End Class
