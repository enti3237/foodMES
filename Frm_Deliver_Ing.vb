﻿
Imports Excel = Microsoft.Office.Interop.Excel

Public Class Frm_Deliver_Ing
    Dim Grid_Display_Ch As Boolean
    Dim Grid_Info_Row As Integer
    Dim Grid_Info_Col As Integer

    Dim Grid_Main_Row As Integer
    Dim Grid_Main_Col As Integer


    Dim xlapp As Excel.Application
    Dim xlbook As Excel.Workbook
    Dim xlsheet As Excel.Worksheet
    Dim Excel_check As Boolean
    Private Sub Frm_Deliver_Ing_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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
        Grid_Info.RowCount = 12





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
        Grid_Info.Item(0, 8).Value = "수주전표"
        Grid_Info.Item(0, 9).Value = "거래명세서양식"
        Grid_Info.Item(0, 10).Value = "수량"
        Grid_Info.Item(0, 11).Value = "비고"


        'Grid_Info.Columns(0).ReadOnly = True
        'Grid_Info.Columns(1).ReadOnly = True

        Grid_Info.Rows(0).ReadOnly = True
        Grid_Info.Rows(1).ReadOnly = True
        Grid_Info.Rows(2).ReadOnly = True
        Grid_Info.Rows(3).ReadOnly = True
        Grid_Info.Rows(10).ReadOnly = True

        Grid_Info_Setup = True
    End Function
    Public Function Grid_Main_Setup() As Boolean
        Grid_Color(Grid_Main)

        Grid_Main.AllowUserToAddRows = False
        Grid_Main.RowTemplate.Height = 24
        Grid_Main.ColumnCount = 11
        Grid_Main.RowCount = 0

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Main.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Grid_Main.RowHeadersWidth = 10


        Grid_Main.Columns(0).HeaderText = "순번"
        Grid_Main.Columns(1).HeaderText = "일자"
        Grid_Main.Columns(2).HeaderText = "제품코드"
        Grid_Main.Columns(3).HeaderText = "제품명"
        Grid_Main.Columns(4).HeaderText = "규격"
        Grid_Main.Columns(5).HeaderText = "단위"
        Grid_Main.Columns(6).HeaderText = "단가"
        Grid_Main.Columns(7).HeaderText = "수량"
        Grid_Main.Columns(8).HeaderText = "금액"
        Grid_Main.Columns(9).HeaderText = "세액"
        Grid_Main.Columns(10).HeaderText = "비고"

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
        Grid_Code_Display()
    End Sub
    Public Function Grid_Code_Display() As Boolean

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Code.RowCount = 1

        Select Case Search_Combo.Text
            Case "코드"
                StrSQL = "Select DL_Code, DL_Date, DL_Customer_Name FROM SC_DELIVER with(NOLOCK) Where DL_Code Like '" & Search_Text.Text & "%'  Order  By DL_Code DESC"
            Case "날짜"
                StrSQL = "Select DL_Code, DL_Date, DL_Customer_Name FROM SC_DELIVER with(NOLOCK) Where DL_Date Like '" & Search_Text.Text & "%' Or DL_Date Is Null  Order By DL_Date"
            Case "거래처"
                StrSQL = "Select DL_Code, DL_Date, DL_Customer_Name FROM SC_DELIVER with(NOLOCK) Where DL_Customer_Name Like '" & Search_Text.Text & "%' Or DL_Customer_Name Is Null  Order By DL_Customer_Name"
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


        StrSQL = "Select DL_Code FROM SC_DELIVER with(NOLOCK) Where DL_Date Like  '" & My_Date & "%' Order By DL_Code Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            JP_Code = JP_Code & "001"
        Else
            JP_Code = JP_Code & Format(Val(Mid(DBT.Rows(0).Item(0), 11, 3)) + 1, "00#")
        End If

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO SC_DELIVER (DL_Code, DL_Date, DL_Time, DL_Name, DL_Specifications, DL_Check)  Values ('CP" & JP_Code & "', '" & My_Date & "','" & My_Time & "','" & Frm_Main.Text_Man_Name.Text & "','기본','' )"
        Re_Count = DB_Execute()
        Grid_Code_Display()
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
            Panel_Main.Visible = True
            Grid_Info_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
            Label_Color(Com_Main, Color_Label, Di_Panel2)
            Grid_Main_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
            Panel_Main.Visible = True
            Panel_Excel.Visible = False
            Com_Excel_Print.Visible = False
        End If
    End Sub
    Private Sub Grid_Code_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Code.CellEnter
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
            Panel_Main.Visible = True
            Grid_Info_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
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
        StrSQL = "Select * FROM SC_DELIVER With(NOLOCK) Where DL_Code = '" & Code & "'"
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
        Grid_Display_Ch = True
    End Function
    Public Function Grid_Main_Display(PL_Code As String) As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display_Ch = False
        Grid_Main.RowCount = 0
        StrSQL = "Select DL_Sun, DL_Day, DL_PL_Code, DL_PL_Name, DL_Standard, DL_Unit, DL_Unit_Amount, DL_Total, DL_Gold, DL_Vat, DL_Bigo  FROM SC_DELIVER_LIST with(NOLOCK)  Where DL_Code = '" & PL_Code & "' Order By Convert(Decimal,DL_Sun)"
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
                    StrSQL = "Select CM_Code  FROM SI_CUSTOMER With(NOLOCK) Order By CM_Code"
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
                    StrSQL = "Select CR_Code  FROM SC_SALES With(NOLOCK) Where CR_Customer_Code = '" & Grid_Info.Item(1, 4).Value & "' And CR_Sel = '진행중' Order By CR_Customer_Code DESC"
                    Re_Count = DB_Select(DBT)
                    Grid_Info_Combo.Items.Clear()
                    If Re_Count <> 0 Then
                        For i = 0 To Re_Count - 1
                            Grid_Info_Combo.Items.Add(DBT.Rows(i)("CR_Code"))
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
                Case 9
                    Grid_Info_Combo.Items.Clear()
                    Grid_Info_Combo.Items.Add("기본")
                    Grid_Info_Combo.Items.Add("강원랜드")

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


    Private Sub Grid_Info_Combo_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Grid_Info_Combo.SelectionChangeCommitted
        Dim DBT As New DataTable
        Grid_Info.Item(Grid_Info_Col, Grid_Info_Row).Value = Grid_Info_Combo.SelectedItem.ToString()

        If Grid_Info_Row = 4 Then
            StrSQL = "Select CM_Code, CM_Name, CM_Man_Name, CM_Man_Tel  FROM SI_CUSTOMER With(NOLOCK) Where CM_Code = '" & Grid_Info_Combo.SelectedItem.ToString() & "'"
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
            StrSQL = "Select CM_Code, CM_Name, CM_Man_Name, CM_Man_Tel  FROM SI_CUSTOMER With(NOLOCK) Where CM_Name = '" & Grid_Info_Combo.SelectedItem.ToString() & "'"
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

    Private Sub Grid_Info_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Info.CellValueChanged
        If Grid_Display_Ch = True Then
            Grid_Info.CurrentRow.HeaderCell.Value = "*"
        End If
    End Sub

    Private Sub Save_Com_Click(sender As Object, e As EventArgs) Handles Save_Com.Click
        'Grid_Info 를 저장 한다.
        'DataBase에서 필드 이름 가지고 오기

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
            Table_Name = "Deliver"
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


        '수량 변경
        '수량이 변경 되었는지 확인 한다.
        Dim Val_Check As String
        Dim PP_G_Sun As String
        Dim PP_G_Amount As String

        Val_Check = ""
        StrSQL = "Select DL_Check FROM SC_DELIVER With(NOLOCK) Where DL_Code = '" & Grid_Info.Item(1, 0).Value & "'"
        Re_Count = DB_Select(DBT)
        Grid_Info_Combo.Items.Clear()
        If Re_Count <> 0 Then
            If DBT.Rows(0)("DL_Check") = "처리" Then
                Val_Check = "처리"
                MsgBox("이미 처리 되었습니다. 삭제 후 다시 처리 하세요")
                Exit Sub
            End If
        End If
        'If Val_Check = "처리" Then
        '    '기존 Data를 삭제 한다.
        '    'Grid의 수량을 가지고 온다.

        '    For i = 0 To Grid_Main.RowCount - 1
        '        PP_G_Sun = "0"
        '        PP_G_Amount = "0"
        '        StrSQL = "Select PP_Sun, PP_Amount FROM SI_Process_List with(NOLOCK) Where PP_Code = '" & Grid_Main.Item(2, i).Value & "' Order By Convert(Decimal,PP_Sun) Desc"
        '        Re_Count = DB_Select(DBT)
        '        If Re_Count > 0 Then
        '            PP_G_Sun = DBT.Rows(0)("PP_Sun")
        '            PP_G_Amount = DBT.Rows(0)("PP_Amount")
        '        End If
        '        If PP_G_Sun = "0" Then
        '        Else
        '            '변경
        '            PP_G_Amount = Val(PP_G_Amount) + Val(Grid_Main.Item(7, i).Value)
        '            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        '            StrSQL = StrSQL & "UPDATE SI_PROCESS_LIST Set PP_Amount = '" & PP_G_Amount & "' Where PP_Code = '" & Grid_Main.Item(2, i).Value & "' And PP_Sun = '" & PP_G_Sun & "'"
        '            Re_Count = DB_Execute()
        '        End If
        '    Next i

        'End If

        '수량을 변경한다.
        For i = 0 To Grid_Main.RowCount - 1
            PP_G_Sun = "0"
            PP_G_Amount = "0"
            StrSQL = "Select PP_Sun, PP_Amount FROM SI_Process_List with(NOLOCK) Where PP_Code = '" & Grid_Main.Item(2, i).Value & "' Order By Convert(Decimal,PP_Sun) Desc"
            Re_Count = DB_Select(DBT)
            If Re_Count > 0 Then
                PP_G_Sun = DBT.Rows(0)("PP_Sun")
                PP_G_Amount = DBT.Rows(0)("PP_Amount")
            End If
            If PP_G_Sun = "0" Then
            Else
                '변경
                PP_G_Amount = Val(PP_G_Amount) - Val(Grid_Main.Item(7, i).Value)
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SI_PROCESS_LIST Set PP_Amount = '" & PP_G_Amount & "' Where PP_Code = '" & Grid_Main.Item(2, i).Value & "' And PP_Sun = '" & PP_G_Sun & "'"
                Re_Count = DB_Execute()
            End If
        Next i

        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "UPDATE SC_DELIVER Set DL_Check = '처리' Where DL_Code = '" & Grid_Info.Item(1, 0).Value & "'"
        Re_Count = DB_Execute()

        Grid_Info.Item(1, 10).Value = "처리"








    End Sub

    Private Sub Grid_Info_Combo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Grid_Info_Combo.SelectedIndexChanged

    End Sub

    Private Sub Com_Main_Click(sender As Object, e As EventArgs) Handles Com_Main.Click
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        Label_Color(sender, Color_Label, Di_Panel2)
        '판넬 보이기
        Panel_Main.Visible = True
        Panel_Excel.Visible = False
        Com_Excel_Print.Visible = False


    End Sub

    Private Sub Insert__Main_Click(sender As Object, e As EventArgs)
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        If Grid_Info.Item(1, 10).Value = "처리" Then
            MsgBox("수량 처리가 되어서 수정 할 수 없습니다.")
            Exit Sub
        End If

        Grid_Display_Ch = False

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


        'Dim DBT As New DataTable
        Dim Db_Sun As Long
        StrSQL = "Select DL_Sun FROM SC_DELIVER_LIST with(NOLOCK) Where DL_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' Order By Convert(Decimal,DL_Sun) Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            Db_Sun = 1
        Else
            Db_Sun = DBT.Rows(0).Item(0) + 1
        End If


        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO SC_DELIVER_LIST (DL_Code, DL_Sun, DL_Day)  Values ('" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "', '" & Db_Sun & "','" & My_Date & "' )"
        Re_Count = DB_Execute()
        Grid_Main_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True
    End Sub

    Private Sub Save_Main_Click(sender As Object, e As EventArgs)
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        If Grid_Info.Item(1, 10).Value = "처리" Then
            MsgBox("수량 처리가 되어서 수정 할 수 없습니다.")
            Exit Sub
        End If

        Dim i As Integer

        Grid_Display_Ch = False
        For i = 0 To Grid_Main.Rows.Count - 1
            If Grid_Main.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SC_DELIVER_LIST Set DL_Day = '" & Grid_Main.Item(1, i).Value & "',DL_PL_Code = '" & Grid_Main.Item(2, i).Value & "', DL_PL_Name = '" & Grid_Main.Item(3, i).Value & "', DL_Standard = '" & Grid_Main.Item(4, i).Value & "', DL_Unit = '" & Grid_Main.Item(5, i).Value & "', DL_Unit_Amount = '" & Grid_Main.Item(6, i).Value & "', DL_Total = '" & Grid_Main.Item(7, i).Value & "', DL_Gold = '" & Grid_Main.Item(8, i).Value & "', DL_Vat = '" & Grid_Main.Item(9, i).Value & "', DL_Bigo = '" & Grid_Main.Item(10, i).Value & "' where DL_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And DL_Sun = '" & Grid_Main.Item(0, i).Value & "'"
                Re_Count = DB_Execute()
                Grid_Main.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display_Ch = True

    End Sub
    Private Sub Grid_Main_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Main.CellValueChanged
        If Grid_Display_Ch = True Then
            Grid_Main.CurrentRow.HeaderCell.Value = "*"
            Grid_Main.Item(8, Grid_Main.CurrentCell.RowIndex).Value = Val(Grid_Main.Item(6, Grid_Main.CurrentCell.RowIndex).Value.ToString) * Val(Grid_Main.Item(7, Grid_Main.CurrentCell.RowIndex).Value.ToString)
            Grid_Main.Item(9, Grid_Main.CurrentCell.RowIndex).Value = Int(Val(Grid_Main.Item(8, Grid_Main.CurrentCell.RowIndex).Value) * 0.1)
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
            Case 2
                StrSQL = "Select SI_Product.PL_Code, SI_Product.PL_Name  FROM SI_PRODUCT with(NOLOCK), SC_SALES_LIST with(NOLOCK) Where SC_SALES_LIST.CL_Code = '" & Grid_Info.Item(1, 8).Value & "' And  SC_SALES_LIST.CL_PL_Code = SI_Product.PL_Code  Order By PL_Code"
                Re_Count = DB_Select(DBT)
                Grid_Main_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        Grid_Main_Combo.Items.Add(DBT.Rows(i).Item(0))
                    Next i
                End If
                Grid_Main_Combo.Top = Grid_Main.Top + Grid_Main.ColumnHeadersHeight + (Grid_Main_Row) * 24
                Grid_Main_Combo.Left = Grid_Main.Left + Grid_Main.RowHeadersWidth + Grid_Main.Columns(0).Width + Grid_Main.Columns(1).Width
                Grid_Main_Combo.Width = Grid_Main.Columns(Grid_Main_Col).Width
                Grid_Main_Combo.Text = Grid_Main.CurrentCell.Value.ToString
                Grid_Main_Combo.BackColor = Grid_Main.RowsDefaultCellStyle.BackColor
                Grid_Main_Combo.Visible = True
                Grid_Main_Combo.Focus.ToString()
            Case 3
                StrSQL = "Select SI_Product.PL_Code, SI_Product.PL_Name  FROM SI_PRODUCT with(NOLOCK), SC_SALES_LIST with(NOLOCK) Where SC_SALES_LIST.CL_Code = '" & Grid_Info.Item(1, 8).Value & "' And  SC_SALES_LIST.CL_PL_Code = SI_Product.PL_Code  Order By PL_Code"
                Re_Count = DB_Select(DBT)
                Grid_Main_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        Grid_Main_Combo.Items.Add(DBT.Rows(i).Item(1))
                    Next i
                End If
                Grid_Main_Combo.Top = Grid_Main.Top + Grid_Main.ColumnHeadersHeight + (Grid_Main_Row) * 24
                Grid_Main_Combo.Left = Grid_Main.Left + Grid_Main.RowHeadersWidth + Grid_Main.Columns(0).Width + Grid_Main.Columns(1).Width + Grid_Main.Columns(2).Width
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
            Case 2
                StrSQL = "Select PL_Code, PL_Name, PL_Standard ,PL_Unit, PL_Unit_Gold FROM SI_PRODUCT with(NOLOCK) Where PL_Code = '" & Grid_Main_Combo.SelectedItem.ToString() & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Main.Item(3, Grid_Main_Row).Value = DBT.Rows(0).Item(1)
                    Grid_Main.Item(4, Grid_Main_Row).Value = DBT.Rows(0).Item(2)
                    Grid_Main.Item(5, Grid_Main_Row).Value = DBT.Rows(0).Item(3)
                    Grid_Main.Item(6, Grid_Main_Row).Value = DBT.Rows(0).Item(4)

                End If
            Case 3
                StrSQL = "Select PL_Code, PL_Name, PL_Standard ,PL_Unit, PL_Unit_Gold  FROM SI_PRODUCT with(NOLOCK) Where PL_Name= '" & Replace(Grid_Main_Combo.SelectedItem.ToString(), "'", "''") & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Main.Item(2, Grid_Main_Row).Value = DBT.Rows(0).Item(0)
                    Grid_Main.Item(4, Grid_Main_Row).Value = DBT.Rows(0).Item(2)
                    Grid_Main.Item(5, Grid_Main_Row).Value = DBT.Rows(0).Item(3)
                    Grid_Main.Item(6, Grid_Main_Row).Value = DBT.Rows(0).Item(4)
                End If
        End Select
        Grid_Main_Combo.Visible = False
        'Grid_Main.Columns(1).HeaderText = "제품코드"
        'Grid_Main.Columns(2).HeaderText = "제품명"
        'Grid_Main.Columns(3).HeaderText = "규격"
        'Grid_Main.Columns(4).HeaderText = "단위"
        'Grid_Main.Columns(5).HeaderText = "단가"
    End Sub

    Private Sub Delete__Main_Click(sender As Object, e As EventArgs)
        '해당 칼럼 삭제
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        If Grid_Info.Item(1, 10).Value = "처리" Then
            MsgBox("수량 처리가 되어서 수정 할 수 없습니다.")
            Exit Sub
        End If

        Grid_Display_Ch = False
        Dim i As Integer
        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SC_DELIVER_LIST Where DL_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And  DL_Sun = '" & Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        For i = Grid_Main.CurrentCell.RowIndex + 1 To Grid_Main.RowCount - 1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "Update  Deliver_List Set DL_Sun = '" & i & "' Where DL_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And  DL_Sun = '" & i + 1 & "'"
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
        Dim j As Long

        If Setup_Data(9) = "정우프린텍" Then
            xlapp = New Excel.Application
            If Grid_Info.Item(1, 9).Value = "기본" Then
                xlbook = xlapp.Workbooks.Open(Application.StartupPath & "\excel\정우프린텍 거래명세표.xlsx")
            End If

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

            If Grid_Info.Item(1, 9).Value = "기본" Then
                '거래 일자
                xlsheet.Cells(15, 1).Value = Grid_Info.Item(1, 2).Value
                '상호
                xlsheet.Cells(9, 9).Value = Grid_Info.Item(1, 5).Value

                '주소
                StrSQL = "Select CM_Add1, CM_Add2  FROM SI_CUSTOMER with(NOLOCK) Where CM_Code= '" & Grid_Info.Item(1, 4).Value & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    xlsheet.Cells(11, 9).Value = DBT.Rows(0).Item(0)
                    xlsheet.Cells(11, 9).Value = DBT.Rows(0).Item(1)
                End If
                '전화번호
                xlsheet.Cells(11, 9).Value = Grid_Info.Item(1, 7).Value
                j = 0
                For i = 0 To Grid_Main.RowCount - 1
                    'xlsheet.Cells(15 + i, 3).Value = Mid(Grid_Main.Item(1, i).Value.ToString, 6, 2)
                    'xlsheet.Cells(15 + i, 4).Value = Mid(Grid_Main.Item(1, i).Value.ToString, 9, 2)
                    'xlsheet.Cells(15 + i, 5).Value = Grid_Main.Item(3, i).Value.ToString
                    '규격
                    xlsheet.Cells(15 + i, 2).Value = Grid_Main.Item(4, i).Value.ToString
                    '수량
                    xlsheet.Cells(15 + i, 7).Value = Grid_Main.Item(7, i).Value.ToString
                    '단가
                    xlsheet.Cells(15 + i, 8).Value = Grid_Main.Item(6, i).Value.ToString
                    '금액
                    xlsheet.Cells(15 + i, 10).Value = Grid_Main.Item(8, i).Value.ToString
                    '세액
                    xlsheet.Cells(15 + i, 11).Value = Grid_Main.Item(9, i).Value.ToString

                Next i
            End If

            If Grid_Info.Item(1, 9).Value = "강원랜드" Then
                '거래 일자
                xlsheet.Cells(5, 2).Value = "납품일자  : " & Grid_Info.Item(1, 2).Value
                '제품명
                xlsheet.Cells(10, 2).Value = "제품명 : " & Grid_Main.Item(3, 0).Value.ToString
                j = 0
                For i = 0 To Grid_Main.RowCount - 1
                    xlsheet.Cells(14 + i, 3).Value = Grid_Main.Item(3, i).Value.ToString
                    '규격
                    xlsheet.Cells(14 + i, 4).Value = Grid_Main.Item(4, i).Value.ToString
                    xlsheet.Cells(14 + i, 5).Value = Grid_Main.Item(5, i).Value.ToString
                    xlsheet.Cells(14 + i, 6).Value = Grid_Main.Item(7, i).Value.ToString
                    xlsheet.Cells(14 + i, 7).Value = Grid_Main.Item(6, i).Value.ToString
                Next i
                '이 하 여 백
                xlsheet.Cells(14 + i, 4).Value = "이 하 여 백"
            End If
        Else

            'If excel_check Then
            'Else
            xlapp = New Excel.Application
            If Grid_Info.Item(1, 9).Value = "기본" Then
                xlbook = xlapp.Workbooks.Open(Application.StartupPath & "\excel\납품.xls")
            End If
            If Grid_Info.Item(1, 9).Value = "강원랜드" Then
                xlbook = xlapp.Workbooks.Open(Application.StartupPath & "\excel\납품강원랜드.xls")
            End If

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

            If Grid_Info.Item(1, 9).Value = "기본" Then
                '거래 일자
                xlsheet.Cells(5, 3).Value = Grid_Info.Item(1, 2).Value
                '상호
                xlsheet.Cells(6, 7).Value = Grid_Info.Item(1, 5).Value

                '주소
                StrSQL = "Select CM_Add1, CM_Add2  FROM SI_CUSTOMER with(NOLOCK) Where CM_Code= '" & Grid_Info.Item(1, 4).Value & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    xlsheet.Cells(8, 7).Value = DBT.Rows(0).Item(0)
                    xlsheet.Cells(9, 7).Value = DBT.Rows(0).Item(1)
                End If
                '전화번호
                xlsheet.Cells(10, 7).Value = Grid_Info.Item(1, 7).Value
                j = 0
                For i = 0 To Grid_Main.RowCount - 1
                    xlsheet.Cells(15 + i, 3).Value = Mid(Grid_Main.Item(1, i).Value.ToString, 6, 2)
                    xlsheet.Cells(15 + i, 4).Value = Mid(Grid_Main.Item(1, i).Value.ToString, 9, 2)
                    xlsheet.Cells(15 + i, 5).Value = Grid_Main.Item(3, i).Value.ToString
                    '규격
                    xlsheet.Cells(15 + i, 11).Value = Grid_Main.Item(4, i).Value.ToString
                    xlsheet.Cells(15 + i, 16).Value = Grid_Main.Item(7, i).Value.ToString
                    xlsheet.Cells(15 + i, 18).Value = Grid_Main.Item(6, i).Value.ToString
                    xlsheet.Cells(15 + i, 23).Value = Grid_Main.Item(8, i).Value.ToString
                Next i
            End If

            If Grid_Info.Item(1, 9).Value = "강원랜드" Then
                '거래 일자
                xlsheet.Cells(5, 2).Value = "납품일자  : " & Grid_Info.Item(1, 2).Value
                '제품명
                xlsheet.Cells(10, 2).Value = "제품명 : " & Grid_Main.Item(3, 0).Value.ToString
                j = 0
                For i = 0 To Grid_Main.RowCount - 1
                    xlsheet.Cells(14 + i, 3).Value = Grid_Main.Item(3, i).Value.ToString
                    '규격
                    xlsheet.Cells(14 + i, 4).Value = Grid_Main.Item(4, i).Value.ToString
                    xlsheet.Cells(14 + i, 5).Value = Grid_Main.Item(5, i).Value.ToString
                    xlsheet.Cells(14 + i, 6).Value = Grid_Main.Item(7, i).Value.ToString
                    xlsheet.Cells(14 + i, 7).Value = Grid_Main.Item(6, i).Value.ToString
                Next i
                '이 하 여 백
                xlsheet.Cells(14 + i, 4).Value = "이 하 여 백"
            End If

        End If



        xlbook.SaveAs(Application.StartupPath & "\Excel\납품\" & Grid_Info.Item(1, 0).Value & Grid_Info.Item(1, 5).Value & ".xls")
        'Label4.Text = Format(Now, "yyyy년 MM월 dd일 HH:mm:ss")

        'xlsheet = Nothing
        'xlbook = Nothing
        'xlapp = Nothing
        'releaseObject(xlsheet)
        'releaseObject(xlbook)
        'releaseObject(xlapp)
        Cursor = Cursors.Default
        Excel_check = True

    End Sub

    Private Sub Com_Excel_Print_Click(sender As Object, e As EventArgs) Handles Com_Excel_Print.Click
        If Excel_check = True Then
            xlsheet.PrintOut(From:=1, To:=1, Copies:=1, Collate:=True)
        End If


    End Sub

    Private Sub Delete_Com_Click(sender As Object, e As EventArgs) Handles Delete_Com.Click

        '해당 칼럼 삭제
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        If MsgBox("납품전표  '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "'를 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "납품전표 삭제") = vbNo Then
            Exit Sub
        End If


        '수량 변경
        '수량이 변경 되었는지 확인 한다.
        Dim Val_Check As String
        Dim PP_G_Sun As String
        Dim PP_G_Amount As String
        Dim DBT As DataTable = Nothing

        Val_Check = ""
        StrSQL = "Select DL_Check FROM SC_DELIVER With(NOLOCK) Where DL_Code = '" & Grid_Info.Item(1, 0).Value & "'"
        Re_Count = DB_Select(DBT)
        Grid_Info_Combo.Items.Clear()
        If Re_Count <> 0 Then
            If DBT.Rows(0)("DL_Check") = "처리" Then
                Val_Check = "처리"
            End If
        End If
        If Val_Check = "처리" Then
            '기존 Data를 삭제 한다.
            'Grid의 수량을 가지고 온다.

            For i = 0 To Grid_Main.RowCount - 1
                PP_G_Sun = "0"
                PP_G_Amount = "0"
                StrSQL = "Select PP_Sun, PP_Amount FROM SI_Process_List with(NOLOCK) Where PP_Code = '" & Grid_Main.Item(2, i).Value & "' Order By Convert(Decimal,PP_Sun) Desc"
                Re_Count = DB_Select(DBT)
                If Re_Count > 0 Then
                    PP_G_Sun = DBT.Rows(0)("PP_Sun")
                    PP_G_Amount = DBT.Rows(0)("PP_Amount")
                End If
                If PP_G_Sun = "0" Then
                Else
                    '변경
                    PP_G_Amount = Val(PP_G_Amount) + Val(Grid_Main.Item(7, i).Value)
                    StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                    StrSQL = StrSQL & "UPDATE SI_PROCESS_LIST Set PP_Amount = '" & PP_G_Amount & "' Where PP_Code = '" & Grid_Main.Item(2, i).Value & "' And PP_Sun = '" & PP_G_Sun & "'"
                    Re_Count = DB_Execute()
                End If
            Next i

        End If

        Grid_Display_Ch = False
        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SC_DELIVER Where DL_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' "
        Re_Count = DB_Execute()

        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SC_DELIVER Where DL_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' "
        Re_Count = DB_Execute()

        Grid_Code_Display()
        Grid_Info_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
        MsgBox("삭제되었습니다")
        Grid_Display_Ch = True

    End Sub

    Private Sub Grid_Main_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Main.CellContentClick

    End Sub

    Private Sub Grid_Info_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Info.CellContentClick

    End Sub

    Private Sub Grid_Code_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Code.CellContentClick

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Dim P_Code As String
        Dim P_Su As String
        Dim i As Integer
        Dim DBT As New DataTable



        Dim JP_Code As Long
        Dim My_Date As String
        Dim My_Time As String



        If Len(TextBox1.Text) < 1 Then
        Else
            If Mid(TextBox1.Text, Len(TextBox1.Text), 1) = "K" Or Mid(TextBox1.Text, Len(TextBox1.Text), 1) = "k" Then
                '수량 과 코드 분리
                For i = Len(TextBox1.Text) To 1 Step -1
                    If Mid(TextBox1.Text, i, 1) = "-" Then
                        P_Code = Mid(TextBox1.Text, 1, i - 1)
                        P_Su = Mid(TextBox1.Text, i + 1, 5)
                        P_Su = Mid(P_Su, 1, Len(P_Su) - 1)
                        Exit For
                    End If
                Next i
                If Len(P_Code) < 4 Or Len(P_Su) < 1 Then
                    TextBox1.Text = ""
                    Exit Sub
                End If
                '코드 추가 한다.
                '코드 정보를 가지고 온다.
                '기존에 코드가 있는가?
                For i = 0 To Grid_Main.Rows.Count - 1
                    If Grid_Main.Item(2, i).Value = P_Code Then
                        Grid_Main.Item(7, i).Value = Val(Grid_Main.Item(7, i).Value) + Val(P_Su)
                        TextBox1.Text = ""
                        Exit Sub
                    End If
                Next i


                '추가 한다.
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

                Dim Db_Sun As Long
                StrSQL = "Select DL_Sun FROM SC_DELIVER_LIST with(NOLOCK) Where DL_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' Order By Convert(Decimal,DL_Sun) Desc "
                Re_Count = DB_Select(DBT)
                If Re_Count = 0 Then
                    Db_Sun = 1
                Else
                    Db_Sun = DBT.Rows(0).Item(0) + 1
                End If


                '추가한다
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO SC_DELIVER_LIST (DL_Code, DL_Sun, DL_Day)  Values ('" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "', '" & Db_Sun & "','" & My_Date & "' )"
                Re_Count = DB_Execute()
                If Len(Grid_Main.Item(0, Grid_Main.RowCount - 1).Value) < 1 Then
                Else
                    Grid_Main.RowCount = Grid_Main.RowCount + 1
                End If
                'DL_Sun, DL_Day, DL_PL_Code, DL_PL_Name, DL_Standard, DL_Unit, DL_Unit_Amount, DL_Total, DL_Gold, DL_Vat, DL_Bigo
                Grid_Main.Item(0, Grid_Main.RowCount - 1).Value = Grid_Main.RowCount
                Grid_Main.Item(1, Grid_Main.RowCount - 1).Value = My_Date

                StrSQL = "Select PL_Code, PL_Name, PL_Standard ,PL_Unit, PL_Unit_Gold FROM SI_PRODUCT with(NOLOCK) Where PL_Code = '" & P_Code & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Main.Item(2, Grid_Main.RowCount - 1).Value = P_Code
                    Grid_Main.Item(3, Grid_Main.RowCount - 1).Value = DBT.Rows(0).Item(1)
                    Grid_Main.Item(4, Grid_Main.RowCount - 1).Value = DBT.Rows(0).Item(2)
                    Grid_Main.Item(5, Grid_Main.RowCount - 1).Value = DBT.Rows(0).Item(3)
                    Grid_Main.Item(6, Grid_Main.RowCount - 1).Value = DBT.Rows(0).Item(4)
                    Grid_Main.Item(7, Grid_Main.RowCount - 1).Value = Val(P_Su)
                    TextBox1.Text = ""
                End If


                'Search_Text.Text = "KK"
            End If
        End If


    End Sub
End Class