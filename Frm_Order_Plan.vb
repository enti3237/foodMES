﻿Imports Excel = Microsoft.Office.Interop.Excel
Public Class Frm_Order_Plan
    Dim Grid_Display_Ch As Boolean
    Dim Grid_Info_Row As Integer
    Dim Grid_Info_Col As Integer

    Dim Grid_Str(1000, 1000) As String
    Dim Je_ToT(1000, 1000) As String

    Dim Grid_Str_CM(1000, 1000) As String


    Dim Grid_Row As Integer
    Dim Grid_Col As Integer
    Dim Grid_Col2 As Integer

    Dim Product_Count_Cnt As Long

    Dim xlapp As Excel.Application
    Dim xlbook As Excel.Workbook
    Dim xlsheet As Excel.Worksheet
    Dim Excel_check As Boolean

    Private Sub Frm_Order_Plan_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()

        Grid_Code_Setup()
        Grid_Info_Setup()






        Search_Combo.Items.Add("코드")
        Search_Combo.Items.Add("날짜")
        Search_Combo.Items.Add("담당자")
        Search_Combo.Text = "코드"


        Panel_Main.Top = 136
        Panel_Main.Left = 389
        Panel_Excel.Top = 136
        Panel_Excel.Left = 389

        Panel_Main.Visible = True
        Panel_Excel.Visible = False

        Com_Excel_Print.Visible = False
        Button3.Visible = False
        Com_Excel.Visible = False
        Search_Com.PerformClick()
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
        Grid_Code.Columns(2).HeaderText = "담당자"

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
        Grid_Info.RowCount = 10





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
        Grid_Info.Item(0, 4).Value = "산출일자(기준)"
        Grid_Info.Item(0, 5).Value = "산출일자(종료)"
        Grid_Info.Item(0, 6).Value = "발주서(시작)"
        Grid_Info.Item(0, 7).Value = "발주서(끝)"
        Grid_Info.Item(0, 8).Value = "구분"
        Grid_Info.Item(0, 9).Value = "비고"


        Grid_Info.Rows(0).ReadOnly = True
        Grid_Info.Rows(1).ReadOnly = True
        Grid_Info.Rows(2).ReadOnly = True
        Grid_Info.Rows(3).ReadOnly = True
        Grid_Info.Rows(6).ReadOnly = True
        Grid_Info.Rows(7).ReadOnly = True
        Grid_Info.Rows(8).ReadOnly = True

        Grid_Info_Setup = True
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


        StrSQL = "Select CL_Code From MC_STOCK with(NOLOCK) Where CL_Date Like  '" & My_Date & "%' Order By CL_Code Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            JP_Code = JP_Code & "001"
        Else
            JP_Code = JP_Code & Format(Val(Mid(DBT.Rows(0).Item(0), 11, 3)) + 1, "00#")
        End If

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert into MC_STOCK (CL_Code, CL_Date, CL_Time, CL_Name, CL_Day, CL_En_Day, CL_Sel)  Values ('CL" & JP_Code & "', '" & My_Date & "','" & My_Time & "','" & Frm_Main.Text_Man_Name.Text & "', '" & My_Date & "', '" & My_Date & "','')"
        Re_Count = DB_Execute()
        Grid_Code_Display()
        Button3.Visible = False
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
    Public Function Grid_Code_Display() As Boolean

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Code.RowCount = 1

        Select Case Search_Combo.Text
            Case "코드"
                StrSQL = "Select CL_Code, CL_Date, CL_Name From MC_STOCK with(NOLOCK) Where CL_Code Like '%" & Search_Text.Text & "%'  Order  By CL_Code DESC"
            Case "날짜"
                StrSQL = "Select CL_Code, CL_Date, CL_Name From MC_STOCK with(NOLOCK) Where CL_Date Like '%" & Search_Text.Text & "%' Or CL_Date Is Null  Order By CL_Date"
            Case "담당자"
                StrSQL = "Select CL_Code, CL_Date, CL_Name From MC_STOCK with(NOLOCK) Where CL_Name Like '%" & Search_Text.Text & "%' Or CL_Name Is Null  Order By CL_Name"
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
    Private Sub Grid_Code_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Code.CellEnter
        Button3.Visible = False
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
            Panel_Main.Visible = True
            Grid_Info_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
            Label_Color(Com_Main, Color_Label, Di_Panel2)
            Grid_Main_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
            Panel_Main.Visible = True
            Panel_Excel.Visible = False
            Com_Excel_Print.Visible = False
            Com_Excel.Visible = True
        End If

    End Sub
    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        Grid_Code_Display()
    End Sub
    Public Function Grid_Info_Display(Code As String) As Boolean
        Grid_Display_Ch = False
        Grid_Info_Display = True
        Dim DBT As New DataTable
        Dim i As Integer
        StrSQL = "Select * From MC_STOCK With(NOLOCK) Where CL_Code = '" & Code & "'"
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
        Dim NullStr(1000) As String
        Dim D_Col As Integer

        Grid_Display_Ch = False

        Grid_Main.RowCount = 0
        StrSQL = "Select CL_Data From MC_STOCK_List with(NOLOCK)  Where CL_Code = '" & PL_Code & "' Order By Convert(Decimal,CL_Sun)"
        ' StrSQL = "Select CL_Data From MC_STOCK_List with(NOLOCK), SI_PRODUCT_CUSTOMER with(NOLOCK)  Where CL_Code = '" & PL_Code & "' Order By Convert(Decimal,CL_Sun)"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            D_Col = Data_Col_Count(DBT.Rows(0)("CL_Data").ToString, NullStr)
            Grid_Main_Setup(Re_Count - 1, D_Col)
            For j = 0 To Grid_Main.ColumnCount - 1
                Grid_Main.Columns(j).HeaderText = NullStr(j + 1)
            Next j
            For i = 0 To Re_Count - 2
                D_Col = Data_Col_Count(DBT.Rows(i + 1)("CL_Data").ToString, NullStr)
                For j = 0 To Grid_Main.ColumnCount - 1
                    Grid_Main.Item(j, i).Value = NullStr(j + 1)
                    If Mid(NullStr(j + 1), 1, 1) = "-" Then
                        Grid_Main.Rows(i).Cells(j).Style.ForeColor = Color.Red
                    Else
                        Grid_Main.Rows(i).Cells(j).Style.ForeColor = Color.Black
                    End If
                Next j
            Next i
        Else
            Grid_Main.RowCount = 0
            'For j = 0 To Grid_Main.ColumnCount - 1
            '    Grid_Main.Item(j, 0).Value = ""
            'Next j
        End If


        Grid_Display_Ch = True
        Grid_Main_Display = True

    End Function
    Public Function Grid_Main_Setup(RowCount As Long, ColumnCount As Long) As Boolean

        Grid_Color(Grid_Main)
        Grid_Main.AllowUserToAddRows = False
        Grid_Main.RowTemplate.Height = 24
        Grid_Main.ColumnCount = ColumnCount
        Grid_Main.RowCount = RowCount

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        'Grid_Main.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Grid_Main.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Grid_Main.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Grid_Main.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'Grid_Main.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        'Grid_Main.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.RowHeadersWidth = 10


        Grid_Main.Columns(0).HeaderText = "제품코드"
        Grid_Main.Columns(1).HeaderText = "제품명"
        Grid_Main.Columns(2).HeaderText = "규격"
        Grid_Main.Columns(3).HeaderText = "단위"
        Grid_Main.Columns(4).HeaderText = "단가"
        Grid_Main.Columns(5).HeaderText = "매입처"
        Grid_Main.Columns(6).HeaderText = "매입처명"

        Dim i As Integer

        'Grid_Main.Columns(0).Width = 40

        For i = 0 To Grid_Main.ColumnCount - 1
            Grid_Main.Columns(i).Width = 80
        Next i

        For i = 7 To Grid_Main.ColumnCount - 1
            Grid_Main.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Next i

        Grid_Main_Setup = True
    End Function



    Private Sub Save_Com_Click(sender As Object, e As EventArgs) Handles Save_Com.Click
        'Grid_Info 를 저장 한다.
        'DataBase에서 필드 이름 가지고 오기
        If Grid_Info.Item(1, 8).Value = "처리" Then
            MsgBox("작업지시서가 처리되어 저장 할 수 없습니다.")
            Exit Sub
        End If
        Dim Table_Name As String
        Dim j As Integer
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
            Table_Name = "MC_STOCK"
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

        'If Button3.Visible = True Then
        'Else
        '    Exit Sub
        'End If


        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Delete From MC_STOCK_LIst Where CL_Code = '" & Grid_Info.Item(1, 0).Value & "'"
        Re_Count = DB_Execute()

        Dim DD As String

        DD = ""
        For j = 0 To Grid_Main.ColumnCount - 1
            If j = Grid_Main.ColumnCount - 1 Then
                DD = DD & Grid_Main.Columns(j).HeaderText
            Else
                DD = DD & Grid_Main.Columns(j).HeaderText & "|"
            End If
        Next j
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert Into MC_STOCK_LIst Values ('" & Grid_Info.Item(1, 0).Value & "','1','" & DD & "')"
        Re_Count = DB_Execute()

        For i = 0 To Grid_Main.Rows.Count - 1
            'Grid_Code.Columns(0).HeaderText
            DD = ""

            For j = 0 To Grid_Main.ColumnCount - 1
                If j = Grid_Main.ColumnCount - 1 Then
                    DD = DD & Grid_Main.Item(j, i).Value
                Else
                    DD = DD & Grid_Main.Item(j, i).Value & "|"

                End If
            Next j
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "Insert Into MC_STOCK_LIst Values ('" & Grid_Info.Item(1, 0).Value & "','" & i + 2 & "','" & DD & "')"
            Re_Count = DB_Execute()
        Next i
    End Sub
    Private Sub Grid_Info_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Info.CellValueChanged
        If Grid_Display_Ch = True Then
            Grid_Info.CurrentRow.HeaderCell.Value = "*"
        End If

    End Sub

    Private Sub Calculation_Com_Click(sender As Object, e As EventArgs) Handles Calculation_Com.Click

        If Grid_Info.Item(1, 8).Value = "처리" Then
            MsgBox("발주서가 처리되어 산출 할 수 없습니다.")
            Exit Sub
        End If
        '수주 리스트에서 제품 코드를 가지고온다.
        '1달을 기준으로 한다.
        '일자별 제품 코드를 가지고 온다/
        Dim ST_Day As String
        Dim EN_Day As String
        Dim Day_Cnt(33) As String
        Dim Je_Sung(1000, 100) As String
        Dim Je_Cnt As Long
        Dim Je_Cnt2 As Long
        Dim i As Integer
        Dim j As Integer
        Dim DBT As New DataTable

        If Len(Grid_Info.Item(1, 0).Value) > 0 Then
        Else
            Exit Sub
        End If

        Button3.Visible = True

        ST_Day = Grid_Info.Item(1, 4).Value
        Day_Cnt(1) = ST_Day
        For i = 1 To 29
            Day_Cnt(1 + i) = DateAdd(DateInterval.Day, i, CDate(ST_Day))
        Next
        '32에다 이후 모든 수량을 저장한다.
        '모든 제품을 가지고 온다.

        Dim Null_Je_Code As String
        Dim Day1 As String
        Dim Day2 As String
        Dim Day3 As String

        Dim Day_Su1 As Long
        Dim Day_Su2 As Long
        Dim Day_Su3 As Long

        Je_Cnt = 0

        StrSQL = "Select * FROM SC_SALES With(NOLOCK), SC_SALES_LIST With(NOLOCK) Where CR_Sel = '진행중' And CR_Code = CL_Code"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            '3Day 를 처리 한다.
            '제품 코드가 등록 되어 있는가?
            For i = 0 To Re_Count - 1
                Null_Je_Code = DBT.Rows(i)("CL_PL_Code").ToString
                Day1 = DBT.Rows(i)("CR_Day1").ToString
                Day2 = DBT.Rows(i)("CR_Day2").ToString
                Day3 = DBT.Rows(i)("CR_Day3").ToString

                Day_Su1 = Val(DBT.Rows(i)("CL_Day01").ToString)
                Day_Su2 = Val(DBT.Rows(i)("CL_Day02").ToString)
                Day_Su3 = Val(DBT.Rows(i)("CL_Day03").ToString)




                Je_Cnt2 = 0
                For j = 1 To Je_Cnt
                    If Je_Sung(j, 0) = Null_Je_Code Then
                        Je_Cnt2 = j
                        Exit For
                    End If
                Next
                If Je_Cnt2 = 0 Then
                    Je_Cnt = Je_Cnt + 1
                    Je_Cnt2 = Je_Cnt
                    Je_Sung(Je_Cnt2, 0) = Null_Je_Code
                End If
                '일자별 수량을 처리 한다.
                For j = 1 To 30
                    If Day1 <= Day_Cnt(j) Then
                        Je_Sung(Je_Cnt2, j) = Val(Je_Sung(Je_Cnt2, j)) + Day_Su1
                    End If
                    If Day2 <= Day_Cnt(j) Then
                        Je_Sung(Je_Cnt2, j) = Val(Je_Sung(Je_Cnt2, j)) + Day_Su2
                    End If
                    If Day3 <= Day_Cnt(j) Then
                        Je_Sung(Je_Cnt2, j) = Val(Je_Sung(Je_Cnt2, j)) + Day_Su3
                    End If
                Next j
                '총수주량
                Je_Sung(Je_Cnt2, 32) = Val(Je_Sung(Je_Cnt2, 32)) + Day_Su1 + Day_Su2 + Day_Su3
                Je_Sung(Je_Cnt2, 31) = Je_Sung(Je_Cnt2, 32)

            Next i
        End If
        '납품량을 계산한다.
        '납품량/33
        For i = 1 To Je_Cnt
            StrSQL = "SELECT Sum(Convert(Decimal,DL_Total)) AS DL_To FROM SC_DELIVER_LIST With(NOLOCK), SC_DELIVER With(NOLOCK), SC_SALES WITH(NOLOCK)  WHERE  DL_PL_Code = '" & Je_Sung(i, 0) & "' And SC_DELIVER.DL_Code  = SC_DELIVER_LIST.DL_Code And SC_DELIVER.DL_Cr_Code = SC_SALES.CR_Code And CR_Sel = '진행중' And Len(DL_Total) > 0  "
            Re_Count = DB_Select(DBT)
            If Re_Count > 0 Then
                Je_Sung(i, 33) = DBT.Rows(0)("DL_To").ToString
            End If
        Next i

        For i = 0 To 1000
            For j = 0 To 1000
                Grid_Str(i, j) = ""
            Next j
        Next i
        For i = 0 To 1000
            For j = 0 To 1000
                Grid_Str_CM(i, j) = ""
            Next j
        Next i


        'Rows와col을 조절한다.
        '제품 표시
        Grid_Col = 0 '8씩 증가
        Grid_Row = 0
        Dim Jee_Sung(33) As Long
        Dim Je_Go As Double
        Grid_Col = Grid_Col ' + 8

        For i = 1 To Je_Cnt
            For j = 1 To 33
                Jee_Sung(j) = Val(Je_Sung(i, j))
            Next j
            Je_Go = 0
            Product_Count(i, Je_Sung(i, 0), Jee_Sung, 1, 1, Je_Go)
            'Exit For
        Next i

        '마지막 제품만 가지고 온다.
        'Grid_Str(Grid_Row, 1000) = "Ok"
        Dim UU As Integer
        Dim UUU As Integer
        Dim k As Integer
        For i = 1 To Grid_Row
            If Grid_Str(i, 1000) <> "" Then
                '앞에 있는 제품코드
                UUU = Val(Grid_Str(i, 1000))
                For j = i To 1 Step -1
                    If Len(Grid_Str(j, UUU)) > 0 Then
                        UU = UU + 1
                        Grid_Str_CM(UU, 1) = Grid_Str(j, UUU)
                        For k = 1 To 32
                            Grid_Str_CM(UU, 7 + k) = Je_ToT(i, 2 + k)
                        Next k
                        Exit For
                    End If
                Next j
                'Grid_Str_CM(Grid_Row, 10)
            End If
        Next
        Grid_Row = UU
        For i = 1 To UU
            StrSQL = "SELECT PL_Name, PL_Standard, SI_Product.PL_Unit, SI_PRODUCT_CUSTOMER.PL_Unit_Gold, PL_CM_Code, PL_CM_Name, PL_Base_Amount  FROM SI_PRODUCT With(NOLOCK), SI_PRODUCT_CUSTOMER With(NOLOCK) WHERE SI_Product.PL_Code = '" & Grid_Str_CM(i, 1) & "' And  SI_Product.PL_Code = SI_PRODUCT_CUSTOMER.PL_Code And SI_PRODUCT_CUSTOMER.PL_Sel = '매입' Order By Convert(Decimal,PL_Sun)"
            Re_Count = DB_Select(DBT)
            If Re_Count > 0 Then
                Grid_Str_CM(i, 2) = DBT.Rows(0)("PL_Name").ToString
                Grid_Str_CM(i, 3) = DBT.Rows(0)("PL_Standard").ToString
                Grid_Str_CM(i, 4) = DBT.Rows(0)("PL_Unit").ToString
                Grid_Str_CM(i, 5) = DBT.Rows(0)("PL_Unit_Gold").ToString
                Grid_Str_CM(i, 6) = DBT.Rows(0)("PL_CM_Code").ToString
                Grid_Str_CM(i, 7) = DBT.Rows(0)("PL_CM_Name").ToString
                Grid_Str_CM(i, 1000) = DBT.Rows(0)("PL_Base_Amount").ToString
            End If
        Next i



        Grid_Main_Setup(Grid_Row, 38)



        For i = 1 To 30
            Grid_Main.Columns(6 + i).HeaderText = Day_Cnt(i)
        Next i
        Grid_Main.Columns(6 + 31).HeaderText = "이후총수량"

        For i = 1 To Grid_Row
            For j = 1 To 38
                Grid_Main.Item(j - 1, i - 1).Value = Grid_Str_CM(i, j)
                If Mid(Grid_Str_CM(i, j), 1, 1) = "-" Then
                    Grid_Main.Rows(i - 1).Cells(j - 1).Style.ForeColor = Color.Red
                Else
                    Grid_Main.Rows(i - 1).Cells(j - 1).Style.ForeColor = Color.Black
                End If

            Next j

            'Grid_Main.Item(Grid_Col, i - 1).Value = Je_ToT(i, 1) & "/" & Je_ToT(i, 2) & "/" & Val(Je_ToT(i, 1)) - Val(Je_ToT(i, 2))
            'For j = 1 To 31
            '    If Mid(Je_ToT(i, 2 + j), 1, 1) = "-" Then
            '        Grid_Main.Rows(i - 1).Cells(Grid_Col + j).Style.ForeColor = Color.Red
            '    Else
            '        Grid_Main.Rows(i - 1).Cells(Grid_Col + j).Style.ForeColor = Color.Black
            '    End If
            '    Grid_Main.Item(Grid_Col + j, i - 1).Value = Je_ToT(i, 2 + j)
            'Next j

        Next i


    End Sub
    Public Function Product_Count(Je_Sub As Long, Je_Code As String, Je_Count() As Long, Qty As String, Grid_Coll As Long, Je_Go As Double) As String
        Dim i As Integer
        Dim j As Integer
        Dim Grid_Co As Integer

        Dim DBT As New DataTable
        '제품 코드 및 이름 추가
        '제품명 표시
        Grid_Co = Grid_Coll * 8
        StrSQL = "Select * FROM SI_PRODUCT With(NOLOCK) Where PL_Code = '" & Je_Code & "'"
        Re_Count = DB_Select(DBT)
        If Re_Count > 0 Then
            Grid_Row = Grid_Row + 1
            Grid_Str(Grid_Row, Grid_Co - 7) = Je_Sub
            Grid_Str(Grid_Row, Grid_Co - 6) = DBT.Rows(0)("PL_Code")
            Grid_Str(Grid_Row, Grid_Co - 5) = DBT.Rows(0)("PL_Name")
            '공정을 call 한다
            Process_Count(DBT.Rows(0)("PL_Code"), Je_Count, Qty, Grid_Coll, Je_Go)
        End If
    End Function
    Public Function Process_Count(Je_Code As String, Je_Count() As Long, Qty As String, Grid_Coll As Long, Je_Go As Double) As String
        Dim i As Integer
        Dim j As Integer
        Dim X As Integer
        Dim Y As Integer
        Dim YY As Integer
        Dim J_Code As String


        Dim DBT As New DataTable
        Dim RR1 As Double
        Dim RR2 As Double
        Dim RR3 As Double

        Dim R1 As Double
        Dim R2 As Double
        Dim R3 As Double
        Dim R4 As Double
        Dim R As Long
        Dim RR As Long
        Dim RRR As String


        Dim Grid_Co As Integer

        Grid_Co = Grid_Coll * 8


        YY = 0
        For Y = 1 To Grid_Row - 1
            For X = 1 To Grid_Co
                If Je_Code = Grid_Str(Y, X) Then
                    '상위에 존제 한다.
                    YY = Y
                    Exit For
                End If
            Next X
            If YY = 0 Then
            Else
                Exit For
            End If
        Next Y


        '제품 코드 및 이름 추가
        '제품명 표시
        StrSQL = "Select PP_Sun, PC_Code, PC_Name, PP_Amount FROM SI_PROCESS With(NOLOCK), SI_Process_List with(NOLOCK) Where PP_PC_Code = PC_Code And PP_Code = '" & Je_Code & "' Order By Convert(Decimal,PP_Sun) Desc"
        Re_Count = DB_Select(DBT)
        If Re_Count > 0 Then
            For i = 0 To Re_Count - 1
                If i = 0 Then
                Else
                    Grid_Row = Grid_Row + 1
                End If
                Grid_Str(Grid_Row, Grid_Co - 4) = DBT.Rows(i)("PP_Sun")
                Grid_Str(Grid_Row, Grid_Co - 3) = DBT.Rows(i)("PC_Code")
                Grid_Str(Grid_Row, Grid_Co - 2) = DBT.Rows(i)("PC_Name")
                '상위에 있는지 확인 한다.

                If YY = 0 Then
                    Grid_Str(Grid_Row, Grid_Co - 1) = DBT.Rows(i)("PP_Amount")
                Else
                    Grid_Str(Grid_Row, Grid_Co - 1) = " - "
                End If

                Grid_Str(Grid_Row, Grid_Co) = Qty
                '수량 계산
                '전체수주
                Je_ToT(Grid_Row, 1) = Je_Count(32) * Val(Qty)
                '전체 납품
                Je_ToT(Grid_Row, 2) = Je_Count(33) * Val(Qty)
                '현재고
                RR1 = Val(DBT.Rows(i)("PP_Amount"))
                Je_Go = Je_Go * Val(Qty)

                If YY = 0 Then
                    Je_Go = Je_Go + RR1
                    For j = 1 To 31
                        '현일짜까지 수주량
                        RR2 = Je_Count(j) * Val(Qty)
                        RR3 = RR2 - Val(Je_ToT(Grid_Row, 2)) - Je_Go
                        RR3 = 0 - RR3
                        Je_ToT(Grid_Row, j + 2) = RR3
                    Next j
                Else
                    Je_Go = Je_Go
                    For j = 1 To 31
                        '현일짜까지 수주량
                        RR2 = Je_Count(j) * Val(Qty)
                        RR3 = RR2 - Val(Je_ToT(Grid_Row, 2)) - Je_Go
                        RR3 = 0 - RR3
                        Je_ToT(YY, j + 2) = Val(Je_ToT(YY, j + 2)) + RR3
                        Je_ToT(Grid_Row, j + 2) = " - "
                    Next j

                End If

            Next i
        End If
        '조합이 있으면 Call 한다.
        StrSQL = "Select PL_Sub_Code, PL_Sun, PL_Qty FROM SI_PRODUCT_RECIPE With(NOLOCK) Where PL_Code = '" & Je_Code & "' Order By Convert(Decimal,PL_Sun) "
        Re_Count = DB_Select(DBT)
        If Re_Count > 0 Then
            'Grid_Coll = Grid_Coll + 8
            If Grid_Col <= Grid_Coll * 8 Then
                Grid_Col = (Grid_Coll + 1) * 8

            End If
            For i = 0 To Re_Count - 1
                Product_Count(DBT.Rows(i)("PL_Sun").ToString, DBT.Rows(i)("PL_Sub_Code").ToString, Je_Count, Val(DBT.Rows(i)("PL_Qty").ToString) * Val(Qty), Grid_Coll + 1, Je_Go)
            Next i

        Else
            '조합이 없으면 발주 한다.
            If YY = 0 Then
                Grid_Str(Grid_Row, 1000) = Grid_Co - 6
            End If

            'Grid_Str(Grid_Row, 999) = Grid_Co - 8
        End If


    End Function

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
        xlbook = xlapp.Workbooks.Open(Application.StartupPath & "\excel\발주계획.xlsx")
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


        '거래 일자
        j = 0
        For j = 0 To Grid_Main.ColumnCount - 1
            xlsheet.Cells(1, j + 1).Value = "'" & Grid_Main.Columns(j).HeaderText
        Next j
        For i = 0 To Grid_Main.RowCount - 1
            '제목표시
            For j = 0 To Grid_Main.ColumnCount - 1
                xlsheet.Cells(i + 2, j + 1).Value = Grid_Main.Item(j, i).Value.ToString
            Next j
        Next i
        xlbook.SaveAs(Application.StartupPath & "\Excel\발주계획\" & Grid_Info.Item(1, 0).Value & Grid_Info.Item(1, 5).Value & ".xls")
        Cursor = Cursors.Default
        Excel_check = True

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

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        '해당 그리드에서 공정별로 생산 수량을 산출 한다.
        '공정별 해당 일자까지 필요 수량을 가지고 온다.
        '        
        If Grid_Info.Item(1, 8).Value = "처리" Then
            MsgBox("발주서가 처리되어 더이상 진행 할 수 없습니다.")
            Exit Sub
        End If
        Dim GO_Code(1000, 5) As String '1 업체코드 2 업체명 
        Dim En_day_Col As Integer
        Dim Go_X As Integer
        Dim J_X As Integer
        Dim J_Y As Integer
        Dim i As Integer
        Dim j As Integer
        Dim NullStr As String
        Dim RR As Long


        ' 해당 일자의 Grid Col를 자지고 온다.
        En_day_Col = 0
        For j = 0 To Grid_Main.ColumnCount - 1
            If Grid_Info.Item(1, 5).Value = Grid_Main.Columns(j).HeaderText Then
                En_day_Col = j
                Exit For
            End If
        Next j
        If En_day_Col = 0 Then
            MsgBox("발주서 기준일자가 1개월 이상 입니다. 확인후 다시 진행 하세요")
            Exit Sub
        End If
        Go_X = 0
        NullStr = ""
        For i = 0 To Grid_Main.RowCount - 1
            If Len(Grid_Main.Item(5, i).Value) > 0 Then
                RR = 0 - Val(Grid_Main.Item(En_day_Col, i).Value)
                If RR < 0 Then
                Else
                    If Go_X = 0 Then
                        Go_X = Go_X + 1
                        GO_Code(Go_X, 1) = Grid_Main.Item(5, i).Value
                        GO_Code(Go_X, 2) = Grid_Main.Item(6, i).Value
                    Else
                        NullStr = ""
                        For j = 1 To Go_X
                            If GO_Code(Go_X, 1) = Grid_Main.Item(5, i).Value Then
                                NullStr = "Ok"
                                Exit For
                            End If
                        Next j
                        If NullStr = "Ok" Then
                        Else
                            Go_X = Go_X + 1
                            GO_Code(Go_X, 1) = Grid_Main.Item(5, i).Value
                            GO_Code(Go_X, 2) = Grid_Main.Item(6, i).Value
                        End If
                    End If

                End If
            End If
        Next i
        '업체별로 발주서를 처리한다.
        '업체코드를 가지고 온다




        Dim G_Co As String

        Dim DBT As New DataTable
        Dim JP_Code As Long
        Dim My_Date As String
        Dim My_Time As String
        Dim De_Code As String
        Dim De_Name As String
        Dim De_Cm_Code As String
        Dim Su_Cnt As Long


        For i = 1 To Go_X
            If Len(GO_Code(i, 1)) > 0 Then
                G_Co = GO_Code(i, 1)
                '발주서 전표 생성

                StrSQL = "Select CM_Man_Name, CM_Man_Tel FROM SI_CUSTOMER with(NOLOCK)  Where CM_Code =  '" & GO_Code(i, 1) & "'  "
                Re_Count = DB_Select(DBT)
                If Re_Count > 0 Then
                    De_Code = DBT.Rows(0)("CM_Man_Name")
                    De_Name = DBT.Rows(0)("CM_Man_Tel")
                    'De_Cm_Code = DBT.Rows(0)("PL_Sel")
                End If
                '
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
                StrSQL = "Select CO_Code From MC_STOCK_ORDER with(NOLOCK) Where CO_Date Like  '" & My_Date & "%' Order By CO_Code Desc "
                Re_Count = DB_Select(DBT)
                If Re_Count = 0 Then
                    JP_Code = JP_Code & "001"
                Else
                    JP_Code = JP_Code & Format(Val(Mid(DBT.Rows(0).Item(0), 11, 3)) + 1, "00#")
                End If
                If Len(Grid_Info.Item(1, 6).Value) < 1 Then
                    Grid_Info.Item(1, 6).Value = "CO" & JP_Code
                End If
                Grid_Info.Item(1, 7).Value = "CO" & JP_Code
                Grid_Info.Item(1, 8).Value = "처리"

                If Grid_Main.Columns(En_day_Col).HeaderText = "이후총수량" Then
                    De_Cm_Code = Grid_Main.Columns(Grid_Main.ColumnCount - 2).HeaderText
                Else
                    De_Cm_Code = Grid_Main.Columns(En_day_Col).HeaderText
                End If
                '추가한다
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert into MC_STOCK_ORDER (CO_Code, CO_Date, CO_Time, CO_Name, CO_Day1, CO_Sel,CO_Customer_Code,CO_Customer_Name,CO_Customer_Man_Name, CO_Customer_Man_Tel,CO_Condition)  Values ('CO" & JP_Code & "', '" & My_Date & "','" & My_Time & "','" & Frm_Main.Text_Man_Name.Text & "', '" & De_Cm_Code & "','자동','" & GO_Code(i, 1) & "','" & GO_Code(i, 2) & "','" & De_Code & "','" & De_Name & "','일반결제' )"
                Re_Count = DB_Execute()
                Su_Cnt = 0
                For j = 0 To Grid_Main.RowCount - 1
                    If Grid_Main.Item(5, j).Value = GO_Code(i, 1) Then
                        RR = 0 - Val(Grid_Main.Item(En_day_Col, j).Value)
                        If RR < 0 Then
                        Else
                            Su_Cnt = Su_Cnt + 1
                            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                            StrSQL = StrSQL & "Insert into MC_STOCK_ORDER_List (CO_Code, CO_Sun,CO_PL_Code, CO_PL_Name, CO_Standard, CO_Unit ,CO_Unit_Amount, CO_day01,CO_Total,CO_Gold)  Values ('CO" & JP_Code & "', '" & Su_Cnt & "','" & Grid_Main.Item(0, j).Value & "','" & Grid_Main.Item(1, j).Value & "','" & Grid_Main.Item(2, j).Value & "','" & Grid_Main.Item(3, j).Value & "','" & Grid_Main.Item(4, j).Value & "','" & RR & "','" & RR & "','" & RR * Val(Grid_Main.Item(4, j).Value) & "')"
                            Re_Count = DB_Execute()

                        End If
                    End If
                Next j
            End If
        Next i
        '저장한다.
        Grid_Info.Rows(6).HeaderCell.Value = "*"
        Grid_Info.Rows(7).HeaderCell.Value = "*"
        Grid_Info.Rows(8).HeaderCell.Value = "*"



        '저장한다.
        Dim Table_Name As String
        Dim Field_Name(500) As String
        j = 0
        For i = 1 To Grid_Info.RowCount - 1
            If Grid_Info.Rows(i).HeaderCell.Value = "*" Then
                j = 1
            End If
        Next i
        If j = 0 Then
        Else
            Table_Name = "MC_STOCK"
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

        '내역 저장

        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Delete From MC_STOCK_LIst Where CL_Code = '" & Grid_Info.Item(1, 0).Value & "'"
        Re_Count = DB_Execute()


        NullStr = ""
        For j = 0 To Grid_Main.Columns.Count - 1
            NullStr = NullStr & Grid_Main.Columns(j).HeaderText
            If j = Grid_Main.Columns.Count - 1 Then
                NullStr = NullStr
            Else
                NullStr = NullStr & "|"
            End If
        Next j
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert into MC_STOCK_List  (CL_Code, CL_Sun, CL_Data)  Values ('" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "', '1','" & Replace(NullStr, "'", "''") & "' )"
        Re_Count = DB_Execute()

        For i = 0 To Grid_Main.Rows.Count - 1
            NullStr = ""
            For j = 0 To Grid_Main.Columns.Count - 1
                NullStr = NullStr & Grid_Main.Item(j, i).Value
                If j = Grid_Main.Columns.Count - 1 Then
                    NullStr = NullStr
                Else
                    NullStr = NullStr & "|"
                End If
            Next j
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "Insert into MC_STOCK_List  (CL_Code, CL_Sun, CL_Data)  Values ('" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "', '" & i + 2 & "','" & Replace(NullStr, "'", "''") & "' )"
            Re_Count = DB_Execute()
        Next i


    End Sub

    Private Sub Delete_Com_Click(sender As Object, e As EventArgs) Handles Delete_Com.Click
        If Grid_Info.Item(1, 8).Value = "처리" Then
            MsgBox("발주서가 처리되어 삭제 할 수 없습니다.")
            Exit Sub
        End If

        '해당 칼럼 삭제
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        If MsgBox("발주량 산출  '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "'를 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "발주량 산출 삭제") = vbNo Then
            Exit Sub
        End If



        Grid_Display_Ch = False
        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Delete MC_STOCK Where CL_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' "
        Re_Count = DB_Execute()

        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Delete MC_STOCK_List Where CL_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' "
        Re_Count = DB_Execute()

        Grid_Code_Display()
        Grid_Info_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True



    End Sub


End Class
