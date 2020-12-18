﻿Public Class Frm_PC_Stock2
    Dim Grid_Display_Ch As Boolean
    Dim Grid_Info_Row As Integer
    Dim Grid_Info_Col As Integer

    Dim Grid_Main_Row As Integer
    Dim Grid_Main_Col As Integer
    Dim Grid_Line_QC_Row As Integer
    Dim Grid_Line_QC_Col As Integer
    Dim Grid_Go_Row As Integer
    Dim Grid_Go_Col As Integer

    Dim Info_Lab() As Button
    Dim Info_Txt() As TextBox
    Dim Info_Com() As ComboBox

    ' Grid_Line_QC
    Private Sub Frm_PC_Stock_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Info_Lab = New Button() {Info_La0, Info_La1, Info_La2, Info_La3, Info_La4, Info_La5, Info_La6, Info_La7, Info_La8, Info_La9, Info_La10, Info_La11}
        Info_Txt = New TextBox() {Info_Tx0, Info_Tx1, Info_Tx2, Info_Tx3, Info_Tx4, Info_Tx5, Info_Tx6, Info_Tx7, Info_Tx8, Info_Tx9, Info_Tx10, Info_Tx11}
        Info_Com = New ComboBox() {Info_Co0, Info_Co1, Info_Co2, Info_Co3, Info_Co4, Info_Co5, Info_Co6, Info_Co7, Info_Co8, Info_Co9, Info_Co10, Info_Co11}

        Me.Dock = DockStyle.Fill
        Me.BringToFront()
        Grid_Code_Setup()
        Grid_Info_Setup()
        Grid_Main_Setup()
        Grid_Error_Setup()
        Process_Line_QC_Setup()
        Grid_Go_Setup()



        Search_Combo.Items.Add("코드")
        Search_Combo.Items.Add("날짜")
        Search_Combo.Items.Add("생산라인")
        Search_Combo.Text = "코드"


        Panel_Main.Top = 136
        Panel_Main.Left = 389
        Panel_Excel.Top = 136
        Panel_Excel.Left = 389

        Panel_Go.Top = 136
        Panel_Go.Left = 389

        Panel_Main.Visible = True
        Panel_Excel.Visible = False
        Panel_Error.Visible = False
        Panel_Go.Visible = False

        Save_Com.Visible = False
        Insert_Com.Visible = False
        Delete_Com.Visible = False
        Panel3.Visible = False

        Search_Com.PerformClick()
        Grid_Code.ClearSelection()
    End Sub
    Public Function Grid_Error_Setup() As Boolean
        Grid_Color(Grid_Error)

        Grid_Error.AllowUserToAddRows = False
        Grid_Error.RowTemplate.Height = 24
        Grid_Error.ColumnCount = 4
        Grid_Error.RowCount = 0

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Error.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Error.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Error.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Error.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Error.RowHeadersWidth = 10


        Grid_Error.Columns(0).HeaderText = "순번"
        Grid_Error.Columns(1).HeaderText = "불량유형"
        Grid_Error.Columns(2).HeaderText = "수량"
        Grid_Error.Columns(3).HeaderText = "비고"

        Dim i As Integer

        Grid_Error.Columns(0).Width = 55
        Grid_Error.Columns(1).Width = 100
        Grid_Error.Columns(2).Width = 55
        Grid_Error.Columns(3).Width = 100
        Grid_Error_Setup = True
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
        Grid_Code.Columns(2).HeaderText = "생산라인"

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
        Grid_Info.Item(0, 4).Value = "지시서코드"
        Grid_Info.Item(0, 5).Value = "공정코드"
        Grid_Info.Item(0, 6).Value = "공정명"
        Grid_Info.Item(0, 7).Value = "라인코드"
        Grid_Info.Item(0, 8).Value = "라인명"
        Grid_Info.Item(0, 9).Value = "구분"
        Grid_Info.Item(0, 10).Value = "수량"
        Grid_Info.Item(0, 11).Value = "비고"

        '버튼 텍스트 변경
        For i = 0 To Grid_Info.RowCount - 1
            Info_Lab(i).Text = Grid_Info.Item(0, i).Value
        Next i

        '콤보박스 숨기기
        For i = 0 To Grid_Info.RowCount - 1
            Info_Com(i).Visible = False
        Next i

        Info_Com(4).Visible = True
        Info_Com(5).Visible = True
        Info_Com(6).Visible = True
        Info_Com(7).Visible = True
        Info_Com(8).Visible = True

        Info_Txt(4).Visible = False
        'Info_Txt(5).Visible = False
        Info_Txt(6).Visible = False
        'Info_Txt(7).Visible = False
        'Info_Txt(8).Visible = False

        Info_Txt(9).ReadOnly = True
        Info_Tx10.ReadOnly = True

        Grid_Info.Rows(0).ReadOnly = True
        Grid_Info.Rows(1).ReadOnly = True
        Grid_Info.Rows(2).ReadOnly = True
        Grid_Info.Rows(3).ReadOnly = True
        Grid_Info.Rows(9).ReadOnly = True
        Grid_Info.Rows(10).ReadOnly = True

        Grid_Info_Setup = True
    End Function

    Public Function Grid_Main_Setup() As Boolean
        Grid_Color(Grid_Main)

        Grid_Main.AllowUserToAddRows = False
        Grid_Main.RowTemplate.Height = 24
        Grid_Main.ColumnCount = 14
        Grid_Main.RowCount = 0

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Main.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.RowHeadersWidth = 10


        Grid_Main.Columns(0).HeaderText = "순번"
        Grid_Main.Columns(1).HeaderText = "제품코드"
        Grid_Main.Columns(2).HeaderText = "제품명"
        Grid_Main.Columns(3).HeaderText = "규격"
        Grid_Main.Columns(4).HeaderText = "단위"
        Grid_Main.Columns(5).HeaderText = "지시수량"
        Grid_Main.Columns(6).HeaderText = "시작일자"
        Grid_Main.Columns(7).HeaderText = "시작시간"
        Grid_Main.Columns(8).HeaderText = "완료일자"
        Grid_Main.Columns(9).HeaderText = "완료시간"
        Grid_Main.Columns(10).HeaderText = "투입수량"
        Grid_Main.Columns(11).HeaderText = "생산수량"
        Grid_Main.Columns(12).HeaderText = "불량수량"
        Grid_Main.Columns(13).HeaderText = "비고"

        Dim i As Integer

        Grid_Main.Columns(0).Width = 55
        Grid_Main.Columns(1).Width = 100
        Grid_Main.Columns(2).Width = 100
        Grid_Main.Columns(3).Width = 100
        Grid_Main.Columns(4).Width = 50
        For i = 5 To Grid_Main.ColumnCount - 1
            Grid_Main.Columns(i).Width = 79
        Next i

        'Grid_Main.Columns(12).ReadOnly = True

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


        StrSQL = "Select PS_Code FROM PC_STOCK with(NOLOCK) Where PS_Date Like  '" & My_Date & "%' Order By PS_Code Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            JP_Code = JP_Code & "001"
        Else
            JP_Code = JP_Code & Format(Val(Mid(DBT.Rows(0).Item(0), 11, 3)) + 1, "00#")
        End If

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert into PC_Stock (PS_Code, PS_Date, PS_Time, PS_Name, PS_Sel)  Values ('PS" & JP_Code & "', '" & My_Date & "','" & My_Time & "','" & Frm_Main.Text_Man_Name.Text & "','관리')"
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
                StrSQL = "Select PS_Code, PS_Date, PS_DE_Name FROM PC_STOCK with(NOLOCK) Where PS_Code Like '%" & Search_Text.Text & "%'  Order  By PS_Code DESC"
            Case "날짜"
                StrSQL = "Select PS_Code, PS_Date, PS_DE_Name FROM PC_STOCK with(NOLOCK) Where PS_Date Like '%" & Search_Text.Text & "%' Or PS_Date Is Null  Order By PS_Date"
            Case "생산라인"
                StrSQL = "Select PS_Code, PS_Date, PS_DE_Name FROM PC_STOCK with(NOLOCK) Where PS_DE_Name Like '%" & Search_Text.Text & "%' Or PS_DE_Name Is Null  Order By PS_DE_Name"
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

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        Grid_Code_Display()
    End Sub
    Private Sub Grid_Code_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Code.CellEnter




        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
            Panel_Main.Visible = True
            'Grid_Info_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
            Grid_Info_Display2(Grid_Info, "Select * FROM PC_STOCK With(NOLOCK) Where PS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "'", Info_Txt, Info_Com)
            Label_Color(Com_Main, Color_Label, Di_Panel2)
            Grid_Main_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
            'Grid_Line_QC_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)

            Panel_Main.Visible = True
            Panel_Excel.Visible = False
            Panel_Go.Visible = False
            Grid_Info_Combo.Visible = False
            Grid_Main_Combo.Visible = False

            'If Grid_Info.Item(1, 9).Value = "현장" Then
            '    Info_Tx10.ReadOnly = True
            'Else
            '    Grid_Info.Item(1, 10).ReadOnly = False
            '    Info_Tx10.ReadOnly = False
            'End If

        End If
    End Sub
    Public Function Grid_Info_Display(Code As String) As Boolean
        Grid_Display_Ch = False
        Grid_Info_Display = True
        Dim DBT As New DataTable
        Dim i As Integer
        StrSQL = "Select * FROM PC_STOCK With(NOLOCK) Where PS_Code = '" & Code & "'"
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
        StrSQL = "Select PS_Sun, PS_PL_Code, PS_PL_Name, PS_Standard, PS_Unit, PS_PO_Total, PS_St_Day, PS_St_Time, PS_En_Day, PS_En_Time, PS_Go, PS_Total, PS_Er, PS_Bigo  FROM PC_STOCK_List with(NOLOCK)  Where PS_Code = '" & PL_Code & "' Order By Convert(Decimal,PS_Sun)"
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


        Grid_Display_Ch = True
        Grid_Main_Display = True

    End Function
    Public Function Grid_Error_Display(PL_Code As String, PL_Count As String) As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display_Ch = False
        Grid_Error.RowCount = 0
        StrSQL = "Select PS_Sun, PS_History, PS_Total, PS_Bigo  FROM PC_STOCK_Er with(NOLOCK)  Where PS_Code = '" & PL_Code & "' And PS_Count = '" & PL_Count & "' Order By Convert(Decimal,PS_Sun)"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Error.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To Grid_Error.ColumnCount - 1
                    Grid_Error.Item(j, i).Value = DBT.Rows(i).Item(j).ToString
                Next j
            Next i
        Else
            Grid_Error.RowCount = 1
            For j = 0 To Grid_Error.ColumnCount - 1
                Grid_Error.Item(j, 0).Value = ""
            Next j
        End If


        Grid_Display_Ch = True
        Grid_Error_Display = True

    End Function

    Public Function Grid_Line_QC_Display(PL_Code As String) As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display_Ch = False
        Grid_Line_QC.RowCount = 0

        StrSQL = "Select PS_Sun, PS_History, PS_Result, PS_Bigo FROM PC_STOCK_Device_QC with(NOLOCK)  Where PS_Device = '" & PL_Code & "' And PS_Day = '" & Format(Now, "yyyy-MM-dd") & "' Order By Convert(Decimal,PS_Sun)"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
        Else
            '추가 한다.

            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "Insert into PC_Stock_Device_QC (PS_Device, PS_Day, PS_Sun, PS_History, PS_Result, PS_Bigo  )  Select '" & PL_Code & "', '" & Format(Now, "yyyy-MM-dd") & "' , PQ_Sun , PQ_History, '',PQ_Bigo From SI_Machine_QC with(NOLOCK) Where PQ_Code = '" & Grid_Info.Item(1, 7).Value & "' "
            Re_Count = DB_Execute()
            'INSERT INTO Man ( Man_Code, Man_Pic ) Select Man_Door.Man_Code, '1224' AS Expr1 From Man_Door Where (((Man_Door.Man_Con_Name) ='123456'));
        End If


        StrSQL = "Select PS_Sun, PS_History, PS_Result, PS_Bigo FROM PC_STOCK_Device_QC with(NOLOCK)  Where PS_Device = '" & PL_Code & "' And PS_Day = '" & Format(Now, "yyyy-MM-dd") & "' Order By Convert(Decimal,PS_Sun)"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Line_QC.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To Grid_Line_QC.ColumnCount - 1
                    Grid_Line_QC.Item(j, i).Value = DBT.Rows(i).Item(j).ToString
                Next j
            Next i
        Else
            'Grid_Line_QC.RowCount = 1
            'For j = 0 To Grid_Line_QC.ColumnCount - 1
            '    Grid_Line_QC.Item(j, 0).Value = ""
            'Next j
        End If


        Grid_Display_Ch = True
        Grid_Line_QC_Display = True

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
                    StrSQL = "Select PO_Code FROM PC_ORDER With(NOLOCK) WHERE PO_PR_CODE IS NOT NULL Order By PO_Code Desc"
                    Re_Count = DB_Select(DBT)
                    Grid_Info_Combo.Items.Clear()
                    If Re_Count <> 0 Then
                        For i = 0 To Re_Count - 1
                            Grid_Info_Combo.Items.Add(DBT.Rows(i)("PO_Code"))
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
                    StrSQL = "Select PC_Code  From SI_Process With(NOLOCK) Order By PC_Code"
                    Re_Count = DB_Select(DBT)
                    Grid_Info_Combo.Items.Clear()
                    If Re_Count <> 0 Then
                        For i = 0 To Re_Count - 1
                            Grid_Info_Combo.Items.Add(DBT.Rows(i)("PC_Code"))
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
                Case 6
                    StrSQL = "Select PC_Name  From SI_Process With(NOLOCK) Order By PC_Name"
                    Re_Count = DB_Select(DBT)
                    Grid_Info_Combo.Items.Clear()
                    If Re_Count <> 0 Then
                        For i = 0 To Re_Count - 1
                            Grid_Info_Combo.Items.Add(DBT.Rows(i)("PC_Name"))
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
                Case 7
                    StrSQL = "Select PL_Code  From SI_Machine With(NOLOCK), SI_Process_Device With(NOLOCK) Where SI_Process_Device.PC_Code = '" & Grid_Info.Item(1, 5).Value.ToString & "' And SI_Process_Device.PC_PL_Code = SI_Machine.PL_Code  Order By Convert(Decimal,SI_Process_Device.PC_Sun) Desc "
                    Re_Count = DB_Select(DBT)
                    Grid_Info_Combo.Items.Clear()
                    If Re_Count <> 0 Then
                        For i = 0 To Re_Count - 1
                            Grid_Info_Combo.Items.Add(DBT.Rows(i)("PL_Code"))
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
                    StrSQL = "Select PL_Name  From SI_Machine With(NOLOCK), SI_Process_Device With(NOLOCK) Where SI_Process_Device.PC_Code = '" & Grid_Info.Item(1, 5).Value & "' And SI_Process_Device.PC_PL_Code = SI_Machine.PL_Code  Order By Convert(Decimal,SI_Process_Device.PC_Sun) Desc "
                    Re_Count = DB_Select(DBT)
                    Grid_Info_Combo.Items.Clear()
                    If Re_Count <> 0 Then
                        For i = 0 To Re_Count - 1
                            Grid_Info_Combo.Items.Add(DBT.Rows(i)("PL_Name"))
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
    Private Sub Grid_Info_Combo_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Grid_Info_Combo.SelectionChangeCommitted
        Dim DBT As New DataTable
        Grid_Info.Item(Grid_Info_Col, Grid_Info_Row).Value = Grid_Info_Combo.SelectedItem.ToString()

        If Grid_Info_Row = 4 Then
            StrSQL = "Select PO_Code,PO_PR_Code, PO_PR_Name, PO_DE_Code, PO_De_Name  FROM PC_ORDER With(NOLOCK) Where PO_Code = '" & Grid_Info_Combo.SelectedItem.ToString() & "'"
            Re_Count = DB_Select(DBT)
            Grid_Info_Combo.Items.Clear()
            If Re_Count <> 0 Then
                Grid_Info.Item(Grid_Info_Col, 5).Value = DBT.Rows(0)("PO_PR_Code")
                Grid_Info.Item(Grid_Info_Col, 6).Value = DBT.Rows(0)("PO_PR_Name")
                Grid_Info.Item(Grid_Info_Col, 7).Value = DBT.Rows(0)("PO_DE_Code")
                Grid_Info.Item(Grid_Info_Col, 8).Value = DBT.Rows(0)("PO_DE_Name")
                Grid_Info.Rows(4).HeaderCell.Value = "*"
                Grid_Info.Rows(5).HeaderCell.Value = "*"
                Grid_Info.Rows(6).HeaderCell.Value = "*"
                Grid_Info.Rows(7).HeaderCell.Value = "*"
                Grid_Info.Rows(8).HeaderCell.Value = "*"
            End If
        End If
        If Grid_Info_Row = 5 Then
            StrSQL = "Select PC_Code, PC_Name  From SI_Process With(NOLOCK) Where PC_Code = '" & Grid_Info_Combo.SelectedItem.ToString() & "'"
            Re_Count = DB_Select(DBT)
            Grid_Info_Combo.Items.Clear()
            If Re_Count <> 0 Then
                Grid_Info.Item(Grid_Info_Col, 6).Value = DBT.Rows(0)("PC_Name")
                Grid_Info.Rows(5).HeaderCell.Value = "*"
                Grid_Info.Rows(6).HeaderCell.Value = "*"
            End If
        End If
        If Grid_Info_Row = 6 Then
            StrSQL = "Select PC_Code, PC_Name  From SI_Process With(NOLOCK) Where PC_Name = '" & Grid_Info_Combo.SelectedItem.ToString() & "'"
            Re_Count = DB_Select(DBT)
            Grid_Info_Combo.Items.Clear()
            If Re_Count <> 0 Then
                Grid_Info.Item(Grid_Info_Col, 5).Value = DBT.Rows(0)("PC_Code")
                Grid_Info.Rows(5).HeaderCell.Value = "*"
                Grid_Info.Rows(6).HeaderCell.Value = "*"
            End If
        End If
        If Grid_Info_Row = 7 Then
            StrSQL = "Select PL_Code, PL_Name, PL_Sel  From SI_Machine With(NOLOCK) Where PL_Code = '" & Grid_Info_Combo.SelectedItem.ToString() & "'"
            Re_Count = DB_Select(DBT)
            Grid_Info_Combo.Items.Clear()
            If Re_Count <> 0 Then
                Grid_Info.Item(Grid_Info_Col, 8).Value = DBT.Rows(0)("PL_Name")
                Grid_Info.Rows(7).HeaderCell.Value = "*"
                Grid_Info.Rows(8).HeaderCell.Value = "*"
            End If
        End If
        If Grid_Info_Row = 8 Then
            StrSQL = "Select PL_Code, PL_Name, PL_Sel  From SI_Machine With(NOLOCK) Where PL_Name = '" & Grid_Info_Combo.SelectedItem.ToString() & "'"
            Re_Count = DB_Select(DBT)
            Grid_Info_Combo.Items.Clear()
            If Re_Count <> 0 Then
                Grid_Info.Item(Grid_Info_Col, 7).Value = DBT.Rows(0)("PL_Code")
                Grid_Info.Rows(7).HeaderCell.Value = "*"
                Grid_Info.Rows(8).HeaderCell.Value = "*"
            End If
        End If

        Grid_Info.Rows(Grid_Info_Row).HeaderCell.Value = "*"
        Grid_Info_Combo.Visible = False


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
            My_Time = My_Time
        End If

        Dim Db_Sun As Long
        StrSQL = "Select PS_Sun FROM PC_STOCK_List with(NOLOCK) Where PS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' Order By Convert(Decimal,PS_Sun) Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            Db_Sun = 1
        Else
            Exit Sub
        End If


        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert into PC_Stock_List  (PS_Code, PS_Sun,PS_St_Day,PS_St_Time,PS_En_Day,PS_En_Time)  Values ('" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "', '" & Db_Sun & "','" & My_Date & "','" & My_Time & "','" & My_Date & "','" & My_Time & "' )"
        Re_Count = DB_Execute()
        Grid_Main_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True

    End Sub

    Private Sub save_info()
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
            Table_Name = "PC_Stock"
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

        Search_Com.PerformClick()
        MsgBox("저장되었습니다")
        '수량 변경 한다.
        '수량 변경
        '수량이 변경 되었는지 확인 한다.
        Dim Val_Check As String
        Dim PP_G_Sun As String
        Dim PP_G_Amount As String

        Val_Check = ""
        StrSQL = "Select PS_Check FROM PC_STOCK With(NOLOCK) Where PS_Code = '" & Grid_Info.Item(1, 0).Value & "'"
        Re_Count = DB_Select(DBT)
        Grid_Info_Combo.Items.Clear()
        If Re_Count <> 0 Then
            If DBT.Rows(0)("PS_Check").ToString = "처리" Then
                Val_Check = "처리"
                MsgBox("이미 처리 되었습니다. 삭제 후 다시 처리 하세요")
                Exit Sub
            End If
        End If
        '수량 변경
        '기존 Data를 삭제 한다.
        'Grid의 수량을 가지고 온다.
        For i = 0 To Grid_Main.RowCount - 1
            Count_Pro(Grid_Main.Item(1, i).Value, Grid_Info.Item(1, 5).Value, Grid_Main.Item(10, i).Value, Grid_Main.Item(11, i).Value, Grid_Main.Item(12, i).Value)
        Next i

        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Update PC_Stock Set PS_Check = '처리' Where PS_Code = '" & Grid_Info.Item(1, 0).Value & "'"
        Re_Count = DB_Execute()

        Grid_Info.Item(1, 10).Value = "처리"
        '  MsgBox("저장되었습니다")
        Search_Com.PerformClick()

    End Sub

    Private Sub Save_Com_Click(sender As Object, e As EventArgs) Handles Save_Com.Click
        save_info()


    End Sub

    Private Sub Insert__Main_Click(sender As Object, e As EventArgs) Handles Insert__Main.Click
        Dim DBT As New DataTable

        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        If Grid_Info.Item(1, 10).Value = "처리" Then
            MsgBox("수량 처리가 되어서 수정 할 수 없습니다.")
            Exit Sub
        End If

        If Len(Grid_Info.Item(1, 7).Value) > 0 Then
            StrSQL = "Select PS_Sun, PS_History, PS_Result, PS_Bigo FROM PC_STOCK_Device_QC with(NOLOCK)  Where PS_Device = '" & Grid_Info.Item(1, 7).Value & "' And PS_Day = '" & Format(Now, "yyyy-MM-dd") & "' Order By Convert(Decimal,PS_Sun)"
            Re_Count = DB_Select(DBT)
            If Re_Count <> 0 Then
            Else
                MsgBox("장비 점검 List 가 처리 되지 않았습니다. 확인 후 처리 하세요")
                Exit Sub
            End If

        End If

        Grid_Display_Ch = False
        '새로운 코드 추가
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
            My_Time = My_Time
        End If
        'Dim DBT As New DataTable
        Dim Db_Sun As Long
        StrSQL = "Select PS_Sun FROM PC_STOCK_List with(NOLOCK) Where PS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' Order By Convert(Decimal,PS_Sun) Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            Db_Sun = 1
        Else
            Db_Sun = DBT.Rows(0).Item(0) + 1
        End If



        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert into PC_Stock_List  (PS_Code, PS_Sun,PS_St_Day,PS_St_Time,PS_En_Day,PS_En_Time)  Values ('" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "', '" & Db_Sun & "','" & My_Date & "','" & My_Time & "','" & My_Date & "','" & My_Time & "' )"
        Re_Count = DB_Execute()
        Grid_Main_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True

    End Sub
    Private Sub Grid_Main_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Main.CellValueChanged
        If Grid_Display_Ch = True Then
            Grid_Main.CurrentRow.HeaderCell.Value = "*"
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
                StrSQL = "Select SI_Product.PL_Code, SI_Product.PL_Name  From SI_Product with(NOLOCK), SI_Process_List with(NOLOCK) Where SI_Process_List.PP_PC_Code = '" & Grid_Info.Item(1, 5).Value & "' And  SI_Process_List.PP_Code = SI_Product.PL_Code  Order By PL_Code"
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
                StrSQL = "Select SI_Product.PL_Code, SI_Product.PL_Name  From SI_Product with(NOLOCK), SI_Process_List with(NOLOCK) Where SI_Process_List.PP_PC_Code = '" & Grid_Info.Item(1, 5).Value & "' And  SI_Process_List.PP_Code = SI_Product.PL_Code  Order By PL_Name"
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
    Private Sub Grid_Main_Combo_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Grid_Main_Combo.SelectionChangeCommitted
        Dim DBT As New DataTable
        Grid_Main.Item(Grid_Main_Col, Grid_Main_Row).Value = Grid_Main_Combo.SelectedItem.ToString()
        'Grid_Main.CurrentRow.HeaderCell.Value = "*"



        Select Case Grid_Main_Col
            Case 1
                StrSQL = "Select PL_Code, PL_Name, PL_Standard ,PL_Unit, PO_Total From SI_Product with(NOLOCK), PC_ORDER_LIST with(NOLOCK) Where PL_Code = '" & Grid_Main_Combo.SelectedItem.ToString() & "' And PO_Code = '" & Grid_Info.Item(1, 4).Value & "' And PO_PL_Code = PL_Code "
                StrSQL = "Select PL_Code, PL_Name, PL_Standard ,PL_Unit, PO_Total From SI_Product with(NOLOCK), PC_ORDER_LIST with(NOLOCK) Where PL_Code = '" & Grid_Main_Combo.SelectedItem.ToString() & "' "
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Main.Item(2, Grid_Main_Row).Value = DBT.Rows(0).Item(1)
                    Grid_Main.Item(3, Grid_Main_Row).Value = DBT.Rows(0).Item(2)
                    Grid_Main.Item(4, Grid_Main_Row).Value = DBT.Rows(0).Item(3)
                    Grid_Main.Item(5, Grid_Main_Row).Value = DBT.Rows(0).Item(4)
                End If
                '지시 수량을 가지고 온다.

            Case 2
                StrSQL = "Select PL_Code, PL_Name, PL_Standard ,PL_Unit, PO_Total From SI_Product with(NOLOCK), PC_ORDER_LIST with(NOLOCK) Where PL_Name = '" & Grid_Main_Combo.SelectedItem.ToString() & "' And PO_Code = '" & Grid_Info.Item(1, 4).Value & "' And PO_PL_Code = PL_Code "
                StrSQL = "Select PL_Code, PL_Name, PL_Standard ,PL_Unit, PO_Total From SI_Product with(NOLOCK), PC_ORDER_LIST with(NOLOCK) Where PL_Name = '" & Grid_Main_Combo.SelectedItem.ToString() & "'  "
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

    Private Sub Com_Excel_Click(sender As Object, e As EventArgs) Handles Com_Excel.Click
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        If Len(Grid_Info.Item(1, 7).Value) > 0 Then
        Else
            Exit Sub
        End If

        Label_Color(sender, Color_Label, Di_Panel2)
        '판넬 보이기
        Panel_Main.Visible = False
        Panel_Excel.Visible = True
        Panel_Go.Visible = False

        '전표에 저장된 내역이 있으면 저장된 Data 없으면 장비 Teabl
        Grid_Line_QC_Display(Grid_Info.Item(1, 7).Value)




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
        Panel_Go.Visible = False

    End Sub

    Public Function Process_Line_QC_Setup() As Boolean
        Grid_Color(Grid_Line_QC)

        Grid_Line_QC.AllowUserToAddRows = False
        Grid_Line_QC.RowTemplate.Height = 24


        '열의 갯수는 하나 더 많아야 함.
        Grid_Line_QC.ColumnCount = 4
        Grid_Line_QC.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Line_QC.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Line_QC.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Line_QC.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Grid_Line_QC.RowHeadersWidth = 40
        Grid_Line_QC.Columns(0).Width = 60
        Grid_Line_QC.Columns(1).Width = 150
        Grid_Line_QC.Columns(2).Width = 100
        Grid_Line_QC.Columns(3).Width = 100
        Grid_Line_QC.Columns(0).HeaderText = "순번"
        Grid_Line_QC.Columns(1).HeaderText = "점검내역"
        Grid_Line_QC.Columns(2).HeaderText = "점검결과"
        Grid_Line_QC.Columns(3).HeaderText = "비고"


        Grid_Line_QC.Columns(0).ReadOnly = True
        Process_Line_QC_Setup = True
    End Function
    Private Sub Grid_Info_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Info.CellValueChanged
        If Grid_Display_Ch = True Then
            Grid_Info.CurrentRow.HeaderCell.Value = "*"



        End If

    End Sub

    Private Sub Grid_Info_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Info.CellContentClick

    End Sub

    Private Sub Grid_Code_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Code.CellContentClick

    End Sub

    Private Sub Save_Main_Click(sender As Object, e As EventArgs) Handles Save_Main.Click
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        If Grid_Info.Item(1, 10).Value = "처리" Then
            MsgBox("수량 처리가 되어서 수정 할 수 없습니다.")
            Exit Sub
        End If
        Dim DBT As New DataTable

        If Len(Grid_Info.Item(1, 7).Value) > 0 Then
            StrSQL = "Select PS_Sun, PS_History, PS_Result, PS_Bigo FROM PC_STOCK_Device_QC with(NOLOCK)  Where PS_Device = '" & Grid_Info.Item(1, 7).Value & "' And PS_Day = '" & Format(Now, "yyyy-MM-dd") & "' Order By Convert(Decimal,PS_Sun)"
            Re_Count = DB_Select(DBT)
            If Re_Count <> 0 Then
            Else
                MsgBox("장비 점검 List 가 처리 되지 않았습니다. 확인 후 처리 하세요")
                Exit Sub
            End If
        End If


        Dim i As Integer

        For i = 0 To Grid_Main.Rows.Count - 1
            If Grid_Main.Rows(i).HeaderCell.Value = "*" Then
                StrSQL = "Select PQ_Sun, PQ_History1, PQ_History2, PQ_Result, PQ_Bigo  FROM PC_STOCK_QC with(NOLOCK)  Where PQ_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And PQ_PL_Code = '" & Grid_Main.Item(1, i).Value & "' And  PQ_PC_Code  = '" & Grid_Info.Item(1, 5).Value & "' Order By Convert(Decimal,PQ_Sun)"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                Else
                    StrSQL = "Select PQ_PC_Code, PQ_Sun, PQ_History1, PQ_History2,  PQ_Bigo  From SI_Product_QC with(NOLOCK) Where PQ_Code = '" & Grid_Main.Item(1, i).Value & "' And PQ_PC_Code = '" & Grid_Info.Item(1, 5).Value & "'"
                    Re_Count = DB_Select(DBT)
                    If Re_Count <> 0 Then
                        MsgBox("'" & Grid_Main.Item(1, i).Value & "' '" & Grid_Main.Item(2, i).Value & "' 공정 점검 List 가 처리 되지 않았습니다. 확인 후 처리 하세요")
                        Exit Sub
                    Else
                    End If
                End If
            End If
        Next i


        Grid_Display_Ch = False
        For i = 0 To Grid_Main.Rows.Count - 1
            If Grid_Main.Rows(i).HeaderCell.Value = "*" Then

                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Update PC_Stock_List Set PS_PL_COde = '" & Grid_Main.Item(1, i).Value & "', PS_PL_Name = '" & Grid_Main.Item(2, i).Value & "', PS_Standard = '" & Grid_Main.Item(3, i).Value & "', PS_Unit = '" & Grid_Main.Item(4, i).Value & "', PS_PO_Total = '" & Grid_Main.Item(5, i).Value & "', PS_St_Day = '" & Grid_Main.Item(6, i).Value & "', PS_St_Time = '" & Grid_Main.Item(7, i).Value & "', PS_En_day = '" & Grid_Main.Item(8, i).Value & "', PS_En_Time = '" & Grid_Main.Item(9, i).Value & "', PS_Go = '" & Grid_Main.Item(10, i).Value & "', PS_Total = '" & Grid_Main.Item(11, i).Value & "', PS_Er = '" & Grid_Main.Item(12, i).Value & "', PS_Bigo = '" & Grid_Main.Item(13, i).Value & "' Where PS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And PS_Sun = '" & Grid_Main.Item(0, i).Value & "'"
                Re_Count = DB_Execute()
                Grid_Main.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display_Ch = True

        save_info()

        '  MsgBox("저장되었습니다")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        Dim i As Integer

        Grid_Display_Ch = False
        For i = 0 To Grid_Line_QC.Rows.Count - 1
            If Grid_Line_QC.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "update PC_Stock_Device_QC Set PS_Result = '" & Grid_Line_QC.Item(2, i).Value & "', PS_Bigo = '" & Grid_Line_QC.Item(3, i).Value & "' where PS_Device = '" & Grid_Info.Item(1, 7).Value & "' And PS_Day = '" & Format(Now, "yyyy-MM-dd") & "' And PS_Sun = '" & Grid_Line_QC.Item(0, i).Value & "'"
                Re_Count = DB_Execute()
                Grid_Line_QC.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display_Ch = True



    End Sub

    Private Sub Grid_Line_QC_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Line_QC.CellValueChanged
        If Grid_Display_Ch = True Then
            Grid_Line_QC.CurrentRow.HeaderCell.Value = "*"
        End If
    End Sub

    Private Sub Grid_Main_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Main.CellContentClick

    End Sub

    Private Sub Grid_Main_Combo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Grid_Main_Combo.SelectedIndexChanged

    End Sub

    Private Sub Grid_Main_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles Grid_Main.MouseDoubleClick
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        If Grid_Info.Item(1, 10).Value = "처리" Then
            MsgBox("수량 처리가 되어서 수정 할 수 없습니다.")
            Exit Sub
        End If
        If Len(Grid_Main.Item(1, Grid_Main.CurrentCell.RowIndex).Value) < 1 Then
            Exit Sub
        End If


        '불량 처리 한다.
        Panel_Error.Visible = True
        Grid_Error_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value, Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value)

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        '수량 Check



        Panel_Error.Visible = False

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click



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
            My_Time = My_Time
        End If
        'Dim DBT As New DataTable
        Dim Db_Sun As Long
        StrSQL = "Select PS_Sun FROM PC_STOCK_Er with(NOLOCK) Where PS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And PS_Count = '" & Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value & "' Order By Convert(Decimal,PS_Sun) Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            Db_Sun = 1
        Else
            Db_Sun = DBT.Rows(0).Item(0) + 1
        End If



        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert into PC_Stock_Er  (PS_Code, PS_Count, PS_Sun)  Values ('" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "', '" & Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value & "', '" & Db_Sun & "' )"
        Re_Count = DB_Execute()
        Grid_Error_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value, Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True



    End Sub

    Private Sub Grid_Error_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Error.CellContentClick

    End Sub

    Private Sub Grid_Error_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Error.CellValueChanged
        If Grid_Display_Ch = True Then
            Grid_Error.CurrentRow.HeaderCell.Value = "*"
        End If

    End Sub

    Private Sub Delete__Main_Click(sender As Object, e As EventArgs) Handles Delete__Main.Click
        '해당 칼럼 삭제
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        If Grid_Info.Item(1, 10).Value = "처리" Then
            MsgBox("수량 처리가 되어서 수정 할 수 없습니다.")
            Exit Sub
        End If

        If MsgBox("순번 '" & Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value & "'을(를) 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "삭제") = vbNo Then
            Exit Sub
        End If

        Grid_Display_Ch = False
        Dim i As Integer
        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Delete PC_Stock_List Where PS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And  PS_Sun = '" & Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()
        '불량유형
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Delete PC_Stock_Er Where PS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And  PS_Count = '" & Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value & "' "
        Re_Count = DB_Execute()

        For i = Grid_Main.CurrentCell.RowIndex + 1 To Grid_Main.RowCount - 1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "Update  PC_Stock_List Set PS_Sun = '" & i & "' Where PS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And  PS_Sun = '" & i + 1 & "'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        Grid_Main_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        Dim i As Integer

        Grid_Display_Ch = False
        For i = 0 To Grid_Error.Rows.Count - 1
            If Grid_Error.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Update PC_Stock_Er Set PS_History = '" & Grid_Error.Item(1, i).Value & "', PS_Total = '" & Grid_Error.Item(2, i).Value & "', PS_Bigo = '" & Grid_Error.Item(3, i).Value & "' Where PS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And PS_Count = '" & Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value & "' And PS_Sun = '" & Grid_Error.Item(0, i).Value & "'" 'Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value, Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value
                Re_Count = DB_Execute()
                Grid_Error.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display_Ch = True

        MsgBox("삭제되었습니다")

    End Sub

    Private Sub Grid_Error_MouseClick(sender As Object, e As MouseEventArgs) Handles Grid_Error.MouseClick
        Grid_Main_Row = Grid_Error.CurrentCell.RowIndex
        Grid_Main_Col = Grid_Error.CurrentCell.ColumnIndex
        Grid_Error_Combo.Visible = False
        If Grid_Main_Row < 0 Then
            Exit Sub
        End If
        If Grid_Main_Col < 0 Then
            Exit Sub
        End If
        Dim DBT As New DataTable

        Select Case Grid_Main_Col
            Case 1
                StrSQL = "Select Base_Name From SI_Base_Code with(NOLOCK) Where Base_Sel = '불량유형' Order By Convert(Decimal,Base_Sun) "
                Re_Count = DB_Select(DBT)
                Grid_Error_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        Grid_Error_Combo.Items.Add(DBT.Rows(i).Item(0))
                    Next i
                End If
                Grid_Error_Combo.Top = Grid_Error.Top + Grid_Error.ColumnHeadersHeight + (Grid_Main_Row) * 24
                Grid_Error_Combo.Left = Grid_Error.Left + Grid_Error.RowHeadersWidth + Grid_Error.Columns(0).Width
                Grid_Error_Combo.Width = Grid_Error.Columns(Grid_Main_Col).Width
                Grid_Error_Combo.Text = Grid_Error.CurrentCell.Value.ToString
                Grid_Error_Combo.BackColor = Grid_Error.RowsDefaultCellStyle.BackColor
                Grid_Error_Combo.Visible = True
                Grid_Error_Combo.Focus.ToString()
        End Select

    End Sub

    Private Sub Grid_Error_Combo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Grid_Error_Combo.SelectedIndexChanged

    End Sub

    Private Sub Grid_Error_Combo_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Grid_Error_Combo.SelectionChangeCommitted
        Dim DBT As New DataTable
        Grid_Error.Item(Grid_Main_Col, Grid_Main_Row).Value = Grid_Error_Combo.SelectedItem.ToString()
        Grid_Error.CurrentRow.HeaderCell.Value = "*"
        Grid_Error_Combo.Visible = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '해당 칼럼 삭제
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        If Len(Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        If Len(Grid_Error.Item(0, Grid_Error.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        Grid_Display_Ch = False
        Dim i As Integer
        '삭제한다


        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Delete PC_Stock_Er Where PS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And  PS_Count = '" & Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value & "' And  PS_Sun = '" & Grid_Error.Item(0, Grid_Error.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        For i = Grid_Main.CurrentCell.RowIndex + 1 To Grid_Main.RowCount - 1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "Update  PC_Stock_Er Set PS_Sun = '" & i & "' Where PS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And  PS_Count = '" & Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value & "' And  PS_Sun = '" & i + 1 & "' "
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        Grid_Error_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value, Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True
    End Sub
    Private Sub Grid_Main_Combo_LostFocus(sender As Object, e As EventArgs) Handles Grid_Main_Combo.LostFocus
        Grid_Main_Combo.Visible = False
    End Sub

    Private Sub Delete_Com_Click(sender As Object, e As EventArgs) Handles Delete_Com.Click

        '수량 변경 한다.
        '수량 변경
        '수량이 변경 되었는지 확인 한다.

        '해당 칼럼 삭제
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        If MsgBox("작업실적  '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "'를 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "작업실적 삭제") = vbNo Then
            Exit Sub
        End If


        Dim Val_Check As String
        Dim PP_G_Sun As String
        Dim PP_G_Amount As String
        Dim DBT As DataTable = Nothing
        Dim i As Integer
        Dim j As Integer



        Val_Check = ""
        StrSQL = "Select PS_Check FROM PC_STOCK With(NOLOCK) Where PS_Code = '" & Grid_Info.Item(1, 0).Value & "'"
        Re_Count = DB_Select(DBT)
        Grid_Info_Combo.Items.Clear()
        If Re_Count <> 0 Then
            If DBT.Rows(0)("PS_Check").ToString = "처리" Then
                Val_Check = "처리"
            End If
        End If
        '수량 변경
        '기존 Data를 삭제 한다.
        'Grid의 수량을 가지고 온다.
        If Val_Check = "처리" Then
            For i = 0 To Grid_Main.RowCount - 1
                Count_Pro_M(Grid_Main.Item(1, i).Value, Grid_Info.Item(1, 5).Value, Grid_Main.Item(10, i).Value, Grid_Main.Item(11, i).Value, Grid_Main.Item(12, i).Value)
            Next i

        End If


        Grid_Display_Ch = False
        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Delete PC_Stock Where PS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' "
        Re_Count = DB_Execute()

        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Delete PC_Stock_List Where PS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' "
        Re_Count = DB_Execute()

        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Delete PC_Stock_Er Where PS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' "
        Re_Count = DB_Execute()

        'StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        'StrSQL = StrSQL & "Delete PC_Stock_Device_QC Where PS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' "
        'Re_Count = DB_Execute()


        Grid_Code_Display()
        Grid_Info_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True

        Grid_Main.RowCount = 0
        MsgBox("삭제되었습니다")
    End Sub
    Public Function Grid_Go_Setup() As Boolean

        Grid_Color(Grid_Go)
        Grid_Go.AllowUserToAddRows = False
        Grid_Go.RowTemplate.Height = 24


        '열의 갯수는 하나 더 많아야 함.
        Grid_Go.ColumnCount = 5
        Grid_Go.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Go.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Go.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Go.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Grid_Go.RowHeadersWidth = 40
        Grid_Go.Columns(0).Width = 60
        Grid_Go.Columns(1).Width = 150
        Grid_Go.Columns(2).Width = 100
        Grid_Go.Columns(3).Width = 100
        Grid_Go.Columns(0).HeaderText = "순번"
        Grid_Go.Columns(1).HeaderText = "점검항목"
        Grid_Go.Columns(2).HeaderText = "기준값"
        Grid_Go.Columns(3).HeaderText = "점검결과"
        Grid_Go.Columns(4).HeaderText = "비고"


        Grid_Line_QC.Columns(0).ReadOnly = True
        Grid_Go_Setup = True
    End Function

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        If Len(Grid_Info.Item(1, 5).Value) > 0 Then
        Else
            Exit Sub
        End If

        Label_Color(sender, Color_Label, Di_Panel2)
        '판넬 보이기
        Panel_Main.Visible = False
        Panel_Excel.Visible = False
        Panel_Go.Visible = True

        '전표에 저장된 내역이 있으면 저장된 Data 없으면 장비 Teabl
        Grid_Go_Display(Grid_Info.Item(1, 5).Value)

    End Sub
    Public Function Grid_Go_Display(PL_Code As String) As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display_Ch = False
        Grid_Go.RowCount = 0

        StrSQL = "Select PQ_Sun, PQ_History1, PQ_History2, PQ_Result, PQ_Bigo  FROM PC_STOCK_QC with(NOLOCK)  Where PQ_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And PQ_PL_Code = '" & Grid_Main.Item(1, Grid_Main.CurrentCell.RowIndex).Value & "' And  PQ_PC_Code  = '" & Grid_Info.Item(1, 5).Value & "' Order By Convert(Decimal,PQ_Sun)"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
        Else
            '추가 한다.
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "Insert into PC_Stock_QC ( PQ_Code, PQ_PL_Code, PQ_PC_Code, PQ_Sun, PQ_History1, PQ_History2, PQ_Result, PQ_Bigo  )  Select '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "', '" & Grid_Main.Item(1, Grid_Main.CurrentCell.RowIndex).Value & "' , PQ_PC_Code, PQ_Sun, PQ_History1, PQ_History2, '', PQ_Bigo  From SI_Product_QC with(NOLOCK) Where PQ_Code = '" & Grid_Main.Item(1, Grid_Main.CurrentCell.RowIndex).Value & "' And PQ_PC_Code = '" & Grid_Info.Item(1, 5).Value & "' "
            '                                   Select          , ' ,        PQ_PC_Code, PQ_Sun, PQ_History1,  PQ_History2, '', PQ_Bigo  From SI_Product_QC with(NOLOCK) Where PQ_Code = '" & Grid_Main.Item(1, Grid_Code.CurrentCell.RowIndex).Value & "' And PQ_PC_Code = '" & Grid_Info.Item(1, 5).Value & "' "
            Re_Count = DB_Execute()
        End If

        StrSQL = "Select PQ_Sun, PQ_History1, PQ_History2, PQ_Result, PQ_Bigo FROM PC_STOCK_QC with(NOLOCK)  Where PQ_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And PQ_PL_Code = '" & Grid_Main.Item(1, Grid_Main.CurrentCell.RowIndex).Value & "' And  PQ_PC_Code  = '" & Grid_Info.Item(1, 5).Value & "' Order By Convert(Decimal,PQ_Sun)"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Go.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To Grid_Go.ColumnCount - 1
                    Grid_Go.Item(j, i).Value = DBT.Rows(i).Item(j).ToString
                Next j
            Next i
        Else
            'Grid_Line_QC.RowCount = 1
            'For j = 0 To Grid_Line_QC.ColumnCount - 1
            '    Grid_Line_QC.Item(j, 0).Value = ""
            'Next j
        End If


        Grid_Display_Ch = True
        Grid_Go_Display = True

    End Function

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        Dim i As Integer

        Grid_Display_Ch = False
        For i = 0 To Grid_Go.Rows.Count - 1
            If Grid_Go.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "update PC_Stock_QC Set PQ_Result = '" & Grid_Go.Item(3, i).Value & "', PQ_Bigo = '" & Grid_Go.Item(4, i).Value & "' Where PQ_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And PQ_PL_Code = '" & Grid_Main.Item(1, Grid_Main.CurrentCell.RowIndex).Value & "' And PQ_PC_Code ='" & Grid_Info.Item(1, 5).Value & "' And PQ_Sun = '" & Grid_Go.Item(0, i).Value & "'"
                Re_Count = DB_Execute()
                Grid_Go.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display_Ch = True


    End Sub

    Private Sub Grid_Line_QC_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Line_QC.CellContentClick

    End Sub

    Private Sub Grid_Go_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Go.CellContentClick

    End Sub

    Private Sub Grid_Go_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Go.CellValueChanged
        If Grid_Display_Ch = True Then
            Grid_Go.CurrentRow.HeaderCell.Value = "*"
        End If
    End Sub

    Private Sub Grid_Line_QC_MouseClick(sender As Object, e As MouseEventArgs) Handles Grid_Line_QC.MouseClick
        '여기처리

        Grid_Line_QC_Row = Grid_Line_QC.CurrentCell.RowIndex
        Grid_Line_QC_Col = Grid_Line_QC.CurrentCell.ColumnIndex
        Grid_Line_QC_Combo.Visible = False
        If Grid_Line_QC_Row < 0 Then
            Exit Sub
        End If
        If Grid_Line_QC_Col < 0 Then
            Exit Sub
        End If
        Dim DBT As New DataTable

        Select Case Grid_Line_QC_Col
            Case 2
                StrSQL = "Select Base_Name From SI_Base_Code with(NOLOCK) Where Base_Sel = '장비점검' Order By Base_Name"
                Re_Count = DB_Select(DBT)
                Grid_Line_QC_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        Grid_Line_QC_Combo.Items.Add(DBT.Rows(i).Item(0))
                    Next i
                End If
                Grid_Line_QC_Combo.Top = Grid_Line_QC.Top + Grid_Line_QC.ColumnHeadersHeight + (Grid_Line_QC_Row) * 24
                Grid_Line_QC_Combo.Left = Grid_Line_QC.Left + Grid_Line_QC.RowHeadersWidth + Grid_Line_QC.Columns(0).Width + Grid_Line_QC.Columns(1).Width
                Grid_Line_QC_Combo.Width = Grid_Line_QC.Columns(Grid_Line_QC_Col).Width
                Grid_Line_QC_Combo.Text = Grid_Line_QC.CurrentCell.Value.ToString
                Grid_Line_QC_Combo.BackColor = Grid_Line_QC.RowsDefaultCellStyle.BackColor
                Grid_Line_QC_Combo.Visible = True
                Grid_Line_QC_Combo.Focus.ToString()
        End Select


    End Sub

    Private Sub Grid_Line_QC_Combo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Grid_Line_QC_Combo.SelectedIndexChanged

    End Sub

    Private Sub Grid_Line_QC_Combo_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Grid_Line_QC_Combo.SelectionChangeCommitted
        Dim DBT As New DataTable
        Grid_Line_QC.Item(Grid_Line_QC_Col, Grid_Line_QC_Row).Value = Grid_Line_QC_Combo.SelectedItem.ToString()
        Grid_Line_QC.CurrentRow.HeaderCell.Value = "*"
        Grid_Line_QC_Combo.Visible = False
    End Sub

    Private Sub Grid_Go_MouseClick(sender As Object, e As MouseEventArgs) Handles Grid_Go.MouseClick
        '여기처리

        Grid_Go_Row = Grid_Go.CurrentCell.RowIndex
        Grid_Go_Col = Grid_Go.CurrentCell.ColumnIndex
        Grid_Go_Combo.Visible = False
        If Grid_Go_Row < 0 Then
            Exit Sub
        End If
        If Grid_Go_Col < 0 Then
            Exit Sub
        End If
        Dim DBT As New DataTable

        Select Case Grid_Go_Col
            Case 3
                StrSQL = "Select Base_Name From SI_Base_Code with(NOLOCK) Where Base_Sel = '장비점검' Order By Base_Name"
                Re_Count = DB_Select(DBT)
                Grid_Go_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        Grid_Go_Combo.Items.Add(DBT.Rows(i).Item(0))
                    Next i
                End If
                Grid_Go_Combo.Top = Grid_Go.Top + Grid_Go.ColumnHeadersHeight + (Grid_Go_Row) * 24
                Grid_Go_Combo.Left = Grid_Go.Left + Grid_Go.RowHeadersWidth + Grid_Go.Columns(0).Width + Grid_Go.Columns(1).Width + Grid_Go.Columns(2).Width
                Grid_Go_Combo.Width = Grid_Go.Columns(Grid_Go_Col).Width
                Grid_Go_Combo.Text = Grid_Go.CurrentCell.Value.ToString
                Grid_Go_Combo.BackColor = Grid_Go.RowsDefaultCellStyle.BackColor
                Grid_Go_Combo.Visible = True
                Grid_Go_Combo.Focus.ToString()
        End Select

    End Sub

    Private Sub Grid_Go_Combo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Grid_Go_Combo.SelectedIndexChanged

    End Sub

    Private Sub Grid_Go_Combo_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Grid_Go_Combo.SelectionChangeCommitted
        Dim DBT As New DataTable
        Grid_Go.Item(Grid_Go_Col, Grid_Go_Row).Value = Grid_Go_Combo.SelectedItem.ToString()
        Grid_Go.CurrentRow.HeaderCell.Value = "*"
        Grid_Go_Combo.Visible = False
    End Sub

    Private Sub Info_Co0_MouseClick(sender As Object, e As MouseEventArgs) Handles Info_Co4.MouseClick, Info_Co5.MouseClick, Info_Co6.MouseClick, Info_Co7.MouseClick, Info_Co8.MouseClick
        Dim index As Integer

        If Grid_Display_Ch = False Then
            Exit Sub
        Else
            'MsgBox(Mid(sender.name.ToString, 8, 2))
        End If

        index = Val(Mid(sender.name.ToString, 8, 2))
        Select Case index
            Case 4
                StrSQL = "Select PO_Code  FROM PC_ORDER With(NOLOCK) WHERE PO_PR_CODE IS NOT NULL Order By PO_Code Desc"
            Case 5
                StrSQL = "Select PC_Code  From SI_Process With(NOLOCK) Order By PC_Code"
            Case 6
                StrSQL = "Select PC_Name  From SI_Process With(NOLOCK) Order By PC_Name"
            Case 7
                StrSQL = "Select PL_Code  From SI_Machine With(NOLOCK), SI_Process_Device With(NOLOCK) Where SI_Process_Device.PC_Code = '" & Grid_Info.Item(1, 5).Value.ToString & "' And SI_Process_Device.PC_PL_Code = SI_Machine.PL_Code  Order By Convert(Decimal,SI_Process_Device.PC_Sun) Desc "
            Case 8
                StrSQL = "Select PL_Name  From SI_Machine With(NOLOCK), SI_Process_Device With(NOLOCK) Where SI_Process_Device.PC_Code = '" & Grid_Info.Item(1, 5).Value & "' And SI_Process_Device.PC_PL_Code = SI_Machine.PL_Code  Order By Convert(Decimal,SI_Process_Device.PC_Sun) Desc "
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

    Private Sub Info_Tx0_TextChanged(sender As Object, e As EventArgs) Handles Info_Tx0.TextChanged, Info_Tx1.TextChanged, Info_Tx2.TextChanged, Info_Tx3.TextChanged, Info_Tx4.TextChanged, Info_Tx5.TextChanged, Info_Tx6.TextChanged, Info_Tx7.TextChanged, Info_Tx8.TextChanged, Info_Tx9.TextChanged, Info_Tx10.TextChanged, Info_Tx11.TextChanged
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

    Private Sub Info_Co0_SelectedValueChanged(sender As Object, e As EventArgs) Handles Info_Co4.SelectedValueChanged, Info_Co4.SelectedValueChanged, Info_Co5.SelectedValueChanged, Info_Co6.SelectedValueChanged, Info_Co7.SelectedValueChanged, Info_Co8.SelectedValueChanged
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
                StrSQL = "Select PO_CODE, PO_PR_CODE, PO_PR_NAME, PO_DE_CODE, PO_DE_NAME FROM PC_ORDER With(NOLOCK) Where PO_Code = '" & sender.text.ToString() & "' AND PO_PR_CODE IS NOT NULL"
                Re_Count = DB_Select(DBT)
                Grid_Info_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    Info_Txt(4).Text = DBT.Rows(0)("PO_CODE")
                    Info_Txt(5).Text = DBT.Rows(0)("PO_PR_CODE")
                    Info_Txt(6).Text = DBT.Rows(0)("PO_PR_NAME")
                    Info_Txt(7).Text = DBT.Rows(0)("PO_DE_CODE")
                    Info_Txt(8).Text = DBT.Rows(0)("PO_DE_NAME")
                    ' Info_Txt(7).Text = DBT.Rows(0)("CM_Man_Tel")
                    Grid_Info.Rows(4).HeaderCell.Value = "*"
                    Grid_Info.Rows(5).HeaderCell.Value = "*"
                    Grid_Info.Rows(6).HeaderCell.Value = "*"
                    Grid_Info.Rows(7).HeaderCell.Value = "*"
                    Grid_Info.Rows(8).HeaderCell.Value = "*"
                    ' Grid_Info.Rows(6).HeaderCell.Value = "*"
                    '  Grid_Info.Rows(7).HeaderCell.Value = "*"
                End If
            Case 5
                StrSQL = "Select PC_Code, PC_Name From SI_Process With(NOLOCK) Where PC_Code = '" & sender.text.ToString() & "'"
                Re_Count = DB_Select(DBT)
                Grid_Info_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    Info_Txt(5).Text = DBT.Rows(0)("PC_Code")
                    Info_Txt(6).Text = DBT.Rows(0)("PC_Name")
                    ' Info_Txt(7).Text = DBT.Rows(0)("CM_Man_Tel")
                    Grid_Info.Rows(5).HeaderCell.Value = "*"
                    Grid_Info.Rows(6).HeaderCell.Value = "*"
                    ' Grid_Info.Rows(6).HeaderCell.Value = "*"
                    '  Grid_Info.Rows(7).HeaderCell.Value = "*"
                End If

            Case 6
                StrSQL = "Select PC_Code, PC_Name From SI_Process With(NOLOCK) Where PC_Name = '" & sender.text.ToString() & "'"
                Re_Count = DB_Select(DBT)
                Grid_Info_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    Info_Txt(5).Text = DBT.Rows(0)("PC_Code")
                    Info_Txt(6).Text = DBT.Rows(0)("PC_Name")
                    ' Info_Txt(7).Text = DBT.Rows(0)("CM_Man_Tel")
                    Grid_Info.Rows(5).HeaderCell.Value = "*"
                    Grid_Info.Rows(6).HeaderCell.Value = "*"
                    ' Grid_Info.Rows(6).HeaderCell.Value = "*"
                    '  Grid_Info.Rows(7).HeaderCell.Value = "*"
                End If

            Case 7
                StrSQL = "Select PL_Code, PL_Name From SI_Machine With(NOLOCK) Where PL_Code = '" & sender.text.ToString() & "'"
                Re_Count = DB_Select(DBT)
                Grid_Info_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    Info_Txt(7).Text = DBT.Rows(0)("PL_Code")
                    Info_Txt(8).Text = DBT.Rows(0)("PL_Name")
                    '  Info_Txt(7).Text = DBT.Rows(0)("CM_Man_Tel")
                    '  Grid_Info.Rows(4).HeaderCell.Value = "*"
                    '  Grid_Info.Rows(5).HeaderCell.Value = "*"
                    Grid_Info.Rows(7).HeaderCell.Value = "*"
                    Grid_Info.Rows(8).HeaderCell.Value = "*"
                End If
            Case 8
                StrSQL = "Select PL_Code, PL_Name From SI_Machine With(NOLOCK) Where PL_Name = '" & sender.text.ToString() & "'"
                Re_Count = DB_Select(DBT)
                Grid_Info_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    Info_Txt(7).Text = DBT.Rows(0)("PL_Code")
                    Info_Txt(8).Text = DBT.Rows(0)("PL_Name")
                    '  Info_Txt(7).Text = DBT.Rows(0)("CM_Man_Tel")
                    '  Grid_Info.Rows(4).HeaderCell.Value = "*"
                    '  Grid_Info.Rows(5).HeaderCell.Value = "*"
                    Grid_Info.Rows(7).HeaderCell.Value = "*"
                    Grid_Info.Rows(8).HeaderCell.Value = "*"
                End If

        End Select

    End Sub

    Private Sub Panel_Menu_Paint(sender As Object, e As PaintEventArgs) Handles Panel_Menu.Paint

    End Sub
End Class
