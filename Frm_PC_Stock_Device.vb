﻿Imports System.ComponentModel

Public Class Frm_PC_Stock_Device
    Dim D_Code As String
    Dim Grid_Display_Ch As Boolean


    Dim Grid_Line_QC_Row As Integer
    Dim Grid_Line_QC_Col As Integer
    Dim Grid_Go_Row As Integer
    Dim Grid_Go_Col As Integer


    Private Sub Frm_PC_Stock_Device_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        End
    End Sub

    Private Sub Frm_PC_Stock_Device_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        FileOpen(1, Application.StartupPath + "\Setup.txt", OpenMode.Input)
        Setup_Data(1) = LineInput(1)

        FileClose(1)
        D_Code = Mid(Setup_Data(1), 8, 3)
        '장비 이름

        Dim i As Integer
        Dim j As Integer
        Dim DBT As New DataTable
        If D_Code = "1" Then
            Label3.Text = "동영산업 생산 관리"
        Else

            StrSQL = "Select * FROM SI_MACHINE with(NOLOCK) Where PL_Code = '" & D_Code & "'"
            Re_Count = DB_Select(DBT)
            If Re_Count <> 0 Then
                Me.Text = D_Code & " " & DBT.Rows(0).Item(1)
                Label3.Text = DBT.Rows(0).Item(1)
            Else
                MsgBox("장비 코드를 확인 하세요.")
                End
            End If

        End If



        Panel_Excel.Visible = False
        Panel_Excel.Top = 70
        Panel_Excel.Left = 70
        Process_Line_QC_Setup()


        Panel_Go.Visible = False
        Panel_Go.Top = 70
        Panel_Go.Left = 70
        Grid_Go_Setup()

        'Panel3.
        Panel3.Visible = False
        Panel3.Top = 70
        Panel3.Left = 70
        Panel6.Visible = False
        Panel6.Top = 70
        Panel6.Left = Panel3.Left + Panel3.Width


        Grid_Main_Setup()
        Grid_Main_Display()

        'Label1.Visible = False

        Grid_Main.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        Grid_Main.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells


    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Day_Label.Text = Format(Now, "yyyy년 MM월 dd일")
        Time_Label.Text = Format(Now, "HH시 mm분 ss초")






        ''생산수량
        'If Grid_Main.Rows.Count < 1 Then
        '    Exit Sub
        'End If

        'Dim DBT As New DataTable
        'Dim i As Integer
        'Dim j As Long
        'For i = 0 To Grid_Main.Rows.Count - 1
        '    '생산수량
        '    If Grid_Main.Item(11, i).Value = "종료" Then
        '        j = 0
        '        StrSQL = "Select Sum(Convert(Decimal,PS_Go)) As E1, Sum(Convert(Decimal,PS_Total)) As E2, Sum(Convert(Decimal,PS_Er)) As E3 FROM PC_STOCK with(NOLOCK), PC_Stock_List with(NOLOCK)  Where PC_Stock.PS_PO_Code  = '" & Grid_Main.Item(1, i).Value & "' And  PC_Stock.PS_Code = PC_STOCK_LIST.PS_Code And  PS_PL_Code =  '" & Grid_Main.Item(2, i).Value & "'  And len(PS_Go) > 0 And PS_Go is Not null And len(PS_Total) > 0 And PS_Total is Not null And len(PS_Er) > 0 And PS_Er is Not null"
        '        Re_Count = DB_Select(DBT)
        '        If Re_Count <> 0 Then
        '            j = Val(DBT.Rows(0)("E1").ToString)
        '        End If

        '        StrSQL = "Select PL_Go FROM IOT with(NOLOCK) Where PL_Code = '" & D_Code & "' "
        '        Re_Count = DB_Select(DBT)
        '        If Re_Count <> 0 Then
        '            If Len(Grid_Main.Item(9, i).Value.ToString) > 0 Then
        '            Else
        '                Grid_Main.Item(8, i).Value = j + DBT.Rows(0)("PL_Go")
        '            End If
        '        End If
        '    End If

        'Next i








    End Sub
    Public Function Grid_Main_Setup() As Boolean


        ' Grid_Color(Grid_Main)

        Grid_Main.GridColor = Color.White
        Grid_Main.GridColor = Color.White
        Grid_Main.BackgroundColor = Color.White
        Grid_Main.EnableHeadersVisualStyles = False
        Grid_Main.ColumnHeadersDefaultCellStyle.BackColor = Color.DarkOrange
        Grid_Main.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        Grid_Main.RowsDefaultCellStyle.BackColor = Color.FromArgb(210, 222, 239)
        Grid_Main.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(234, 239, 247)

        'Grid_Main.AllowUserToAddRows = False
        'Grid_Main.RowTemplate.Height = 96
        Grid_Main.ColumnHeadersHeight = 100
        Grid_Main.RowHeadersVisible = False



        Grid_Main.ColumnCount = 12


        Grid_Main.RowCount = 0
        Dim i As Integer

        For i = 0 To Grid_Main.ColumnCount - 1
            Grid_Main.Columns(i).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Next



        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Main.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Main.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Main.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Grid_Main.Columns(0).HeaderText = "순번"
        Grid_Main.Columns(1).HeaderText = "생산지시코드"
        Grid_Main.Columns(1).Visible = False
        Grid_Main.Columns(2).HeaderText = "제품코드"
        Grid_Main.Columns(2).Visible = False

        Grid_Main.Columns(3).HeaderText = "제품명"
        Grid_Main.Columns(4).HeaderText = "규격"
        Grid_Main.Columns(4).Visible = False
        Grid_Main.Columns(5).HeaderText = "지시수량"
        Grid_Main.Columns(6).HeaderText = "시작"
        Grid_Main.Columns(6).Visible = False

        Grid_Main.Columns(7).HeaderText = "완료"
        Grid_Main.Columns(7).Visible = False
        Grid_Main.Columns(8).HeaderText = "투입량"
        Grid_Main.Columns(9).HeaderText = "생산량"
        Grid_Main.Columns(10).HeaderText = "불량"
        Grid_Main.Columns(11).HeaderText = "구분"




        Grid_Main.Columns(0).Width = 50
        Grid_Main.Columns(1).Width = 350

        For i = 2 To 7
            Grid_Main.Columns(i).Width = 130
        Next i

        'Grid_Main.Columns(1).Width = 220
        'Grid_Main.Columns(2).Width = 200
        'Grid_Main.Columns(3).Width = 200
        'Grid_Main.Columns(4).Width = 200
        'Grid_Main.Columns(5).Width = 170
        'Grid_Main.Columns(6).Width = 170

        For i = 8 To 10
            Grid_Main.Columns(i).Width = 130
        Next i





        Grid_Main.Columns(11).Width = 150

        Grid_Main_Setup = True

        For i = 0 To Grid_Main.ColumnCount - 1
            Grid_Main.Columns(i).ReadOnly = True
        Next i


    End Function
    Public Function Grid_Main_Display() As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Main.RowCount = 0
        If D_Code = "1" Then
            StrSQL = "Select  Max(PC_ORDER.PO_Code), PO_PL_Code, PO_PL_Name, PO_Standard,  Sum(Convert(Decimal,PO_Total)) As E1  FROM PC_ORDER_LIST with(NOLOCK), PC_ORDER WITH(NOLOCK)  Where PC_ORDER.PO_Code = PC_ORDER_LIST.PO_Code  And PO_Day = '" & Format(Now, "yyyy-MM-dd") & "' And Len(PO_Total) > 0   GROUP BY PC_ORDER.PO_Code, PO_PL_Code, PO_PL_Name, PO_Standard Order By PC_ORDER.PO_Code"
        Else
            StrSQL = "Select  Max(PC_ORDER.PO_Code), PO_PL_Code, PO_PL_Name, PO_Standard,  Sum(Convert(Decimal,PO_Total)) As E1  FROM PC_ORDER_LIST with(NOLOCK), PC_ORDER WITH(NOLOCK)  Where PC_ORDER.PO_Code = PC_ORDER_LIST.PO_Code And PO_DE_Code = '" & D_Code & "' And PO_Day = '" & Format(Now, "yyyy-MM-dd") & "' And Len(PO_Total) > 0   GROUP BY PC_ORDER.PO_Code, PO_PL_Code, PO_PL_Name, PO_Standard Order By PC_ORDER.PO_Code"
        End If
        'StrSQL = "Select  Max(PO.PO_Code), PO_PL_Code, PO_PL_Name, PO_Standard,  Sum(Convert(Decimal,PO_Total)) As E1  FROM PC_ORDER_LIST with(NOLOCK), PO with(NOLOCK)  Where PO.PO_Code = PC_ORDER_LIST.PO_Code And PO_DE_Code = '" & D_Code & "' And PO_Day = '" & Format(Now, "yyyy-MM-dd") & "' And Len(PO_Total) > 0   GROUP BY PO.PO_Code, PO_PL_Code, PO_PL_Name, PO_Standard Order By PO.PO_Code"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Main.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                Grid_Main.Item(0, i).Value = i + 1
                Grid_Main.Item(1, i).Value = DBT.Rows(i).Item(0).ToString
                Grid_Main.Item(2, i).Value = DBT.Rows(i).Item(1).ToString
                Grid_Main.Item(3, i).Value = DBT.Rows(i).Item(2).ToString
                Grid_Main.Item(4, i).Value = DBT.Rows(i).Item(3).ToString
                Grid_Main.Item(5, i).Value = DBT.Rows(i).Item(4).ToString
                'Grid_Main.Item(11, i).Value = "시작"
            Next i
        Else
            ' MsgBox("설비와 관련된 작업지시서가 존재하지 않습니다")
            ' Me.Close()
            Grid_Main.RowCount = 1
            For j = 0 To Grid_Main.ColumnCount - 1
                Grid_Main.Item(j, 0).Value = ""
            Next j

        End If
        '시작 표시







        Dim 작업지시서 As String = ""

        For i = 0 To Grid_Main.Rows.Count - 1
            'IOT db에 따라 시작/종료시간 적어주기
            StrSQL = " select ps_po_code FROM IOT with(nolock) join PC_STOCK on PC_Stock.PS_Code = IOT.PL_PS_CODE WHERE PL_NAME ='" & Label3.Text & "' AND PL_PL_CODE='" & Grid_Main.Item(2, i).Value & "'"
            Re_Count = DB_Select(DBT)
            If Re_Count <> 0 Then
                작업지시서 = DBT.Rows(0)("ps_po_code").ToString
            End If

            Dim IOT코드 As String

            StrSQL = "Select PL_St_Time, PL_PS_CODE FROM IOT with(NOLOCK) join PC_STOCK on PC_Stock.PS_Code = IOT.PL_PS_CODE join PC_STOCK_LIST on PC_STOCK_LIST.PS_Code = IOT.PL_PS_CODE  Where PL_PL_CODE = '" & Grid_Main.Item(2, i).Value & "' And Len(PL_St_Day) > 0 And PL_Ing = 'ing' AND PS_PO_CODE = '" & 작업지시서 & "' AND  PS_PO_CODE = '" & Grid_Main.Item(1, i).Value & "' and PL_PL_NAME ='" & Grid_Main.Item(3, i).Value & "' and PS_STANDARD ='" & Grid_Main.Item(4, i).Value & "'"
            Re_Count = DB_Select(DBT)

            If Re_Count <> 0 Then
                '시작시간
                Grid_Main.Item(6, i).Value = DBT.Rows(0)("PL_St_Time")
                IOT코드 = DBT.Rows(0)("PL_PS_CODE")
                Grid_Main.Item(11, i).Value = "종료"


                '생산수량
                StrSQL = "Select PL_Go FROM IOT with(NOLOCK) Where PL_Code  = '" & D_Code & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Main.Item(8, i).Value = DBT.Rows(0)("PL_Go").ToString

                End If
            Else

                StrSQL = "Select Sum(Convert(Decimal,PS_Go)) As E1, Sum(Convert(Decimal,PS_Total)) As E2, Sum(Convert(Decimal,PS_Er)) As E3, Min(PS_St_Time) As E4, Max(PS_En_Time) As E5  FROM PC_STOCK with(NOLOCK), PC_Stock_List with(NOLOCK)  Where PC_Stock.PS_PO_Code  = '" & Grid_Main.Item(1, i).Value & "' And PS_STANDARD = '" & Grid_Main.Item(4, i).Value & "' and  PC_Stock.PS_Code = PC_STOCK_LIST.PS_Code And  PS_PL_Code =  '" & Grid_Main.Item(2, i).Value & "'  And len(PS_Go) > 0 And PS_Go is Not null And len(PS_Total) > 0 And PS_Total is Not null And len(PS_Er) > 0 And PS_Er is Not null"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    If Len(DBT.Rows(0)("E4").ToString) > 0 Then
                        Grid_Main.Item(6, i).Value = DBT.Rows(0)("E4")
                        Grid_Main.Item(7, i).Value = DBT.Rows(0)("E5")
                        Grid_Main.Item(8, i).Value = DBT.Rows(0)("E1")
                        Grid_Main.Item(9, i).Value = DBT.Rows(0)("E2")
                        Grid_Main.Item(10, i).Value = DBT.Rows(0)("E3")
                        Grid_Main.Item(11, i).Value = "완료"
                    Else
                        Grid_Main.Item(6, i).Value = ""
                        Grid_Main.Item(7, i).Value = ""
                        Grid_Main.Item(11, i).Value = "시작"
                        Grid_Main.Item(8, i).Value = ""
                        Grid_Main.Item(9, i).Value = ""
                        Grid_Main.Item(10, i).Value = ""
                    End If
                    '종료시간 검색
                Else
                    Grid_Main.Item(6, i).Value = ""
                    Grid_Main.Item(7, i).Value = ""
                    Grid_Main.Item(11, i).Value = "시작"
                    Grid_Main.Item(8, i).Value = ""
                    Grid_Main.Item(9, i).Value = ""
                    Grid_Main.Item(10, i).Value = ""

                End If
            End If
        Next i



        ''현까지 생산 수량 표시
        'For i = 0 To Grid_Main.Rows.Count - 1
        '    '생산수량
        '    StrSQL = "Select Sum(Convert(Decimal,PS_Go)) As E1, Sum(Convert(Decimal,PS_Total)) As E2, Sum(Convert(Decimal,PS_Er)) As E3 FROM PC_STOCK with(NOLOCK), PC_Stock_List with(NOLOCK)  Where PC_Stock.PS_PO_Code  = '" & Grid_Main.Item(1, i).Value & "' And  PC_Stock.PS_Code = PC_STOCK_LIST.PS_Code And  PS_PL_Code =  '" & Grid_Main.Item(2, i).Value & "'  And len(PS_Go) > 0 And PS_Go is Not null And len(PS_Total) > 0 And PS_Total is Not null And len(PS_Er) > 0 And PS_Er is Not null"
        '    Re_Count = DB_Select(DBT)
        '    If Re_Count <> 0 Then
        '        Grid_Main.Item(8, i).Value = DBT.Rows(0)("E1")
        '        Grid_Main.Item(9, i).Value = DBT.Rows(0)("E2")
        '        Grid_Main.Item(10, i).Value = DBT.Rows(0)("E3")
        '    End If
        'Next i




        Grid_Main_Display = True
        Grid_Main.ClearSelection()

    End Function

    Private Sub Grid_Main_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Main.CellContentClick

    End Sub

    Private Sub Grid_Main_MouseClick(sender As Object, e As MouseEventArgs) Handles Grid_Main.MouseClick

    End Sub

    Private Sub Grid_Main_MouseDown(sender As Object, e As MouseEventArgs) Handles Grid_Main.MouseDown

    End Sub
    Public Function Process_Line_QC_Setup() As Boolean

        '  Grid_Color(Grid_Line_QC)


        Grid_Line_QC.GridColor = Color.White
        Grid_Line_QC.GridColor = Color.White
        Grid_Line_QC.BackgroundColor = Color.White
        Grid_Line_QC.EnableHeadersVisualStyles = False
        Grid_Line_QC.ColumnHeadersDefaultCellStyle.BackColor = Color.DarkOrange
        Grid_Line_QC.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        Grid_Line_QC.RowsDefaultCellStyle.BackColor = Color.FromArgb(210, 222, 239)
        Grid_Line_QC.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(234, 239, 247)




        Grid_Line_QC.ColumnCount = 4
        Grid_Line_QC.RowCount = 1

        Grid_Line_QC.AllowUserToAddRows = False
        Grid_Line_QC.ColumnHeadersHeight = 100
        Grid_Line_QC.RowTemplate.Height = 100
        Grid_Line_QC.RowHeadersVisible = False


        Grid_Line_QC.RowCount = 0
        Dim i As Integer
        For i = 0 To Grid_Line_QC.ColumnCount - 1
            Grid_Line_QC.Columns(i).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Next



        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Line_QC.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Line_QC.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Line_QC.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft


        Grid_Line_QC.Columns(0).Width = 50
        Grid_Line_QC.Columns(1).Width = 230
        Grid_Line_QC.Columns(2).Width = 250
        Grid_Line_QC.Columns(3).Width = 200
        Grid_Line_QC.Columns(0).HeaderText = "순번"
        Grid_Line_QC.Columns(1).HeaderText = "점검내역"
        Grid_Line_QC.Columns(2).HeaderText = "점검결과"
        Grid_Line_QC.Columns(3).HeaderText = "비고"

        Grid_Line_QC.Columns(0).ReadOnly = True
        Grid_Line_QC.Columns(1).ReadOnly = True
        Process_Line_QC_Setup = True

    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim i As Integer

        Grid_Display_Ch = False
        For i = 0 To Grid_Line_QC.Rows.Count - 1
            If Grid_Line_QC.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "update PC_STOCK_DEVICE_QC Set PS_Result = '" & Grid_Line_QC.Item(2, i).Value & "', PS_Bigo = '" & Grid_Line_QC.Item(3, i).Value & "' where PS_Device = '" & D_Code & "' And PS_Day = '" & Format(Now, "yyyy-MM-dd") & "' And PS_Sun = '" & Grid_Line_QC.Item(0, i).Value & "'"
                Re_Count = DB_Execute()
                Grid_Line_QC.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display_Ch = True
        Panel_Excel.Visible = False

    End Sub
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
            StrSQL = StrSQL & "Insert into PC_Stock_Device_QC (PS_Device, PS_Day, PS_Sun, PS_History, PS_Result, PS_Bigo  )  Select '" & PL_Code & "', '" & Format(Now, "yyyy-MM-dd") & "' , PQ_Sun , PQ_History, '',PQ_Bigo FROM SI_MACHINE_QC with(NOLOCK) Where PQ_Code = '" & PL_Code & "' "
            Re_Count = DB_Execute()
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
        End If


        Grid_Display_Ch = True
        Grid_Line_QC_Display = True

    End Function

    Public Function Grid_Go_Setup() As Boolean
        '  Grid_Color(Grid_Go)

        Grid_Go.GridColor = Color.White
        Grid_Go.GridColor = Color.White
        Grid_Go.BackgroundColor = Color.White
        Grid_Go.EnableHeadersVisualStyles = False
        Grid_Go.ColumnHeadersDefaultCellStyle.BackColor = Color.DarkOrange
        Grid_Go.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        Grid_Go.RowsDefaultCellStyle.BackColor = Color.FromArgb(210, 222, 239)
        Grid_Go.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(234, 239, 247)



        Grid_Go.AllowUserToAddRows = False
        Grid_Go.RowTemplate.Height = 24



        Grid_Go.AllowUserToAddRows = False
        Grid_Go.RowHeadersVisible = False
        Grid_Go.ColumnHeadersHeight = 100
        Grid_Go.RowTemplate.Height = 100

        Grid_Go.ColumnCount = 5
        Grid_Go.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Go.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Go.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Go.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Grid_Go.RowHeadersWidth = 40
        Grid_Go.Columns(0).Width = 50
        Grid_Go.Columns(1).Width = 200
        Grid_Go.Columns(2).Width = 200
        Grid_Go.Columns(3).Width = 200
        Grid_Go.Columns(4).Width = 200
        Grid_Go.Columns(0).HeaderText = "순번"
        Grid_Go.Columns(1).HeaderText = "점검항목"
        Grid_Go.Columns(2).HeaderText = "기준값"
        Grid_Go.Columns(3).HeaderText = "점검결과"
        Grid_Go.Columns(4).HeaderText = "비고"


        Grid_Line_QC.Columns(0).ReadOnly = True
        Grid_Go_Setup = True
    End Function

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim i As Integer
        Dim 공정코드 As String
        Dim 공정명 As String

        Dim DBT As New DataTable

        StrSQL = "Select * FROM PC_ORDER with(NOLOCK) Where PO_Code = '" & Grid_Main.Item(1, Grid_Main.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            공정코드 = DBT.Rows(0)("PO_PR_Code")
            공정명 = DBT.Rows(0)("PO_PR_Name")
        End If


        Grid_Display_Ch = False
        For i = 0 To Grid_Go.Rows.Count - 1
            If Grid_Go.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "update PC_STOCK_QC Set PQ_Result = '" & Grid_Go.Item(3, i).Value & "', PQ_Bigo = '" & Grid_Go.Item(4, i).Value & "' Where PQ_Code = '" & Grid_Main.Item(1, Grid_Main.CurrentCell.RowIndex).Value & "' And PQ_PL_Code = '" & Grid_Main.Item(2, Grid_Main.CurrentCell.RowIndex).Value & "' And PQ_PC_Code ='" & 공정코드 & "' And PQ_Sun = '" & Grid_Go.Item(0, i).Value & "'"
                Re_Count = DB_Execute()
                Grid_Go.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display_Ch = True
        Panel_Go.Visible = False
    End Sub
    Public Function Grid_Go_Display(PL_Code As String, P_Code As String, Go_Code As String) As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display_Ch = False
        Grid_Go.RowCount = 0
        StrSQL = "Select PQ_Sun, PQ_History1, PQ_History2, PQ_Result, PQ_Bigo  FROM PC_STOCK_QC with(NOLOCK)  Where PQ_Code = '" & PL_Code & "' And PQ_PL_Code = '" & P_Code & "' And  PQ_PC_Code  = '" & Go_Code & "' Order By Convert(Decimal,PQ_Sun)"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
        Else
            '추가 한다.
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "Insert into PC_Stock_QC ( PQ_Code, PQ_PL_Code, PQ_PC_Code, PQ_Sun, PQ_History1, PQ_History2, PQ_Result, PQ_Bigo  )  Select '" & PL_Code & "', '" & P_Code & "' , PQ_PC_Code, PQ_Sun, PQ_History1, PQ_History2, '', PQ_Bigo  FROM SI_PRODUCT_QC with(NOLOCK) Where PQ_Code = '" & P_Code & "' And PQ_PC_Code = '" & Go_Code & "' "
            Re_Count = DB_Execute()
        End If

        StrSQL = "Select PQ_Sun, PQ_History1, PQ_History2, PQ_Result, PQ_Bigo FROM PC_STOCK_QC with(NOLOCK)  Where PQ_Code = '" & PL_Code & "' And PQ_PL_Code = '" & P_Code & "' And  PQ_PC_Code  = '" & Go_Code & "' Order By Convert(Decimal,PQ_Sun)"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Go.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To Grid_Go.ColumnCount - 1
                    Grid_Go.Item(j, i).Value = DBT.Rows(i).Item(j).ToString
                Next j
            Next i
        Else
        End If


        Grid_Display_Ch = True
        Grid_Go_Display = True

    End Function

    Private Sub Grid_Main_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles Grid_Main.MouseDoubleClick

        If Grid_Main.Item(0, 0).Value = Nothing Then
            Exit Sub
        End If


        Dim 지시서번호 As String
        Dim 실적번호 As String
        Dim 공정코드 As String
        Dim 공정명 As String
        Dim 규격 As String
        'Dim 공정코드 As String
        'Dim 공정코드 As String

        Dim JP_Code As Long
        Dim My_Date As String
        Dim My_Time As String



        Dim DBT As New DataTable
        If Grid_Main.CurrentCell.ColumnIndex = 11 Then
            지시서번호 = Grid_Main.Item(1, Grid_Main.CurrentCell.RowIndex).Value
            규격 = Grid_Main.Item(4, Grid_Main.CurrentCell.RowIndex).Value
            StrSQL = "Select * FROM PC_ORDER  with(NOLOCK) Where PO_Code = '" & 지시서번호 & "'"
            Re_Count = DB_Select(DBT)
            If Re_Count <> 0 Then
                공정코드 = DBT.Rows(0)("PO_PR_Code")
                공정명 = DBT.Rows(0)("PO_PR_Name")
            End If
            If Grid_Main.Item(11, Grid_Main.CurrentCell.RowIndex).Value = "시작" Then
                For i = 0 To Grid_Main.Rows.Count - 1
                    If Grid_Main.Item(11, i).Value.ToString = "종료" Then
                        MsgBox("끝나지 않은 작업이 있습니다. 확인 후 처리 하세요")
                        Exit Sub
                    End If
                Next i

                StrSQL = "Select PS_Sun, PS_History, PS_Result, PS_Bigo FROM PC_STOCK_Device_QC with(NOLOCK)  Where PS_Device = '" & D_Code & "' And PS_Day = '" & Format(Now, "yyyy-MM-dd") & "' Order By Convert(Decimal,PS_Sun)"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                Else
                    MsgBox("장비 점검 List 가 처리 되지 않았습니다. 확인 후 처리 하세요")
                    '장비 점검 List 처리

                    Panel_Excel.Visible = True
                    Grid_Line_QC_Display(D_Code)
                    Exit Sub
                End If
                'Select PQ_PC_Code, PQ_Sun, PQ_History1, PQ_History2,  PQ_Bigo  From PL_QC with(NOLOCK) Where PQ_PL_Code = '300-100' And PQ_PC_Code = ''

                Dim 현재작업지시서 As String = ""
                Dim 현재제품코드 As String = ""

                현재작업지시서 = Grid_Main.Item(1, Grid_Main.CurrentCell.RowIndex).Value
                현재제품코드 = Grid_Main.Item(2, Grid_Main.CurrentCell.RowIndex).Value

                StrSQL = "Select PQ_Sun, PQ_History1, PQ_History2, PQ_Result, PQ_Bigo  FROM PC_STOCK_QC with(NOLOCK)  Where PQ_Code = '" & Grid_Main.Item(1, Grid_Main.CurrentCell.RowIndex).Value & "' And PQ_PL_Code = '" & Grid_Main.Item(2, Grid_Main.CurrentCell.RowIndex).Value & "' And  PQ_PC_Code  = '" & 공정코드 & "' Order By Convert(Decimal,PQ_Sun)"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                Else
                    StrSQL = "Select PQ_PC_Code, PQ_Sun, PQ_History1, PQ_History2,  PQ_Bigo  FROM SI_PRODUCT_QC with(NOLOCK) Where PQ_Code = '" & Grid_Main.Item(2, Grid_Main.CurrentCell.RowIndex).Value & "' And PQ_PC_Code = '" & 공정코드 & "'"
                    Re_Count = DB_Select(DBT)
                    If Re_Count <> 0 Then
                        MsgBox("'" & Grid_Main.Item(1, Grid_Main.CurrentCell.RowIndex).Value & "' '" & Grid_Main.Item(2, Grid_Main.CurrentCell.RowIndex).Value & "' 공정 점검 List 가 처리 되지 않았습니다. 확인 후 처리 하세요")
                        Panel_Go.Visible = True
                        ' MsgBox((Grid_Main.Item(1, Grid_Main.CurrentCell.RowIndex).Value))
                        Grid_Go_Display(현재작업지시서, 현재제품코드, 공정코드)
                        ' Grid_Go_Display(Grid_Main.Item(1, Grid_Main.CurrentCell.RowIndex).Value, Grid_Main.Item(2, Grid_Main.CurrentCell.RowIndex).Value, 공정코드)
                        ' MsgBox((Grid_Main.Item(1, Grid_Main.CurrentCell.RowIndex).Value))
                        Exit Sub
                    Else
                    End If
                End If
                '시작 처리 한다.
                'IOT 처리 한다.
                'IOT 레코드가 있는가?
                StrSQL = "Select *  FROM IOT with(NOLOCK)  Where PL_Code = '" & D_Code & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                Else
                    '추가 한다.
                    StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                    ' StrSQL = StrSQL & "Insert INTO IOT (PL_Code, PL_Name, PL_Group, PL_ID, PL_Port, PL_Sel, PL_Bigo)  Select * FROM SI_MACHINE with(NOLOCK) Where PL_Code = '" & D_Code & "' "
                    StrSQL = StrSQL & "Insert INTO IOT (PL_Code, PL_Name, PL_Group, PL_ID, PL_Port, PL_Sel)  Select PL_Code, PL_NAME, PL_GROUP, PL_ID, PL_PORT, PL_SEL FROM SI_MACHINE with(NOLOCK) Where PL_Code = '" & D_Code & "' "

                    Re_Count = DB_Execute()
                End If
                '작업 지서서 번호


                '작업 실적 번호
                StrSQL = "Select GETDATE()"
                Re_Count = DB_Select(DBT)
                If Re_Count = 0 Then
                    Exit Sub
                Else
                    My_Date = DBT.Rows(0).Item(0)
                    JP_Code = Mid(My_Date, 1, 4) & Mid(My_Date, 6, 2) & Mid(My_Date, 9, 2)
                    'My_Time = Mid(My_Date, 12, 8)
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

                If Re_Count <= 0 Then
                    MsgBox("생산실적에서 오늘 날짜에 관한 전표가 하나 이상 있는지 확인하세요")
                    Exit Sub
                End If

                Dim zbun As String
                zbun = ""

                zbun = DBT.Rows(0).Item(0)
                실적번호 = "PS" & JP_Code

                '추가 및 변경 코드
                'PS_PO_Code와 PS_Code와 PS_Date를 일치시켜서 PC_ORDER_LIST의 순번을 실적번호 맨뒤에 '-순번'으로 붙이면 된다
                Dim sunbun As String
                sunbun = ""

                '순번가져옴
                StrSQL = "Select PO_Sun FROM PC_STOCK, PC_ORDER_LIST where PC_Stock.PS_PO_Code = PC_ORDER_LIST.PO_Code and  PS_Code = '" & zbun & "' order by PO_Sun"

                Re_Count = DB_Select(DBT)
                For i = 0 To Re_Count - 1
                    sunbun = "-" & DBT.Rows(i)("PO_Sun")
                Next i
                실적번호 = "PS" & JP_Code & sunbun


                '생산 실적에 추가한다
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert into PC_Stock (PS_Code, PS_Date, PS_Time, PS_Name, PS_Sel, PS_PO_Code,PS_PR_Code,PS_PR_Name,PS_DE_Code,PS_DE_Name, PS_Check)    Select '" & 실적번호 & "', '" & My_Date & "','" & My_Time & "','테블릿','현장', PO_Code, PO_PR_Code,PO_PR_Name,PO_DE_Code,PO_DE_Name,'처리' FROM PC_ORDER  with(NOLOCK) Where PO_Code = '" & 지시서번호 & "' "
                Re_Count = DB_Execute()

                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert into PC_Stock_List (PS_Code, PS_Sun, PS_PL_Code, PS_PL_Name, PS_Standard, PS_Unit, PS_Po_Total, PS_St_Day, PS_St_Time, PS_Total )   Select '" & 실적번호 & "', '1', PO_PL_Code, PO_PL_Name, PO_Standard, PO_Unit, PO_Total , '" & My_Date & "', '" & My_Time & "','0' FROM PC_ORDER_LIST  with(NOLOCK) Where PO_Code = '" & 지시서번호 & "' And  PO_PL_Code = '" & Grid_Main.Item(2, Grid_Main.CurrentCell.RowIndex).Value & "' and PO_STANDARD ='" & 규격 & "' "
                Re_Count = DB_Execute()


                Grid_Main.Item(6, Grid_Main.CurrentCell.RowIndex).Value = My_Time



                'Update 공정코드 공정명
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE IOT Set PL_Ing = 'ing', PL_PS_Code = '" & 실적번호 & "', PL_PL_CODE = '" & Grid_Main.Item(2, Grid_Main.CurrentCell.RowIndex).Value & "', PL_PL_Name = '" & Grid_Main.Item(3, Grid_Main.CurrentCell.RowIndex).Value & "', PL_STANDARD='" & 규격 & "', PL_PC_Code = '" & 공정코드 & "', PL_PC_Name = '" & 공정명 & "', PL_GO = '0', PL_Total = '0', PL_Er = '0' where PL_Code = '" & D_Code & "'"
                Re_Count = DB_Execute()


                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE IOT Set PL_St_Day = '" & My_Date & "', PL_St_Time = '" & My_Time & "',PL_En_Day = '', PL_En_Time = '' where PL_Code = '" & D_Code & "'"
                Re_Count = DB_Execute()

                ''  StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                'StrSQL = "SELECT PL_GO FROM IOT WHERE PL_PS_CODE 4"
                'Re_Count = DB_Execute()


                ''PC_STOCK_LIST total 업데이트(같이)
                'StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                'StrSQL = StrSQL & "update PC_STOCK_LIST Set PS_TOTAL = '&  &', PL_PS_Code = '" & 실적번호 & "', PL_CODE = '" & Grid_Main.Item(2, Grid_Main.CurrentCell.RowIndex).Value & "', PL_PL_Name = '" & Grid_Main.Item(3, Grid_Main.CurrentCell.RowIndex).Value & "', PL_PC_Code = '" & 공정코드 & "', PL_PC_Name = '" & 공정명 & "', PL_GO = '0', PL_Total = '0', PL_Er = '0'    where PL_Code = '" & D_Code & "'"
                'Re_Count = DB_Execute()


                ''시작 시간이 이미 입력 되어 있으면 처리 하지 않고 없으면 처리한다.
                'StrSQL = "Select * FROM IOT with(NOLOCK) where PL_Code = '" & D_Code & "'"
                'Re_Count = DB_Select(DBT)
                'If Re_Count <> 0 Then
                '    NullStr = DBT.Rows(0)("PL_St_Day").ToString
                '    NullStr1 = DBT.Rows(0)("PL_St_Time").ToString
                'End If
                'If Len(NullStr) < 1 Then
                '    StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                '    StrSQL = StrSQL & "UPDATE IOT Set PL_St_Day = '" & My_Date & "', PL_St_Time = '" & My_Time & "'   where PL_Code = '" & D_Code & "'"
                '    Re_Count = DB_Execute()

                'Else
                '    If NullStr = My_Date Then
                '        If NullStr1 < My_Time Then
                '        Else
                '            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                '            StrSQL = StrSQL & "UPDATE IOT Set PL_St_Day = '" & My_Date & "', PL_St_Time = '" & My_Time & "'   where PL_Code = '" & D_Code & "'"
                '            Re_Count = DB_Execute()
                '        End If
                '    Else
                '        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                '        StrSQL = StrSQL & "UPDATE IOT Set PL_St_Day = '" & My_Date & "', PL_St_Time = '" & My_Time & "'   where PL_Code = '" & D_Code & "'"
                '        Re_Count = DB_Execute()
                '    End If
                'End If



                Grid_Main.Item(11, Grid_Main.CurrentCell.RowIndex).Value = "종료"
                Exit Sub
            End If
            If Grid_Main.Item(11, Grid_Main.CurrentCell.RowIndex).Value = "종료" Then
                '투입 수량과 불량수량 완료 수량 Check
                Text1.Text = Grid_Main.Item(8, Grid_Main.CurrentCell.RowIndex).Value
                Text2.Text = Grid_Main.Item(9, Grid_Main.CurrentCell.RowIndex).Value
                Text3.Text = Grid_Main.Item(10, Grid_Main.CurrentCell.RowIndex).Value


                If Len(Text1.Text) < 1 Then
                    Text1.Text = "0"
                End If
                If Len(Text2.Text) < 1 Then
                    Text2.Text = "0"
                End If
                If Len(Text3.Text) < 1 Then
                    Text3.Text = "0"
                End If
                TextBox1.Text = Grid_Main.CurrentCell.RowIndex

                Timer1.Enabled = False
                Grid_Main.Enabled = False
                Panel3.Visible = True
                Panel6.Visible = True




            End If

        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '수량 Check
        If Val(Text1.Text) < Val(Text2.Text) Then
            MsgBox("수량을 확인 하세요")
            Exit Sub
        End If


        Dim 입력수량1 As String
        Dim 입력수량2 As String
        Dim 입력수량3 As String

        입력수량1 = Text1.Text
        입력수량2 = Text2.Text
        입력수량3 = Text3.Text

        Dim JP_Code As Long

        Dim My_Date As String
        Dim My_Time As String
        Dim DBT As New DataTable

        Dim NullStr As String
        Dim NullStr1 As String
        Dim NullStr2 As String

        Dim 공정코드 As String
        Dim 공정명 As String

        Dim 제품코드 As String
        Dim 규격 As String

        Dim 지시서번호 As String
        Dim Gr_Row As Integer

        'Gr_Row = Val(TextBox1.Text)

        ' 지시서번호 = Grid_Main.Item(1, Gr_Row).Value

        StrSQL = "Select * FROM IOT with(NOLOCK)  Where PL_Code = '" & D_Code & "'"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            NullStr = DBT.Rows(0)("PL_PS_Code").ToString
            제품코드 = DBT.Rows(0)("PL_PL_CODE").ToString
            규격 = DBT.Rows(0)("PL_STANDARD").ToString

        End If


        StrSQL = "Select PS_PO_CODE FROM PC_STOCK with(NOLOCK) WHERE PS_CODE = '" & NullStr & "'"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            지시서번호 = DBT.Rows(0)("PS_PO_CODE").ToString
        End If

        StrSQL = "Select * FROM PC_ORDER with(NOLOCK) Where PO_Code = '" & 지시서번호 & "'"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            공정코드 = DBT.Rows(0)("PO_PR_Code")
            공정명 = DBT.Rows(0)("PO_PR_Name")

        End If

        StrSQL = "UPDATE PC_ORDER SET PO_ING='완료' Where PO_Code = '" & 지시서번호 & "'"
        Re_Count = DB_Execute()

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
        Grid_Main.Item(7, Gr_Row).Value = My_Time
        Grid_Main.Item(8, Gr_Row).Value = Text1.Text
        Grid_Main.Item(9, Gr_Row).Value = Text2.Text
        Grid_Main.Item(10, Gr_Row).Value = Text3.Text



        '작업실적 전표 Update
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        'StrSQL = StrSQL & "update PC_STOCK_LIST Set  PS_En_day = '" & My_Date & "', PS_En_Time = '" & My_Time & "', PS_Go = '" & Grid_Main.Item(8, Grid_Main.CurrentCell.RowIndex).Value & "', PS_Total = '" & Grid_Main.Item(9, Grid_Main.CurrentCell.RowIndex).Value & "', PS_Er = '" & Grid_Main.Item(10, Grid_Main.CurrentCell.RowIndex).Value & "' Where PS_Code = '" & NullStr & "' And PS_Sun = '1' AND PS_PL_CODE='" & 설비코드 & "' AND PS_STANDARD ='" & 규격 & "'"
        'Re_Count = DB_Execute()
        StrSQL = StrSQL & "update PC_STOCK_LIST Set PS_En_day = '" & My_Date & "', PS_En_Time = '" & My_Time & "', PS_Go = '" & 입력수량1 & "', PS_Total = '" & 입력수량2 & "', PS_Er = '" & 입력수량3 & "' Where PS_Code = '" & NullStr & "' And PS_Sun = '1' AND PS_PL_CODE='" & 제품코드 & "' AND PS_STANDARD ='" & 규격 & "'"
        Re_Count = DB_Execute()

        '수량 변경
        Count_Pro(제품코드, 공정코드, 입력수량1, 입력수량2, 입력수량3)

        '  Count_Pro(Grid_Main.Item(2, Grid_Main.CurrentCell.RowIndex).Value, 공정코드, Grid_Main.Item(8, Grid_Main.CurrentCell.RowIndex).Value, Grid_Main.Item(9, Grid_Main.CurrentCell.RowIndex).Value, Grid_Main.Item(10, Grid_Main.CurrentCell.RowIndex).Value)


        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "UPDATE IOT Set PL_Ing = '', PL_En_Day = '" & My_Date & "', PL_En_Time = '" & My_Time & "'   where PL_Code = '" & D_Code & "'"
        Re_Count = DB_Execute()


        Dim count As String

        '기존타수 알아내기
        StrSQL = "Select PL_COUNT FROM SI_MACHINE with(NOLOCK) Where PL_Code = '" & D_Code & "'"
        Re_Count = DB_Select(DBT)
        count = DBT.Rows(0)("PL_count")

        If count = "" Then
            count = 입력수량2
        Else
            count = (Integer.Parse(count) + Integer.Parse(입력수량2)).ToString()
        End If

        '타수 더해주기 / PL
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "UPDATE SI_MACHINE Set PL_COUNT = '" & count & "' where PL_Code = '" & D_Code & "'"
        Re_Count = DB_Execute()

        Grid_Main.Item(11, Grid_Main.CurrentCell.RowIndex).Value = "완료"
        Timer1.Enabled = True
        Grid_Main.Enabled = True
        Panel3.Visible = False
        Panel6.Visible = False




    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Timer1.Enabled = True
        Grid_Main.Enabled = True
        Panel3.Visible = False
        Panel6.Visible = False



    End Sub

    Private Sub Text1_TextChanged(sender As Object, e As EventArgs) Handles Text1.TextChanged
        Text3.Text = Val(Text1.Text) - Val(Text2.Text)
    End Sub

    Private Sub Text2_TextChanged(sender As Object, e As EventArgs) Handles Text2.TextChanged
        Text3.Text = Val(Text1.Text) - Val(Text2.Text)
    End Sub

    Private Sub Grid_Line_QC_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Line_QC.CellContentClick

    End Sub

    Private Sub Grid_Line_QC_MouseClick(sender As Object, e As MouseEventArgs) Handles Grid_Line_QC.MouseClick
        '여기처리
        If IsDBNull(Grid_Line_QC.Rows.Count) Then

            MsgBox("설비 점검 list를 추가해주세요")
            Exit Sub
        End If
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
                StrSQL = "Select Base_Name FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '장비점검' Order By Base_Name"
                Re_Count = DB_Select(DBT)
                Grid_Line_QC_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        Grid_Line_QC_Combo.Items.Add(DBT.Rows(i).Item(0))
                    Next i
                End If
                Grid_Line_QC_Combo.Top = Grid_Line_QC.Top + Grid_Line_QC.ColumnHeadersHeight + (Grid_Line_QC_Row) * 100
                Grid_Line_QC_Combo.Left = Grid_Line_QC.Left + Grid_Line_QC.Columns(0).Width + Grid_Line_QC.Columns(1).Width
                Grid_Line_QC_Combo.Width = Grid_Line_QC.Columns(Grid_Line_QC_Col).Width
                Grid_Line_QC_Combo.Height = 105
                Grid_Line_QC_Combo.Text = Grid_Line_QC.CurrentCell.Value.ToString
                Grid_Line_QC_Combo.BackColor = Grid_Line_QC.RowsDefaultCellStyle.BackColor
                Grid_Line_QC_Combo.Visible = True
                Grid_Line_QC_Combo.Focus.ToString()
        End Select


    End Sub
    Private Sub Grid_Line_QC_Combo_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Grid_Line_QC_Combo.SelectionChangeCommitted
        Dim DBT As New DataTable
        Grid_Line_QC.Item(Grid_Line_QC_Col, Grid_Line_QC_Row).Value = Grid_Line_QC_Combo.SelectedItem.ToString()
        Grid_Line_QC.CurrentRow.HeaderCell.Value = "*"
        Grid_Line_QC_Combo.Visible = False
    End Sub
    Private Sub Grid_Go_MouseClick(sender As Object, e As MouseEventArgs) Handles Grid_Go.MouseClick
        '여기처리

        If IsDBNull(Grid_Go.Rows.Count) Then
            MsgBox("점검 list를 추가해주세요")
            Exit Sub
        End If

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
                StrSQL = "Select Base_Name FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '장비점검' Order By Base_Name"
                Re_Count = DB_Select(DBT)
                Grid_Go_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        Grid_Go_Combo.Items.Add(DBT.Rows(i).Item(0))
                    Next i
                End If
                Grid_Go_Combo.Top = Grid_Go.Top + Grid_Go.ColumnHeadersHeight + (Grid_Go_Row) * 100
                Grid_Go_Combo.Left = Grid_Go.Left + Grid_Go.Columns(0).Width + Grid_Go.Columns(1).Width + Grid_Go.Columns(2).Width
                Grid_Go_Combo.Width = Grid_Go.Columns(Grid_Go_Col).Width
                Grid_Go_Combo.Height = 105
                Grid_Go_Combo.Text = Grid_Go.CurrentCell.Value.ToString
                Grid_Go_Combo.BackColor = Grid_Go.RowsDefaultCellStyle.BackColor
                Grid_Go_Combo.Visible = True
                Grid_Go_Combo.Focus.ToString()
        End Select

    End Sub

    Private Sub Grid_Go_Combo_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Grid_Go_Combo.SelectionChangeCommitted
        Dim DBT As New DataTable
        Grid_Go.Item(Grid_Go_Col, Grid_Go_Row).Value = Grid_Go_Combo.SelectedItem.ToString()
        Grid_Go.CurrentRow.HeaderCell.Value = "*"
        Grid_Go_Combo.Visible = False
    End Sub

    Private Sub Text1_MouseClick(sender As Object, e As MouseEventArgs) Handles Text1.MouseClick
        Text1.SelectAll()
        TextBox2.Text = 1
    End Sub
    Private Sub Text2_MouseClick(sender As Object, e As MouseEventArgs) Handles Text2.MouseClick
        Text2.SelectAll()
        TextBox2.Text = 2
    End Sub
    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles C0.Click

    End Sub

    Private Sub C1_Click(sender As Object, e As EventArgs) Handles C0.Click, C1.Click, C2.Click, C3.Click, C4.Click, C5.Click, C6.Click, C7.Click, C8.Click, C9.Click, C10.Click, C11.Click
        Dim Text As TextBox

        Select Case TextBox2.Text
            Case "1"
                Text = Text1
            Case "2"
                Text = Text2
            Case Else
                Exit Sub
        End Select
        If sender.text = "완료" Then
            Exit Sub
        End If

        If sender.text = "삭제" Then
            If Len(Text.Text) = 0 Then
                Exit Sub
            End If
            Text.Text = Mid(Text.Text, 1, Len(Text.Text) - 1)
            Exit Sub

        End If

        If Text.SelectionLength > 0 Then
            Text.Text = sender.text
        Else
            Text.Text = Text.Text & sender.text
        End If
    End Sub

    Private Sub Panel_Menu_Paint(sender As Object, e As PaintEventArgs) Handles Panel_Menu.Paint

    End Sub

    Private Sub Time_Label_Click(sender As Object, e As EventArgs) Handles Time_Label.Click
        Me.Hide()

    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Grid_Main_Display()
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        Me.Hide()

    End Sub
    Public Sub New()

        ' 디자이너에서 이 호출이 필요합니다.
        InitializeComponent()

        ' InitializeComponent() 호출 뒤에 초기화 코드를 추가하세요.

    End Sub

    Private Sub Day_Label_Click(sender As Object, e As EventArgs) Handles Day_Label.Click
        Me.Hide()

    End Sub
End Class