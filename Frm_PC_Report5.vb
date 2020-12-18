﻿Imports Excel = Microsoft.Office.Interop.Excel
Public Class Frm_PC_Report5
    Dim xlapp As Excel.Application
    Dim xlbook As Excel.Workbook
    Dim xlsheet As Excel.Worksheet
    Dim Excel_Check As Boolean

    Private Sub Frm_PC_Report2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()

        S_Text_Setup()
        Grid_Main_Setup()

        Panel_Main.Top = 136
        Panel_Main.Left = 12
        Panel_Excel.Top = 136
        Panel_Excel.Left = 12

        Panel_Main.Visible = True
        Panel_Excel.Visible = False

        Com_Excel_Print.Visible = False
        Search_Com.PerformClick()

    End Sub
    Public Function S_Text_Setup() As Boolean
        E_Day.Text = Format(Now, "yyyy-MM-dd")
        S_Day.Text = Mid(E_Day.Text, 1, 8) & "01"
        S_Time.Text = "00:00"
        E_Time.Text = "23:59"



        S_Text_Setup = True

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
        Grid_Main.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Grid_Main.RowHeadersWidth = 10


        Grid_Main.Columns(0).HeaderText = "순번"
        Grid_Main.Columns(1).HeaderText = "공정코드"
        Grid_Main.Columns(2).HeaderText = "공정"
        Grid_Main.Columns(3).HeaderText = "지시수량"
        Grid_Main.Columns(4).HeaderText = "투입수량"
        Grid_Main.Columns(5).HeaderText = "생산수량"
        Grid_Main.Columns(6).HeaderText = "불량수량"
        Grid_Main.Columns(7).HeaderText = "비고"


        Dim i As Integer
        Grid_Main.Columns(0).Width = 60
        Grid_Main.Columns(1).Width = 100
        Grid_Main.Columns(2).Width = 180
        Grid_Main.Columns(3).Width = 100
        Grid_Main.Columns(4).Width = 100
        Grid_Main.Columns(5).Width = 100
        Grid_Main.Columns(6).Width = 100
        Grid_Main.Columns(7).Width = 100
        Grid_Main_Setup = True
    End Function

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        Grid_Main.Columns(0).HeaderText = "순번"
        Grid_Main.Columns(1).HeaderText = "공정코드"
        Grid_Main.Columns(2).HeaderText = "공정"
        Grid_Main.Columns(3).HeaderText = "지시수량"
        Grid_Main.Columns(4).HeaderText = "투입수량"
        Grid_Main.Columns(5).HeaderText = "생산수량"
        Grid_Main.Columns(6).HeaderText = "불량수량"
        Grid_Main.Columns(7).HeaderText = "비고"

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Main.RowCount = 0


        StrSQL = "Select PC_Code, PC_Name FROM si_process with(NOLOCK) "
        StrSQL = StrSQL & " Order By PC_Code"

        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Main.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                Grid_Main.Item(0, i).Value = i + 1
                For j = 1 To 2
                    Grid_Main.Item(j, i).Value = DBT.Rows(i).Item(j - 1).ToString
                Next j
            Next i
        Else
            Grid_Main.RowCount = 1
            For j = 0 To Grid_Main.ColumnCount - 1
                Grid_Main.Item(j, 0).Value = ""
            Next j
        End If




        For i = 0 To Grid_Main.RowCount - 1
                '지시수량
                StrSQL = "Select Sum(Convert(Decimal,PO_Total)) AS F1  FROM PC_ORDER with(NOLOCK), PC_ORDER_LIST with(NOLOCK)  "
                StrSQL = StrSQL & " Where PC_ORDER.PO_Code = PC_ORDER_LIST.PO_Code And PO_Date >= '" & S_Day.Text & "' And PO_Date <= '" & E_Day.Text & "' And Len(PO_Total) > 0 And Len(PO_PR_CODE) > 0 And PO_PR_CODE = '" & Grid_Main.Item(1, i).Value & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Main.Item(3, i).Value = DBT.Rows(0).Item(0).ToString
                Else
                    Grid_Main.Item(3, i).Value = "0"
                End If

            Next i

            '생산수량 

            For i = 0 To Grid_Main.RowCount - 1
                '생산수량
                StrSQL = "Select Sum(Convert(Decimal,PS_Go)) AS F1, Sum(Convert(Decimal,PS_Total)) AS F2, Sum(Convert(Decimal,PS_Er)) AS F3 FROM PC_STOCK with(NOLOCK), PC_Stock_List with(NOLOCK) "
                StrSQL = StrSQL & " Where PC_Stock.PS_Code = PC_STOCK_LIST.PS_Code And PS_St_Day >= '" & S_Day.Text & "' And PS_En_Day <= '" & E_Day.Text & "' And Len(PS_Go) > 0 And Len(PS_Total) > 0 And Len(PS_Er) > 0 "
                StrSQL = StrSQL & " And PS_PR_CODE = '" & Grid_Main.Item(1, i).Value & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Main.Item(4, i).Value = DBT.Rows(0).Item(0).ToString
                    Grid_Main.Item(5, i).Value = DBT.Rows(0).Item(1).ToString
                    Grid_Main.Item(6, i).Value = DBT.Rows(0).Item(2).ToString
                End If
                If Len(Grid_Main.Item(4, i).Value) = 0 Then
                    Grid_Main.Item(4, i).Value = "0"
                End If
                If Len(Grid_Main.Item(5, i).Value) = 0 Then
                    Grid_Main.Item(5, i).Value = "0"
                End If
                If Len(Grid_Main.Item(6, i).Value) = 0 Then
                    Grid_Main.Item(6, i).Value = "0"
                End If

            Next i



    End Sub

    Private Sub Com_Excel_Click(sender As Object, e As EventArgs) Handles Com_Excel.Click
        Label_Color(sender, Color_Label, Di_Panel2)
        '판넬 보이기
        Panel_Main.Visible = False
        Panel_Excel.Visible = True
        Com_Excel_Print.Visible = True

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
        xlbook = xlapp.Workbooks.Open(Application.StartupPath & "\excel\null.xlsx")
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


        SetParent(xlapp.Hwnd, Panel_Excel.Handle)
        SendMessage(xlapp.Hwnd, WM_SYSCOMMAND, SC_MAXIMIZE, 0)



        '제목

        For j = 0 To Grid_Main.ColumnCount - 1
            xlsheet.Cells(1, j + 1).Value = Grid_Main.Columns(j).HeaderText
            xlsheet.Columns(j + 1).ColumnWidth = Grid_Main.Columns(j).Width / 10

        Next j

        For i = 0 To Grid_Main.RowCount - 1
            For j = 0 To Grid_Main.ColumnCount - 1
                xlsheet.Cells(i + 2, j + 1).Value = "'" & Grid_Main.Item(j, i).Value
            Next j
        Next i



        xlbook.SaveAs(Application.StartupPath & "\Excel\내역서\" & Format(Now, "yyyyMMddHHmmss") & ".xlsx")
        Cursor = Cursors.Default
        Excel_Check = True
    End Sub

    Private Sub Panel_Excel_Paint(sender As Object, e As PaintEventArgs) Handles Panel_Excel.Paint

    End Sub

    Private Sub Excel_Timer_Tick(sender As Object, e As EventArgs)

    End Sub

    Private Sub Com_Excel_Print_Click(sender As Object, e As EventArgs) Handles Com_Excel_Print.Click
        If Excel_Check = True Then
            xlsheet.PrintOut(From:=1, To:=1, Copies:=1, Collate:=True)
        End If

    End Sub

    Private Sub Frm_CC_Report1_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        Dim i As Long
        i = i
    End Sub

    Private Sub Frm_CC_Report1_VisibleChanged(sender As Object, e As EventArgs) Handles Me.VisibleChanged
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

    Private Sub Com_Excel_Print_VisibleChanged(sender As Object, e As EventArgs) Handles Com_Excel_Print.VisibleChanged
    End Sub

    Private Sub Com_Contract_Click(sender As Object, e As EventArgs) Handles Com_Contract.Click
        Label_Color(sender, Color_Label, Di_Panel2)
        '판넬 보이기
        Panel_Main.Visible = True
        Panel_Excel.Visible = False
        Com_Excel_Print.Visible = False
    End Sub

End Class
