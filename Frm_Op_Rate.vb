﻿Imports Excel = Microsoft.Office.Interop.Excel

Public Class Frm_Op_Rate
    Dim xlapp As Excel.Application
    Dim xlbook As Excel.Workbook
    Dim xlsheet As Excel.Worksheet
    Dim Excel_Check As Boolean
    Private Sub Frm_FL_Report2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

        ' Com_Excel_Print.Visible = False
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
        Grid_Main.ColumnCount = 3
        Grid_Main.RowCount = 0

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Main.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Main.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Main.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter



        Grid_Main.RowHeadersWidth = 10



        Grid_Main.Columns(0).HeaderText = "설비코드"
        Grid_Main.Columns(1).HeaderText = "설비명"
        Grid_Main.Columns(2).HeaderText = "가동시간(h)"




        Dim i As Integer
        Grid_Main.Columns(0).Width = 180
        Grid_Main.Columns(1).Width = 180
        Grid_Main.Columns(2).Width = 180

        Grid_Main.Columns(0).ReadOnly = True
        Grid_Main.Columns(1).ReadOnly = True
        Grid_Main.Columns(2).ReadOnly = True

        Grid_Main_Setup = True
    End Function

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Dim s As String

        StrSQL = "Select PL_Code, PL_Name FROM SI_MACHINE with(NOLOCK) "
        StrSQL = StrSQL & " Order By PL_Code"

        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Main.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                '  Grid_Main.Item(0, i).Value = i + 1
                For j = 0 To 1
                    Grid_Main.Item(j, i).Value = DBT.Rows(i).Item(j).ToString
                Next j
            Next i
        Else
            Grid_Main.RowCount = 1
            For j = 0 To Grid_Main.ColumnCount - 1
                Grid_Main.Item(j, 0).Value = ""
            Next j
        End If


        For i = 0 To Grid_Main.RowCount - 1
            '시간
            StrSQL = "Select Sum((CASE WHEN [PS_En_Time] IS NULL THEN '0' ELSE
	                  DATEDIFF(MI, concat([PS_St_Day], ' ', [PS_St_Time]), concat([PS_En_Day], ' ', [PS_En_Time]))/60.0 END)) AS F1
                    FROM PC_STOCK with(NOLOCK), PC_Stock_List with(NOLOCK) "
            StrSQL = StrSQL & " Where PC_Stock.PS_Code = PC_STOCK_LIST.PS_Code And PS_St_Day >= '" & S_Day.Text & "' And PS_En_Day <= '" & E_Day.Text & "' And (Len(PS_Go) > 0 or Len(PS_Total) > 0 or Len(PS_Er) > 0) "
            StrSQL = StrSQL & " And PS_DE_Code = '" & Grid_Main.Item(0, i).Value & "' and ps_de_code is not null"
            Re_Count = DB_Select(DBT)

            If IsDBNull(DBT.Rows(0).Item(0)) Then
                Exit Sub
            End If

            If Re_Count <> 0 Then
                Grid_Main.Item(2, i).Value = DBT.Rows(0).Item(0).ToString
                ' MsgBox(Grid_Main.Item(2, i).Value)
                If IsDBNull(Grid_Main.Item(2, i).Value) Then
                    Exit Sub
                End If
                '소수점 첫째자리 변환
                Grid_Main.Item(2, i).Value = Math.Round(CDbl(Grid_Main.Item(2, i).Value), 1)

                End If



        Next i
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
        Label_Color(sender, Color_Label, Di_Panel2)
        '판넬 보이기
        Panel_Main.Visible = True
        Panel_Excel.Visible = False
        '   Com_Excel_Print.Visible = False
    End Sub

    Private Sub Com_Excel_Click(sender As Object, e As EventArgs)
        Label_Color(sender, Color_Label, Di_Panel2)
        '판넬 보이기
        Panel_Main.Visible = False
        Panel_Excel.Visible = True
        '  Com_Excel_Print.Visible = True

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

    Private Sub Com_Excel_Print_Click(sender As Object, e As EventArgs)
        If Excel_Check = True Then
            xlsheet.PrintOut(From:=1, To:=1, Copies:=1, Collate:=True)
        End If
    End Sub
End Class