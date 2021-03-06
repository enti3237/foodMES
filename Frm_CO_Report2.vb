﻿Imports Excel = Microsoft.Office.Interop.Excel
Public Class Frm_CO_Report2
    Dim xlapp As Excel.Application
    Dim xlbook As Excel.Workbook
    Dim xlsheet As Excel.Worksheet
    Dim Excel_Check As Boolean

    Dim Grid_Display_Ch As Boolean

    Private Sub Frm_CO_Report2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()
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
        Grid_Main.ClearSelection()
    End Sub
    Public Function S_Text_Setup() As Boolean

        Dim DBT As New DataTable

        Search_Combo.Items.Clear()
        Search_Combo.Items.Add("전체")
        StrSQL = "Select Base_Name  FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '제품종류' Order By Base_Name"
        Re_Count = DB_Select(DBT)
        'Grid_Info_Combo.Items.Clear()
        If Re_Count <> 0 Then
            For i = 0 To Re_Count - 1
                Search_Combo.Items.Add(DBT.Rows(i)("Base_Name"))
            Next i
        End If
        Search_Combo.Text = "전체"

        S_Text_Setup = True

    End Function
    Public Function Grid_Main_Setup() As Boolean
        Grid_Color(Grid_Main)


        Grid_Main.AllowUserToAddRows = False
        Grid_Main.RowTemplate.Height = 24
        Grid_Main.ColumnCount = 4
        Grid_Main.RowCount = 0

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Main.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Grid_Main.RowHeadersWidth = 10

        Grid_Main.Columns(0).HeaderText = "구분"
        Grid_Main.Columns(1).HeaderText = "코드"
        Grid_Main.Columns(2).HeaderText = "품목명"
        Grid_Main.Columns(3).HeaderText = "수량"


        Dim i As Integer
        Grid_Main.Columns(0).Width = 60
        Grid_Main.Columns(1).Width = 120
        Grid_Main.Columns(2).Width = 150
        Grid_Main.Columns(3).Width = 150

        Grid_Main_Setup = True


    End Function

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Main.RowCount = 0

        If Search_Combo.Text = "전체" Then
            StrSQL = "Select PL_SEL, PL_CODE, PL_NAME  FROM SI_PRODUCT with(NOLOCK) Where PL_Name Like '%" & S_Day.Text & "%' Order By PL_Code "
        Else
            StrSQL = "Select PL_SEL, PL_CODE, PL_NAME  FROM SI_PRODUCT with(NOLOCK) Where PL_Name Like '%" & S_Day.Text & "%' And Pl_Sel = '" & Search_Combo.Text & "' Order By PL_Code "
        End If
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Main.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                '   Grid_Main.Item(0, i).Value = i + 1
                Grid_Main.Item(0, i).Value = DBT.Rows(i)("PL_Sel").ToString
                Grid_Main.Item(1, i).Value = DBT.Rows(i)("PL_Code").ToString
                Grid_Main.Item(2, i).Value = DBT.Rows(i)("PL_Name").ToString
            Next i
        Else
            Grid_Main.RowCount = 1
            For j = 0 To Grid_Main.ColumnCount - 1
                Grid_Main.Item(j, 0).Value = ""
            Next j
        End If
        '재고수량

        For i = 0 To Grid_Main.RowCount - 1
            StrSQL = "Select PL_QTY AS F1 FROM STOCK with(NOLOCK) "
            StrSQL = StrSQL & " Where PL_CODE = '" & Grid_Main.Item(1, i).Value & "'"
            Re_Count = DB_Select(DBT)
            If Re_Count <> 0 Then
                Grid_Main.Item(3, i).Value = DBT.Rows(0)("F1").ToString
            Else
                Grid_Main.Item(3, i).Value = "0"

            End If

        Next i

        Grid_Main.MultiSelect = False
        Grid_Main.ClearSelection()
        Grid_Display_Ch = True
    End Sub

    Private Sub Com_Excel_Print_Click(sender As Object, e As EventArgs) Handles Com_Excel_Print.Click
        If Excel_Check = True Then
            xlsheet.PrintOut(From:=1, To:=1, Copies:=1, Collate:=True)
        End If
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
        Com_Excel_Print.Visible = False
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

    'Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
    '    '재고수정
    '    Grid_Display_Ch = False

    '    Dim Table_Name As String
    '    Dim i As Integer
    '    Dim j As Long
    '    Dim k As Integer
    '    Dim Field_Name(500) As String
    '    Dim DBT As DataTable = Nothing
    '    Dim RRR As String
    '    j = 0

    '    '현재 일자
    '    Dim tDate As Date
    '    Dim tDateTime As Date
    '    tDate = Now()
    '    tDate = Format(Now(), "yyyy-MM-dd")
    '    tDateTime = Format(Now(), "yyyy-MM-dd HH:mm:ss")

    '    For i = 0 To Grid_Main.Rows.Count - 1
    '        If Grid_Main.Rows(i).HeaderCell.Value = "*" Then
    '            ' MsgBox(tDateTime)
    '            'Update

    '            StrSQL = "SELECT COUNT(PP_SUN) FROM SI_PROCESS_LIST WHERE PP_CODE='" & Grid_Main.Item(2, i).Value & "'"
    '            Re_Count = DB_Select(DBT)

    '            If Re_Count = 0 Then
    '                Continue For
    '            Else
    '            End If

    '            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
    '            StrSQL = StrSQL & "UPDATE SI_PROCESS_LIST SET PP_AMOUNT='" & Grid_Main.Item(7, i).Value & "' "
    '            StrSQL = StrSQL & "WHERE PP_SUN='1' AND PP_CODE='" & Grid_Main.Item(2, i).Value & "'"

    '            'StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
    '            'StrSQL = StrSQL & "INSERT INTO INVADJHIS(TRDDATE, ITMCD, ADJQTY, REMARK, REGUSER, EDTUSER, REGDATIME, EDTDATIME) "
    '            'StrSQL = StrSQL & "values('" & Mid(tDate.ToString(), 1, 10) & "', '" & Grid_Main.Item(1, i).Value & "', '" & Grid_Main.Item(4, i).Value - Grid_Main.Item(3, i).Value & "', '" & Grid_Main.Item(5, i).Value & "', '" & Frm_Main.Text_Man_Code.Text
    '            'StrSQL = StrSQL & "', '" & Frm_Main.Text_Man_Code.Text & "', '" & tDateTime.ToString("yyyy-MM-dd HH:MM:ss") & "', '" & tDateTime.ToString("yyyy-MM-dd HH:MM:ss") & "')"
    '            Re_Count = DB_Execute()
    '            Grid_Main.Rows(i).HeaderCell.Value = ""
    '        End If
    '    Next i
    '    MsgBox("저장되었습니다")
    '    Grid_Display_Ch = True

    '    Search_Com.PerformClick()

    'End Sub

    Private Sub Grid_Main_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Main.CellValueChanged
        If Grid_Display_Ch = True Then
            Grid_Main.CurrentRow.HeaderCell.Value = "*"
        End If
    End Sub
End Class
