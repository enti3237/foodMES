﻿Imports Excel = Microsoft.Office.Interop.Excel
Public Class Frm_CC_Report2
    Dim xlapp As Excel.Application
    Dim xlbook As Excel.Workbook
    Dim xlsheet As Excel.Worksheet
    Dim Excel_Check As Boolean

    Private Sub Frm_CC_Report1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()

        S_Text_Setup()
        Grid_Main_Setup()
        'Search_Combo.Items.Add("코드")
        'Search_Combo.Items.Add("날짜")
        'Search_Combo.Items.Add("거래처")
        'Search_Combo.Text = "코드"
        '12, 135

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

        Dim DBT As New DataTable
        Dim i As Integer
        ComboBox2.Items.Clear()
        ComboBox2.Items.Add("전체")
        StrSQL = "Select *  FROM SI_PRODUCT with(NOLOCK) WHERE PL_SEL='제품' Order By PL_Code"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            For i = 0 To Re_Count - 1
                ComboBox2.Items.Add(DBT.Rows(i)("PL_Code"))  'DBT.Rows(0)("PL_Name")
            Next i
        End If

        ComboBox1.Items.Clear()
        ComboBox1.Items.Add("전체")
        StrSQL = "Select *  FROM SI_PRODUCT with(NOLOCK) WHERE PL_SEL='제품' Order By PL_Name"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            For i = 0 To Re_Count - 1
                ComboBox1.Items.Add(DBT.Rows(i)("PL_Name"))  'DBT.Rows(0)("PL_Name")
            Next i
        End If

        ComboBox2.Text = "전체"
        ComboBox1.Text = "전체"


        Search_Combo.Items.Clear()
        Search_Combo.Items.Add("전체")
        StrSQL = "Select *  FROM SI_CUSTOMER with(NOLOCK) Order By CM_Code"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            For i = 0 To Re_Count - 1
                Search_Combo.Items.Add(DBT.Rows(i)("CM_Code"))  'DBT.Rows(0)("CM_Name")
            Next i
        End If

        Search_Combo_Name.Items.Clear()
        Search_Combo_Name.Items.Add("전체")
        StrSQL = "Select *  FROM SI_CUSTOMER with(NOLOCK) Order By CM_Name"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            For i = 0 To Re_Count - 1
                Search_Combo_Name.Items.Add(DBT.Rows(i)("CM_Name"))  'DBT.Rows(0)("CM_Name")
            Next i
        End If

        Search_Combo.Text = "전체"
        Search_Combo_Name.Text = "전체"


        S_Text_Setup = True

    End Function


    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "전체" Then
            ComboBox2.Text = "전체"
            Exit Sub
        End If

        Dim DBT As New DataTable
        StrSQL = "Select *  FROM SI_PRODUCT with(NOLOCK) Where PL_Name = '" & ComboBox1.Text & "'"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            ComboBox2.Text = DBT.Rows(0)("PL_Code")
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.Text = "전체" Then
            ComboBox1.Text = "전체"
            Exit Sub
        End If

        Dim DBT As New DataTable
        StrSQL = "Select *  FROM SI_PRODUCT with(NOLOCK) Where PL_Code = '" & ComboBox2.Text & "' Order By PL_Code"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            ComboBox1.Text = DBT.Rows(0)("PL_Name")
        End If
    End Sub

    Private Sub Search_Combo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Search_Combo.SelectedIndexChanged

        If Search_Combo.Text = "전체" Then
            Search_Combo_Name.Text = "전체"
            Exit Sub
        End If

        Dim DBT As New DataTable
        StrSQL = "Select *  FROM SI_CUSTOMER with(NOLOCK) Where Cm_Code = '" & Search_Combo.Text & "' Order By CM_Code"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Search_Combo_Name.Text = DBT.Rows(0)("CM_Name")
        End If
    End Sub

    Private Sub Search_Combo_Name_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Search_Combo_Name.SelectedIndexChanged
        If Search_Combo_Name.Text = "전체" Then
            Search_Combo.Text = "전체"
            Exit Sub
        End If

        Dim DBT As New DataTable
        StrSQL = "Select *  FROM SI_CUSTOMER with(NOLOCK) Where CM_Name = '" & Search_Combo_Name.Text & "'"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Search_Combo.Text = DBT.Rows(0)("CM_Code")
        End If

    End Sub
    Public Function Grid_Main_Setup() As Boolean
        Grid_Color(Grid_Main)

        Grid_Main.AllowUserToAddRows = False
        Grid_Main.RowTemplate.Height = 24
        Grid_Main.ColumnCount = 9
        Grid_Main.RowCount = 0

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Main.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Grid_Main.RowHeadersWidth = 10


        Grid_Main.Columns(0).HeaderText = "납품일자"
        Grid_Main.Columns(1).HeaderText = "거래처명"
        Grid_Main.Columns(2).HeaderText = "제품명"
        Grid_Main.Columns(3).HeaderText = "규격"
        Grid_Main.Columns(4).HeaderText = "단가"
        Grid_Main.Columns(5).HeaderText = "수량"
        Grid_Main.Columns(6).HeaderText = "금액"
        Grid_Main.Columns(7).HeaderText = "세금"
        Grid_Main.Columns(8).HeaderText = "비고"

        Dim i As Integer
        Grid_Main.Columns(0).Width = 60
        Grid_Main.Columns(1).Width = 120
        Grid_Main.Columns(2).Width = 150
        Grid_Main.Columns(3).Width = 150
        For i = 4 To Grid_Main.ColumnCount - 1
            Grid_Main.Columns(i).Width = 90
        Next i
        Grid_Main_Setup = True
    End Function

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        Dim S_T As String

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        ' S_T = Mid(S_Day.Text, 1, 4) & Mid(S_Day.Text, 6, 2) & Mid(S_Day.Text, 9, 2) & Mid(S_Time.Text, 1, 2)

        Grid_Main.RowCount = 0

        If Search_Combo_Name.Text = "전체" And ComboBox2.Text = "전체" Then
            StrSQL = "SELECT DL_Day, CR_CM_NAME, DL_PL_Name, DL_Standard, DL_Unit_Amount, DL_QTY, DL_Gold, DL_Vat, SC_DELIVER_LIST.DL_Bigo  FROM SC_DELIVER with(NOLOCK), SC_DELIVER_LIST with(NOLOCK)  Where SC_DELIVER.DR_Code = SC_DELIVER_LIST.DL_DR_Code And DL_Day >= '" & S_Day.Text & "' And DL_Day <= '" & E_Day.Text & "' Order By DL_Day"
        ElseIf Search_Combo_Name.Text = "전체" And ComboBox2.Text <> "전체" Then
            StrSQL = "Select DL_Day, CR_CM_NAME, DL_PL_Name, DL_Standard, DL_Unit_Amount, DL_QTY, DL_Gold, DL_Vat, SC_DELIVER_LIST.DL_Bigo  FROM SC_DELIVER with(NOLOCK), SC_DELIVER_LIST with(NOLOCK)  Where SC_DELIVER.DR_Code = SC_DELIVER_LIST.DL_DR_Code And DL_Day >= '" & S_Day.Text & "' And DL_Day <= '" & E_Day.Text & "' and DL_PL_NAME ='" & ComboBox1.Text & "' Order By DL_Day"
        ElseIf Search_Combo_Name.Text <> "전체" And ComboBox2.Text = "전체" Then
            StrSQL = "Select DL_Day, CR_CM_NAME, DL_PL_Name, DL_Standard, DL_Unit_Amount, DL_QTY, DL_Gold, DL_Vat, SC_DELIVER_LIST.DL_Bigo  FROM SC_DELIVER with(NOLOCK), SC_DELIVER_LIST with(NOLOCK)  Where SC_DELIVER.DR_Code = SC_DELIVER_LIST.DL_DR_Code And DL_Day >= '" & S_Day.Text & "' And DL_Day <= '" & E_Day.Text & "' And CR_CM_NAME = '" & Search_Combo.Text & "'  Order By DL_Day"
        Else
            StrSQL = "Select DL_Day, CR_CM_NAME, DL_PL_Name, DL_Standard, DL_Unit_Amount, DL_QTY, DL_Gold, DL_Vat, SC_DELIVER_LIST.DL_Bigo  FROM SC_DELIVER with(NOLOCK), SC_DELIVER_LIST with(NOLOCK)  Where SC_DELIVER.DR_Code = SC_DELIVER_LIST.DL_DR_Code And DL_Day >= '" & S_Day.Text & "' And DL_Date <= '" & E_Day.Text & "' And CR_CM_NAME = '" & Search_Combo.Text & "' and DL_PL_NAME ='" & ComboBox1.Text & "' Order By DL_Day"
        End If


        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Main.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                Grid_Main.Item(0, i).Value = i + 1
                For j = 1 To Grid_Main.ColumnCount - 1
                    Grid_Main.Item(j, i).Value = DBT.Rows(i).Item(j - 1).ToString
                Next j
            Next i
        Else
            Grid_Main.RowCount = 1
            For j = 0 To Grid_Main.ColumnCount - 1
                Grid_Main.Item(j, 0).Value = ""
            Next j
        End If

        Grid_Main.MultiSelect = False
        Grid_Main.ClearSelection()

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
                xlsheet.Cells(i + 2, j + 1).Value = "'" & Grid_Main.Item(j, i).Value.ToString
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
