Imports Excel = Microsoft.Office.Interop.Excel

Public Class Frm_FL_Report1
    Dim xlapp As Excel.Application
    Dim xlbook As Excel.Workbook
    Dim xlsheet As Excel.Worksheet
    Dim Excel_Check As Boolean
    Private Sub Frm_FL_Report1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

        Button1.Visible = False
        ' Com_Excel_Print.Visible = False
        Search_Com.PerformClick()
    End Sub
    Public Function S_Text_Setup() As Boolean
        E_Day.Text = Format(Now, "yyyy-MM-dd")
        S_Day.Text = Mid(E_Day.Text, 1, 8) & "01"



        S_Text_Setup = True

    End Function

    Public Function Grid_Main_Setup() As Boolean
        Grid_Color(Grid_Main)

        Grid_Main.AllowUserToAddRows = False
        Grid_Main.RowTemplate.Height = 24
        Grid_Main.ColumnCount = 3
        Grid_Main.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Main.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Main.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Main.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter


        Grid_Main.RowHeadersWidth = 10


        Grid_Main.Columns(0).HeaderText = "설비코드"
        Grid_Main.Columns(1).HeaderText = "설비명"
        Grid_Main.Columns(2).HeaderText = "가동여부"
        'Grid_Main.Columns(3).HeaderText = "총 생산량"
        ' Grid_Main.Columns(4).HeaderText = "총 생산량"
        '   Grid_Main.Columns(5).HeaderText = "시간당 생산량"

        Dim i As Integer

        'Grid_Main.Columns(0).Width = 40

        For i = 0 To Grid_Main.ColumnCount - 1
            Grid_Main.Columns(i).Width = 180
        Next i


        Grid_Main_Setup = True
    End Function

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        Log_Display()

        Dim DBT As New DataTable
        Dim DBT2 As New DataTable
        Dim DBT3 As New DataTable

        Dim i As Integer
        Dim j As Integer
        Dim k As Integer

        Dim Re_count2 As Integer
        Dim Re_count3 As Integer

        Grid_Main.RowCount = 0
        ' StrSQL = "Select FL_Date, FL_Time, FL_Name, FL_DE_Code, FL_DE_Name, FL_DE_History, FL_Gold, FC_RECORD_List.FL_Bigo From FC_RECORD with(NOLOCK), FC_RECORD_List with(NOLOCK) Where FC_RECORD.FL_Code = FC_RECORD_List.FL_Code And  FL_Date + FL_Time >= '" & S_Day.Text & S_Time.Text & "' And FL_Date + FL_Time <= '" & E_Day.Text & E_Time.Text & "'   Order By FL_Date, FL_Time "
        StrSQL = "SELECT PD_CODE, PD_NAME FROM SI_MACHINE  "
        Re_Count = DB_Select(DBT)

        If Re_Count <> 0 Then
            Grid_Main.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To DBT.Columns.Count - 1
                    Grid_Main.Item(j, i).Value = DBT.Rows(i).Item(j).ToString

                    '가동 비가동여부
                    StrSQL = "select status from SI_MACHINE join IOT_MACHINE on SI_MACHINE.PD_Name=IOT_MACHINE.MACHINE_NM 
WHERE PD_CODE ='" & Grid_Main.Item(0, i).Value & "'"
                    Re_count2 = DB_Select(DBT2)

                    If IsDBNull(Re_count2) Then
                        Continue For
                    End If

                    If Re_count2 <> 0 Then

                        Grid_Main.Item(2, i).Value = "가동중"

                    Else
                        Grid_Main.Item(2, i).Value = "비가동"
                    End If
                    '  Grid_Main.Item(j, 0).Value = DBT.Rows(i).Item(j).ToString

                    ''총 생산량 
                    'StrSQL = "SELECT SUM(IIF(PS_TOTAL is null, 0,  CONVERT(decimal, PS_TOTAL))) as Total 
                    '            FROM PC_STOCK 
                    '           WHERE PS_DE_NAME ='" & Grid_Main.Item(1, i).Value & "' AND PS_St_Day BETWEEN '" & S_Day.Text & "' and '" & E_Day.Text & "'"
                    'Re_count3 = DB_Select(DBT3)

                    'If IsDBNull(DBT3.Rows(0).Item(0)) Then
                    '    Grid_Main.Item(3, i).Value = "0"
                    '    Continue For
                    'End If

                    'If Re_count3 <> 0 Then

                    '    Grid_Main.Item(3, i).Value = DBT3.Rows(0).Item(0).ToString
                    'Else
                    '    Grid_Main.Item(3, i).Value = "-"
                    'End If
                    ''  Grid_Main.Item(j, 0).Value = DBT.Rows(i).Item(j).ToString

                Next j
            Next i
        Else
            Grid_Main.RowCount = 1
            For j = 0 To Grid_Main.ColumnCount - 1
                Grid_Main.Item(j, 0).Value = ""
            Next j
        End If



        If Re_Count <> 0 Then
            Grid_Main.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To DBT.Columns.Count - 2
                    Grid_Main.Item(j, i).Value = DBT.Rows(i).Item(j).ToString
                Next j
            Next i
        Else
            Grid_Main.RowCount = 1
            For j = 0 To Grid_Main.ColumnCount - 1
                Grid_Main.Item(j, 0).Value = ""
            Next j
        End If



    End Sub
    Public Function Log_Setup() As Boolean
        Grid_Color(DataGridView1)

        DataGridView1.AllowUserToAddRows = False
        DataGridView1.RowTemplate.Height = 20
        DataGridView1.ColumnCount = 5
        DataGridView1.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.RowHeadersWidth = 10

        DataGridView1.Columns(0).HeaderText = "작업일자"
        DataGridView1.Columns(1).HeaderText = "작업화면"
        DataGridView1.Columns(2).HeaderText = "작업"

        DataGridView1.Columns(3).HeaderText = "작업자"
        DataGridView1.Columns(4).HeaderText = "작업비고"


        Dim i As Integer
        For i = 0 To DataGridView1.ColumnCount - 1
            DataGridView1.Columns(i).ReadOnly = True
        Next i
        Log_Setup = True
    End Function
    Public Function Log_Display() As Boolean

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Log_Setup()

        StrSQL = "Select * FROM MAN_LOG with(NOLOCK) order By ACT_DATE DESC"
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

        Log_Display = True
        DataGridView1.ClearSelection()


    End Function
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
        ' Com_Excel_Print.Visible = False
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
