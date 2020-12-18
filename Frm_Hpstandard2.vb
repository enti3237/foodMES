Imports Excel = Microsoft.Office.Interop.Excel

Public Class Frm_Hpstandard2
    Dim xlapp As Excel.Application
    Dim xlbook As Excel.Workbook
    Dim xlsheet As Excel.Worksheet
    Dim Excel_Check As Boolean
    Dim cycle As Integer

    Private Sub Frm_CC_Report1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()
        Grid_Main_Setup()
        'Search_Combo.Items.Add("코드")
        'Search_Combo.Items.Add("날짜")
        'Search_Combo.Items.Add("거래처")
        'Search_Combo.Text = "코드"
        '12, 135
        DataGridView1_Setup()
        S_Text_Setup()


        Panel_Main.Top = 136
        Panel_Main.Left = 12
        Panel_Excel.Top = 136
        Panel_Excel.Left = 12

        Panel_Main.Visible = True
        Panel_Excel.Visible = False
        Com_Excel.Visible = False
        Com_Excel_Print.Visible = False


    End Sub

    Public Function S_Text_Setup() As Boolean
        E_Day.Text = Format(Now, "yyyy-MM-dd")
        S_Day.Text = Mid(E_Day.Text, 1, 8) & "01"
        S_Time.Text = "00:00"
        E_Time.Text = "23:59"

        S_Text_Setup = True

    End Function


    Public Function DataGridView1_Setup() As Boolean
        Grid_Color(DataGridView1)

        DataGridView1.AllowUserToAddRows = False
        DataGridView1.RowTemplate.Height = 24
        DataGridView1.ColumnCount = 2
        DataGridView1.RowCount = 0

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter


        DataGridView1.RowHeadersWidth = 10


        DataGridView1.Columns(0).HeaderText = "문서명"
        DataGridView1.Columns(1).HeaderText = "주기"
        DataGridView1.Columns(1).Visible = False

        Dim i As Integer
        DataGridView1.Columns(0).Width = 60
        DataGridView1.Columns(1).Width = 120



        DataGridView1_Setup = True
    End Function
    Public Function Grid_Main_Setup() As Boolean
        Grid_Color(Grid_Main)

        Grid_Main.AllowUserToAddRows = False
        Grid_Main.RowTemplate.Height = 24
        Grid_Main.ColumnCount = 2
        Grid_Main.RowCount = 0

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Main.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Main.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        '   Grid_Main.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Grid_Main.RowHeadersWidth = 10


        Grid_Main.Columns(0).HeaderText = "일자"
        Grid_Main.Columns(1).HeaderText = "작성여부"
        ' Grid_Main.Columns(2).HeaderText = "작성여부"

        Dim i As Integer
        Grid_Main.Columns(0).Width = 60
        Grid_Main.Columns(1).Width = 120
        '  Grid_Main.Columns(2).Width = 150


        Grid_Main_Setup = True
    End Function

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        '  Dim S_T As String

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        '  S_T = Mid(S_Day.Text, 1, 4) & Mid(S_Day.Text, 6, 2) & Mid(S_Day.Text, 9, 2) & Mid(S_Time.Text, 1, 2)


        Grid_Main.RowCount = 0


        DataGridView1.RowCount = 0

        StrSQL = "select ST_Doc,ST_Date from HP_STANDARD_LIST"

        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            DataGridView1.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                '    Grid_Main.Item(0, i).Value = i + 1
                For j = 0 To DataGridView1.ColumnCount - 1

                    DataGridView1.Item(j, i).Value = DBT.Rows(i).Item(j).ToString

                Next j
            Next i
        Else
            DataGridView1.RowCount = 1
            For j = 0 To DataGridView1.ColumnCount - 1
                DataGridView1.Item(j, 0).Value = ""
            Next j
        End If


        DataGridView1.MultiSelect = False
        DataGridView1.ClearSelection()

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
        'StrSQL = "Select CM_Add1, CM_Add2  FROM HP_STANDARD_LIST with(NOLOCK) Where ST_DOC= '" & Grid_Info.Item(1, 4).Value & "'"
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

    Private Sub DataGridView1_Click(sender As Object, e As EventArgs) Handles DataGridView1.Click
        '    MsgBox((DataGridView1.Item(0, DataGridView1.CurrentCell.RowIndex).Value))
        Dim DBT As New DataTable
        Dim Mok As Integer
        Dim DBT2 As New DataTable
        Dim k As Integer
        Dim interval As Integer '구하려는 날짜 차이
        Dim d1 As Date = S_Day.Text
        Dim d2 As Date = E_Day.Text

        ' MsgBox(DateDiff("d", d1, d2))

        '해당 문서 주기 구하기
        cycle = DataGridView1.Item(1, DataGridView1.CurrentCell.RowIndex).Value
        '  MsgBox(cycle)

        '달력 날짜가 같으면 오늘날짜에만 있는 지 확인
        If S_Day.Text.ToString.Equals(E_Day.Text) Then
            StrSQL = "SELECT HP_DATE, HP_TIME, HP_BIGO 
                    FROM HP_LIST
                   WHERE HP_DATE ='" & S_Day.Text & "'"

            Re_Count = DB_Select(DBT)

            If Re_Count = 0 Then
                '오늘날짜에 써진 게 없음
                Grid_Main.RowCount = 1
                Grid_Main.Item(0, 0).Value = S_Day.Text
                Grid_Main.Item(1, 0).Value = "X"
            Else
                Grid_Main.RowCount = 1
                Grid_Main.Item(0, 0).Value = S_Day.Text
                Grid_Main.Item(1, 0).Value = "O"
            End If


        Else
            'PICK 1 날짜부터 PICK2 날짜까지 주기 계산해서 뿌리기 
            Dim d As Date
            interval = DateDiff("d", d1, d2)
            Mok = interval / cycle
            d = d1
            ' Grid_Main.RowCount = Mok + 1



            If interval >= cycle Then '날짜 차이가 주기보다 클때 
                Grid_Main.RowCount = Mok + 1
                For k = 0 To Mok
                    If Mid(d, 1, 10) > d2 Then
                        Continue For
                    Else
                        Grid_Main.Item(0, k).Value = Mid(d, 1, 10)


                        If d.DayOfWeek = 6 Then
                            Grid_Main.Item(0, k).Style.ForeColor = Color.Blue
                        End If
                        If d.DayOfWeek = 0 Then
                            Grid_Main.Item(0, k).Style.ForeColor = Color.Red
                        End If



                        d = Mid(d.AddDays(cycle), 1, 10)

                        StrSQL = "SELECT * 
                                FROM HP_LIST
                               WHERE HP_DATE ='" & Grid_Main.Item(0, k).Value & "' and HP_Name = '" & DataGridView1.Item(0, DataGridView1.CurrentCell.RowIndex).Value & "'"

                        Re_Count = DB_Select(DBT)

                        If Re_Count = 0 Then
                            ' Grid_Main.Item(0, k).Value = Mid(d, 1, 10)
                            Grid_Main.Item(1, k).Value = "X"
                        Else
                            ' Grid_Main.Item(0, 0).Value = Mid(d, 1, 10)
                            Grid_Main.Item(1, k).Value = "O"
                        End If
                    End If






                Next
            Else
                Grid_Main.RowCount = 1
                StrSQL = "SELECT * 
                                FROM HP_LIST
                               WHERE HP_DATE ='" & Grid_Main.Item(0, 0).Value & "'"

                Re_Count = DB_Select(DBT)

                If Re_Count = 0 Then
                    Grid_Main.Item(0, 0).Value = Mid(d, 1, 10)
                    Grid_Main.Item(1, 0).Value = "X"
                Else
                    Grid_Main.Item(0, 0).Value = Mid(d, 1, 10)
                    Grid_Main.Item(1, 0).Value = "O"
                End If
            End If


            'For k = 1 To interval

            '    '  Grid_Main.RowCount = DateDiff("d", d1, d2)
            '    '  MsgBox(DateDiff("d", d1, d2))
            '    ' MsgBox("d1 :" & d1)
            '    '  MsgBox(Mid(d1.AddDays(cycle), 1, 10))


            '    d1 = Mid(d1.AddDays(cycle), 1, 10)
            '    Grid_Main.Item(0, k).Value = Mid(d1, 1, 10)
            '    StrSQL = "SELECT * 
            '                FROM HP_LIST
            '               WHERE HP_DATE ='" & Grid_Main.Item(0, k).Value & "'"

            '    Re_Count2 = DB_Select(DBT2)

            '    If Re_Count2 = 0 Then
            '        Grid_Main.Item(1, k).Value = "X"
            '    Else
            '        Grid_Main.Item(1, k).Value = "O"
            '    End If

            '    d1 = d1.AddDays(cycle)
            '    If d1 > d2 Then
            '        MsgBox("d2 : " & d2)
            '        Exit Sub
            '    End If
            'Next
        End If



        'Re_Count = DB_Select(DBT)
        'If Re_Count <> 0 Then
        '    Grid_Main.RowCount = Re_Count
        '    For i = 0 To Re_Count - 1
        '        '    Grid_Main.Item(0, i).Value = i + 1
        '        For j = 0 To Grid_Main.ColumnCount - 1
        '            Grid_Main.Item(j, i).Value = DBT.Rows(i).Item(j).ToString
        '        Next j
        '    Next i
        'Else
        '    Grid_Main.RowCount = 1
        '    For j = 0 To Grid_Main.ColumnCount - 1
        '        Grid_Main.Item(j, 0).Value = ""
        '    Next j
        'End If

        'Grid_Main.MultiSelect = False
        'Grid_Main.ClearSelection()



    End Sub
End Class
