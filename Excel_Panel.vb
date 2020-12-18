﻿Imports Microsoft.Office.Interop

Public Class Excel_Panel

    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Public Declare Auto Function SetParent Lib "user32.dll" (ByVal hWndChild As IntPtr, ByVal hWndNewParent As IntPtr) As Integer
    Public Declare Auto Function SendMessage Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Public Const WM_SYSCOMMAND As Integer = 274
    Public Const SC_MAXIMIZE As Integer = 61488
    Public Const GWL_STYLE = -16
    Public Const WS_CAPTION = &HC00000
    Public Const WS_THICKFRAME = &H40000


    Dim xlapp As Excel.Application
    Dim xlbook As Excel.Workbook
    Dim xlsheet As Excel.Worksheet
    Dim xlwin As Excel.Window
    Dim ex_start As Date
    Dim fn As String

    Private Sub Excel_Panel_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Visible = False
        Label2.Visible = False
        Label3.Visible = False
        txtValue.Visible = False
        txtX.Visible = False
        txtY.Visible = False
        btnRead.Visible = False
        btnWrite.Visible = False
    End Sub

    Private Sub Excel_panel_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        xlsheet = Nothing
        xlbook = Nothing
        xlapp.Quit()
        xlapp = Nothing
        '다른 프로그램에서 실행한 엑셀도 모두 닫힘
        'Dim xlp() As Process = Process.GetProcessesByName("EXCEL")
        'For Each Process As Process In xlp
        '    If Process.StartTime >= ex_start And Process.StartTime <= Date.Now Then
        '        Process.Kill()
        '        Exit For
        '    End If
        'Next
    End Sub

    Public Sub New(panelName As Panel, fileName As String)

        ' 디자이너에서 이 호출이 필요합니다.
        InitializeComponent()

        ' InitializeComponent() 호출 뒤에 초기화 코드를 추가하세요.
        ex_start = Date.Now
        panelName.Controls.Add(Me)
        Me.Dock = DockStyle.Fill
        Me.BringToFront()


        xlapp = New Excel.Application

        'filename은 hp1에서 보낸 이름 그대로 받음 

        ' MsgBox((Application.StartupPath & "\" & fileName)
        ' MsgBox(fileName)
        ' xlbook = xlapp.Workbooks.Open(fileName)
        xlbook = xlapp.Workbooks.Open(Application.StartupPath & "\" & fileName)
        fn = fileName
        '  xlbook = xlapp.Workbooks.Open(Application.StartupPath & "\" & fileName)
        xlsheet = xlbook.ActiveSheet
        xlapp.DisplayAlerts = False ' 닫을때 저장할 것인가 묻기
        xlapp.WindowState = Excel.XlWindowState.xlMinimized
        xlapp.Visible = False
        xlapp.DisplayFormulaBar = False
        xlapp.DisplayStatusBar = True

        'PInvoke 에러시 프로젝트 속성에서 컴파일을 x64로 하거나 Any CPU일 경우 32비트 기본 사용 체크를 해제한다.
        SetWindowLong(xlapp.Hwnd, GWL_STYLE, GetWindowLong(xlapp.Hwnd, GWL_STYLE) And Not WS_CAPTION)

        xlwin = xlapp.ActiveWindow
        xlwin.DisplayWorkbookTabs = False
        xlwin.DisplayHeadings = False
        xlwin.Zoom = TrackBar1.Value

        SetParent(xlapp.Hwnd, Panel1.Handle)
        SendMessage(xlapp.Hwnd, WM_SYSCOMMAND, SC_MAXIMIZE, 0)

    End Sub

    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        Try
            xlwin.Zoom = TrackBar1.Value
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub btnWrite_Click(sender As Object, e As EventArgs) Handles btnWrite.Click
        Try
            MsgBox("쓰기준비완료")
            '   xlsheet.Cells(Int(txtX.Text), Int(txtY.Text)).Value = txtValue.Text
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub btnRead_Click(sender As Object, e As EventArgs) Handles btnRead.Click
        Try
            txtValue.Text = xlsheet.Cells(Int(txtX.Text), Int(txtY.Text)).Value
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        '파일저장
        Try
            '  MsgBox("\Excel\HACCP\" & DateTime.Now.ToString("yyyyMMddHHmmss") & fn)
            'xlbook.SaveAs()
            xlbook.SaveAs(Application.StartupPath & "\Excel\HACCP\" & DateTime.Now.ToString("yyyy-MM-dd-HH-mm-") & fn)
            'xlbook.Save() '\그 자리 바로 저장
            '파라메터를 주면 다른이름으로 저장이 가능. 예를 들어서 연습.xlsx 을 연습202006241413.xlsx 이렇게 바꿀수있음.
            MsgBox("저장되었습니다")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        'DB에 저장 
        Try
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "Insert INTO HP_LIST (HP_DATE, HP_TIME, HP_NAME)  Values ('" & DateTime.Today.ToString("yyyy-MM-dd") & "','" & DateTime.Now.ToString("HH:mm:ss") & "','" & fn & "')"
            Re_Count = DB_Execute()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        xlsheet.PrintOut(From:=1, To:=1, Copies:=1, Collate:=True)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        xlsheet = Nothing
        xlbook = Nothing
        xlapp.Quit()
        xlapp = Nothing

        Dim xlp() As Process = Process.GetProcessesByName("EXCEL")
        For Each Process As Process In xlp
            Process.Kill()
        Next
    End Sub
End Class
