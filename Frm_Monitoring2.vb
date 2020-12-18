Public Class Frm_Monitoring2
    Dim De_CC(100) As String
    Dim De_Su(100) As String
    Dim count As Integer = 0
    Dim today As String = ""


    Dim c() As Label '창고명 (DB에 들어가면 바로씀)
    Dim ii() As Label '온도
    Dim p() As Label 'ip 또는 그 외적인 것  


    Private Sub Frm_Monitoring_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim DBT As New DataTable
        Dim k As Integer
        Dim i As Integer



        PIc_Pic(Picture_Logo, Application.StartupPath & "\Picture\Logo2_" & Setup_Data(9) & ".png")
        'PIc_Pic(Picture_Logo, Application.StartupPath & "\Picture\Logo2.png")

        c = New Label() {C1, C2, C3, C4, C5}
        ii = New Label() {i1, i2, i3, i4, i5}
        p = New Label() {p1, p2, p3, p4, p5}

        Me.Dock = DockStyle.Fill
        Me.BringToFront()

        StrSQL = "SELECT GETDATE()"

        Re_Count = DB_Select(DBT)

        If Re_Count = 0 Then
            Exit Sub
        Else

            today = DBT.Rows(0).Item(0)
            today = Mid(today, 1, 4) & Mid(today, 6, 2) & Mid(today, 9, 2)
        End If


        c(0).Text = "투입수량"
        c(1).Text = "검출수량"
        c(2).Text = "생산수량"
        c(3).Text = "검출율"
        c(4).Text = "불량률"

        For i = 0 To 4
            p(i).Visible = False
            ii(i).Text = "0"
            '예시데이터 넣어보기
            ' ii(0).Text = "100"
            ' ii(1).Text = "-"
        Next
        '  Display()

        c(0).Visible = False
        ii(0).Visible = False
        p(0).Visible = False

        c(1).Visible = False
        ii(1).Visible = False
        p(1).Visible = False

        c(3).Visible = False
        ii(3).Visible = False
        p(3).Visible = False

        c(4).Visible = False
        ii(4).Visible = False
        p(4).Visible = False

    End Sub

    Private Sub Display()

        Dim DBT As New DataTable
        Dim DBT2 As New DataTable

        Dim Re_Count2 As Long
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer



        ''설비명으로 찾을 경우 -> 설비컬럼으로 바꿔주기
        'StrSQL = "SELECT C_TEMP
        '        FROM IOT_METAL
        '        WHERE C_DATE = '" & today & "' and C_Name ='" & c(i).Text & "'"


        'ip로 찾아올 경우 -> ip가 있는 정보를 따로 저장해서 비교할것
        'StrSQL = "SELECT C_TEMP
        '        FROM IOT_METAL
        '        WHERE C_DATE = '" & today & "' and C_IP ='" & p(i).Text & "'"

        '투입수량
        StrSQL = "SELECT count(C_DATA)
                    FROM IOTLOG
                    WHERE C_IP = '230'"

        Re_Count = DB_Select(DBT)

        If Re_Count = 0 Then
            ii(2).Text = "-"
        Else
            ii(2).Text = DBT.Rows(0).Item(0)
        End If


        ''불량 검출
        'StrSQL = "SELECT COUNT
        '            FROM IOT_MACHINE
        '            WHERE MACHINE_IP = '230"

        'Re_Count = DB_Select(DBT)

        '    If IsDBNull(Re_Count) Then
        '    ii(1).Text = "-"
        'Else
        '    ii(1).Text = DBT.Rows(0).Item(0)
        'End If



        'Dim 투입수량 As Integer
        'Dim 검출수량 As Integer
        'Dim 생산수량 As Integer

        ''생산수량
        'If ii(0).Text = "-" Or ii(0).Text = "0" Then
        '    ii(2).Text = "-"
        '    ii(3).Text = "-"
        '    Exit Sub
        'Else

        '    If ii(1).Text = "-" Then
        '        ii(1).Text = "0"
        '    End If

        '    투입수량 = Val(ii(0).Text)
        '    검출수량 = Val(ii(1).Text)
        '    생산수량 = Val((투입수량 - 검출수량))

        '    ii(2).Text = 생산수량.ToString
        '    ii(3).Text = (Math.Truncate(검출수량 * 100 / 투입수량)).ToString & "%"
        'End If


    End Sub

    Private Sub Frm_Monitoring_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles Me.MouseDoubleClick
        Dim dualMonitor() As Screen = Screen.AllScreens
        If dualMonitor.Length = 2 Then
            If Me.Bounds = dualMonitor(1).Bounds Then
                Me.Bounds = dualMonitor(0).WorkingArea
            Else
                Me.Bounds = dualMonitor(1).WorkingArea
            End If
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        Display()

    End Sub


    Private Sub Picture_Logo_Click(sender As Object, e As EventArgs) Handles Picture_Logo.Click
        Me.Close()
        ' End
    End Sub

    Private Sub Label21_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label40_Click(sender As Object, e As EventArgs)

    End Sub
End Class