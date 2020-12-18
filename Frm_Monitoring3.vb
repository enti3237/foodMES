Public Class Frm_Monitoring3
    Dim De_CC(100) As String
    Dim De_Su(100) As String
    Dim count As Integer = 0
    Dim today As String = ""
    Dim max As Integer

    Dim c() As Label '창고명 (DB에 들어가면 바로씀)
    Dim ii() As Label '온도
    Dim p() As Label 'ip 또는 그 외적인 것  


    Private Sub Frm_Monitoring_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim DBT As New DataTable
        Dim k As Integer
        Dim i As Integer



        PIc_Pic(Picture_Logo, Application.StartupPath & "\Picture\Logo2_" & Setup_Data(9) & ".png")
        'PIc_Pic(Picture_Logo, Application.StartupPath & "\Picture\Logo2.png")

        c = New Label() {C1, C2, C3, C4, C5, C6}
        ii = New Label() {i1, i2, i3, i4, i5, i6}
        p = New Label() {p1, p2, p3, p4, p5, P6}

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


        StrSQL = "SELECT COUNT(FL_CODE) AS A FROM SI_STOCK"
        Re_Count = DB_Select(DBT)
        max = DBT.Rows(0)("A")
        If Re_Count = 0 Then
            MsgBox("창고정보가 존재하지 않습니다")
            Exit Sub
        Else
            If Re_Count < 7 Then
                For i = 0 To 5

                    p(i).Visible = False

                    '설비명 붙이기
                    StrSQL = "SELECT FL_NAME, FL_IP FROM SI_STOCK"
                    Re_Count = DB_Select(DBT)

                    If Re_Count <> 0 Then
                        If i > max Then
                            ' c(i).Visible = False
                            ' ii(i).Visible = False
                            c(i).Text = DBT.Rows(i)("FL_NAME")
                            ii(i).Text = "-"
                            p(i).Text = DBT.Rows(i)("FL_IP")

                        Else
                            If IsDBNull(DBT.Rows(i)("FL_NAME")) Then
                                c(i).Text = "-"
                                ii(i).Text = "-"
                            ElseIf IsDBNull(DBT.Rows(i)("FL_NAME")) Then
                                p(i).Text = "-"
                                ii(i).Text = "-"
                            Else
                                c(i).Text = DBT.Rows(i)("FL_NAME")
                                ii(i).Text = "-"
                                p(i).Text = DBT.Rows(i)("FL_IP")
                            End If

                        End If

                    Else
                        Exit Sub
                    End If
                Next
            Else
                For i = 0 To 5
                    p(i).Visible = False

                    '설비명 붙이기
                    StrSQL = "SELECT FL_NAME, FL_IP FROM SI_STOCK"
                    Re_Count = DB_Select(DBT)

                    If Re_Count <> 0 Then
                        c(i).Text = DBT.Rows(i)("FL_NAME")
                        ii(i).Text = ""
                        p(i).Text = DBT.Rows(i)("FL_IP")
                    Else
                        Exit Sub
                    End If
                Next
            End If

        End If


        'ip 설정해주기 p(i).text

        Display()

        '

    End Sub

    Private Sub Display()

        Dim DBT As New DataTable
        Dim DBT2 As New DataTable

        Dim Re_Count2 As Long
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer

        For i = 0 To max - 1
            '설비명으로 찾을 경우 -> 설비컬럼으로 바꿔주기
            'StrSQL = "SELECT C_TEMP
            '        FROM IOTLOG
            '        WHERE C_DATE = '" & today & "' and C_Name ='" & c(i).Text & "'"

            'MsgBox(max)
            'ip로 찾아올 경우 -> ip가 있는 정보를 따로 저장해서 비교할것
            StrSQL = "SELECT count(C_DATA) AS A
                    FROM IOTLOG
                    WHERE C_DATE = '" & today & "' and C_IP ='" & p(i).Text & "' GROUP BY C_DATE, C_TIME, C_DATA ORDER BY C_DATE DESC,C_TIME DESC"

            Re_Count = DB_Select(DBT)

            If Re_Count = 0 Then
                ii(i).Text = "-"
                Continue For
            Else
                If DBT.Rows(0)("A") = 0 Then
                    ii(i).Text = "-"
                Else
                    ''ip로 찾아올 경우 -> ip가 있는 정보를 따로 저장해서 비교할것
                    StrSQL = "SELECT C_DATA
                    FROM IOTLOG
                    WHERE C_DATE = '" & today & "' and C_IP ='" & p(i).Text & "' GROUP BY C_DATE, C_TIME, C_DATA ORDER BY C_DATE DESC,C_TIME DESC"
                    Re_Count = DB_Select(DBT)
                    If IsDBNull(DBT.Rows(i).Item("C_DATA")) Then
                        ii(i).Text = "--"
                    Else
                        ii(i).Text = DBT.Rows(i).Item("C_DATA")
                    End If

                End If




            End If
        Next



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
        '  End
    End Sub

    Private Sub Label21_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label40_Click(sender As Object, e As EventArgs)

    End Sub
End Class