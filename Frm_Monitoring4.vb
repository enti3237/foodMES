﻿Public Class Frm_Monitoring4
    Dim De_CC(100) As String
    Dim De_Su(100) As String
    Dim DBT As New DataTable
    Dim max As Integer
    Dim mok As Integer '몫
    Dim tail As Integer '나머지
    Dim count As Integer = 0
    Dim ten As Integer = 0
    Dim today As String = ""

    Private Sub Frm_Monitoring_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim DBT As New DataTable
        Dim k As Integer
        Dim i As Integer
        Label2.Text = DateTime.Now.ToString("yyyy-MM-dd")
        Label3.Text = DateTime.Now.ToString("HH:mm:ss")

        PIc_Pic(Picture_Logo, Application.StartupPath & "\Picture\Logo2_" & Setup_Data(9) & ".png")
        'PIc_Pic(Picture_Logo, Application.StartupPath & "\Picture\Logo2.png")


        Me.Dock = DockStyle.Fill
        Me.BringToFront()
        Grid_Main_Setup()
        Display()
    End Sub

    Private Sub Display()

        '총 보여줄개수 구하기

        StrSQL = "Select count(PL_CODE) AS a
                    FROM SI_PRODUCT"
        Re_Count = DB_Select(DBT)

        If Re_Count = 0 Then
            MsgBox("제품정보가 없습니다")
            Exit Sub
        Else
            max = DBT.Rows(0)("a")
        End If

        If ten > max Then
            ten = 0
        Else

        End If

        '10개씩 끊어서 보여주기
        Dim i As Integer
        If max < 11 Then
            '10개까지

            StrSQL = "Select distinct TOP 10 PL_CODE, PL_NAME, PL_STANDARD, PL_UNIT, PL_UNIT_GOLD, SUM(CONVERT(DECIMAL,pp_amount)) AS PP_AMOUNT
                  FROM SI_PRODUCT a
                  left outer join SI_PROCESS_LIST b on a.PL_Code = b.PP_Code 
                    GROUP BY PL_Code, PL_NAME, PL_Standard, PL_Unit, PL_Unit_Gold 
                  order by PL_CODE"
            Re_Count = DB_Select(DBT)

            If Re_Count <> 0 Then
                Grid_Main.RowCount = Re_Count
                For i = 0 To Re_Count - 1
                    For j = 0 To 1
                        Grid_Main.Item(0, i).Value = i + 1
                        Grid_Main.Item(1, i).Value = DBT.Rows(i).Item(0).ToString
                        Grid_Main.Item(2, i).Value = DBT.Rows(i).Item(1).ToString
                        Grid_Main.Item(3, i).Value = DBT.Rows(i).Item(2).ToString
                        Grid_Main.Item(4, i).Value = DBT.Rows(i).Item(3).ToString
                        Grid_Main.Item(5, i).Value = DBT.Rows(i).Item(4).ToString


                        If IsDBNull(DBT.Rows(i).Item(5)) Or DBT.Rows(i).Item(5) Is Nothing Or DBT.Rows(i).Item(0) Is "" Then
                            '시스템 내에서 값이 없을 경우
                            Grid_Main.Item(6, i).Value = "0"
                            Grid_Main.Rows(i).DefaultCellStyle.BackColor = Color.LightGreen
                        Else
                            '값 있음
                            Grid_Main.Item(6, i).Value = DBT.Rows(i).Item(5).ToString
                            If Grid_Main.Item(6, i).Value < 10 Or Grid_Main.Item(6, i).Value < 0 Then
                                Grid_Main.Rows(i).DefaultCellStyle.BackColor = Color.Red
                                Grid_Main.Rows(i).DefaultCellStyle.ForeColor = Color.Black
                            Else
                                Grid_Main.Rows(i).DefaultCellStyle.BackColor = Color.White
                                Grid_Main.Rows(i).DefaultCellStyle.ForeColor = Color.Black
                            End If

                        End If

                    Next j
                Next i
            Else
                MsgBox("제품이 존재하지 않습니다")
            End If
        Else
            '10개 이상
            mok = max / 10
            tail = max Mod 10


            StrSQL = "Select distinct TOP 10 PL_CODE, PL_NAME, PL_STANDARD, PL_UNIT, PL_UNIT_GOLD, SUM(CONVERT(DECIMAL,pp_amount)) AS PP_AMOUNT
                  FROM SI_PRODUCT a
                  left outer join SI_PROCESS_LIST b on a.PL_Code = b.PP_Code 
                       where PL_CODE not in (select top " & ten & " PL_CODE from si_product order by PL_Code)
                        GROUP BY PL_Code, PL_NAME, PL_Standard, PL_Unit, PL_Unit_Gold 
                        order by PL_CODE"
            Re_Count = DB_Select(DBT)

            Dim k As Integer
            If Re_Count <> 0 Then
                Grid_Main.RowCount = Re_Count
                'MsgBox(Grid_Main.RowCount)
                For k = 0 To Re_Count - 1
                    For j = 0 To 1
                        Grid_Main.Item(0, k).Value = k + 1
                        Grid_Main.Item(1, k).Value = DBT.Rows(k).Item(0).ToString
                        Grid_Main.Item(2, k).Value = DBT.Rows(k).Item(1).ToString
                        Grid_Main.Item(3, k).Value = DBT.Rows(k).Item(2).ToString
                        Grid_Main.Item(4, k).Value = DBT.Rows(k).Item(3).ToString
                        Grid_Main.Item(5, k).Value = DBT.Rows(k).Item(4).ToString


                        If IsDBNull(DBT.Rows(k).Item(5)) Or DBT.Rows(k).Item(5) Is Nothing Or DBT.Rows(k).Item(0) Is "" Then
                            Grid_Main.Item(6, k).Value = "0"
                            Grid_Main.Rows(k).DefaultCellStyle.BackColor = Color.LightGreen
                        Else
                            '값 있음
                            Grid_Main.Item(6, k).Value = DBT.Rows(k).Item(5).ToString
                            If Grid_Main.Item(6, k).Value < 10 Or Grid_Main.Item(6, k).Value < 0 Then
                                Grid_Main.Rows(k).DefaultCellStyle.BackColor = Color.Red
                                Grid_Main.Rows(k).DefaultCellStyle.ForeColor = Color.Black
                            Else
                                Grid_Main.Rows(k).DefaultCellStyle.BackColor = Color.White
                                Grid_Main.Rows(k).DefaultCellStyle.ForeColor = Color.Black
                            End If
                        End If

                    Next j
                Next k


            Else
                MsgBox("제품이 존재하지 않습니다")
            End If


        End If
        ten = ten + 10
        'StrSQL = "Select PL_CODE, PL_NAME, PL_STANDARD, PL_UNIT, PL_UNIT_GOLD FROM SI_PRODUCT ORDER BY PL_CODE"
        'Re_Count = DB_Select(DBT)
        'Dim i As Integer
        'If Re_Count <> 0 Then
        '    Grid_Main.RowCount = Re_Count
        '    For i = 0 To Re_Count - 1
        '        For j = 0 To 1
        '            Grid_Main.Item(0, i).Value = i + 1
        '            Grid_Main.Item(1, i).Value = DBT.Rows(i).Item(0).ToString
        '            Grid_Main.Item(2, i).Value = DBT.Rows(i).Item(1).ToString
        '            Grid_Main.Item(3, i).Value = DBT.Rows(i).Item(2).ToString
        '            Grid_Main.Item(4, i).Value = DBT.Rows(i).Item(3).ToString
        '            Grid_Main.Item(5, i).Value = DBT.Rows(i).Item(4).ToString
        '        Next j
        '    Next i
        'Else
        '    MsgBox("제품이 존재하지 않습니다")
        'End If

        'Dim k As Integer

        'For k = 0 To Grid_Main.Rows.Count - 1
        '    StrSQL = "Select PP_AMOUNT FROM SI_PROCESS_LIST WHERE PP_CODE = '" & Grid_Main.Item(1, k).Value & "' ORDER BY PP_CODE "
        '    Re_Count = DB_Select(DBT)

        '    If Re_Count <> 0 Then
        '        Grid_Main.RowCount = Re_Count
        '        For i = 0 To Re_Count - 1
        '            If IsDBNull(DBT.Rows(i).Item(0)) Or DBT.Rows(i).Item(0) Is Nothing Or DBT.Rows(i).Item(0) Is "" Then
        '                Grid_Main.Item(6, i).Value = "0"
        '                Grid_Main.Rows(i).DefaultCellStyle.BackColor = Color.LightGreen
        '                Continue For
        '            End If

        '            Grid_Main.Item(6, i).Value = DBT.Rows(i).Item(0).ToString
        '            If DBT.Rows(i).Item(0).ToString < 10 Then
        '                Grid_Main.Rows(i).DefaultCellStyle.BackColor = Color.Red
        '                Grid_Main.Rows(i).DefaultCellStyle.ForeColor = Color.White
        '            Else
        '                Grid_Main.Rows(i).DefaultCellStyle.BackColor = Color.White
        '                Grid_Main.Rows(i).DefaultCellStyle.ForeColor = Color.Black
        '            End If
        '        Next i
        '    Else
        '        'MsgBox("제품이 존재하지 않습니다")
        '    End If
        'Next


        Grid_Main.MultiSelect = False
        Grid_Main.ClearSelection()

    End Sub

    Public Function Grid_Main_Setup() As Boolean
        Grid_Color(Grid_Main)

        Grid_Main.AllowUserToAddRows = False
        Grid_Main.RowTemplate.Height = 60
        Grid_Main.ColumnHeadersHeight = 70
        Grid_Main.ColumnCount = 7
        Grid_Main.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Main.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Main.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Main.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Main.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Main.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter


        Grid_Main.RowHeadersWidth = 10


        Grid_Main.Columns(0).HeaderText = "순번"
        Grid_Main.Columns(1).HeaderText = "제품코드"
        Grid_Main.Columns(1).Visible = False
        Grid_Main.Columns(2).HeaderText = "제품명"
        Grid_Main.Columns(3).HeaderText = "규격"
        Grid_Main.Columns(4).HeaderText = "단위"
        Grid_Main.Columns(4).Visible = False
        Grid_Main.Columns(5).HeaderText = "단가"
        Grid_Main.Columns(5).Visible = False
        Grid_Main.Columns(6).HeaderText = "수량"




        Dim i As Integer

        Grid_Main.Columns(0).Width = 120
        Grid_Main.Columns(1).Width = 120
        Grid_Main.Columns(2).Width = 200

        Grid_Main.Columns(3).Width = 120

        Grid_Main.Columns(6).Width = 200

        'For i = 3 To Grid_Main.ColumnCount - 1

        '    Grid_Main.Columns(i).Width = 120
        'Next i

        Grid_Main_Setup = True
    End Function

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
        Label3.Text = DateTime.Now.ToString("HH:mm:ss")
    End Sub


    Private Sub Picture_Logo_Click(sender As Object, e As EventArgs) Handles Picture_Logo.Click
        Me.Close()

        'End

    End Sub

    Private Sub Label21_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        '  MsgBox(max)
        Display()
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        End
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        End
    End Sub

    Private Sub panel33_Paint(sender As Object, e As PaintEventArgs) Handles panel33.Paint

    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs)
        '10개씩 보여줄 시간 30초씩
    End Sub
End Class