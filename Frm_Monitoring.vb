﻿Public Class Frm_Monitoring
    Dim De_CC(100) As String
    Dim De_Su(100) As String
    Dim count As Integer = 0
    Dim today As String = ""

    Private Sub Frm_Monitoring_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        PIc_Pic(Picture_Logo, Application.StartupPath & "\Picture\Logo2_" & Setup_Data(9) & ".png")
        'PIc_Pic(Picture_Logo, Application.StartupPath & "\Picture\Logo2.png")

        Grid_Main_Setup()
        Grid_Main_Display()

        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.WindowState = FormWindowState.Maximized

        '  Label5.Text = ""
        Dim DBT As New DataTable


        ' Label7.Text = ""
        '  Panel1.Visible = False
        ' Panel2.Visible = False
        'Panel3.Visible = False


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
        Day_Label.Text = Format(Now, "yyyy년 MM월 dd일")
        Time_Label.Text = Format(Now, "HH시 mm분 ss초")
        Grid_Main_Display()
        '공지사항
        Dim DBT As New DataTable
        StrSQL = "Select * From QC_Info with(NOLOCK) Where QC_Title = '종합상황판 공지' Order By QC_CONTENT"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then

            '예외 처리 추가 코드  
            If Re_Count = 1 Then
                If DBT.Rows(0)("QC_MAN") = "" Then
                    DBT.Rows(0)("QC_MAN") = ""
                End If
            End If
            If Re_Count = 2 Then
                If DBT.Rows(0)("QC_MAN") = "" Then
                    DBT.Rows(0)("QC_MAN") = ""
                End If
                If DBT.Rows(1)("QC_MAN") = "" Then
                    DBT.Rows(1)("QC_MAN") = ""
                End If
            End If
            If Re_Count = 3 Then
                If DBT.Rows(0)("QC_MAN") = "" Then
                    DBT.Rows(0)("QC_MAN") = ""
                End If
                If DBT.Rows(1)("QC_MAN") = "" Then
                    DBT.Rows(1)("QC_MAN") = ""
                End If
                If DBT.Rows(2)("QC_MAN") = "" Then
                    DBT.Rows(2)("QC_MAN") = ""
                End If
            End If
            If Re_Count >= 4 Then
                If DBT.Rows(0)("QC_MAN") = "" Then
                    DBT.Rows(0)("QC_MAN") = ""
                End If
                If DBT.Rows(1)("QC_MAN") = "" Then
                    DBT.Rows(1)("QC_MAN") = ""
                End If
                If DBT.Rows(2)("QC_MAN") = "" Then
                    DBT.Rows(2)("QC_MAN") = ""
                End If
                If DBT.Rows(3)("QC_MAN") = "" Then
                    DBT.Rows(3)("QC_MAN") = ""
                End If
            End If

            If Re_Count = 1 Then
                TextBox1.Text = DBT.Rows(0)("QC_MAN")
            ElseIf Re_Count = 2 Then
                TextBox1.Text = DBT.Rows(0)("QC_MAN")
                TextBox2.Text = DBT.Rows(1)("QC_MAN")
            ElseIf Re_Count = 3 Then
                TextBox1.Text = DBT.Rows(0)("QC_MAN")
                TextBox2.Text = DBT.Rows(1)("QC_MAN")
                TextBox3.Text = DBT.Rows(2)("QC_MAN")
            ElseIf Re_Count >= 4 Then
                TextBox1.Text = DBT.Rows(0)("QC_MAN")
                TextBox2.Text = DBT.Rows(1)("QC_MAN")
                TextBox3.Text = DBT.Rows(2)("QC_MAN")
                TextBox4.Text = DBT.Rows(3)("QC_MAN")
            End If

        End If

        StrSQL = "Select isnull( QC_MAN, '') as QC_MAN From QC_Info with(NOLOCK) Where QC_Title = '종합상황판 품질' Order By QC_CONTENT" '종합상황판 품질
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then

            '예외 처리 추가 코드
            If Re_Count = 1 Then
                If DBT.Rows(0)("QC_MAN") = "" Then
                    DBT.Rows(0)("QC_MAN") = ""
                End If
            End If
            If Re_Count = 2 Then
                If DBT.Rows(0)("QC_MAN") = "" Then
                    DBT.Rows(0)("QC_MAN") = ""
                End If
                If DBT.Rows(1)("QC_MAN") = "" Then
                    DBT.Rows(1)("QC_MAN") = ""
                End If
            End If
            If Re_Count = 3 Then
                If DBT.Rows(0)("QC_MAN") = "" Then
                    DBT.Rows(0)("QC_MAN") = ""
                End If
                If DBT.Rows(1)("QC_MAN") = "" Then
                    DBT.Rows(1)("QC_MAN") = ""
                End If
                If DBT.Rows(2)("QC_MAN") = "" Then
                    DBT.Rows(2)("QC_MAN") = ""
                End If
            End If
            If Re_Count >= 4 Then
                If DBT.Rows(0)("QC_MAN") = "" Then
                    DBT.Rows(0)("QC_MAN") = ""
                End If
                If DBT.Rows(1)("QC_MAN") = "" Then
                    DBT.Rows(1)("QC_MAN") = ""
                End If
                If DBT.Rows(2)("QC_MAN") = "" Then
                    DBT.Rows(2)("QC_MAN") = ""
                End If
                If DBT.Rows(3)("QC_MAN") = "" Then
                    DBT.Rows(3)("QC_MAN") = ""
                End If
            End If


            If Re_Count = 1 Then
                TextBox5.Text = DBT.Rows(0)("QC_MAN")
            ElseIf Re_Count = 2 Then
                TextBox5.Text = DBT.Rows(0)("QC_MAN")
                TextBox6.Text = DBT.Rows(1)("QC_MAN")
            ElseIf Re_Count = 3 Then
                TextBox5.Text = DBT.Rows(0)("QC_MAN")
                TextBox6.Text = DBT.Rows(1)("QC_MAN")
                TextBox7.Text = DBT.Rows(2)("QC_MAN")
            ElseIf Re_Count >= 4 Then
                TextBox5.Text = DBT.Rows(0)("QC_MAN")
                TextBox6.Text = DBT.Rows(1)("QC_MAN")
                TextBox7.Text = DBT.Rows(2)("QC_MAN")
                TextBox8.Text = DBT.Rows(3)("QC_MAN")
            End If

            'TextBox5.Text = DBT.Rows(0)("QC_MAN")
            'TextBox6.Text = DBT.Rows(1)("QC_MAN")
            'TextBox7.Text = DBT.Rows(2)("QC_MAN")
            'TextBox8.Text = DBT.Rows(3)("QC_MAN")
        End If


        '    Label7.Text = Format(Now, "yyyy-MM-dd")

        ' Label5.Text = "무사고 " & DateDiff("d", Label6.Text, Label7.Text) & "일"


    End Sub

    Public Function Grid_Main_Setup() As Boolean

        '  Grid_Color(Grid_Main)

        'Grid_Main.GridColor = Color.White
        Grid_Main.GridColor = Color.DarkBlue
        Grid_Main.CellBorderStyle = DataGridViewCellBorderStyle.None
        Grid_Main.RowHeadersBorderStyle = DataGridViewCellBorderStyle.None

        Grid_Main.BackgroundColor = Color.DarkBlue
        Grid_Main.EnableHeadersVisualStyles = False
        Grid_Main.ColumnHeadersDefaultCellStyle.BackColor = Color.DarkBlue
        Grid_Main.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        Grid_Main.RowsDefaultCellStyle.BackColor = Color.DarkBlue
        Grid_Main.RowsDefaultCellStyle.ForeColor = Color.White
        Grid_Main.AlternatingRowsDefaultCellStyle.BackColor = Color.DarkBlue
        Grid_Main.AlternatingRowsDefaultCellStyle.ForeColor = Color.White
        Grid_Main.AllowUserToAddRows = False


        Select Case Setup_Data(9)
            Case "동영산업"
                Grid_Main.RowTemplate.Height = 42
                Grid_Main.ColumnHeadersHeight = 70
                Me.BackColor = Color.DarkBlue
            Case "정우프린텍"
                Grid_Main.RowTemplate.Height = 84
                Grid_Main.ColumnHeadersHeight = 113
            Case "동아프로세스"
                Grid_Main.RowTemplate.Height = 100
                Grid_Main.ColumnHeadersHeight = 113
                Me.BackColor = Color.LightSeaGreen
            Case Else
                Grid_Main.RowTemplate.Height = 53
                Grid_Main.ColumnHeadersHeight = 76
        End Select

        Grid_Main.RowHeadersVisible = False



        Grid_Main.ColumnCount = 10


        Grid_Main.RowCount = 0
        Dim i As Integer

        For i = 0 To Grid_Main.ColumnCount - 1
            Grid_Main.Columns(i).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Next



        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Main.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Main.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Main.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Main.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Main.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Grid_Main.Columns(0).HeaderText = "장비명"
        Grid_Main.Columns(1).HeaderText = "생산실적번호"
        Grid_Main.Columns(2).HeaderText = "제품코드"
        Grid_Main.Columns(3).HeaderText = "제품명"
        Grid_Main.Columns(4).HeaderText = "시작"
        Grid_Main.Columns(5).HeaderText = "완료"
        Grid_Main.Columns(6).HeaderText = "투입수량"
        Grid_Main.Columns(7).HeaderText = "생산수량"
        Grid_Main.Columns(8).HeaderText = "불량수량"
        Grid_Main.Columns(9).HeaderText = "구분"




        Grid_Main.Columns(0).Width = 150
        Grid_Main.Columns(1).Width = 250
        Grid_Main.Columns(2).Width = 180
        Grid_Main.Columns(3).Width = 270
        Grid_Main.Columns(4).Width = 100
        Grid_Main.Columns(5).Width = 100
        Grid_Main.Columns(6).Width = 92
        Grid_Main.Columns(7).Width = 92
        Grid_Main.Columns(8).Width = 92
        Grid_Main.Columns(9).Width = 100


        'For i = 0 To Grid_Main.ColumnCount - 1
        '    Grid_Main.Columns(i).Width = 50
        'Next i

        Grid_Main_Setup = True
    End Function
    Public Function Grid_Main_Display() As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Main.RowCount = 0

        '생산수량 가져올 iot 확인

        StrSQL = "Select * From IOT With(NOLOCK) Order By PL_Code"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Main.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                De_CC(i) = DBT.Rows(i)("PL_Code")
                Grid_Main.Item(0, i).Value = DBT.Rows(i)("PL_Name")
                Grid_Main.Item(1, i).Value = DBT.Rows(i)("PL_PS_Code")
                Grid_Main.Item(2, i).Value = DBT.Rows(i)("PL_PL_CODE")
                Grid_Main.Item(3, i).Value = DBT.Rows(i)("PL_PL_NAME")
                Grid_Main.Item(4, i).Value = Mid(DBT.Rows(i)("PL_ST_Time").ToString, 1, 5)
                Grid_Main.Item(5, i).Value = Mid(DBT.Rows(i)("PL_En_Time").ToString, 1, 5)
                De_Su(i) = DBT.Rows(i)("PL_Go").ToString
                If Len(De_Su(i)) < 1 Then
                    De_Su(i) = "0"
                End If

                If DBT.Rows(i)("PL_Ing").ToString = "Ing" Or DBT.Rows(i)("PL_Ing").ToString = "ing" Then
                    Grid_Main.Item(9, i).Value = "가동중"
                Else
                    Grid_Main.Item(9, i).Value = "비가동"
                End If
            Next i
        Else
            Grid_Main.RowCount = 1
            For j = 0 To Grid_Main.ColumnCount - 1
                Grid_Main.Item(j, 0).Value = ""
            Next j
        End If


        For i = 0 To Grid_Main.Rows.Count - 1
            '생산수량
            StrSQL = "Select Sum(Convert(Decimal,PS_Go)) As E1, Sum(Convert(Decimal,PS_Total)) As E2, Sum(Convert(Decimal,PS_Er)) As E3 From PC_STOCK With(NOLOCK), PC_STOCK_List With(NOLOCK)  Where PC_STOCK.PS_DE_Code  = '" & De_CC(i) & "' And  PC_STOCK.PS_Code = PC_STOCK_List.PS_Code And  len(PS_Go) > 0 And PS_Go is Not null And len(PS_Total) > 0 And PS_Total is Not null And len(PS_Er) > 0 And PS_Er is Not null And PS_Date = '" & Format(Now, "yyyy-MM-dd") & "'"
            Re_Count = DB_Select(DBT)
            If Re_Count <> 0 Then
                'j = DBT.Rows(0)("E1")
                'Grid_Main.Item(6, i).Value = Val(DBT.Rows(0)("E1").ToString)
                Grid_Main.Item(6, i).Value = Val(DBT.Rows(0)("E1").ToString) + Val(De_Su(i))
                If Grid_Main.Item(6, i).Value = "0" Then
                    Grid_Main.Item(6, i).Value = De_Su(i)
                End If
                Grid_Main.Item(7, i).Value = Val(DBT.Rows(0)("E2").ToString)
                Grid_Main.Item(8, i).Value = Val(DBT.Rows(0)("E3").ToString)
            End If


        Next i


        Grid_Main_Display = True
        Grid_Main.ClearSelection()
    End Function

    Private Sub Picture_Logo_Click(sender As Object, e As EventArgs) Handles Picture_Logo.Click
        Me.Close()

        'End

    End Sub

End Class