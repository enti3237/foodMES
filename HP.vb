Public Class HP
    Private Sub HP_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        S_Text_Setup()

        Label1.Width = Panel_Menu.Width * 0.3
        Label1.Height = Panel_Menu.Height * 0.6
        Label1.Top = Panel_Menu.Height / 2 - Label1.Height / 2
        Label1.Left = Panel_Menu.Width / 2 - Label1.Width / 2

        DataGridView1_Setup()
        DataGridView1_Display()
        DataGridView1.Width = Panel2.Width * 0.9
        DataGridView1.Height = Panel2.Height * 0.7
        DataGridView1.Top = Panel2.Height * 0.25
        DataGridView1.Left = Panel2.Width * 0.05


        Day_Label.Width = Panel_Menu.Width * 0.3
        Day_Label.Height = Panel_Menu.Height * 0.4
        Day_Label.Top = Panel_Menu.Height / 2 - Day_Label.Height
        Day_Label.Left = Day_Label.Left + Day_Label.Width - Time_Label.Width / 3

        Time_Label.Width = Panel_Menu.Width * 0.25
        Time_Label.Height = Panel_Menu.Height * 0.4
        Time_Label.Top = Panel_Menu.Height / 2
        Time_Label.Left = Day_Label.Left * 1.09

        'DataGridView1.Width = Panel2.Width * 0.9
        'DataGridView1.Height = Panel2.Height * 0.7
        'DataGridView1.Top = Panel2.Height * 0.1
        'DataGridView1.Left = Panel2.Width * 0.3
        Button2.Left = Panel2.Width * 0.05
        Button2.Width = Panel2.Width * 0.6
        Button2.Height = Panel2.Width * 0.025

        TextBox1.Width = Panel2.Width * 0.2
        TextBox1.Height = Panel2.Width * 0.025
        TextBox1.Left = Panel2.Width * 0.05 + Button2.Width


        Button1.Left = Panel2.Width * 0.05 + Button2.Width
        Button1.Width = Panel2.Width * 0.2
        Button1.Height = Panel2.Width * 0.025




        S_Day.Width = Panel2.Width * 0.19
        S_Day.Height = Panel2.Width * 0.025
        S_Day.Top = Panel2.Width * 0.11
        S_Day.Left = Panel2.Width * 0.05

        S_Time.Height = Panel2.Width * 0.025
        S_Time.Width = Panel2.Width * 0.09
        S_Time.Top = Panel2.Width * 0.11
        S_Time.Left = Panel2.Width * 0.05 + Panel2.Width * 0.2

        Label2.Width = Panel2.Width * 0.01
        Label2.Height = Panel2.Width * 0.01
        Label2.Top = Panel2.Width * 0.12
        Label2.Left = Panel2.Width * 0.061 + S_Day.Width + S_Time.Width

        E_Day.Width = Panel2.Width * 0.19
        E_Day.Height = Panel2.Width * 0.025
        E_Day.Top = Panel2.Width * 0.11
        E_Day.Left = Panel2.Width * 0.05 + Panel2.Width * 0.2 + Panel2.Width * 0.1



        E_Time.Width = Panel2.Width * 0.09
        E_Time.Height = Panel2.Width * 0.025
        E_Time.Top = Panel2.Width * 0.11
        E_Time.Left = Panel2.Width * 0.05 + Panel2.Width * 0.2 + Panel2.Width * 0.1 + Panel2.Width * 0.2


        Search_Com.Width = Panel2.Width * 0.1
        Search_Com.Height = Panel2.Width * 0.05
        Search_Com.Left = Panel2.Width * 0.85
    End Sub

    Public Function DataGridView1_Setup() As Boolean
        Grid_Color(DataGridView1)

        DataGridView1.AllowUserToAddRows = False
        DataGridView1.RowTemplate.Height = 100
        DataGridView1.ColumnCount = 3
        DataGridView1.RowCount = 0

        DataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Blue
        DataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.Gold

        DataGridView1.RowHeadersWidth = 5

        DataGridView1.Columns(0).HeaderText = "창고명"
        DataGridView1.Columns(1).HeaderText = "기준온도"
        DataGridView1.Columns(2).HeaderText = "실제온도"


        'DataGridView1.Columns(0).Width = 10



        DataGridView1_Setup = True
    End Function

    Private Sub DataGridView1_Display()

        Dim row As Integer
        Dim DBT As New DataTable

        Dim i, j As Integer


        DataGridView1.Visible = True

        StrSQL = "Select TOP 50 FL_NAME, FL_ST, N_TIME, C_TIME, C_DATA
                            FROM SI_STOCK with(NOLOCK) join iotlog b on SI_STOCK.FL_IP = b.C_IP 
                            Order By N_TIME DESC, C_TIME DESC"


        Re_Count = DB_Select(DBT)

        If Re_Count <> 0 Then
            DataGridView1.RowCount = Re_Count


            For i = 0 To Re_Count - 1
                If IsDBNull(DBT.Rows(i)("FL_NAME")) Then
                    MsgBox("값이 없습니다")
                    Continue For
                End If
                DataGridView1.Item(0, i).Value = DBT.Rows(i)("FL_NAME")
                DataGridView1.Item(1, i).Value = DBT.Rows(i)("FL_ST")
                DataGridView1.Item(2, i).Value = DBT.Rows(i)("C_DATA")
                'DataGridView1.Item(4, i).Value = DBT.Rows(i)("C_TIME")
                'aGridView1.Item(2, i).Value = DBT.Rows(i)("C_DATA")
            Next i
        End If


    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub
    Public Function S_Text_Setup() As Boolean
        E_Day.Text = Format(Now, "yyyy-MM-dd")
        '  S_Day.Text = Format(Now, "yyyyMM01")
        S_Day.Text = Mid(E_Day.Text, 1, 8) & "01"
        S_Time.Text = "00:00"
        E_Time.Text = "23:59"

        S_Text_Setup = True

    End Function

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        Dim row As Integer
        Dim DBT As New DataTable

        Dim i, j As Integer


        DataGridView1.Visible = True

        StrSQL = "Select TOP 50 FL_NAME, FL_ST, N_TIME, C_TIME, C_DATA
                            FROM SI_STOCK with(NOLOCK) join iotlog b on SI_STOCK.FL_IP = b.C_IP 
where N_TIME BETWEEN '" & S_Day.Text & "' and '" & E_Day.Text & "'
                            Order By N_TIME DESC, C_TIME DESC"


        Re_Count = DB_Select(DBT)

        If Re_Count <> 0 Then
            DataGridView1.RowCount = Re_Count


            For i = 0 To Re_Count - 1
                If IsDBNull(DBT.Rows(i)("FL_NAME")) Then
                    MsgBox("값이 없습니다")
                    Continue For
                End If
                DataGridView1.Item(0, i).Value = DBT.Rows(i)("FL_NAME")
                DataGridView1.Item(1, i).Value = DBT.Rows(i)("FL_ST")
                DataGridView1.Item(2, i).Value = DBT.Rows(i)("N_TIME")
                'DataGridView1.Item(4, i).Value = DBT.Rows(i)("C_TIME")
                'aGridView1.Item(2, i).Value = DBT.Rows(i)("C_DATA")
            Next i
        End If
    End Sub
End Class