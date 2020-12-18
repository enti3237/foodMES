Public Class Gong
    Dim D_Code As String
    Dim Grid_Display_Ch As Boolean

    Dim Grid_Main_Row As Integer
    Dim Grid_Main_Col As Integer
    Dim Grid_Line_QC_Row As Integer
    Dim Grid_Line_QC_Col As Integer
    Dim Grid_Go_Row As Integer
    Dim Grid_Go_Col As Integer
    Private Sub Gong_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill


        Label1.Width = Panel_Menu.Width * 0.3
        Label1.Height = Panel_Menu.Height * 0.6
        Label1.Top = Panel_Menu.Height / 2 - Label1.Height / 2
        Label1.Left = Panel_Menu.Width / 2 - Label1.Width / 2

        Day_Label.Width = Panel_Menu.Width * 0.3
        Day_Label.Height = Panel_Menu.Height * 0.4
        Day_Label.Top = Panel_Menu.Height / 2 - Day_Label.Height
        Day_Label.Left = Panel_Menu.Width * 0.7

        Time_Label.Width = Panel_Menu.Width * 0.25
        Time_Label.Height = Panel_Menu.Height * 0.4
        Time_Label.Top = Panel_Menu.Height / 2
        Time_Label.Left = Day_Label.Left + Day_Label.Width / 2 - Time_Label.Width / 2

        Grid_Main.Width = Panel2.Width * 0.9
        Grid_Main.Height = Panel2.Height * 0.9
        Grid_Main.Top = Panel2.Height * 0.15
        Grid_Main.Left = Panel2.Width * 0.05
        Grid_Main.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)

        Grid_Main_Setup()
        Grid_Main_Display()
        Panel3.Visible = False
    End Sub

    Public Function Grid_Main_Setup() As Boolean
        Grid_Color(Grid_Main)

        ' Grid_Color(Grid_Main)

        'Grid_Main.GridColor = Color.White
        'Grid_Main.GridColor = Color.White
        'Grid_Main.BackgroundColor = Color.White
        'Grid_Main.EnableHeadersVisualStyles = False
        'Grid_Main.ColumnHeadersDefaultCellStyle.BackColor = Color.DarkOrange
        'Grid_Main.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        'Grid_Main.RowsDefaultCellStyle.BackColor = Color.FromArgb(210, 222, 239)
        'Grid_Main.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(234, 239, 247)

        'Grid_Main.AllowUserToAddRows = False
        Grid_Main.RowTemplate.Height = 120
        Grid_Main.ColumnHeadersHeight = 100
        Grid_Main.RowHeadersVisible = False

        Grid_Main.Font = New Font("맑은 고딕", 20, FontStyle.Bold)

        Grid_Main.ColumnCount = 16


        'Grid_Main.RowCount = 0
        Dim i As Integer

        For i = 0 To Grid_Main.ColumnCount - 1
            Grid_Main.Columns(i).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Next



        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Main.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Main.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Main.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Main.Columns(14).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Grid_Main.Columns(0).HeaderText = ""

        Grid_Main.Columns(1).HeaderText = "2"
        Grid_Main.Columns(1).Visible = False
        Grid_Main.Columns(2).HeaderText = "3"
        Grid_Main.Columns(2).Visible = False
        Grid_Main.Columns(3).HeaderText = "제품코드"
        Grid_Main.Columns(3).Visible = False
        Grid_Main.Columns(4).HeaderText = "제품명"

        Grid_Main.Columns(5).HeaderText = "6"
        Grid_Main.Columns(5).Visible = False
        Grid_Main.Columns(6).HeaderText = "7"
        Grid_Main.Columns(6).Visible = False
        Grid_Main.Columns(7).HeaderText = "수량"
        Grid_Main.Columns(8).HeaderText = "시작날짜"


        Grid_Main.Columns(9).HeaderText = "시작시간"
        Grid_Main.Columns(10).HeaderText = "종료날짜"
        Grid_Main.Columns(10).Visible = False
        Grid_Main.Columns(11).HeaderText = "종료시간"
        Grid_Main.Columns(11).Visible = False
        Grid_Main.Columns(12).HeaderText = "지시수량"
        Grid_Main.Columns(13).HeaderText = "총수량"
        Grid_Main.Columns(14).HeaderText = "실패수량"
        Grid_Main.Columns(14).Visible = False
        Grid_Main.Columns(15).HeaderText = "16"
        Grid_Main.Columns(15).Visible = False

        Grid_Main.Columns(14).DefaultCellStyle.ForeColor = Color.Green


        Grid_Main.Columns(0).Width = 60
        'Grid_Main.Columns(1).Width = 350

        'For i = 2 To 7
        '    Grid_Main.Columns(i).Width = 130
        'Next i

        'Grid_Main.Columns(1).Width = 220
        'Grid_Main.Columns(2).Width = 200
        'Grid_Main.Columns(3).Width = 200
        'Grid_Main.Columns(4).Width = 200
        'Grid_Main.Columns(5).Width = 170
        'Grid_Main.Columns(6).Width = 170

        'For i = 8 To 10
        '    Grid_Main.Columns(i).Width = 130
        'Next i


        'Grid_Main.Columns(11).Width = 150
        '    Grid_Main.DefaultCellStyle
        Grid_Main_Setup = True

        For i = 0 To Grid_Main.ColumnCount - 1
            Grid_Main.Columns(i).ReadOnly = True

        Next i
    End Function

    Public Function Grid_Main_Display() As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer

        StrSQL = "select * from PC_Stock_List where Check_G is NULL"
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


    End Function

    Private Sub Grid_Main_CellEnter(sender As Object, e As DataGridViewCellEventArgs)

        ' Panel3.Visible = True

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Panel3.Visible = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Grid_Main_Row = Grid_Main.CurrentCell.RowIndex
        Grid_Main_Col = Grid_Main.CurrentCell.ColumnIndex

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer

        StrSQL = "UPDATE PC_Stock_List SET Check_G='처리'
        where PS_Code = '" & Grid_Main.Item(1, Grid_Main.CurrentCell.RowIndex).Value & "' AND PS_PL_Code= '" & Grid_Main.Item(3, Grid_Main.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()


        Panel3.Visible = False
        Grid_Main_Display()
    End Sub
    Public Function AutoFontSize(label As Label, text As String) As Font
        Dim ft As Font
        Dim gp As Graphics
        Dim sz As SizeF
        Dim Faktor, FaktorX, FaktorY As Single
        gp = label.CreateGraphics()
        sz = gp.MeasureString(text, label.Font)
        gp.Dispose()

        FaktorX = (label.Width) / sz.Width
        FaktorY = (label.Height) / sz.Height
        If FaktorX > FaktorY Then
            Faktor = FaktorY
        Else
            Faktor = FaktorX
            ft = label.Font
        End If

        If Faktor > 0 Then
            Return New Font(ft.Name, ft.SizeInPoints * (Faktor))
        Else
            Return New Font(ft.Name, ft.SizeInPoints * (0.1F))
        End If

    End Function
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Day_Label.Text = Format(Now, "yyyy년 MM월 dd일")
        Time_Label.Text = Format(Now, "HH시 mm분")
    End Sub

    Private Sub Grid_Main_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        Panel3.Visible = True
    End Sub
End Class