Public Class Out3
    Dim D_Code As String
    Dim Grid_Display_Ch As Boolean

    Dim DataGridView1_Row As Integer
    Dim DataGridView1_Col As Integer

    Dim Grid_Line_QC_Row As Integer
    Dim Grid_Line_QC_Col As Integer
    Dim Grid_Go_Row As Integer
    Dim Grid_Go_Col As Integer

    Private Sub Out3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        FileOpen(1, Application.StartupPath + "\Setup.txt", OpenMode.Input)
        Setup_Data(1) = LineInput(1)

        FileClose(1)
        D_Code = Mid(Setup_Data(1), 8, 3)
        '장비 이름
        DataGridView1_Setup()
        DataGridView1_Display()
        DataGridView1.Width = Panel2.Width * 0.9
        DataGridView1.Height = Panel2.Height * 0.9
        DataGridView1.Top = Panel2.Height * 0.15
        DataGridView1.Left = Panel2.Width * 0.05

        DataGridView1.Font = New Font("맑은 고딕", 20, FontStyle.Bold)
        Panel3.Top = DataGridView1.Height * 0.4
        Panel3.Left = DataGridView1.Width * 0.4


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

        Panel3.Visible = False

    End Sub

    Public Function DataGridView1_Display() As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer


        DataGridView1.RowCount = 0
        StrSQL = "select PC_SHIP_LIST.DL_Code , DL_Customer_Name , DL_PL_Code , DL_PL_Name , DL_Total , PC_SHIP_LIST.DL_Bigo , PC_SHIP_LIST.DL_Management 
from PC_SHIP JOIN PC_SHIP_LIST ON PC_SHIP.DL_Code = PC_SHIP_LIST.DL_Code where PC_SHIP_LIST.DL_Management='진행중'"

        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            DataGridView1.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                DataGridView1.Item(0, i).Value = i + 1
                For j = 1 To DataGridView1.ColumnCount - 1
                    DataGridView1.Item(j, i).Value = DBT.Rows(i).Item(j - 1).ToString
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

    End Function

    Public Function DataGridView1_Setup() As Boolean
        Grid_Color(DataGridView1)

        DataGridView1.AllowUserToAddRows = False
        DataGridView1.RowTemplate.Height = 100
        DataGridView1.ColumnCount = 8
        DataGridView1.RowCount = 0

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        'DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'DataGridView1.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        'DataGridView1.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        'DataGridView1.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        'DataGridView1.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        DataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Blue
        DataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.Gold

        DataGridView1.RowHeadersWidth = 5

        DataGridView1.Columns(0).HeaderText = ""
        DataGridView1.Columns(1).HeaderText = "코드번호"
        DataGridView1.Columns(1).Visible = False
        DataGridView1.Columns(2).HeaderText = "거래처명"


        DataGridView1.Columns(3).HeaderText = "제품코드"
        DataGridView1.Columns(3).Visible = False

        DataGridView1.Columns(4).HeaderText = "제품명"
        DataGridView1.Columns(5).HeaderText = "Total"
        DataGridView1.Columns(6).HeaderText = "비고"
        DataGridView1.Columns(7).HeaderText = "관리"
        DataGridView1.Columns(7).Visible = False


        'DataGridView1.Columns(8).HeaderText = "단위"

        'DataGridView1.Columns(9).HeaderText = "단가"

        'DataGridView1.Columns(10).HeaderText = "Total"
        'DataGridView1.Columns(11).HeaderText = "금액"

        'DataGridView1.Columns(12).HeaderText = "VAT"

        'DataGridView1.Columns(13).HeaderText = "비고"
        'DataGridView1.Columns(14).HeaderText = "관리"

        DataGridView1.Columns(0).Width = 10
        'DataGridView1.Columns(1).Width = 20
        'DataGridView1.Columns(2).Width = 1
        'DataGridView1.Columns(3).Width = 1
        'DataGridView1.Columns(4).Width = 1
        'DataGridView1.Columns(5).Width = 1
        'DataGridView1.Columns(6).Width = 20
        'DataGridView1.Columns(7).Width = 1
        'DataGridView1.Columns(8).Width = 1
        'DataGridView1.Columns(9).Width = 1
        'DataGridView1.Columns(10).Width = 1
        'DataGridView1.Columns(11).Width = 20
        'DataGridView1.Columns(12).Width = 1
        'DataGridView1.Columns(13).Width = 1
        'DataGridView1.Columns(14).Width = 1


        DataGridView1_Setup = True
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Panel3.Visible = False
    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1_Row = DataGridView1.CurrentCell.RowIndex
        DataGridView1_Col = DataGridView1.CurrentCell.ColumnIndex

        Dim check1 As Integer
        Dim check2 As Integer
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer

        StrSQL = "UPDATE PC_SHIP_LIST SET  DL_Management='완료'
        where DL_Code = '" & DataGridView1.Item(1, DataGridView1.CurrentCell.RowIndex).Value & "' AND DL_PL_Code= '" & DataGridView1.Item(3, DataGridView1.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        '완료다된애들 마스터에서 완료시키기
        StrSQL = "select count(DL_Management) from PC_SHIP_LIST where  DL_Code='" & DataGridView1.Item(1, DataGridView1.CurrentCell.RowIndex).Value & "' AND DL_Management='완료'"
        Re_Count = DB_Select(DBT)
        check1 = DBT.Rows(0)(0).ToString()
        StrSQL = "select count(DL_Management) from PC_SHIP_LIST where  DL_Code='" & DataGridView1.Item(1, DataGridView1.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Select(DBT)
        check2 = DBT.Rows(0)(0).ToString()

        If (check1 = check2) Then

            StrSQL = "UPDATE PC_SHIP SET  DL_Management='완료'
        where DL_Code = '" & DataGridView1.Item(1, DataGridView1.CurrentCell.RowIndex).Value & "'"
            Re_Count = DB_Execute()
        Else
            StrSQL = "UPDATE PC_SHIP SET  DL_Management='진행중'
        where DL_Code = '" & DataGridView1.Item(1, DataGridView1.CurrentCell.RowIndex).Value & "'"
            Re_Count = DB_Execute()

        End If


        Panel3.Visible = False
        DataGridView1_Display()

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        DataGridView1_Row = DataGridView1.CurrentCell.RowIndex
        DataGridView1_Col = DataGridView1.CurrentCell.ColumnIndex
        Panel3.Visible = True
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        Dim From As Frm_Tmain = New Frm_Tmain
        Me.Hide()
        From.ShowDialog()
        From.StartPosition = FormStartPosition.CenterScreen
        Me.Close()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Day_Label.Text = Format(Now, "yyyy년 MM월 dd일")
        Time_Label.Text = Format(Now, "HH시 mm분")
    End Sub
End Class