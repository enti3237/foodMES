Public Class Manform

    Public Man_1 As String '코드
    Public Man_2 As String '성명
    Public Man_4 As String '부서
    Public Man_5 As String '직급
    Public Man_6 As String '성별
    Public Man_8 As String '내/외국인
    Public Man_9 As String '전화번호
    Public Man_12 As String '암호
    Public Man_13 As String '사용Level
    Public Man_15 As String '비고
    Public Man_16 As String '보건용 갱신일자


    Private Sub Manform_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Man_1 = Info_Tx0.Text
        Man_2 = Info_Tx1.Text

        Man_4 = ComboBox3.Text
        Man_5 = ComboBox2.Text
        Man_6 = ComboBox5.Text

        Man_8 = ComboBox6.Text
        Man_9 = Info_Tx12.Text

        Man_12 = Info_Tx16.Text
        Man_13 = ComboBox4.Text

        Man_15 = Info_Tx21.Text
        Man_16 = Info_Tx22.Text
        ComboBox5.Items.Clear()
        ComboBox6.Items.Clear()

    End Sub

    Private Sub Btn_Save_Click(sender As Object, e As EventArgs) Handles Btn_Save.Click
        Dim DBT As New DataTable


        Man_1 = Info_Tx0.Text
        Man_2 = Info_Tx1.Text

        Man_4 = ComboBox3.Text
        Man_5 = ComboBox2.Text
        Man_6 = ComboBox5.Text

        Man_8 = ComboBox6.Text
        Man_9 = Info_Tx12.Text

        Man_12 = Info_Tx16.Text
        Man_13 = ComboBox4.Text

        Man_15 = Info_Tx21.Text
        Man_16 = Info_Tx22.Text

        StrSQL = "select * from Si_MAn where Man_Code='" & Man_1 & "'"
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "Insert INTO SI_MAN  Values ('" & Man_1 & "','" & Man_2 & "','" & Man_4 & "','" & Man_5 & "','" & Man_6 & "','" & Man_8 & "','" & Man_9 & "','" & Man_12 & "','" & Man_13 & "','" & Man_15 & "','" & Man_16 & "')"
            Re_Count = DB_Execute()
        Else
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "update SI_MAN set MAN_Name='" & Man_2 & "',Man_Bu='" & Man_4 & "',Man_Gik='" & Man_5 & "',Man_Sex='" & Man_6 & "',
Man_Foreigner='" & Man_8 & "',MAN_Tel='" & Man_9 & "',Man_Pass='" & Man_12 & "',Man_Level='" & Man_13 & "',Man_Bigo='" & Man_15 & "',Man_Update='" & Man_16 & "'
where Man_Code='" & Man_1 & "'"
            Re_Count = DB_Execute()
        End If





        MsgBox("저장되었습니다")

        Info_Tx0.Text = ""
        Info_Tx1.Text = ""

        ComboBox3.Text = ""
        ComboBox2.Text = ""
        ComboBox5.Text = ""

        ComboBox6.Text = ""
        Info_Tx12.Text = ""

        Info_Tx16.Text = ""
        ComboBox4.Text = ""

        Info_Tx21.Text = ""
        Info_Tx22.Text = ""
        ComboBox5.Items.Clear()
        ComboBox6.Items.Clear()

        ComboBox2.Items.Clear()
        ComboBox3.Items.Clear()
        ComboBox4.Items.Clear()

        Me.Close()
    End Sub


    Private Sub ComboBox2_Click(sender As Object, e As EventArgs) Handles ComboBox2.Click
        Dim DBT As New DataTable
        StrSQL = "Select Base_Name  FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '직급' Order By Base_Name"
        Re_Count = DB_Select(DBT)
        ComboBox2.Items.Clear()
        If Re_Count <> 0 Then
            For i = 0 To Re_Count - 1
                ComboBox2.Items.Add(DBT.Rows(i).Item(0))
            Next i
        End If
    End Sub

    Private Sub ComboBox3_Click(sender As Object, e As EventArgs) Handles ComboBox3.Click
        Dim DBT As New DataTable
        StrSQL = "Select Base_Name  FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '부서' Order By Base_Name"
        Re_Count = DB_Select(DBT)
        ComboBox3.Items.Clear()
        If Re_Count <> 0 Then
            For i = 0 To Re_Count - 1
                ComboBox3.Items.Add(DBT.Rows(i).Item(0))
            Next i
        End If
    End Sub

    Private Sub ComboBox4_Click(sender As Object, e As EventArgs) Handles ComboBox4.Click
        Dim DBT As New DataTable
        StrSQL = "Select Le_Code  FROM SI_MAN_Level with(NOLOCK) Order By Le_Code"
        Re_Count = DB_Select(DBT)
        ComboBox4.Items.Clear()
        If Re_Count <> 0 Then
            For i = 0 To Re_Count - 1
                ComboBox4.Items.Add(DBT.Rows(i).Item(0))
            Next i
        End If
    End Sub

    Private Sub ComboBox5_Click(sender As Object, e As EventArgs) Handles ComboBox5.Click
        ComboBox5.Items.Clear()
        ComboBox5.Items.Add("남")
        ComboBox5.Items.Add("여")
    End Sub

    Private Sub ComboBox6_Click(sender As Object, e As EventArgs) Handles ComboBox6.Click
        ComboBox6.Items.Clear()
        ComboBox6.Items.Add("내국인")
        ComboBox6.Items.Add("외국인")
    End Sub
End Class