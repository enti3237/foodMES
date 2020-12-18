Public Class popup2

    Public type As String = ""


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        DialogResult = DialogResult.OK
        Close()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DialogResult = DialogResult.No
        Close()

    End Sub

    Private Sub popup1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Combobox_Setup()
    End Sub

    Private Sub Combobox_Setup()
        Dim i As Integer
        Dim DBT As New DataTable
        StrSQL = "select * from si_process"

        Re_Count = DB_Select(DBT)

        ComboBox1.Items.Add("제품생산완료")
        If Re_Count <> 0 Then
            For i = 0 To DBT.Rows.Count - 1
                ComboBox1.Items.Add(DBT.Rows(i)("PC_Name"))
            Next
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        DialogResult = DialogResult.OK
        Close()
    End Sub

    Private Sub popup1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If type = "one" Then
            Button3.Visible = True
            Button1.Visible = False
            Button2.Visible = False
        End If
        If type = "two" Then

            Button3.Visible = False
            Button1.Visible = True
            Button2.Visible = True
        End If
    End Sub
End Class