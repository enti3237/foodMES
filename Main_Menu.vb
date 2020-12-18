Public Class Main_Menu
    Public Event Button_Menu_Click_Data(ByVal sender As System.Object, ByVal e As System.EventArgs)

    Public Sub Set_Button_Text(title As String)
        Button_Menu.Text = title
    End Sub

    Public Sub Set_Label_Color(co As Color)
        Label_Menu.BackColor = co
    End Sub

    Public Sub Set_Label_Color_Won(co As Color)
        OBJ_Back_Color(Label_Menu)
        'Label_Menu.BackColor = co

    End Sub


    Public Sub Button_Color(co As Color)
        Button_Menu.BackColor = co
    End Sub

    Public Sub Fore_Color(co As Color)
        Button_Menu.ForeColor = co

    End Sub

    Public Function Label_Get_Text() As String
        Return Button_Menu.Text
    End Function


    Private Sub Button_Menu_Click(sender As Object, e As EventArgs) Handles Button_Menu.MouseClick
        RaiseEvent Button_Menu_Click_Data(sender, e)
    End Sub

    Private Sub Button_Menu_MouseDown(sender As Object, e As MouseEventArgs) Handles Button_Menu.MouseDown
        sender.forecolor = Color.Black
    End Sub

    Private Sub Main_Menu_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        OBJ_Back_Color(Label_Menu)
    End Sub
End Class
