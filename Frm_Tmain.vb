Imports System.ComponentModel

Public Class Frm_Tmain
    Dim D_Code As String
    Dim Grid_Display_Ch As Boolean


    Dim Grid_Line_QC_Row As Integer
    Dim Grid_Line_QC_Col As Integer
    Dim Grid_Go_Row As Integer
    Dim Grid_Go_Col As Integer


    Private Sub Frm_PC_Stock_Device_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        End
    End Sub

    Private Sub Frm_PC_Stock_Device_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        FileOpen(1, Application.StartupPath + "\Setup.txt", OpenMode.Input)
        Setup_Data(1) = LineInput(1)

        FileClose(1)
        D_Code = Mid(Setup_Data(1), 8, 3)
        '장비 이름
        Button5.Visible = False
        Button2.Visible = False
        Button3.Visible = False
        Button7.Visible = False
        'Button8.Visible = False
        Button9.Visible = False

        Dim i As Integer
        Dim j As Integer
        Dim DBT As New DataTable
        If D_Code = "1" Then
            '   Label3.Text = "동영산업 생산 관리"
        Else

            StrSQL = "Select * From SI_MACHINE with(NOLOCK) Where PL_Code = '" & D_Code & "'"
            Re_Count = DB_Select(DBT)
            If Re_Count <> 0 Then
                '  Me.Text = D_Code & " " & DBT.Rows(0).Item(1)
                '  Label3.Text = DBT.Rows(0).Item(1)
            Else
                MsgBox("장비 코드를 확인 하세요.")
                End
            End If
        End If '



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

        Button8.Width = Panel2.Width * 0.2
        Button8.Height = Panel2.Width * 0.15
        Button8.Top = Panel2.Width * 0.08
        Button8.Left = Panel2.Width * 0.09

        Button4.Width = Panel2.Width * 0.2
        Button4.Height = Panel2.Width * 0.15
        Button4.Top = Panel2.Width * 0.08
        Button4.Left = Panel2.Width * 0.1 + Button8.Width

        Button1.Width = Panel2.Width * 0.2
        Button1.Height = Panel2.Width * 0.15
        Button1.Top = Panel2.Width * 0.08
        Button1.Left = Panel2.Width * 0.11 + Button8.Width + Button4.Width

        Button11.Width = Panel2.Width * 0.2
        Button11.Height = Panel2.Width * 0.15
        Button11.Top = Panel2.Width * 0.08
        Button11.Left = Panel2.Width * 0.12 + Button8.Width + Button4.Width + Button1.Width

        '밑에버튼4개
        Button12.Width = Panel2.Width * 0.2
        Button12.Height = Panel2.Width * 0.15
        Button12.Top = Panel2.Width * 0.24
        Button12.Left = Panel2.Width * 0.09

        Button13.Width = Panel2.Width * 0.2
        Button13.Height = Panel2.Width * 0.15
        Button13.Top = Panel2.Width * 0.24
        Button13.Left = Panel2.Width * 0.1 + Button8.Width

        Button14.Width = Panel2.Width * 0.2
        Button14.Height = Panel2.Width * 0.15
        Button14.Top = Panel2.Width * 0.24
        Button14.Left = Panel2.Width * 0.11 + Button8.Width + Button4.Width

        Button6.Width = Panel2.Width * 0.2
        Button6.Height = Panel2.Width * 0.15
        Button6.Top = Panel2.Width * 0.24
        Button6.Left = Panel2.Width * 0.12 + Button8.Width + Button4.Width + Button1.Width

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Day_Label.Text = Format(Now, "yyyy년 MM월 dd일")
        Time_Label.Text = Format(Now, "HH시 mm분 ss초")


        ''생산수량
        'If Grid_Main.Rows.Count < 1 Then
        '    Exit Sub
        'End If

        'Dim DBT As New DataTable
        'Dim i As Integer
        'Dim j As Long
        'For i = 0 To Grid_Main.Rows.Count - 1
        '    '생산수량
        '    If Grid_Main.Item(11, i).Value = "종료" Then
        '        j = 0
        '        StrSQL = "Select Sum(Convert(Decimal,PS_Go)) As E1, Sum(Convert(Decimal,PS_Total)) As E2, Sum(Convert(Decimal,PS_Er)) As E3 From PO_Stock with(NOLOCK), PO_Stock_List with(NOLOCK)  Where PO_Stock.PS_PO_Code  = '" & Grid_Main.Item(1, i).Value & "' And  PO_Stock.PS_Code = PO_Stock_List.PS_Code And  PS_PL_Code =  '" & Grid_Main.Item(2, i).Value & "'  And len(PS_Go) > 0 And PS_Go is Not null And len(PS_Total) > 0 And PS_Total is Not null And len(PS_Er) > 0 And PS_Er is Not null"
        '        Re_Count = DB_Select(DBT)
        '        If Re_Count <> 0 Then
        '            j = Val(DBT.Rows(0)("E1").ToString)
        '        End If

        '        StrSQL = "Select PL_Go From Device_Ing with(NOLOCK) Where PL_Code = '" & D_Code & "' "
        '        Re_Count = DB_Select(DBT)
        '        If Re_Count <> 0 Then
        '            If Len(Grid_Main.Item(9, i).Value.ToString) > 0 Then
        '            Else
        '                Grid_Main.Item(8, i).Value = j + DBT.Rows(0)("PL_Go")
        '            End If
        '        End If
        '    End If

        'Next i








    End Sub


    Private Sub Time_Label_Click(sender As Object, e As EventArgs) Handles Time_Label.Click
        Me.Close()

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub
    Public Sub New()

        ' 디자이너에서 이 호출이 필요합니다.
        InitializeComponent()

        ' InitializeComponent() 호출 뒤에 초기화 코드를 추가하세요.

    End Sub

    Private Sub Day_Label_Click(sender As Object, e As EventArgs) Handles Day_Label.Click
        Me.Close()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Frm_PC_Stock_Device.Show()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        '냉동고
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '냉장고

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs)
        '납품등록

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        '가공실

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        '포장실

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        '출고등록
        Out2.Show()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        '입고등록
        Go.Show()

    End Sub

    Private Function GetForm(ByVal NamespaceName As String, ByVal FormName As String) As Object  ' 변수명으로 해당폼을 반환한다.

        Dim asm As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly

        Dim mdls As System.Reflection.Module() = asm.GetModules()

        Dim t As Type = mdls(0).GetType(NamespaceName & "." & FormName)

        If Not t Is Nothing Then

            Return Activator.CreateInstance(t)

        End If

        Return Nothing

    End Function

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Close()

    End Sub

    Private Sub Button8_Click_1(sender As Object, e As EventArgs) Handles Button8.Click
        Print.Show()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs)
        Frm_PC_Stock_Device.Show()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        HP.Show()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Chul.Show()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Gong.Show()
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Out2.Show()
    End Sub
End Class