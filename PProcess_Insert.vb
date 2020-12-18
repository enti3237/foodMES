Public Class PProcess_Insert
    Dim JP_Code As String
    Private Sub Process_Insert_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim DBT As New DataTable
        Label1.Visible = False
        TextBox1.Visible = False
        If Label1.Text = "insert2" Then
            'BOM추가
            process_code.Enabled = False
            process_name.Enabled = False
            Process_bigo.Enabled = False
            Insert_Com.Enabled = False

            Dim Db_Sun As Long
            StrSQL = "Select PP_Sun FROM SI_PROCESS_LIST with(NOLOCK) Where PP_PC_Code = '" & process_code.Text & "' Order By Convert(Decimal,PP_Sun) Desc "
            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                Db_Sun = 1
            Else
                Db_Sun = DBT.Rows(0).Item(0) + 1

            End If

            TextBox3.Text = Db_Sun

        ElseIf Label1.Text = "update2" Then
            process_code.Enabled = False
            process_name.Enabled = False
            Process_bigo.Enabled = False
            Insert_Com.Enabled = False

        End If

    End Sub

    Private Sub Btn_Save_Click(sender As Object, e As EventArgs) Handles Btn_Save.Click
        If Label1.Text = "insert2" Then
            Dim DBT As New DataTable
            Try
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO SI_PROCESS_LIST (PP_PC_CODE, PP_SUN, PP_CODE) Values ('" & process_code.Text & "', '" & TextBox3.Text & "', '" & TextBox1.Text & "')"
                Re_Count = DB_Execute()
                MsgBox("저장되었습니다")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Me.Close()

        ElseIf Label1.Text = "update2" Then

            'UPDATE문 실행
            Try
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SI_PROCESS_LIST
                                          SET PP_CODE = '" & TextBox1.Text & "'     
                                        WHERE PP_CODE ='" & process_code.Text & "' AND PP_SUN='" & TextBox3.Text & "'"
                Re_Count = DB_Execute()
                MsgBox("수정되었습니다")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Me.Close()

        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs)
        Dim popup_process As New popup_process
        popup_process.ShowDialog()

        If popup_process.DialogResult = DialogResult.OK Then
            TextBox4.Text = popup_process.PROCESS_Name
            TextBox1.Text = popup_process.PROCESS_Code

        End If
    End Sub

    Private Sub btnJobSearch_Click(sender As Object, e As EventArgs) Handles btnJobSearch.Click
        Dim popup_process As New popup_process
        popup_process.ShowDialog()

        If popup_process.DialogResult = DialogResult.OK Then
            TextBox4.Text = popup_process.PROCESS_Name
            TextBox1.Text = popup_process.PROCESS_Code
        End If
    End Sub
End Class