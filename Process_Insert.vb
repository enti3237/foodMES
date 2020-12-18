Public Class Process_Insert
    Dim JP_Code As String
    Private Sub Process_Insert_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim DBT As New DataTable
        Label1.Visible = False
        If Label1.Text = "insert" Then

            '번호자동생성
            StrSQL = "Select PC_CODE FROM SI_PROCESS with(NOLOCK)  Order By Convert(Decimal,PC_CODE) Desc "
            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                JP_Code = "01"
            Else
                'JP_Code = Val(DBT.Rows(0).Item(0)) + 1
                JP_Code = Format(DBT.Rows(0).Item(0) + 1, "0#")

                process_code.Text = JP_Code
            End If

            TextBox3.Enabled = False
            TextBox4.Enabled = False
            TextBox5.Enabled = False
            TextBox6.Enabled = False
            Btn_Save.Enabled = False

        ElseIf Label1.Text = "update" Then
            TextBox3.Enabled = False
            TextBox4.Enabled = False
            TextBox5.Enabled = False
            TextBox6.Enabled = False
            Btn_Save.Enabled = False
            'DB 데이터 불러오기
            StrSQL = "Select * FROM SI_process with(NOLOCK) Where PC_CODE = '" & process_code.Text & "' Order By PC_CODE"
            Re_Count = DB_Select(DBT)

            If Re_Count = 0 Then
                Me.Close()
                Exit Sub
            Else
                process_name.Text = DBT.Rows(0)("PC_Name")
                Process_bigo.Text = DBT.Rows(0)("Pc_bigo")
            End If

        ElseIf Label1.Text = "insert2" Then

            process_code.Enabled = False
            process_name.Enabled = False
            Process_bigo.Enabled = False
            Insert_Com.Enabled = False

            StrSQL = "Select PC_SUN FROM SI_process_device with(NOLOCK) WHERE PC_CODE = '" & process_code.Text & "' Order By Convert(Decimal, PC_SUN) Desc "
            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                JP_Code = "1"
            Else
                'JP_Code = Val(DBT.Rows(0).Item(0)) + 1
                JP_Code = Format(DBT.Rows(0).Item(0) + 1, "#")
            End If

            TextBox3.Text = JP_Code

        ElseIf Label1.Text = "update2" Then
            process_code.Enabled = False
            process_name.Enabled = False
            Process_bigo.Enabled = False
            Insert_Com.Enabled = False

            'DB 데이터 불러오기
            'StrSQL = "Select * FROM SI_process_device with(NOLOCK) Where PC_CODE = '" & process_code.Text & "' Order By PC_CODE"
            'Re_Count = DB_Select(DBT)

            'If Re_Count = 0 Then
            '    Me.Close()
            '    Exit Sub
            'Else
            '    TextBox3.Text = DBT.Rows(0)("PC_Sun")
            '    TextBox4.Text = DBT.Rows(0)("PC_PL_Code")
            '    TextBox6.Text = DBT.Rows(0)("PC_Bigo")
            'End If
        End If

    End Sub
    Private Sub Insert_Com_Click(sender As Object, e As EventArgs) Handles Insert_Com.Click
        '유효성 검사
        If process_code.Text = "" Then
            MsgBox("제품코드는 공백일 수 없습니다")
            Exit Sub
        End If

        If Label1.Text = "insert" Then

            '기존 코드가 존재하는지 확인 후 INSERT문 실행
            Dim DBT As New DataTable

            StrSQL = "Select PC_CODE FROM SI_process with(NOLOCK) Where PC_CODE = '" & process_code.Text & "' Order By PC_CODE"
            Re_Count = DB_Select(DBT)

            If Re_Count = 1 Then
                MsgBox("현재 같은 코드가 존재합니다")
                Exit Sub
            Else
            End If
            If Len(process_code.Text) = 2 Then
            Else
                MsgBox("코드값은 2자리 숫자로 입력하세요.")
                Exit Sub
            End If
            Try
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO SI_Process Values ('" & process_code.Text & "', '" & process_name.Text & "', '" & Process_bigo.Text & "')"
                Re_Count = DB_Execute()
                MsgBox("저장되었습니다")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Me.Close()

        ElseIf Label1.Text = "update" Then

            'UPDATE문 실행
            Try
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SI_process
                                          SET PC_NAME = '" & process_name.Text & "',
                                              PC_Bigo = '" & Process_bigo.Text & "'
                                        WHERE PC_CODE ='" & process_code.Text & "'"
                Re_Count = DB_Execute()
                MsgBox("수정되었습니다")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Me.Close()

        End If
    End Sub

    Private Sub Btn_Save_Click(sender As Object, e As EventArgs) Handles Btn_Save.Click
        If Label1.Text = "insert2" Then
            Dim DBT As New DataTable
            Try
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO SI_process_device Values ('" & TextBox3.Text & "', '" & TextBox4.Text & "', '" & TextBox5.Text & "', '" & TextBox6.Text & "')"
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
                StrSQL = StrSQL & "UPDATE SI_machine_QC
                                      SET PQ_SUN = '" & TextBox3.Text & "',
                                          PQ_History = '" & TextBox4.Text & "',
                                          PL_Bigo = '" & TextBox6.Text & "'
                                    WHERE PQ_CODE ='" & process_code.Text & "'"
                Re_Count = DB_Execute()
                MsgBox("수정되었습니다")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Me.Close()

        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim popup_device As New popup_device
        popup_device.ShowDialog()

        If popup_device.DialogResult = DialogResult.OK Then
            TextBox4.Text = popup_device.DEVICE_NAME
            TextBox5.Text = popup_device.DEVICE_CODE

        End If
    End Sub

    Private Sub btnJobSearch_Click(sender As Object, e As EventArgs) Handles btnJobSearch.Click
        Dim popup_device As New popup_device
        popup_device.ShowDialog()

        If popup_device.DialogResult = DialogResult.OK Then
            TextBox4.Text = popup_device.DEVICE_NAME
            TextBox5.Text = popup_device.DEVICE_CODE
        End If
    End Sub
End Class