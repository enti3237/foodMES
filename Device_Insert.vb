Public Class Device_Insert
    Dim JP_Code As String
    Private Sub Product_Insert_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim DBT As New DataTable
        Label1.Visible = False
        If Label1.Text = "insert" Then
            TextBox3.Enabled = False
            TextBox4.Enabled = False
            TextBox6.Enabled = False
            Btn_Save.Enabled = False

            '번호자동생성
            StrSQL = "Select PD_CODE FROM SI_MACHINE with(NOLOCK)  Order By Convert(Decimal,PD_CODE) Desc "
            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                JP_Code = "001"
            Else
                'JP_Code = Val(DBT.Rows(0).Item(0)) + 1
                JP_Code = Format(DBT.Rows(0).Item(0) + 1, "00#")

                device_code.Text = JP_Code
            End If

        ElseIf Label1.Text = "update" Then
            TextBox3.Enabled = False
            TextBox4.Enabled = False
            TextBox6.Enabled = False
            Btn_Save.Enabled = False
            'DB 데이터 불러오기
            StrSQL = "Select * FROM SI_machine with(NOLOCK) Where PD_CODE = '" & device_code.Text & "' Order By PD_CODE"
            Re_Count = DB_Select(DBT)

            If Re_Count = 0 Then
                Me.Close()
                Exit Sub
            Else
                Device_name.Text = DBT.Rows(0)("PD_Name")
                Device_ID.Text = DBT.Rows(0)("PD_ID")
                Device_Port.Text = DBT.Rows(0)("PD_Port")
                Device_bigo.Text = DBT.Rows(0)("PD_Bigo")
            End If

        ElseIf Label1.Text = "insert2" Then

            device_code.Enabled = False
            Device_name.Enabled = False
            Device_ID.Enabled = False
            Device_Port.Enabled = False
            Device_bigo.Enabled = False
            Insert_Com.Enabled = False

            StrSQL = "Select PQ_SUN FROM SI_machine_QC with(NOLOCK) WHERE PQ_CODE='" & device_code.Text & "' Order By Convert(Decimal, PQ_SUN) Desc "
            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                JP_Code = "1"
            Else
                'JP_Code = Val(DBT.Rows(0).Item(0)) + 1
                JP_Code = Format(DBT.Rows(0).Item(0) + 1, "#")
            End If

            TextBox3.Text = JP_Code

        ElseIf Label1.Text = "update2" Then
            device_code.Enabled = False
            Device_name.Enabled = False
            Device_ID.Enabled = False
            Device_Port.Enabled = False
            Device_bigo.Enabled = False
            Insert_Com.Enabled = False
            ''DB 데이터 불러오기
            'StrSQL = "Select * FROM SI_machine_QC with(NOLOCK) Where PQ_CODE = '" & device_code.Text & "' AND PQ_SUN='" & TextBox3.Text & "' Order By PQ_CODE"
            'Re_Count = DB_Select(DBT)

            'If Re_Count = 0 Then
            '    ' Me.Close()
            '    ' Exit Sub
            'Else
            '    TextBox4.Text = DBT.Rows(0)("PQ_History")
            '    TextBox6.Text = DBT.Rows(0)("PQ_Bigo")
            'End If
        End If

    End Sub
    Private Sub Btn_Save_Click(sender As Object, e As EventArgs) Handles Btn_Save.Click
        If Label1.Text = "insert2" Then
            Dim DBT As New DataTable
            Try
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO SI_machine_QC Values ('" & device_code.Text & "', '" & JP_Code & "', '" & TextBox4.Text & "', '" & TextBox6.Text & "')"
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
                                      SET PQ_History = '" & TextBox4.Text & "',
                                          PD_Bigo = '" & TextBox6.Text & "'
                                    WHERE PQ_CODE ='" & device_code.Text & "' AND PQ_SUN = '" & TextBox3.Text & "'"
                Re_Count = DB_Execute()
                MsgBox("수정되었습니다")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Me.Close()

        End If
    End Sub

    Private Sub Insert_Com_Click(sender As Object, e As EventArgs) Handles Insert_Com.Click
        '유효성 검사
        If device_code.Text = "" Then
            MsgBox("제품코드는 공백일 수 없습니다")
            Exit Sub
        End If

        If Label1.Text = "insert" Then

            '기존 코드가 존재하는지 확인 후 INSERT문 실행
            Dim DBT As New DataTable

            StrSQL = "Select PD_CODE FROM SI_machine with(NOLOCK) Where PD_CODE = '" & device_code.Text & "' Order By PD_CODE"
            Re_Count = DB_Select(DBT)

            If Re_Count = 1 Then
                MsgBox("현재 같은 코드가 존재합니다")
                Exit Sub
            Else
            End If
            If Len(device_code.Text) = 3 Then
            Else
                MsgBox("코드값은 3자리 숫자로 입력하세요.")
                Exit Sub
            End If
            Try
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO SI_machine Values ('" & device_code.Text & "', '" & Device_name.Text & "', '" & Device_ID.Text & "', '" & Device_Port.Text & "', '" & Device_bigo.Text & "')"
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
                StrSQL = StrSQL & "UPDATE SI_machine
                                      SET PD_NAME = '" & Device_name.Text & "',
                                          PD_ID = '" & Device_ID.Text & "',
                                          PD_Port = '" & Device_Port.Text & "',
                                          PD_Bigo = '" & Device_bigo.Text & "'
                                    WHERE PD_CODE ='" & device_code.Text & "'"
                Re_Count = DB_Execute()
                MsgBox("수정되었습니다")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Me.Close()

        End If
    End Sub
End Class