Public Class Claim_Insert



    Private Sub Btn_Save_Click(sender As Object, e As EventArgs) Handles Btn_Save.Click


        Dim DBT As New DataTable
        If Label1.Text = "insert" Then

            Try

                StrSQL = "insert into QC_CLAIM values ('" & QC_CODE.Text & "','" & QC_CM_NAME.Text & "','" & QC_CONTENT.Text & "',
'" & QC_MAN.Text & "','" & QC_DATE.Text & "','" & QC_BIGO.Text & "')"

                Re_Count = DB_Execute()

                MsgBox("저장되었습니다")

                '로그
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO MAN_LOG Values ('" & DateTime.Now & "', '" & Me.Name & "', 'Claim 추가', '" & Frm_Main.Text_Man_Name.Text & "',
                                                                     '" & QC_CODE.Text & " Claim 입력')"
                Re_Count = DB_Execute()

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Me.Close()

        ElseIf Label1.Text = "update" Then

            'UPDATE문 실행
            Try


                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "update QC_CLAIM Set QC_CM_NAME = '" & QC_CM_NAME.Text &
                                       "',QC_CONTENT = '" & QC_CONTENT.Text &
                                       "', QC_MAN = '" & QC_MAN.Text &
                                       "', QC_BIGO = '" & QC_BIGO.Text &
                                       "' where QC_CODE = '" & QC_CODE.Text & "'"

                Re_Count = DB_Execute()

                MsgBox("수정되었습니다")


                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO MAN_LOG Values ('" & DateTime.Now & "', '" & Me.Name & "', 'Claim 수정', '" & Frm_Main.Text_Man_Name.Text & "',
                                                                     '" & QC_CODE.Text & " Claim 수정')"
                Re_Count = DB_Execute()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Me.Close()


            StrSQL = "SELECT * FROM MAN_LOG"
            Re_Count = DB_Select(DBT)
        End If
        Me.Close()

    End Sub



    Private Sub Import_Insert_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Dim DBT As New DataTable
        Label1.Visible = False
        If Label1.Text = "insert" Then

            Dim JP_Code As Long
            Dim My_Date As String
            Dim My_Time As String

            StrSQL = "Select GETDATE() "
            Re_Count = DB_Select(DBT)


            If Re_Count = 0 Then
                Exit Sub
            Else
                My_Date = DBT.Rows(0).Item(0)
                JP_Code = Mid(My_Date, 1, 4) & Mid(My_Date, 6, 2) & Mid(My_Date, 9, 2)
                My_Time = DateTime.Now.ToString("HH:mm:ss")
                My_Date = Mid(My_Date, 1, 4) & "-" & Mid(My_Date, 6, 2) & "-" & Mid(My_Date, 9, 2)
            End If


            StrSQL = "Select QC_CODE From QC_Claim with(NOLOCK) Where QC_CODE LIKE 'QC" & JP_Code & "%' Order By QC_CODE Desc "
            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                JP_Code = JP_Code & "001"
            Else
                JP_Code = JP_Code & Format(Val(Mid(DBT.Rows(0).Item(0), 11, 3)) + 1, "00#")
            End If


            QC_CODE.Text = "QC" & JP_Code
            QC_DATE.Text = My_Date

        ElseIf Label1.Text = "update" Then
            'DB 데이터 불러오기
            StrSQL = "Select * FROM QC_Claim with(NOLOCK) Where QC_CODE = '" & QC_CODE.Text & "'"
            Re_Count = DB_Select(DBT)

            If Re_Count = 0 Then
                Me.Close()
                Exit Sub
            Else
                QC_CM_NAME.Text = DBT.Rows(0)("QC_CM_NAME")
                QC_CONTENT.Text = DBT.Rows(0)("QC_CONTENT")
                QC_MAN.Text = DBT.Rows(0)("QC_MAN")
                QC_DATE.Text = DBT.Rows(0)("QC_DATE")
                QC_BIGO.Text = DBT.Rows(0)("QC_BIGO")

            End If
        End If

    End Sub



    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim POPUP_ACCOUNT As New popup_account
        POPUP_ACCOUNT.ShowDialog()

        If POPUP_ACCOUNT.DialogResult = DialogResult.OK Then
            QC_CM_NAME.Text = POPUP_ACCOUNT.Customer_Name

        End If
    End Sub
End Class