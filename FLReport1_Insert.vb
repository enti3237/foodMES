Public Class FLReport1_Insert



    Private Sub Btn_Save_Click(sender As Object, e As EventArgs) Handles Btn_Save.Click


        Dim DBT As New DataTable
        If Label1.Text = "insert" Then

            Try

                StrSQL = "insert into FC_RECORD values ('" & FL_CODE.Text & "','" & FL_DATE.Text & "','" & FL_TIME.Text & "',
'" & FL_NAME.Text & "','" & FL_PD_Code.Text & "','" & FL_PD_Name.Text & "','" & FL_History.Text & "','" & FL_Gold.Text & "','" & FL_Bigo.Text & "')"

                Re_Count = DB_Execute()

                MsgBox("저장되었습니다")

                '로그
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO MAN_LOG Values ('" & DateTime.Now & "', '" & Me.Name & "', '설비이력 추가', '" & Frm_Main.Text_Man_Name.Text & "',
                                                                     '" & FL_CODE.Text & " 설비이력 입력')"
                Re_Count = DB_Execute()

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Me.Close()

        ElseIf Label1.Text = "update" Then

            'UPDATE문 실행
            Try


                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "update FC_RECORD Set FL_DATE = '" & FL_DATE.Text &
                                       "',FL_TIME = '" & FL_TIME.Text &
                                       "', FL_NAME = '" & FL_NAME.Text &
                                       "', FL_PD_Code = '" & FL_PD_Code.Text &
                                       "',FL_PD_Name = '" & FL_PD_Name.Text &
                                       "', FL_History = '" & FL_History.Text &
                                       "',FL_Gold = '" & FL_Gold.Text &
                                       "', FL_Bigo = '" & FL_Bigo.Text &
                                       "' where FL_Code = '" & FL_CODE.Text & "'"

                Re_Count = DB_Execute()

                MsgBox("수정되었습니다")


                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO MAN_LOG Values ('" & DateTime.Now & "', '" & Me.Name & "', '설비이력 수정', '" & Frm_Main.Text_Man_Name.Text & "',
                                                                     '" & FL_CODE.Text & " 설비이력 수정')"
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


            StrSQL = "Select FL_CODE From FC_RECORD with(NOLOCK) Where FL_CODE LIKE 'FL" & JP_Code & "%' Order By FL_CODE Desc "
            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                JP_Code = JP_Code & "001"
            Else
                JP_Code = JP_Code & Format(Val(Mid(DBT.Rows(0).Item(0), 11, 3)) + 1, "00#")
            End If


            FL_CODE.Text = "FL" & JP_Code
            FL_DATE.Text = My_Date
            FL_TIME.Text = My_Time

        ElseIf Label1.Text = "update" Then
            'DB 데이터 불러오기
            StrSQL = "Select * FROM FC_RECORD with(NOLOCK) Where FL_CODE = '" & FL_CODE.Text & "'"
            Re_Count = DB_Select(DBT)

            If Re_Count = 0 Then
                Me.Close()
                Exit Sub
            Else
                FL_DATE.Text = DBT.Rows(0)("FL_DATE")
                FL_TIME.Text = DBT.Rows(0)("FL_TIME")
                FL_NAME.Text = DBT.Rows(0)("FL_NAME")
                FL_PD_Code.Text = DBT.Rows(0)("FL_PD_Code")
                FL_PD_Name.Text = DBT.Rows(0)("FL_PD_NAme")
                FL_History.Text = DBT.Rows(0)("FL_HISTORY")
                FL_Gold.Text = DBT.Rows(0)("FL_GOLD")
                FL_Bigo.Text = DBT.Rows(0)("FL_BIGO")

            End If
        End If

    End Sub



    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Dim POPUP_ACCOUNT As New popup_account
        'POPUP_ACCOUNT.ShowDialog()

        'If POPUP_ACCOUNT.DialogResult = DialogResult.OK Then
        '    QC_CM_NAME.Text = POPUP_ACCOUNT.Customer_Name

        'End If
    End Sub

    Private Sub btnJobSearch_Click(sender As Object, e As EventArgs) Handles btnJobSearch.Click
        Dim POPUP_ACCOUNT As New popup_FL
        POPUP_ACCOUNT.ShowDialog()

        If POPUP_ACCOUNT.DialogResult = DialogResult.OK Then
            FL_PD_Code.Text = POPUP_ACCOUNT.Code
            FL_PD_Name.Text = POPUP_ACCOUNT.Name
        End If
    End Sub
End Class