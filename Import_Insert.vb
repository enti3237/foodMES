Public Class Import_Insert
    Dim XO1 As String
    Dim XO2 As String
    Dim XO3 As String
    Dim XO4 As String
    Dim XO5 As String
    Public number As Integer


    Private Sub Btn_Save_Click(sender As Object, e As EventArgs) Handles Btn_Save.Click
        Dim i As Integer

        XO1 = ""
        XO2 = ""
        XO3 = ""
        XO4 = ""
        XO5 = ""

        Dim DBT As New DataTable
        If Label1.Text = "insert" Then

            Try

                StrSQL = "insert into QUALITY VALUES ('" & QC_CODE.Text & "','" & QC_DATE.Text & "',
        '" & QC_TYPE.Text & "','" & QC_MAN.Text & "','" & CODE.Text & "','" & QC_PL_CODE.Text & "',
        '" & QC_PL_NAME.Text & "','" & QC_ITEM.Text & "','" & QC_METHOD.Text & "','" & QC_LEVEL.Text & "',
        '','','','','','','" & QC_BIGO.Text & "')"

                Re_Count = DB_Execute()

                For i = 0 To CheckedListBox1.Items.Count - 1
                    If (CheckedListBox1.GetItemChecked(i)) Then
                        'Label1.Text += i.ToString + ", "
                        If (i = 0) Then
                            XO1 = "ok"
                        End If
                        If (i = 1) Then
                            XO2 = "ok"
                        End If
                        If (i = 2) Then
                            XO3 = "ok"
                        End If
                        If (i = 3) Then
                            XO4 = "ok"
                        End If
                        If (i = 4) Then
                            XO5 = "ok"
                        End If
                    End If

                Next

                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "update QUALITY Set QC_XO1 = '" & XO1 & "' , QC_XO2 = '" & XO2 & "', QC_XO3 = '" & XO3 & "', QC_XO4 = '" & XO4 & "', QC_XO5 = '" & XO5 & "' 
        where QC_CODE = '" & QC_CODE.Text & "'"

                Re_Count = DB_Execute()
                XO1 = ""
                XO2 = ""
                XO3 = ""
                XO4 = ""
                XO5 = ""
                MsgBox("저장되었습니다")

                '로그
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO MAN_LOG Values ('" & DateTime.Now & "', '" & Me.Name & "', '입고검사 추가', '" & Frm_Main.Text_Man_Name.Text & "',
                                                                     '" & QC_CODE.Text & " 입고검사 입력')"
                Re_Count = DB_Execute()

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Me.Close()

        ElseIf Label1.Text = "update" Then

            'UPDATE문 실행
            Try

                 For i = 0 To CheckedListBox1.Items.Count - 1
                    If (CheckedListBox1.GetItemChecked(i)) Then
                        'Label1.Text += i.ToString + ", "
                        If (i = 0) Then
                            XO1 = "ok"
                        End If

                        If (i = 1) Then
                            XO2 = "ok"
                        End If
                        If (i = 2) Then
                            XO3 = "ok"
                        End If
                        If (i = 3) Then
                            XO4 = "ok"
                        End If
                        If (i = 4) Then
                            XO5 = "ok"
                        End If

                    End If

                Next


                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "update QUALITY Set QC_DATE = '" & QC_DATE.Text &
                                       "',QC_TYPE = '" & QC_TYPE.Text &
                                       "', QC_MAN = '" & QC_MAN.Text &
                                       "', CODE = '" & CODE.Text &
                                       "', QC_PL_CODE = '" & QC_PL_CODE.Text &
                                       "', QC_PL_NAME = '" & QC_PL_NAME.Text &
                                       "', QC_ITEM = '" & QC_ITEM.Text &
                                       "', QC_METHOD = '" & QC_METHOD.Text &
                                       "', QC_LEVEL = '" & QC_LEVEL.Text &
                                       "', QC_AQL = '" & QC_AQL.Text &
                                       "', QC_XO1 = '"& XO1 &
                                       "', QC_XO2 = '" & XO2 &
                                       "', QC_XO3 = '" & XO3 &
                                       "', QC_XO4 = '"& XO4 &
                                       "', QC_XO5 = '" & XO5 &
                                       "', QC_BIGO = '" & QC_BIGO.Text &
                                       "' where QC_CODE = '" & QC_CODE.Text & "'"

                Re_Count = DB_Execute()

                MsgBox("수정되었습니다")


                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO MAN_LOG Values ('" & DateTime.Now & "', '" & Me.Name & "', '입고검사 수정', '" & Frm_Main.Text_Man_Name.Text & "',
                                                                     '" & QC_CODE.Text & " 입고검사 수정')"
                Re_Count = DB_Execute()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Me.Close()


            StrSQL = "SELECT * FROM MAN_LOG"
            Re_Count = DB_Select(DBT)
        End If








        XO1 = ""
        XO2 = ""
        XO3 = ""
        XO4 = ""
        XO5 = ""
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


            StrSQL = "Select QC_CODE From QUALITY with(NOLOCK) Where QC_CODE LIKE 'QC" & JP_Code & "%' Order By QC_CODE Desc "
            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                JP_Code = JP_Code & "001"
            Else
                JP_Code = JP_Code & Format(Val(Mid(DBT.Rows(0).Item(0), 11, 3)) + 1, "00#")
            End If


            QC_CODE.Text = "QC" & JP_Code
            QC_DATE.Text = My_Date
            QC_TYPE.Text = "수입검사"
            CODE.Text = "입고"

        ElseIf Label1.Text = "update" Then
            'DB 데이터 불러오기
            StrSQL = "Select * FROM QUALITY with(NOLOCK) Where QC_CODE = '" & QC_CODE.Text & "'"
            Re_Count = DB_Select(DBT)

            If Re_Count = 0 Then
                Me.Close()
                Exit Sub
            Else
                QC_DATE.Text = DBT.Rows(0)("QC_DATE")
                QC_TYPE.Text = DBT.Rows(0)("QC_TYPE")
                QC_MAN.Text = DBT.Rows(0)("QC_MAN")
                CODE.Text = DBT.Rows(0)("CODE")
                QC_PL_CODE.Text = DBT.Rows(0)("QC_PL_CODE")
                QC_PL_NAME.Text = DBT.Rows(0)("QC_PL_NAME")
                QC_ITEM.Text = DBT.Rows(0)("QC_ITEM")
                QC_METHOD.Text = DBT.Rows(0)("QC_METHOD")
                QC_LEVEL.Text = DBT.Rows(0)("QC_LEVEL")
                QC_AQL.Text = DBT.Rows(0)("QC_AQL")
                XO1 = DBT.Rows(0)("QC_XO1")
                XO2 = DBT.Rows(0)("QC_XO2")
                XO3 = DBT.Rows(0)("QC_XO3")
                XO4 = DBT.Rows(0)("QC_XO4")
                XO5 = DBT.Rows(0)("QC_XO5")
                QC_BIGO.Text = DBT.Rows(0)("QC_BIGO")

                If (XO1 = "ok") Then
                    CheckedListBox1.SetItemChecked(0, True)
                Else
                    CheckedListBox1.SetItemChecked(0, False)
                End If
                If (XO2 = "ok") Then
                    CheckedListBox1.SetItemChecked(1, True)
                Else
                    CheckedListBox1.SetItemChecked(1, False)
                End If
                If (XO3 = "ok") Then
                    CheckedListBox1.SetItemChecked(2, True)
                Else
                    CheckedListBox1.SetItemChecked(2, False)
                End If
                If (XO4 = "ok") Then
                    CheckedListBox1.SetItemChecked(3, True)
                Else
                    CheckedListBox1.SetItemChecked(3, False)
                End If
                If (XO5 = "ok") Then
                    CheckedListBox1.SetItemChecked(4, True)
                Else
                    CheckedListBox1.SetItemChecked(4, False)
                End If

            End If
        End If

    End Sub

    Private Sub btnJobSearch_Click(sender As Object, e As EventArgs) Handles btnJobSearch.Click
        Dim POPUP_PRODUCT As New popup_product
        POPUP_PRODUCT.ShowDialog()

        If POPUP_PRODUCT.DialogResult = DialogResult.OK Then
            QC_PL_NAME.Text = POPUP_PRODUCT.Product_Name
            QC_PL_CODE.Text = POPUP_PRODUCT.Product_Code

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim POPUP_PRODUCT As New popup_product
        POPUP_PRODUCT.ShowDialog()

        If POPUP_PRODUCT.DialogResult = DialogResult.OK Then
            QC_PL_NAME.Text = POPUP_PRODUCT.Product_Name
            QC_PL_CODE.Text = POPUP_PRODUCT.Product_Code
        End If
    End Sub
End Class