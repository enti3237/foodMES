Public Class PCOrder_Insert

    '작업지시서 생성=작업실적까지 한번에 들어가기
    Dim JP_Code As String

    Private Sub Process_Insert_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim DBT As New DataTable
        Dim JP_Code As Long
        Dim My_Date As String
        Dim My_Time As String

        Label1.Visible = False
        If Label1.Text = "insert" Then
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
            StrSQL = "Select CR_Code FROM SC_SALES with(NOLOCK) Where CR_Date Like  '" & My_Date & "%' Order By CR_Code Desc "
            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                JP_Code = JP_Code & "001"
            Else
                JP_Code = JP_Code & Format(Val(Mid(DBT.Rows(0).Item(0), 11, 3)) + 1, "00#")
            End If

            Txt_plName.Text = Frm_Main.Text_Man_Name.Text
            Txt_cmName.Text = "CC" & JP_Code
            Txt_crDate.Text = Today
            Txt_standard.Text = TimeString
            'CR_day.Text = Today

            'Grid_Code_Display()
            'Contract_bigo.Enabled = False
            'TextBox4.Enabled = False
            'TextBox5.Enabled = False
            'TextBox6.Enabled = False
            'Btn_Save.Enabled = False

        ElseIf Label1.Text = "update" Then
            Txt_poCode.Enabled = False
            Txt_poDay.Enabled = False
            Txt_poBigo.Enabled = False
            Txt_psCode.Enabled = False
            Txt_psStart.Enabled = False
            Txt_psEnd.Enabled = False
            Txt_psBigo.Enabled = False
            Btn_Save.Enabled = False
            'DB 데이터 불러오기
            StrSQL = "Select * FROM sc_sales with(NOLOCK) Where CR_CODE = '" & Txt_cmName.Text & "' Order By CR_CODE"
            Re_Count = DB_Select(DBT)

            If Re_Count = 0 Then
                Me.Close()
                Exit Sub
            Else
                'process_name.Text = DBT.Rows(0)("PC_Name")
                'Process_bigo.Text = DBT.Rows(0)("Pc_bigo")
            End If

        ElseIf Label1.Text = "insert2" Then

            'process_code.Enabled = False
            'process_name.Enabled = False
            'Process_bigo.Enabled = False
            'Insert_Com.Enabled = False



            StrSQL = "Select PP_Code FROM PC_PLAN with(NOLOCK) Order By Convert(Decimal, PP_Code) Desc "
            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                JP_Code = "0001"
            Else
                'JP_Code = Val(DBT.Rows(0).Item(0)) + 1
                JP_Code = Format(DBT.Rows(0).Item(0) + 1, "000#")
            End If

            Txt_poCode.Text = JP_Code

        ElseIf Label1.Text = "update2" Then
            'Insert_Com.Enabled = False
            'DB 데이터 불러오기
            StrSQL = "Select * FROM sc_sales_list with(NOLOCK) Where CL_CODE = '" & inv_CRCode.Text & "' Order By CL_CODE"
            Re_Count = DB_Select(DBT)

            If Re_Count = 0 Then
                Me.Close()
                Exit Sub
            Else

            End If

        End If

    End Sub
    Private Sub Insert_Com_Click(sender As Object, e As EventArgs)
        '유효성 검사
        If Txt_cmName.Text = "" Then
            MsgBox("제품코드는 공백일 수 없습니다")
            Exit Sub
        End If

        If Label1.Text = "insert" Then

            '기존 코드가 존재하는지 확인 후 INSERT문 실행
            Dim DBT As New DataTable

            StrSQL = "Select CR_CODE FROM SC_Sales with(NOLOCK) Where CR_CODE = '" & Txt_qty.Text & "' Order By CR_CODE"
            Re_Count = DB_Select(DBT)

            If Re_Count = 1 Then
                MsgBox("현재 같은 코드가 존재합니다")
                Exit Sub
            Else
            End If
            Try
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO MAN_LOG Values ('" & DateTime.Now & "', '" & Me.Name & "', '수주 추가', '" & Frm_Main.Text_Man_Name.Text & "',
                                                                     '" & Txt_cmName.Text & " 수주 입력')"
                Re_Count = DB_Execute()

                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO SC_SALES Values ('" & Txt_cmName.Text & "', '" & Txt_plName.Text & "', '" & Txt_crDate.Text & "', '" & Txt_standard.Text & "', '" & Txt_qty.Text & "','" & Txt_crDay.Text & "','" & CR_day.Text & "','" & Contract_bigo.Text & "' )"
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
                StrSQL = StrSQL & "Insert INTO MAN_LOG Values ('" & DateTime.Now & "', '" & Me.Name & "', '수주 수정', '" & Frm_Main.Text_Man_Name.Text & "',
                                                                     '" & Txt_cmName.Text & " 수주 수정')"
                Re_Count = DB_Execute()

                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SC_SALES
                                          SET CR_NAME = '" & Txt_plName.Text & "',
                                              CR_Bigo = '" & Contract_bigo.Text & "'
                                        WHERE CR_CODE ='" & Txt_cmName.Text & "'"
                Re_Count = DB_Execute()
                MsgBox("수정되었습니다")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Me.Close()

        End If
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs)
        Dim popup_customer As New popup_customer
        popup_customer.ShowDialog()

        If popup_customer.DialogResult = DialogResult.OK Then
            Txt_crDay.Text = popup_customer.CUSTOMER_NAME
            Txt_qty.Text = popup_customer.CUSTOMER_CODE
        End If
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs)
        Dim popup_customer As New popup_customer
        popup_customer.ShowDialog()

        If popup_customer.DialogResult = DialogResult.OK Then
            Txt_crDay.Text = popup_customer.CUSTOMER_NAME
            Txt_qty.Text = popup_customer.CUSTOMER_CODE
        End If
    End Sub

    Private Sub Btn_Save_Click(sender As Object, e As EventArgs) Handles Btn_Save.Click
        If Label1.Text = "insert2" Then
            Dim DBT As New DataTable
            Try
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO MAN_LOG Values ('" & DateTime.Now & "', '" & Me.Name & "', '수주 리스트 추가', '" & Frm_Main.Text_Man_Name.Text & "',
                                                                     '" & Txt_cmName.Text & " 수주 리스트 입력')"
                Re_Count = DB_Execute()

                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO sc_sales_list Values ('" & TextBox1.Text & "', '" & Txt_cmName.Text & "', '" & Txt_poCode.Text & "', '" & Txt_poDay.Text & "', '" & Txt_poBigo.Text & "', '" & Txt_psCode.Text & "', '" & Txt_psStart.Text & "', '" & Txt_psEnd.Text & "', '" & Txt_psBigo.Text & "')"
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
                StrSQL = StrSQL & "Insert INTO MAN_LOG Values ('" & DateTime.Now & "', '" & Me.Name & "', '수주 리스트 수정', '" & Frm_Main.Text_Man_Name.Text & "',
                                                                     '" & Txt_cmName.Text & " 수주 리스트수정')"
                Re_Count = DB_Execute()
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE sc_sales_list
                                      SET CL_PL_CODE = '" & Txt_poCode.Text & "',
                                          CL_PL_NAME = '" & Txt_poDay.Text & "',
                                          CL_STANDARD = '" & Txt_poBigo.Text & "',
                                          CL_Unit_Amount = '" & Txt_psCode.Text & "',
                                          CL_Qty = '" & Txt_psStart.Text & "',
                                          CL_Gold = '" & Txt_psEnd.Text & "',
                                          CL_Bigo = '" & Txt_psBigo.Text & "'
                                    WHERE CL_CODE ='" & TextBox1.Text & "'"
                Re_Count = DB_Execute()
                MsgBox("수정되었습니다")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Me.Close()

        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim POPUP_PRODUCT As New popup_product
        POPUP_PRODUCT.ShowDialog()

        If POPUP_PRODUCT.DialogResult = DialogResult.OK Then
            Txt_poDay.Text = POPUP_PRODUCT.Product_Name
            Txt_poCode.Text = POPUP_PRODUCT.Product_Code
            Txt_poBigo.Text = POPUP_PRODUCT.Product_Standard
        End If
    End Sub

    Private Sub btnJobSearch_Click(sender As Object, e As EventArgs) Handles btnJobSearch.Click
        Dim POPUP_PRODUCT As New popup_product
        POPUP_PRODUCT.ShowDialog()

        If POPUP_PRODUCT.DialogResult = DialogResult.OK Then
            Txt_poDay.Text = POPUP_PRODUCT.Product_Name
            Txt_poCode.Text = POPUP_PRODUCT.Product_Code
        End If
    End Sub

End Class