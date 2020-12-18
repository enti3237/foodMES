Public Class Contract_Insert
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

            contract_name.Text = Frm_Main.Text_Man_Name.Text
            contract_code.Text = "CC" & JP_Code
            contract_day.Text = Today
            contract_time.Text = TimeString
            CR_day.Text = Today

            'Grid_Code_Display()
            'Contract_bigo.Enabled = False
            'TextBox4.Enabled = False
            'TextBox5.Enabled = False
            'TextBox6.Enabled = False
            'Btn_Save.Enabled = False

        ElseIf Label1.Text = "update" Then
            TextBox5.Enabled = False
            TextBox4.Enabled = False
            TextBox6.Enabled = False
            TextBox7.Enabled = False
            TextBox8.Enabled = False
            TextBox10.Enabled = False
            TextBox9.Enabled = False
            Btn_Save.Enabled = False
            'DB 데이터 불러오기
            StrSQL = "Select * FROM sc_sales with(NOLOCK) Where CR_CODE = '" & contract_code.Text & "' Order By CR_CODE"
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
            Insert_Com.Enabled = False

            StrSQL = "Select CL_CODE FROM sc_sales_list with(NOLOCK)  Order By Convert(Decimal,CL_CODE) Desc "
            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                JP_Code = "0001"
            Else
                'JP_Code = Val(DBT.Rows(0).Item(0)) + 1
                JP_Code = Format(DBT.Rows(0).Item(0) + 1, "000#")
            End If

            TextBox1.Text = JP_Code

        ElseIf Label1.Text = "update2" Then
            'Insert_Com.Enabled = False
            'DB 데이터 불러오기
            StrSQL = "Select * FROM sc_sales_list with(NOLOCK) Where CL_CODE = '" & TextBox1.Text & "' Order By CL_CODE"
            Re_Count = DB_Select(DBT)

            If Re_Count = 0 Then
                Me.Close()
                Exit Sub
            Else

            End If

        End If

    End Sub
    Private Sub Insert_Com_Click(sender As Object, e As EventArgs) Handles Insert_Com.Click
        '유효성 검사
        If contract_code.Text = "" Then
            MsgBox("제품코드는 공백일 수 없습니다")
            Exit Sub
        End If

        If Label1.Text = "insert" Then

            '기존 코드가 존재하는지 확인 후 INSERT문 실행
            Dim DBT As New DataTable

            StrSQL = "Select CR_CODE FROM SC_Sales with(NOLOCK) Where CR_CODE = '" & customer_code.Text & "' Order By CR_CODE"
            Re_Count = DB_Select(DBT)

            If Re_Count = 1 Then
                MsgBox("현재 같은 코드가 존재합니다")
                Exit Sub
            Else
            End If
            Try
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO MAN_LOG Values ('" & DateTime.Now & "', '" & Me.Name & "', '수주 추가', '" & Frm_Main.Text_Man_Name.Text & "',
                                                                     '" & contract_code.Text & " 수주 입력')"
                Re_Count = DB_Execute()

                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO SC_SALES Values ('" & contract_code.Text & "', '" & contract_name.Text & "', '" & contract_day.Text & "', '" & contract_time.Text & "', '" & customer_code.Text & "','" & customer_name.Text & "','" & CR_day.Text & "','" & Contract_bigo.Text & "' )"
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
                                                                     '" & contract_code.Text & " 수주 수정')"
                Re_Count = DB_Execute()

                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SC_SALES
                                          SET CR_NAME = '" & contract_name.Text & "',
                                              CR_Bigo = '" & Contract_bigo.Text & "'
                                        WHERE CR_CODE ='" & contract_code.Text & "'"
                Re_Count = DB_Execute()
                MsgBox("수정되었습니다")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Me.Close()

        End If
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Dim popup_customer As New popup_customer
        popup_customer.ShowDialog()

        If popup_customer.DialogResult = DialogResult.OK Then
            customer_name.Text = popup_customer.CUSTOMER_NAME
            customer_code.Text = popup_customer.CUSTOMER_CODE
        End If
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Dim popup_customer As New popup_customer
        popup_customer.ShowDialog()

        If popup_customer.DialogResult = DialogResult.OK Then
            customer_name.Text = popup_customer.CUSTOMER_NAME
            customer_code.Text = popup_customer.CUSTOMER_CODE
        End If
    End Sub

    Private Sub Btn_Save_Click(sender As Object, e As EventArgs) Handles Btn_Save.Click
        If Label1.Text = "insert2" Then
            Dim DBT As New DataTable
            Try
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO MAN_LOG Values ('" & DateTime.Now & "', '" & Me.Name & "', '수주 리스트 추가', '" & Frm_Main.Text_Man_Name.Text & "',
                                                                     '" & contract_code.Text & " 수주 리스트 입력')"
                Re_Count = DB_Execute()

                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                    StrSQL = StrSQL & "Insert INTO sc_sales_list Values ('" & TextBox1.Text & "', '" & contract_code.Text & "', '" & TextBox5.Text & "', '" & TextBox4.Text & "', '" & TextBox6.Text & "', '" & TextBox7.Text & "', '" & TextBox8.Text & "', '" & TextBox10.Text & "', '" & TextBox9.Text & "')"
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
                                                                     '" & contract_code.Text & " 수주 리스트수정')"
                Re_Count = DB_Execute()
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE sc_sales_list
                                      SET CL_PL_CODE = '" & TextBox5.Text & "',
                                          CL_PL_NAME = '" & TextBox4.Text & "',
                                          CL_STANDARD = '" & TextBox6.Text & "',
                                          CL_Unit_Amount = '" & TextBox7.Text & "',
                                          CL_Qty = '" & TextBox8.Text & "',
                                          CL_Gold = '" & TextBox10.Text & "',
                                          CL_Bigo = '" & TextBox9.Text & "'
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
            TextBox4.Text = POPUP_PRODUCT.Product_Name
            TextBox5.Text = POPUP_PRODUCT.Product_Code
            TextBox6.Text = POPUP_PRODUCT.Product_Standard
        End If
    End Sub

    Private Sub btnJobSearch_Click(sender As Object, e As EventArgs) Handles btnJobSearch.Click
        Dim POPUP_PRODUCT As New popup_product
        POPUP_PRODUCT.ShowDialog()

        If POPUP_PRODUCT.DialogResult = DialogResult.OK Then
            TextBox4.Text = POPUP_PRODUCT.Product_Name
            TextBox5.Text = POPUP_PRODUCT.Product_Code
        End If
    End Sub

    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown

    End Sub
End Class