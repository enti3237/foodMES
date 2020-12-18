Public Class Deliver_Insert
    Dim JP_Code As String
    Private Sub Deliver_Insert_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
            StrSQL = "Select DR_Code FROM SC_deliver with(NOLOCK) Where DR_Date Like  '" & My_Date & "%' Order By DR_Code Desc "
            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                JP_Code = JP_Code & "001"
            Else
                JP_Code = JP_Code & Format(Val(Mid(DBT.Rows(0).Item(0), 11, 3)) + 1, "00#")
            End If

            deliver_name.Text = Frm_Main.Text_Man_Name.Text
            deliver_code.Text = "CP" & JP_Code
            deliver_day.Text = Today
            deliver_time.Text = TimeString
            'CR_day.Text = Today

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
            StrSQL = "Select * FROM SC_deliver with(NOLOCK) Where DR_CODE = '" & deliver_code.Text & "' Order By DR_CODE"
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

            StrSQL = "Select DL_CODE FROM SC_deliver_list with(NOLOCK)  Order By Convert(Decimal,DL_CODE) Desc "
            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                JP_Code = "0001"
            Else
                'JP_Code = Val(DBT.Rows(0).Item(0)) + 1
                JP_Code = Format(DBT.Rows(0).Item(0) + 1, "000#")
            End If

            TextBox3.Text = JP_Code
            TextBox5.Text = Today

        ElseIf Label1.Text = "update2" Then
            'Insert_Com.Enabled = False
            'DB 데이터 불러오기
            StrSQL = "Select * FROM sc_deliver_list with(NOLOCK) Where DL_CODE = '" & TextBox5.Text & "' Order By DL_CODE"
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
        If deliver_code.Text = "" Then
            MsgBox("제품코드는 공백일 수 없습니다")
            Exit Sub
        End If

        If Label1.Text = "insert" Then

            '기존 코드가 존재하는지 확인 후 INSERT문 실행
            Dim DBT As New DataTable

            StrSQL = "Select DR_CODE FROM SC_deliver with(NOLOCK) Where DR_CODE = '" & deliver_code.Text & "' Order By DR_CODE"
            Re_Count = DB_Select(DBT)

            If Re_Count = 1 Then
                MsgBox("현재 같은 코드가 존재합니다")
                Exit Sub
            Else
            End If
            Try
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO MAN_LOG Values ('" & DateTime.Now & "', '" & Me.Name & "', '납품 추가', '" & Frm_Main.Text_Man_Name.Text & "',
                                                                     '" & deliver_code.Text & " 납품 입력')"
                Re_Count = DB_Execute()

                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO SC_deliver Values ('" & deliver_code.Text & "', '" & deliver_name.Text & "', '" & deliver_day.Text & "', '" & deliver_time.Text & "', '" & customer_code.Text & "','" & customer_name.Text & "','" & deliver_bigo.Text & "' )"
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
                StrSQL = StrSQL & "Insert INTO MAN_LOG Values ('" & DateTime.Now & "', '" & Me.Name & "', '납품 수정', '" & Frm_Main.Text_Man_Name.Text & "',
                                                                     '" & deliver_code.Text & " 납품 수정')"
                Re_Count = DB_Execute()

                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SC_deliver
                                          SET DR_NAME = '" & deliver_name.Text & "',
                                              DR_Bigo = '" & deliver_bigo.Text & "'
                                        WHERE DR_CODE ='" & deliver_code.Text & "'"
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
                StrSQL = StrSQL & "Insert INTO MAN_LOG Values ('" & DateTime.Now & "', '" & Me.Name & "', '납품 리스트 추가', '" & Frm_Main.Text_Man_Name.Text & "',
                                                                     '" & deliver_code.Text & " 납품 리스트 입력')"
                Re_Count = DB_Execute()

                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO sc_deliver_list Values ('" & TextBox3.Text & "','" & TextBox11.Text & "','" & TextBox12.Text & "','" & TextBox5.Text & "', '" & TextBox4.Text & "', '" & TextBox6.Text & "', '" & TextBox7.Text & "', '" & TextBox8.Text & "', '" & TextBox10.Text & "', '" & TextBox9.Text & "', '" & TextBox1.Text & "', '" & TextBox2.Text & "')"
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
                StrSQL = StrSQL & "Insert INTO MAN_LOG Values ('" & DateTime.Now & "', '" & Me.Name & "', '납품 리스트 수정', '" & Frm_Main.Text_Man_Name.Text & "',
                                                                     '" & deliver_code.Text & " 납품 리스트수정')"
                Re_Count = DB_Execute()
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE sc_deliver_list
                                      SET DL_CODE = '" & TextBox3.Text & "',
                                          DL_DR_NAME = '" & TextBox11.Text & "',
                                          DL_CR_NAME = '" & TextBox12.Text & "',
                                          DL_DAY = '" & TextBox5.Text & "',
                                          DL_PL_CODE = '" & TextBox4.Text & "',
                                          DL_PL_NAME = '" & TextBox6.Text & "',
                                          DL_STANDARD = '" & TextBox7.Text & "',
                                          DL_Unit_Amount = '" & TextBox8.Text & "',
                                          DL_QTY = '" & TextBox10.Text & "',
                                          DL_GOLD = '" & TextBox9.Text & "',
                                          DL_VAT = '" & TextBox1.Text & "',
                                          DL_BIGO = '" & TextBox2.Text & "'
                                    WHERE DL_CODE ='" & TextBox3.Text & "'"
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
End Class