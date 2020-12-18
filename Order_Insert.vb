Public Class Order_Insert
    Dim JP_Code As String
    Private Sub order_Insert_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim DBT As New DataTable
        Label1.Visible = False
        If Label1.Text = "insert" Then

            '번호자동생성
            StrSQL = "Select MO_CODE FROM MC_STOCK_ORDER with(NOLOCK) Order By MO_CODE Desc "
            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                JP_Code = "001"
            Else
                'JP_Code = Val(DBT.Rows(0).Item(0)) + 1
                JP_Code = Replace(DBT.Rows(0).Item(0), "MO", "")
                JP_Code = Format(JP_Code + 1, "00#")
            End If

            STANDARD.Text = STANDARD.Text & "KG"
            Order_Code.Text = "MO" & JP_Code
            DAY.Text = Today

        ElseIf Label1.Text = "update" Then
            Order_Code.Enabled = False
            CM_Code.Enabled = False
            CM_Name.Enabled = False
            PL_CODE.Enabled = False
            PL_NAME.Enabled = False
            ' BIGO.Enabled = False
            '    Btn_Save.Enabled = False
            'DB 데이터 불러오기
            StrSQL = "Select * FROM MC_STOCK_ORDER with(NOLOCK) Where MO_CODE = '" & Order_Code.Text & "' Order By MO_CODE"
            Re_Count = DB_Select(DBT)

            If Re_Count = 0 Then
                Me.Close()
                Exit Sub
            Else
                CM_Code.Text = DBT.Rows(0)("MO_CM_CODE")
                CM_Name.Text = DBT.Rows(0)("MO_CM_NAME")
                PL_CODE.Text = DBT.Rows(0)("MO_PL_CODE")
                PL_NAME.Text = DBT.Rows(0)("MO_PL_NAME")
                DAY.Text = DBT.Rows(0)("MO_DAY")
                STANDARD.Text = DBT.Rows(0)("MO_STANDARD")
                QTY.Text = DBT.Rows(0)("MO_QTY")
                TOTAL.Text = DBT.Rows(0)("MO_TOTAL")
                GOLD.Text = DBT.Rows(0)("MO_GOLD")
                BIGO.Text = DBT.Rows(0)("MO_BIGO")

            End If
        End If



    End Sub
    'Private Sub Insert_Com_Click(sender As Object, e As EventArgs)
    '    '유효성 검사
    '    If Order_Code.Text = "" Then
    '        MsgBox("제품코드는 공백일 수 없습니다")
    '        Exit Sub
    '    End If

    '    If Label1.Text = "insert" Then

    '        '기존 코드가 존재하는지 확인 후 INSERT문 실행
    '        Dim DBT As New DataTable

    '        StrSQL = "Select MO_CODE FROM MC_STOCK_ORDER with(NOLOCK) Where MO_CODE = '" & Order_Code.Text & "' Order By MO_CODE"
    '        Re_Count = DB_Select(DBT)

    '        If Re_Count = 1 Then
    '            MsgBox("현재 같은 코드가 존재합니다")
    '            Exit Sub
    '        Else
    '        End If
    '        If Len(Order_Code.Text) = 2 Then
    '        Else
    '            MsgBox("코드값은 2자리 숫자로 입력하세요.")
    '            Exit Sub
    '        End If
    '        Try
    '            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
    '            StrSQL = StrSQL & "Insert INTO MC_STOCK_ORDER Values ('" & Order_Code.Text & "', '" & order_name.Text & "', '" & order_bigo.Text & "')"
    '            Re_Count = DB_Execute()
    '            MsgBox("저장되었습니다")
    '        Catch ex As Exception
    '            MsgBox(ex.Message)
    '        End Try
    '        Me.Close()

    '    ElseIf Label1.Text = "update" Then

    '        'UPDATE문 실행
    '        Try
    '            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
    '            StrSQL = StrSQL & "UPDATE MC_STOCK_ORDER
    '                                      SET MO_NAME = '" & order_name.Text & "',
    '                                          MO_Bigo = '" & order_bigo.Text & "'
    '                                    WHERE MO_CODE ='" & Order_Code.Text & "'"
    '            Re_Count = DB_Execute()
    '            MsgBox("수정되었습니다")
    '        Catch ex As Exception
    '            MsgBox(ex.Message)
    '        End Try
    '        Me.Close()

    '    End If
    'End Sub

    Private Sub Btn_Save_Click(sender As Object, e As EventArgs) Handles Btn_Save.Click

        Dim DBT As New DataTable
        If Label1.Text = "insert" Then

            If GOLD.Text = "" Then
                GOLD.Text = 0
            End If

            Try
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO MC_STOCK_ORDER Values ('" & Order_Code.Text & "', '" & CM_Code.Text & "', '" & CM_Name.Text & "', '" & DAY.Text & "',
                                                                      '" & PL_CODE.Text & "', '" & PL_NAME.Text & "', '" & STANDARD.Text & "', '" & UNIT.Text & "', 
                                                                      '" & QTY.Text & "', '" & TOTAL.Text & "', '" & GOLD.Text & "', '" & BIGO.Text & "','진행중')"
                Re_Count = DB_Execute()
                MsgBox("저장되었습니다")


                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO MAN_LOG Values ('" & DateTime.Now & "', '" & Me.Name & "', '발주 추가', '" & Frm_Main.Text_Man_Name.Text & "',
                                                                     '" & Order_Code.Text & " 발주 입력')"
                Re_Count = DB_Execute()







            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Me.Close()

        ElseIf Label1.Text = "update" Then

            'UPDATE문 실행
            Try
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE MC_STOCK_ORDER
                                      SET MO_CM_CODE = '" & CM_Code.Text & "',
                                          MO_CM_NAME = '" & CM_Name.Text & "',
                                          MO_DAY = '" & DAY.Text & "',
                                          MO_PL_CODE = '" & PL_CODE.Text & "',
                                          MO_PL_NAME = '" & PL_NAME.Text & "',
                                          MO_STANDARD = '" & STANDARD.Text & "',
                                          MO_UNIT_AMOUNT = '" & UNIT.Text & "',
                                          MO_QTY = '" & QTY.Text & "',
                                          MO_TOTAL = '" & TOTAL.Text & "',
                                          MO_GOLD = '" & GOLD.Text & "',
                                          MO_BIGO = '" & BIGO.Text & "'
                                    WHERE MO_CODE ='" & Order_Code.Text & "'"
                Re_Count = DB_Execute()
                MsgBox("수정되었습니다")


                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO MAN_LOG Values ('" & DateTime.Now & "', '" & Me.Name & "', '발주 수정', '" & Frm_Main.Text_Man_Name.Text & "',
                                                                     '" & Order_Code.Text & " 발주 수정')"
                Re_Count = DB_Execute()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Me.Close()


            StrSQL = "SELECT * FROM MAN_LOG"
            Re_Count = DB_Select(DBT)
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim POPUP_ACCOUNT As New popup_account
        POPUP_ACCOUNT.ShowDialog()

        If POPUP_ACCOUNT.DialogResult = DialogResult.OK Then
            CM_Name.Text = POPUP_ACCOUNT.Customer_Name
            CM_Code.Text = POPUP_ACCOUNT.Customer_Code

        End If
    End Sub

    Private Sub btnJobSearch_Click(sender As Object, e As EventArgs) Handles btnJobSearch.Click
        Dim POPUP_ACCOUNT As New popup_account
        POPUP_ACCOUNT.ShowDialog()

        If POPUP_ACCOUNT.DialogResult = DialogResult.OK Then
            CM_Name.Text = POPUP_ACCOUNT.Customer_Name
            CM_Code.Text = POPUP_ACCOUNT.Customer_Code
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles BIGO.TextChanged

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim POPUP_PRODUCT As New popup_product
        POPUP_PRODUCT.ShowDialog()

        If POPUP_PRODUCT.DialogResult = DialogResult.OK Then
            PL_NAME.Text = POPUP_PRODUCT.Product_Name
            PL_CODE.Text = POPUP_PRODUCT.Product_Code

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim POPUP_PRODUCT As New popup_product
        POPUP_PRODUCT.ShowDialog()

        If POPUP_PRODUCT.DialogResult = DialogResult.OK Then
            PL_NAME.Text = POPUP_PRODUCT.Product_Name
            PL_CODE.Text = POPUP_PRODUCT.Product_Code

        End If
    End Sub

    Private Sub QTY_KeyUp(sender As Object, e As KeyEventArgs) Handles QTY.KeyUp
        If QTY.Text = "" Or UNIT.Text = "" Then
            Exit Sub
        Else
            TOTAL.Text = Val(QTY.Text) * Val(UNIT.Text)
        End If
    End Sub

    Private Sub UNIT_KeyUp(sender As Object, e As KeyEventArgs) Handles UNIT.KeyUp
        If QTY.Text = "" Or UNIT.Text = "" Then
            Exit Sub
        Else
            TOTAL.Text = Val(QTY.Text) * Val(UNIT.Text)
        End If
    End Sub
End Class