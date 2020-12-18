Public Class Order2_Insert
    Dim JP_Code As String
    Private Sub order_Insert_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim DBT As New DataTable
        Label1.Visible = False
        If Label1.Text = "insert" Then

            '번호자동생성
            StrSQL = "Select MI_CODE FROM MC_STOCK_IN with(NOLOCK) Order By MI_CODE Desc "
            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                JP_Code = "001"
            Else
                'JP_Code = Val(DBT.Rows(0).Item(0)) + 1
                JP_Code = Replace(DBT.Rows(0).Item(0), "MI", "")
                JP_Code = Format(JP_Code + 1, "00#")
            End If

            STANDARD.Text = STANDARD.Text & "KG"
            Order_Code.Text = "MI" & JP_Code
            DAY.Text = Today

        ElseIf Label1.Text = "update" Then
            Order_Code.Enabled = False
            CM_Code.Enabled = False
            CM_Name.Enabled = False
            PL_CODE.Enabled = False
            PL_NAME.Enabled = False


            STANDARD.ReadOnly = True
            QTY.ReadOnly = True

            Search_Btn1.Enabled = False
            Search_Btn2.Enabled = False
            Search_Btn3.Enabled = False
            Search_Btn4.Enabled = False
            Search_Btn5.Enabled = False

            ' BIGO.Enabled = False
            '    Btn_Save.Enabled = False
            'DB 데이터 불러오기
            StrSQL = "Select * FROM MC_STOCK_IN with(NOLOCK) Where MI_CODE = '" & Order_Code.Text & "' Order By MI_CODE"
            Re_Count = DB_Select(DBT)

            If Re_Count = 0 Then
                Me.Close()
                Exit Sub
            Else
                CM_Code.Text = DBT.Rows(0)("MI_CM_CODE")
                CM_Name.Text = DBT.Rows(0)("MI_CM_NAME")
                PL_CODE.Text = DBT.Rows(0)("MI_PL_CODE")
                PL_NAME.Text = DBT.Rows(0)("MI_PL_NAME")
                DAY.Text = DBT.Rows(0)("MI_DAY")
                STANDARD.Text = DBT.Rows(0)("MI_STANDARD")
                QTY.Text = DBT.Rows(0)("MI_QTY")
                TOTAL.Text = DBT.Rows(0)("MI_TOTAL")
                GOLD.Text = DBT.Rows(0)("MI_GOLD")
                BIGO.Text = DBT.Rows(0)("MI_BIGO")

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

    '        StrSQL = "Select MI_CODE FROM MC_STOCK_IN with(NOLOCK) Where MI_CODE = '" & Order_Code.Text & "' Order By MI_CODE"
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
    '            StrSQL = StrSQL & "Insert INTO MC_STOCK_IN Values ('" & Order_Code.Text & "', '" & order_name.Text & "', '" & order_bigo.Text & "')"
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
    '            StrSQL = StrSQL & "UPDATE MC_STOCK_IN
    '                                      SET MI_NAME = '" & order_name.Text & "',
    '                                          MI_Bigo = '" & order_bigo.Text & "'
    '                                    WHERE MI_CODE ='" & Order_Code.Text & "'"
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
                '입고입력
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO MC_STOCK_IN Values ('" & Order_Code.Text & "', '사무실',  '" & MO_CODE.Text & "','" & CM_Code.Text & "', '" & CM_Name.Text & "', '" & DAY.Text & "',
                                                                     '" & Frm_Main.Text_Man_Name.Text & "', '" & Shelf.Text & "', '" & PL_CODE.Text & "', '" & PL_NAME.Text & "', '" & STANDARD.Text & "', '" & UNIT.Text & "', 
                                                                      '" & QTY.Text & "', '" & TOTAL.Text & "', '" & GOLD.Text & "', '" & BIGO.Text & "')"
                Re_Count = DB_Execute()

                '발주개수 확인
                StrSQL = "SELECT MO_QTY FROM MC_STOCK_ORDER WHERE MO_CODE='" & MO_CODE.Text & "'"
                Re_Count = DB_Select(DBT)

                If Re_Count = 0 Then
                    Exit Sub
                Else
                    If DBT.Rows(0)("MO_QTY") = QTY.Text Then
                        '발주 진행중 -> 완료
                        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                        StrSQL = StrSQL & "UPDATE MC_STOCK_ORDER SET
                                          MO_ING ='완료'
                                   WHERE MO_CODE ='" & MO_CODE.Text & "'"
                        Re_Count = DB_Execute()
                    Else
                    End If


                End If


                '재고 테이블 추가
                Dim stock As Integer = 0
                StrSQL = "SELECT PL_QTY FROM STOCK WHERE PL_CODE='" & PL_CODE.Text & "'"
                Re_Count = DB_Select(DBT)

                If Re_Count = 0 Then
                    Exit Sub
                Else
                    stock = DBT.Rows(0)("PL_QTY")
                End If

                stock = stock + Val(QTY.Text)


                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE STOCK SET
                                          PL_QTY ='" & stock & "'
                                    WHERE PL_CODE ='" & PL_CODE.Text & "'"
                Re_Count = DB_Execute()


                Dim SR_CODE As String

                '재고이력코드 번호자동생성
                StrSQL = "Select SR_CODE FROM STOCK_RECORD with(NOLOCK) Order By SR_CODE Desc "
                Re_Count = DB_Select(DBT)
                If Re_Count = 0 Then
                    SR_CODE = "001"
                Else
                    'JP_Code = Val(DBT.Rows(0).Item(0)) + 1
                    SR_CODE = Replace(DBT.Rows(0).Item(0), "SR", "")
                    SR_CODE = Format(SR_CODE + 1, "00#")
                End If


                '재고이력 테이블 추가
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO STOCK_RECORD Values ('SR" & SR_CODE & "','" & Today & "','" & Order_Code.Text & "','입고','" & PL_CODE.Text & "','" & QTY.Text & "'
                                                                , '" & Frm_Main.Text_Man_Name.Text & "', '" & Today & "','','')"
                Re_Count = DB_Execute()


                'MANLOG 이력입력
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO MAN_LOG Values ('" & DateTime.Now & "', '" & Me.Name & "', '입고 추가', '" & Frm_Main.Text_Man_Name.Text & "',
                                                                     '" & Order_Code.Text & " 입고 입력')"
                Re_Count = DB_Execute()

                MsgBox("저장되었습니다")

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Me.Close()

        ElseIf Label1.Text = "update" Then

            'UPDATE문 실행 - 수량 등 기본 항목 변경 불가능*일부 항목만 수정 가능
            Try
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE MC_STOCK_IN
                                      SET 
                                          MI_SHELF = '" & Shelf.Text & "',
                                          MI_UNIT_AMOUNT = '" & UNIT.Text & "',
                                          MI_TOTAL = '" & TOTAL.Text & "',
                                          MI_GOLD = '" & GOLD.Text & "',
                                          MI_BIGO = '" & BIGO.Text & "'
                                    WHERE MI_CODE ='" & Order_Code.Text & "'"
                Re_Count = DB_Execute()

                'StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                'StrSQL = StrSQL & "UPDATE MC_STOCK_IN
                '                      SET MI_MO_CODE = '" & MO_CODE.Text & "',
                '                          MI_CM_CODE = '" & CM_Code.Text & "',
                '                          MI_CM_NAME = '" & CM_Name.Text & "',
                '                          MI_DAY = '" & DAY.Text & "',
                '                          MI_SHELF = '" & Shelf.Text & "',
                '                          MI_PL_CODE = '" & PL_CODE.Text & "',
                '                          MI_PL_NAME = '" & PL_NAME.Text & "',
                '                          MI_STANDARD = '" & STANDARD.Text & "',
                '                          MI_UNIT_AMOUNT = '" & UNIT.Text & "',
                '                          MI_QTY = '" & QTY.Text & "',
                '                          MI_TOTAL = '" & TOTAL.Text & "',
                '                          MI_GOLD = '" & GOLD.Text & "',
                '                          MI_BIGO = '" & BIGO.Text & "'
                '                    WHERE MI_CODE ='" & Order_Code.Text & "'"
                'Re_Count = DB_Execute()
                'MsgBox("수정되었습니다")


                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO MAN_LOG Values ('" & DateTime.Now & "', '" & Me.Name & "', '입고 수정', '" & Frm_Main.Text_Man_Name.Text & "',
                                                                     '" & Order_Code.Text & " 입고 수정')"
                Re_Count = DB_Execute()

                MsgBox("수정되었습니다")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Me.Close()


            StrSQL = "SELECT * FROM MAN_LOG"
            Re_Count = DB_Select(DBT)
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Search_Btn3.Click
        Dim POPUP_ACCOUNT As New popup_account
        POPUP_ACCOUNT.ShowDialog()

        If POPUP_ACCOUNT.DialogResult = DialogResult.OK Then
            CM_Name.Text = POPUP_ACCOUNT.Customer_Name
            CM_Code.Text = POPUP_ACCOUNT.Customer_Code

        End If
    End Sub

    Private Sub btnJobSearch_Click(sender As Object, e As EventArgs) Handles Search_Btn2.Click
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

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Search_Btn5.Click
        Dim POPUP_PRODUCT As New popup_product
        POPUP_PRODUCT.ShowDialog()

        If POPUP_PRODUCT.DialogResult = DialogResult.OK Then
            PL_NAME.Text = POPUP_PRODUCT.Product_Name
            PL_CODE.Text = POPUP_PRODUCT.Product_Code

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Search_Btn4.Click
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

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Search_Btn1.Click
        Dim form As New popup_order
        form.ShowDialog()

        If form.DialogResult = DialogResult.OK Then
            MO_CODE.Text = form.MO_CODE
            CM_Code.Text = form.MO_CM_CODE
            CM_Name.Text = form.MO_CM_NAME

            PL_CODE.Text = form.MO_PL_CODE
            PL_NAME.Text = form.MO_PL_NAME
            STANDARD.Text = form.MO_STANDARD

            UNIT.Text = form.MO_UNIT_AMOUNT
            QTY.Text = form.MO_QTY
            TOTAL.Text = form.MO_TOTAL

            GOLD.Text = form.MO_GOLD
        Else

        End If
    End Sub
End Class