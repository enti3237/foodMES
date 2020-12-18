﻿Public Class Out2
    Private Sub Go_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TopMost = True
        setup()
        setup2()
        Panel1.Visible = False
        D.ReadOnly = True
    End Sub

    Public Sub setup()
        Grid_Color(D)

        D.AllowUserToAddRows = False
        D.RowTemplate.Height = 70
        D.ColumnCount = 7
        D.RowCount = 0
        ' D.RowCount = 7

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        D.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        D.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        D.Columns(0).Width = 240
        D.Columns(1).Width = 300
        D.RowHeadersWidth = 10
        D.Columns(0).HeaderText = "라벨코드"
        D.Columns(1).HeaderText = "제품명"
        D.Columns(2).HeaderText = "수량"
        D.Columns(3).HeaderText = "입고일자"
        D.Columns(4).HeaderText = "입고시간"
        D.Columns(5).HeaderText = "거래처코드"
        D.Columns(6).HeaderText = "거래처명"

        D.ColumnHeadersHeight = 40


        D.ColumnHeadersHeight = 40
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '출고처리

        If D.Rows.Count < 1 Then
            Exit Sub
        End If

        If D.Item(0, 0).Value = "" Then
            Exit Sub
        End If

        Dim DBT As New DataTable
        Dim i As Integer
        Dim My_Date As String
        Dim My_Time As String
        Dim JP_Code As String = ""


        For i = 0 To D.RowCount - 1

            StrSQL = "Select GETDATE() "
            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                Exit Sub
            Else
                My_Date = DBT.Rows(0).Item(0)
                My_Time = DateTime.Now.ToString("HH:mm")
                JP_Code = Mid(My_Date, 1, 4) & Mid(My_Date, 6, 2) & Mid(My_Date, 9, 2)
                My_Time = Replace(My_Time, ":", "-")
                My_Date = Mid(My_Date, 1, 4) & "-" & Mid(My_Date, 6, 2) & "-" & Mid(My_Date, 9, 2)
            End If

            StrSQL = "UPDATE LABEL SET C_CHECK = '출고', OUT_DATE = '" & My_Date & "', OUT_TIME = '" & My_Time & "', 
                            CM_CODE = '" & D.Item(5, i).Value & "', CM_NAME = '" & D.Item(6, i).Value & "'
                   WHERE CODE ='" & D.Item(0, i).Value & "'"

            Re_Count = DB_Execute()

            '납품처리 

            Dim CM, CT As String
            '거래처 입력여부
            If D.Item(5, i).Value = "" Then
                '거래처없을때
                '납품전표 추가
                StrSQL = "Select DL_Code FROM SC_DELIVER with(NOLOCK) Where DL_Date Like  '%" & My_Date & "%' Order By DL_Code Desc "
                Re_Count = DB_Select(DBT)
                If Re_Count = 0 Then
                    JP_Code = JP_Code & "001"
                Else
                    JP_Code = JP_Code & Format(Val(Mid(DBT.Rows(0).Item(0), 11, 3)) + 1, "00#")
                End If

                '추가한다
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO SC_DELIVER (DL_Code, DL_Date, DL_Time, DL_Name, DL_Specifications, DL_Check)  
                                Values ('CP" & JP_Code & "', '" & My_Date & "','" & My_Time & "','태블릿자동','기본','' )"
                Re_Count = DB_Execute()
            Else

                StrSQL = "Select CM_MAN_NAME, CM_MAN_TEL FROM SI_CUSTOMER with(NOLOCK) Where CM_CODE =  '" & D.Item(5, i).Value & "' Order By CM_CODE Desc "
                Re_Count = DB_Select(DBT)

                If Re_Count <> 0 Then
                    CM = DBT.Rows(0)("CM_MAN_NAME")
                    CT = DBT.Rows(0)("CM_MAN_TEL")
                End If

                StrSQL = "Select DL_Code FROM SC_DELIVER with(NOLOCK) Where DL_Date Like  '%" & My_Date & "%' Order By DL_Code Desc "
                Re_Count = DB_Select(DBT)
                If Re_Count = 0 Then
                    JP_Code = JP_Code & "001"
                Else
                    JP_Code = JP_Code & Format(Val(Mid(DBT.Rows(0).Item(0), 11, 3)) + 1, "00#")
                End If

                '추가한다
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO SC_DELIVER (DL_Code, DL_Date, DL_Time, DL_Name, DL_Specifications, DL_Check, DL_CUSTOMER_CODE, DL_CUSTOMER_NAME, DL_CUSTOMER_MAN_NAME, DL_CUSTOMER_MAN_TEL)  
                                Values ('CP" & JP_Code & "', '" & My_Date & "','" & My_Time & "','태블릿자동','기본','','" & D.Item(5, i).Value & "','" & D.Item(6, i).Value & "', '" & CM & "', '" & CT & "' )"
                Re_Count = DB_Execute()

            End If

            Dim NCODE As String
            NCODE = "CP" & JP_Code

            Dim PC, PS, PU, PUG, GOL, VAT As String

            '제품정보 가져오기
            StrSQL = "Select PL_CODE, PL_STANDARD, PL_UNIT, PL_UNIT_GOLD FROM SI_PRODUCT with(NOLOCK) Where PL_NAME =  '" & D.Item(1, i).Value & "' Order By PL_NAME Desc "
            Re_Count = DB_Select(DBT)

            If Re_Count <> 0 Then
                'PC = DBT.Rows(0)("PL_CODE")
                'PS = DBT.Rows(0)("PL_STANDARD")
                'PU = DBT.Rows(0)("PL_UNIT") '단위
                'PUG = DBT.Rows(0)("PL_UNIT_GOLD") '단가

                If IsDBNull(DBT.Rows(0)("PL_CODE")) Then
                    PC = ""
                Else
                    PC = DBT.Rows(0)("PL_CODE")
                End If


                If IsDBNull(DBT.Rows(0)("PL_STANDARD")) Then
                    PS = ""
                Else
                    PS = DBT.Rows(0)("PL_STANDARD")
                End If


                If IsDBNull(DBT.Rows(0)("PL_UNIT")) Then
                    PU = ""
                Else
                    PU = DBT.Rows(0)("PL_UNIT")
                End If


                If IsDBNull(DBT.Rows(0)("PL_UNIT_GOLD")) Then
                    PUG = 0
                Else
                    PUG = DBT.Rows(0)("PL_UNIT_GOLD")
                End If
            End If

            ' MsgBox(PUG)
            ' MsgBox(D.Item(2, i).Value)

            '  Grid_Main.Item(8, Grid_Main.CurrentCell.RowIndex).Value = Val(Grid_Main.Item(6, Grid_Main.CurrentCell.RowIndex).Value.ToString) * Val(Grid_Main.Item(7, Grid_Main.CurrentCell.RowIndex).Value.ToString)
            '  Grid_Main.Item(9, Grid_Main.CurrentCell.RowIndex).Value = Int(Val(Grid_Main.Item(8, Grid_Main.CurrentCell.RowIndex).Value) * 0.1)
            GOL = (Val(D.Item(2, i).Value) * Val(PUG)).ToString
            VAT = (Val(GOL) * 0.1).ToString



            '그리드뷰 내용 세부내역으로 추가

            Dim name, qty As String

            Dim Db_Sun As Long
            StrSQL = "Select DL_Sun FROM SC_DELIVER_LIST with(NOLOCK) Where DL_Code = '" & NCODE & "' Order By Convert(Decimal,DL_Sun) Desc "
            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                Db_Sun = 1
            Else
                Db_Sun = DBT.Rows(0).Item(0) + 1
            End If



            '추가한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "Insert INTO SC_DELIVER_LIST Values ('" & NCODE & "', '" & Db_Sun & "','" & My_Date & "','" & PC & "', '" & D.Item(1, i).Value & "', '" & PS & "', '" & PU & "', '" & PUG & "', '" & D.Item(2, i).Value & "', '" & GOL & "','" & VAT & "','','')"
            Re_Count = DB_Execute()



            '개수처리

            Dim Val_Check As String
            Dim PP_G_Sun As String
            Dim PP_G_Amount As String

            Val_Check = ""

            StrSQL = "Select DL_Check FROM SC_DELIVER With(NOLOCK) Where DL_Code = '" & NCODE & "'"
            Re_Count = DB_Select(DBT)
            '  Grid_Info_Combo.Items.Clear()
            If Re_Count <> 0 Then
                If DBT.Rows(0)("DL_Check") = "처리" Then
                    Val_Check = "처리"
                    MsgBox("이미 처리 되었습니다. 삭제 후 다시 처리 하세요")
                    Exit Sub
                End If
            End If
            If Val_Check = "처리" Then
                '기존 Data를 삭제 한다.
                'Grid의 수량을 가지고 온다.


                PP_G_Sun = "0"
                PP_G_Amount = "0"
                StrSQL = "Select PP_Sun, PP_Amount FROM SI_Process_List with(NOLOCK) Where PP_Code = '" & PC & "' Order By Convert(Decimal,PP_Sun) Desc"
                Re_Count = DB_Select(DBT)
                If Re_Count > 0 Then
                    PP_G_Sun = DBT.Rows(0)("PP_Sun")
                    PP_G_Amount = DBT.Rows(0)("PP_Amount")
                End If
                If PP_G_Sun = "0" Then
                Else
                    '변경
                    PP_G_Amount = Val(PP_G_Amount) + Val(D.Item(2, i).Value)
                    StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                    StrSQL = StrSQL & "UPDATE SI_PROCESS_LIST Set PP_Amount = '" & PP_G_Amount & "' Where PP_Code = '" & PC & "' And PP_Sun = '" & PP_G_Sun & "'"
                    Re_Count = DB_Execute()
                End If


            End If

            '수량을 변경한다.

            PP_G_Sun = "0"
            PP_G_Amount = "0"
            StrSQL = "Select PP_Sun, PP_Amount FROM SI_Process_List with(NOLOCK) Where PP_Code = '" & PC & "' Order By Convert(Decimal,PP_Sun) Desc"
            Re_Count = DB_Select(DBT)
            If Re_Count > 0 Then
                PP_G_Sun = DBT.Rows(0)("PP_Sun")
                PP_G_Amount = DBT.Rows(0)("PP_Amount")
            End If
            If PP_G_Sun = "0" Then
            Else
                '변경
                PP_G_Amount = Val(PP_G_Amount) - Val(D.Item(2, i).Value)
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SI_PROCESS_LIST Set PP_Amount = '" & PP_G_Amount & "' Where PP_Code = '" & PC & "' And PP_Sun = '" & PP_G_Sun & "'"
                Re_Count = DB_Execute()
            End If


            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "UPDATE SC_DELIVER Set DL_Check = '처리' Where DL_Code = '" & NCODE & "'"
            Re_Count = DB_Execute()


            JP_Code = ""
            NCODE = ""

        Next i
        Me.Close()

        MsgBox("처리되었습니다")

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles BARCODE.KeyDown
        If e.KeyCode = Keys.Return Then
            BARCODE.Text = Trim(BARCODE.Text)

            If BARCODE.Text = "" Then
                Exit Sub
            End If

            ' MsgBox(BARCODE.Text)


            '바코드 확인
            Dim DBT As New DataTable
            Dim i As Integer
            Dim j As Integer

            '  Dim BAR, CODE As String

            ' MsgBox(BAR)
            ' MsgBox(CODE)
            StrSQL = "Select CODE
                        FROM LABEL
                       WHERE CODE = '" & BARCODE.Text & "'"

            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                MsgBox("입력하신 바코드는 유효하지 않습니다")
                BARCODE.Clear()
                Exit Sub
            Else
            End If

            StrSQL = "Select C_CHECK
                        FROM LABEL
                       WHERE CODE = '" & BARCODE.Text & "'"

            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                Exit Sub
            Else
                If DBT.Rows(0)("C_CHECK").Equals("출고") Then
                    MsgBox("이미 출고처리 되었습니다")
                    BARCODE.Clear()
                    Exit Sub
                ElseIf DBT.Rows(0)("C_CHECK").Equals("출력") Then
                    MsgBox("입고처리를 먼저 진행해주세요")
                    BARCODE.Clear()
                    Exit Sub

                End If
            End If


            'Dim Db_Sun As Long
            'StrSQL = "Select MATERIAL_OUT_Sun From MATERIAL_OUT_LIST with(NOLOCK) Where MATERIAL_OUT_Code = '" & BAR & "' Order By Convert(Decimal,MATERIAL_OUT_Sun) Desc "
            'Re_Count = DB_Select(DBT)
            'If Re_Count = 0 Then
            '    Db_Sun = 1
            'Else
            '    Db_Sun = DBT.Rows(0).Item(0) + 1
            'End If


            ''상세내역 추가
            'StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            'StrSQL = StrSQL & "Insert into MATERIAL_OUT_LIST (MATERIAL_OUT_Code, MATERIAL_OUT_Sun, DATIME)  Values ('" & BAR & "', '" & Db_Sun & "', '" & My_Date & "'/'" & My_Time & "')"
            'Re_Count = DB_Execute()

            Dim My_Date As String
            Dim My_Time As String

            StrSQL = "Select GETDATE() "
            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                Exit Sub
            Else
                My_Date = DBT.Rows(0).Item(0)
                My_Time = DateTime.Now.ToString("HH:mm")
                My_Time = Replace(My_Time, ":", "-")
                My_Date = Mid(My_Date, 1, 4) & "-" & Mid(My_Date, 6, 2) & "-" & Mid(My_Date, 9, 2)
            End If

            '     D.RowCount = D.RowCount + 1

            '중복체크
            '   MsgBox(BARCODE.Text)
            For i = 0 To D.Rows.Count - 1
                ' MsgBox("이미 같은 데이터가 들어가 있습니다.")
                If D.Item(0, i).Value.ToString.Equals(BARCODE.Text) Then
                    MsgBox("이미 같은 데이터가 들어가 있습니다.")
                    Exit Sub
                End If

            Next


            '그리드뷰에 옮기고
            StrSQL = "Select CODE, PL_NAME, QTY, GO_DATE, GO_TIME FROM LABEL WHERE CODE = '" & BARCODE.Text & "'"
            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                Exit Sub
            Else
                D.Rows.Add(DBT.Rows(0)("CODE"), DBT.Rows(0)("PL_NAME"), DBT.Rows(0)("QTY"), DBT.Rows(0)("GO_DATE"), DBT.Rows(0)("GO_TIME"))
            End If

            D.MultiSelect = False
            D.ClearSelection()
            BARCODE.Clear()
            ' Panel2.Visible = True
        End If
    End Sub

    Public Sub setup2()
        Grid_Color(DataGridView1)

        DataGridView1.AllowUserToAddRows = False
        DataGridView1.RowTemplate.Height = 50
        DataGridView1.ColumnCount = 2
        DataGridView1.RowCount = 5

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView1.RowHeadersWidth = 10
        DataGridView1.Columns(0).Width = 100
        DataGridView1.Columns(1).Width = 340
        DataGridView1.Columns(0).HeaderText = "거래처코드"
        DataGridView1.Columns(1).HeaderText = "거래처명"
        DataGridView1.ColumnHeadersHeight = 40


        'Datagridview1.Columns(0).ReadOnly = True
        'Datagridview1.Columns(1).ReadOnly = True

    End Sub

    Private Sub D_DoubleClick(sender As Object, e As EventArgs) Handles D.DoubleClick

        If D.RowCount = 0 Then
            Exit Sub
        End If

        If D.Item(0, 0).Value = "" Then
            Exit Sub
        End If

        If D.Item(0, D.CurrentCell.RowIndex).Value = Nothing Then
            Exit Sub
        End If

        Panel1.Visible = True
        display()
    End Sub

    Public Sub display()

        Dim DBT As New DataTable
        Dim i As Integer
        StrSQL = "Select CM_CODE, CM_NAME From SI_CUSTOMER With(NOLOCK) ORDER BY CM_CODE"
        Re_Count = DB_Select(DBT)
        If Re_Count > 0 Then
            DataGridView1.RowCount = Re_Count
            For i = 0 To DBT.Rows.Count - 1
                DataGridView1.Item(0, i).Value = DBT.Rows(i).Item(0).ToString
                DataGridView1.Item(1, i).Value = DBT.Rows(i).Item(1).ToString
            Next i
        Else
            DataGridView1.RowCount = 1
        End If
    End Sub

    Private Sub DataGridView1_DoubleClick(sender As Object, e As EventArgs) Handles DataGridView1.DoubleClick
        If DataGridView1.Item(0, DataGridView1.CurrentCell.RowIndex).Value = "" Then
            Exit Sub
        End If

        If DataGridView1.Item(0, DataGridView1.CurrentCell.RowIndex).Value = Nothing Then
            Exit Sub
        End If
        Dim code = DataGridView1.Item(0, DataGridView1.CurrentCell.RowIndex).Value
        Dim name As String = ""


        Dim DBT As New DataTable
        Dim i As Integer
        StrSQL = "Select CM_CODE, CM_NAME From SI_CUSTOMER With(NOLOCK) Where CM_Code = '" & code & "'"
        Re_Count = DB_Select(DBT)
        If Re_Count > 0 Then
            For i = 0 To DBT.Rows.Count - 1
                name = DBT.Rows(i)("CM_NAME")

            Next i
        Else
            MsgBox("등록되지 않은 거래처코드입니다.")
            Panel1.Visible = False
            Exit Sub
        End If



        D.Item(5, D.CurrentCell.RowIndex).Value = code
        D.Item(6, D.CurrentCell.RowIndex).Value = name


        Panel1.Visible = False
    End Sub
End Class