Public Class Material_Insert
    Private Sub Insert_Marterial_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim DBT As New DataTable
        Label1.Visible = False
        Dim JP_Code As String
        If Label1.Text = "insert" Then
            '콤보박스 미리 띄워놓기

            txt_Sel.Text = "원부재료"
            'ComboBox_Standard.Text = "-"

            ''구분(완제품/원부재료 - 원부재료입력칸이라 만들 필요가 없었다
            'StrSQL = "Select Base_Name FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '제품종류' Order By Base_Sun"
            'Re_Count = DB_Select(DBT)

            'If Re_Count <> 0 Then
            '    ComboBox_Sel.Text = (DBT.Rows(0)("Base_Name"))
            'End If

            '규격(3kg, 1박스 - 직접입력으로 변경
            'StrSQL = "Select Base_Name FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '단위' Order By Base_Sun"
            'Re_Count = DB_Select(DBT)

            'If Re_Count <> 0 Then
            '    ComboBox_Standard.Text = (DBT.Rows(0)("Base_Name"))
            'End If

            '번호자동생성
            StrSQL = "Select PL_CODE FROM SI_PRODUCT with(NOLOCK)  Order By Convert(Decimal,PL_CODE) Desc "
            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                JP_Code = "001"
            Else
                'JP_Code = Val(DBT.Rows(0).Item(0)) + 1
                JP_Code = Format(DBT.Rows(0).Item(0) + 1, "00#")
            End If

            txt_Matecode.Text = JP_Code

        End If


        If Label1.Text = "update" Then

            'DB 데이터 불러오기
            StrSQL = "Select * FROM SI_PRODUCT with(NOLOCK) Where PL_CODE = '" & txt_Matecode.Text & "' Order By PL_CODE"
            Re_Count = DB_Select(DBT)

            If Re_Count = 0 Then
                Me.Close()
                Exit Sub
            Else
                txt_Matecode.Text = DBT.Rows(0)("PL_Code")
                txt_Matename.Text = DBT.Rows(0)("PL_Name")
                txt_Sel.Text = "원부재료"
                txt_Standard.Text = DBT.Rows(0)("PL_Standard")
                txt_Matebigo.Text = DBT.Rows(0)("PL_Bigo")
            End If

        End If

    End Sub

    Private Sub Btn_Save_Click(sender As Object, e As EventArgs) Handles Btn_Save.Click

        '유효성 검사
        If txt_Matecode.Text = "" Then
            MsgBox("제품코드는 공백일 수 없습니다")
            Exit Sub
        End If

        If Label1.Text = "insert" Then

            '기존 코드가 존재하는지 확인 후 INSERT문 실행
            Dim DBT As New DataTable

            StrSQL = "Select PL_CODE FROM SI_PRODUCT with(NOLOCK) Where PL_CODE = '" & txt_Matecode.Text & "' Order By PL_CODE"
            Re_Count = DB_Select(DBT)

            If Re_Count = 1 Then
                MsgBox("현재 같은 코드가 존재합니다")
                Exit Sub
            Else
            End If

            Try
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO SI_PRODUCT (PL_CODE, PL_NAME, PL_SEL, PL_STANDARD, PL_BIGO) Values ('" & txt_Matecode.Text & "', '" & txt_Matename.Text & "', '" & "원부재료" & "', '" & txt_Standard.Text & "', '" & txt_Matebigo.Text & "')"
                Re_Count = DB_Execute()
                MsgBox("저장되었습니다")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            '재고테이블에 추가
            Try
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO STOCK (PL_SEL, PL_CODE, PL_NAME, PL_QTY) Values ('" & txt_Sel.Text & "','" & txt_Matecode.Text & "', '" & txt_Matename.Text & "', '0')"
                Re_Count = DB_Execute()
                '  MsgBox("저장되었습니다")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try


            Me.Close()

        ElseIf Label1.Text = "update" Then

            'UPDATE문 실행
            Try
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SI_PRODUCT 
                                      SET PL_NAME = '" & txt_Matename.Text & "',
                                          PL_SEL = '" & txt_Sel.Text & "',
                                          PL_STANDARD = '" & txt_Standard.Text & "',
                                          PL_BIGO = '" & txt_Matebigo.Text & "'
                                    WHERE PL_CODE ='" & txt_Matecode.Text & "'"
                Re_Count = DB_Execute()
                MsgBox("수정되었습니다")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Me.Close()


        End If

    End Sub


    Private Sub ComboBox_Sel_SelectionChangeCommitted(sender As Object, e As EventArgs)
        'MsgBox("야호")
    End Sub
End Class