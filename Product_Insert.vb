Public Class Product_Insert
    '작성자 : 태승업
    Private Sub Product_Insert_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim DBT As New DataTable
        Label1.Visible = False
        If Label1.Text = "insert" Then
            '콤보박스 미리 띄워놓기

            ComboBox1.Text = "-"
            ComboBox2.Text = "-"
            ComboBox3.Text = "-"
            ComboBox4.Text = "-"
            ComboBox5.Text = "-"
            ComboBox6.Text = "-"

            StrSQL = "Select Base_Name FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '제품종류' Order By Base_Sun"
            Re_Count = DB_Select(DBT)

            If Re_Count <> 0 Then
                ComboBox7.Text = (DBT.Rows(0)("Base_Name"))
            End If

        ElseIf Label1.Text = "update" Then
            Insert_Com.Enabled = False
            ComboBox1.Enabled = False
            ComboBox2.Enabled = False
            ComboBox3.Enabled = False
            ComboBox4.Enabled = False
            ComboBox5.Enabled = False
            ComboBox6.Enabled = False


            'DB 데이터 불러오기
            StrSQL = "Select * FROM SI_PRODUCT with(NOLOCK) Where PL_CODE = '" & product_code.Text & "' Order By PL_CODE"
            Re_Count = DB_Select(DBT)

            If Re_Count = 0 Then
                Me.Close()
                Exit Sub
            Else
                product_name.Text = DBT.Rows(0)("PL_Name")
                ComboBox7.Text = DBT.Rows(0)("PL_SEL")
                standard.Text = DBT.Rows(0)("PL_STANDARD")
                shelf_date.Text = DBT.Rows(0)("PL_Date")
                bigo.Text = DBT.Rows(0)("PL_Bigo")
            End If
        End If

    End Sub

    Private Sub Insert_Com_Click(sender As Object, e As EventArgs) Handles Insert_Com.Click
        '조합식
        Dim DBT As New DataTable
        Dim A, B, C, D As String
        Dim name1, name2, name3, name4 As String

        A = ComboBox1.Text
        B = ComboBox2.Text
        C = ComboBox3.Text
        D = A & "" & B & "" & C

        product_code.Text = D

        name1 = ComboBox4.Text
        name2 = ComboBox5.Text
        name3 = ComboBox6.Text
        name4 = name1 & "" & name2 & "" & name3

        product_name.Text = name4

    End Sub

    Private Sub Btn_Save_Click(sender As Object, e As EventArgs) Handles Btn_Save.Click

        '유효성 검사
        If product_code.Text = "" Then
            MsgBox("제품코드는 공백일 수 없습니다")
            Exit Sub
        End If

        If Label1.Text = "insert" Then

            '기존 코드가 존재하는지 확인 후 INSERT문 실행
            Dim DBT As New DataTable

            StrSQL = "Select PL_CODE FROM SI_PRODUCT with(NOLOCK) Where PL_CODE = '" & product_code.Text & "' Order By PL_CODE"
            Re_Count = DB_Select(DBT)

            If Re_Count = 1 Then
                MsgBox("현재 같은 코드가 존재합니다")
                Exit Sub
            Else
            End If

            Try
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO SI_PRODUCT Values ('" & product_code.Text & "', '" & product_name.Text & "', '" & ComboBox7.Text & "', '" & standard.Text & "','" & shelf_date.Text & "', '" & bigo.Text & "')"
                Re_Count = DB_Execute()
                MsgBox("저장되었습니다")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Me.Close()


            '재고테이블에 추가
            Try
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO STOCK (PL_SEL, PL_CODE, PL_NAME, PL_QTY) Values ('" & ComboBox7.Text & "','" & product_code.Text & "', '" & product_name.Text & "', '0')"
                Re_Count = DB_Execute()
                ' MsgBox("저장되었습니다")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

        ElseIf Label1.Text = "update" Then

            'UPDATE문 실행
            Try
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SI_PRODUCT 
                                      SET PL_NAME = '" & product_name.Text & "',
                                          PL_SEL = '" & ComboBox7.Text & "',
                                          PL_STANDARD = '" & standard.Text & "',
                                          PL_DATE = '" & shelf_date.Text & "',
                                          PL_BIGO = '" & bigo.Text & "'
                                    WHERE PL_CODE ='" & product_code.Text & "'"
                Re_Count = DB_Execute()
                MsgBox("수정되었습니다")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Me.Close()


        End If



    End Sub

    Private Sub ComboBox7_Click(sender As Object, e As EventArgs) Handles ComboBox7.Click
        Dim DBT As New DataTable

        StrSQL = "Select Base_Name FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '제품종류' Order By Base_Sun"
        Re_Count = DB_Select(DBT)
        ComboBox7.Items.Clear()
        If Re_Count <> 0 Then
            For i = 0 To Re_Count - 1
                ComboBox7.Items.Add(DBT.Rows(i)("Base_Name"))
            Next i
        End If
    End Sub

    Private Sub ComboBox7_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox7.SelectionChangeCommitted
        ComboBox7.Text = ComboBox7.SelectedItem.ToString()
    End Sub

    Private Sub ComboBox1_Click(sender As Object, e As EventArgs) Handles ComboBox1.Click
        Dim DBT As New DataTable

        StrSQL = "Select Base_Code, Base_Name FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '대' Order By Base_Name"
        Re_Count = DB_Select(DBT)
        ComboBox1.Items.Clear()
        If Re_Count <> 0 Then
            For i = 0 To Re_Count - 1
                ComboBox1.Items.Add(DBT.Rows(i)("Base_Code"))
            Next i
        End If

    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
        Dim DBT As New DataTable

        ComboBox1.Text = ComboBox1.SelectedItem.ToString()

        StrSQL = "Select Base_Name FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '대' and BASE_CODE ='" & ComboBox1.Text & "' Order By Base_Name"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            ComboBox4.Text = DBT.Rows(0)("Base_Name")
        End If

    End Sub

    Private Sub ComboBox2_Click(sender As Object, e As EventArgs) Handles ComboBox2.Click
        Dim DBT As New DataTable

        StrSQL = "Select Base_Code, Base_Name FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '중' Order By Base_Name"
        Re_Count = DB_Select(DBT)
        ComboBox2.Items.Clear()
        If Re_Count <> 0 Then
            For i = 0 To Re_Count - 1
                ComboBox2.Items.Add(DBT.Rows(i)("Base_Code"))
            Next i
        End If
    End Sub

    Private Sub ComboBox3_Click(sender As Object, e As EventArgs) Handles ComboBox3.Click
        Dim DBT As New DataTable

        StrSQL = "Select Base_Code, Base_Name FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '소' Order By Base_Name"
        Re_Count = DB_Select(DBT)
        ComboBox3.Items.Clear()
        If Re_Count <> 0 Then
            For i = 0 To Re_Count - 1
                ComboBox3.Items.Add(DBT.Rows(i)("Base_Code"))
            Next i
        End If
    End Sub

    Private Sub ComboBox2_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox2.SelectionChangeCommitted
        Dim DBT As New DataTable

        ComboBox2.Text = ComboBox2.SelectedItem.ToString()

        StrSQL = "Select Base_Name FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '중' and BASE_CODE ='" & ComboBox2.Text & "' Order By Base_Name"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            ComboBox5.Text = DBT.Rows(0)("Base_Name")
        End If

    End Sub

    Private Sub ComboBox3_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox3.SelectionChangeCommitted
        Dim DBT As New DataTable

        ComboBox3.Text = ComboBox3.SelectedItem.ToString()

        StrSQL = "Select Base_Name FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '소' and BASE_CODE ='" & ComboBox3.Text & "' Order By Base_Name"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            ComboBox6.Text = DBT.Rows(0)("Base_Name")
        End If
    End Sub

    Private Sub ComboBox4_Click(sender As Object, e As EventArgs) Handles ComboBox4.Click
        Dim DBT As New DataTable

        StrSQL = "Select Base_Code, Base_Name FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '대' Order By Base_Name"
        Re_Count = DB_Select(DBT)
        ComboBox4.Items.Clear()
        If Re_Count <> 0 Then
            For i = 0 To Re_Count - 1
                ComboBox4.Items.Add(DBT.Rows(i)("Base_Name"))
            Next i
        End If
    End Sub

    Private Sub ComboBox5_Click(sender As Object, e As EventArgs) Handles ComboBox5.Click
        Dim DBT As New DataTable

        StrSQL = "Select Base_Code, Base_Name FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '중' Order By Base_Name"
        Re_Count = DB_Select(DBT)
        ComboBox5.Items.Clear()
        If Re_Count <> 0 Then
            For i = 0 To Re_Count - 1
                ComboBox5.Items.Add(DBT.Rows(i)("Base_Name"))
            Next i
        End If
    End Sub

    Private Sub ComboBox6_Click(sender As Object, e As EventArgs) Handles ComboBox6.Click
        Dim DBT As New DataTable

        StrSQL = "Select Base_Code, Base_Name FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '소' Order By Base_Name"
        Re_Count = DB_Select(DBT)
        ComboBox6.Items.Clear()
        If Re_Count <> 0 Then
            For i = 0 To Re_Count - 1
                ComboBox6.Items.Add(DBT.Rows(i)("Base_Name"))
            Next i
        End If
    End Sub

    Private Sub ComboBox4_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox4.SelectionChangeCommitted
        Dim DBT As New DataTable

        ComboBox4.Text = ComboBox4.SelectedItem.ToString()

        StrSQL = "Select Base_Code FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '대' and BASE_NAME ='" & ComboBox4.Text & "' Order By Base_Name"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            ComboBox1.Text = DBT.Rows(0)("Base_Code")
        End If
    End Sub

    Private Sub ComboBox5_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox5.SelectionChangeCommitted
        Dim DBT As New DataTable

        ComboBox5.Text = ComboBox5.SelectedItem.ToString()

        StrSQL = "Select Base_Code FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '중' and BASE_NAME ='" & ComboBox5.Text & "' Order By Base_Name"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            ComboBox2.Text = DBT.Rows(0)("Base_Code")
        End If
    End Sub

    Private Sub ComboBox6_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox6.SelectionChangeCommitted
        Dim DBT As New DataTable

        ComboBox6.Text = ComboBox6.SelectedItem.ToString()

        StrSQL = "Select Base_Code FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '소' and BASE_NAME ='" & ComboBox6.Text & "' Order By Base_Name"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            ComboBox3.Text = DBT.Rows(0)("Base_Code")
        End If
    End Sub


End Class