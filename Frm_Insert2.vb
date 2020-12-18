Public Class Frm_Insert2

    Private Sub Cancel_Com_Click(sender As Object, e As EventArgs) Handles Cancel_Com.Click

        Me.Close()

    End Sub

    Private Sub Insert_Com_Click(sender As Object, e As EventArgs) Handles Insert_Com.Click
        Dim DBT As New DataTable
        '추가 버튼 클릭

        Dim A, B, C, D As String
        Dim name1, name2, name3, name4 As String

        A = ComboBox1.Text
        B = ComboBox2.Text
        C = ComboBox3.Text
        D = A & "" & B & "" & C

        name1 = ComboBox4.Text
        name2 = ComboBox5.Text
        name3 = ComboBox6.Text

        name4 = name1 & "" & name2 & "" & name3

        If Me.Text = "대/중/소 코드 추가" Then

            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "Insert INTO SI_PRODUCT (PL_Code, PL_NAME)  Values ('" & D & "', '" & name4 & "')"
            Re_Count = DB_Execute()

            Me.Close()
            MsgBox("저장되었습니다")
            'If Len(t1.Text) = 3 Then
            'Else
            '    MsgBox("코드값은 3자리 숫자로 입력하세요.")
            '    Exit Sub
            'End If
            'StrSQL = "Select * FROM SI_MACHINE with(NOLOCK) Where PL_Code = '" & t1.Text & "'"
            'Re_Count = DB_Select(DBT)
            'If Re_Count = 0 Then
            '    '추가한다
            '    StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            '    StrSQL = StrSQL & "Insert INTO SI_MACHINE (PL_Code, PL_COUNT)  Values ('" & t1.Text & "', '0')"
            '    Re_Count = DB_Execute()
            '    '다시 보여 준다
            '    Me.Close()
            '    
            'Else
            '    '생산
            '    MsgBox("생산라인 코드 '" & t1.Text & "'는 '" & DBT.Rows(0).Item(1) & "' 로 등록 되어 있습니다.")
            '    Exit Sub
            'End If
        End If


    End Sub


    Private Sub Frm_Insert_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Frm_Insert_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert into Temp (Temp)  Values ('취소')"
        Re_Count = DB_Execute()
        Frm_Main.Enabled = True
    End Sub

    Private Sub Search_Text_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
            MsgBox("문자는 입력할 수 없습니다.")
        End If
    End Sub

    Private Sub ComboBox1_Click(sender As Object, e As EventArgs) Handles ComboBox1.Click
        Dim DBT As New DataTable

        StrSQL = "Select Base_Code, Base_Name FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '대' Order By Base_Name"
        Re_Count = DB_Select(DBT)
        ComboBox1.Items.Clear()
        If Re_Count <> 0 Then
            For i = 0 To Re_Count - 1
                ComboBox1.Items.Add(DBT.Rows(i)("Base_Code"))
                '   ComboBox4.Items.Add(DBT.Rows(i)("Base_Code"))
            Next i 'DBT.Rows(i)("Base_Bigo")
        End If

    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
        Dim DBT As New DataTable

        ComboBox1.Text = ComboBox1.SelectedItem.ToString()

        StrSQL = "Select Base_Name FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '대' and BASE_CODE ='" & ComboBox1.Text & "' Order By Base_Name"
        Re_Count = DB_Select(DBT)
        '  ComboBox1.Items.Clear()
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
                '   ComboBox4.Items.Add(DBT.Rows(i)("Base_Code"))
            Next i 'DBT.Rows(i)("Base_Bigo")
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
                '   ComboBox4.Items.Add(DBT.Rows(i)("Base_Code"))
            Next i 'DBT.Rows(i)("Base_Bigo")
        End If
    End Sub

    Private Sub ComboBox2_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox2.SelectionChangeCommitted
        Dim DBT As New DataTable

        ComboBox2.Text = ComboBox2.SelectedItem.ToString()

        StrSQL = "Select Base_Name FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '중' and BASE_CODE ='" & ComboBox2.Text & "' Order By Base_Name"
        Re_Count = DB_Select(DBT)
        '  ComboBox1.Items.Clear()
        If Re_Count <> 0 Then
            ComboBox5.Text = DBT.Rows(0)("Base_Name")
        End If

    End Sub

    Private Sub ComboBox3_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox3.SelectionChangeCommitted
        Dim DBT As New DataTable

        ComboBox3.Text = ComboBox3.SelectedItem.ToString()

        StrSQL = "Select Base_Name FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '소' and BASE_CODE ='" & ComboBox3.Text & "' Order By Base_Name"
        Re_Count = DB_Select(DBT)
        '  ComboBox1.Items.Clear()
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
                '   ComboBox4.Items.Add(DBT.Rows(i)("Base_Code"))
            Next i 'DBT.Rows(i)("Base_Bigo")
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
                '   ComboBox4.Items.Add(DBT.Rows(i)("Base_Code"))
            Next i 'DBT.Rows(i)("Base_Bigo")
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
                '   ComboBox4.Items.Add(DBT.Rows(i)("Base_Code"))
            Next i 'DBT.Rows(i)("Base_Bigo")
        End If
    End Sub

    Private Sub ComboBox4_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox4.SelectionChangeCommitted
        Dim DBT As New DataTable

        ComboBox4.Text = ComboBox4.SelectedItem.ToString()

        StrSQL = "Select Base_Code FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '대' and BASE_NAME ='" & ComboBox4.Text & "' Order By Base_Name"
        Re_Count = DB_Select(DBT)
        '  ComboBox1.Items.Clear()
        If Re_Count <> 0 Then
            ComboBox1.Text = DBT.Rows(0)("Base_Code")
        End If
    End Sub

    Private Sub ComboBox5_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox5.SelectionChangeCommitted
        Dim DBT As New DataTable

        ComboBox5.Text = ComboBox5.SelectedItem.ToString()

        StrSQL = "Select Base_Code FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '중' and BASE_NAME ='" & ComboBox5.Text & "' Order By Base_Name"
        Re_Count = DB_Select(DBT)
        '  ComboBox1.Items.Clear()
        If Re_Count <> 0 Then
            ComboBox2.Text = DBT.Rows(0)("Base_Code")
        End If
    End Sub

    Private Sub ComboBox6_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox6.SelectionChangeCommitted
        Dim DBT As New DataTable

        ComboBox6.Text = ComboBox6.SelectedItem.ToString()

        StrSQL = "Select Base_Code FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '소' and BASE_NAME ='" & ComboBox6.Text & "' Order By Base_Name"
        Re_Count = DB_Select(DBT)
        '  ComboBox1.Items.Clear()
        If Re_Count <> 0 Then
            ComboBox3.Text = DBT.Rows(0)("Base_Code")
        End If
    End Sub


End Class