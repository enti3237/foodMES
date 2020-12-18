Public Class Out
    Private Sub Out_Load(sender As Object, e As EventArgs)
        Me.TopMost = True
        setup()
        setup2()
        D.ReadOnly = True
        Panel1.Visible = False
        BARCODE.Select()
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

    Public Sub setup()
        Grid_Color(D)

        D.AllowUserToAddRows = False
        D.RowTemplate.Height = 50
        D.ColumnCount = 7
        D.RowCount = 0
        ' D.RowCount = 7

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        D.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        D.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        D.Columns(0).Width = 240
        D.Columns(1).Width = 240
        D.RowHeadersWidth = 10
        D.Columns(0).HeaderText = "라벨코드"
        D.Columns(1).HeaderText = "제품명"
        D.Columns(2).HeaderText = "수량"
        D.Columns(3).HeaderText = "입고일자"
        D.Columns(4).HeaderText = "입고시간"
        D.Columns(5).HeaderText = "거래처코드"
        D.Columns(6).HeaderText = "거래처명"

        D.ColumnHeadersHeight = 40

        'Datagridview1.Columns(0).ReadOnly = True
        'Datagridview1.Columns(1).ReadOnly = True

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

        For i = 0 To D.RowCount - 1
            StrSQL = "UPDATE LABEL SET C_CHECK = '출고', OUT_DATE = '" & My_Date & "', OUT_TIME = '" & My_Time & "',
                                    CM_CODE = '" & D.Item(5, i).Value & ", CM_NAME = '" & D.Item(6, i).Value & "'
                   WHERE CODE ='" & D.Item(0, i).Value & "'"

            Re_Count = DB_Execute()
        Next


        MsgBox("출고처리 되었습니다")
        Me.Hide()


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
        DataGridView1.Columns(1).Width = 240
        DataGridView1.Columns(0).HeaderText = "거래처코드"
        DataGridView1.Columns(1).HeaderText = "거래처명"
        DataGridView1.ColumnHeadersHeight = 40


        'Datagridview1.Columns(0).ReadOnly = True
        'Datagridview1.Columns(1).ReadOnly = True

    End Sub

    Private Sub D_DoubleClick(sender As Object, e As EventArgs) Handles D.DoubleClick

        If D.Item(0, 0).Value = "" Then
            Exit Sub
        End If

        If D.Item(0, DataGridView1.CurrentCell.RowIndex).Value = Nothing Then
            Exit Sub
        End If

        Panel1.Visible = True
        display()

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


        DataGridView1.Item(5, DataGridView1.CurrentCell.RowIndex).Value = code
        DataGridView1.Item(6, DataGridView1.CurrentCell.RowIndex).Value = name


        Panel1.Visible = False

    End Sub

    Private Sub BARCODE_KeyDown(sender As Object, e As KeyEventArgs) Handles BARCODE.KeyDown
        If e.KeyCode = Keys.Return Then

            If BARCODE.Text = "" Then
                Exit Sub
            End If
            BARCODE.Text = Trim(BARCODE.Text)
            ' MsgBox(BARCODE.Text)


            '바코드 확인
            Dim DBT As New DataTable
            Dim i As Integer
            Dim j As Integer

            '  Dim BAR, CODE As String

            ' MsgBox(BAR)
            ' MsgBox(CODE)

            StrSQL = "Select CODE FROM LABEL WHERE CODE = '" & BARCODE.Text & "'"
            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                MsgBox("입력하신 바코드는 유효하지 않습니다")
                BARCODE.Clear()
                Exit Sub
            Else
            End If

            'Dim My_Date As String
            'Dim My_Time As String

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
                End If
            End If

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


            '그리드뷰에 옮기고
            StrSQL = "Select CODE, PL_NAME, QTY, GO_DATE, GO_TIME FROM LABEL"
            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                Exit Sub
            Else
                D.Rows.Add(DBT.Rows(0)("CODE"), DBT.Rows(0)("PL_NAME"), DBT.Rows(0)("QTY"), DBT.Rows(0)("GO_DATE"), DBT.Rows(0)("GO_TIME"))
            End If
        End If

        BARCODE.Clear()
    End Sub

    Private Sub S(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class