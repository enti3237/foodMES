﻿Public Class Go
    Private Sub Go_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TopMost = True
        setup()
        D.ReadOnly = True
    End Sub

    Public Sub setup()
        Grid_Color(D)

        D.AllowUserToAddRows = False
        D.RowTemplate.Height = 50
        D.ColumnCount = 3
        D.RowCount = 0

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        D.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        D.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        D.RowHeadersWidth = 10
        D.Columns(0).Width = 240
        D.Columns(1).Width = 240
        D.Columns(2).Width = 120

        D.Columns(0).HeaderText = "라벨코드"
        D.Columns(1).HeaderText = "제품명"
        D.Columns(2).HeaderText = "수량"


        D.ColumnHeadersHeight = 40
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '입고처리

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
            StrSQL = "UPDATE LABEL SET C_CHECK = '입고', GO_DATE = '" & My_Date & "', GO_TIME = '" & My_Time & "'
                   WHERE CODE ='" & D.Item(0, i).Value & "'"
            Re_Count = DB_Execute()

        Next

        MsgBox("입고처리 되었습니다")
        Me.Close()


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
                If DBT.Rows(0)("C_CHECK").Equals("입고") Then
                    MsgBox("이미 입고처리 되었습니다")
                    BARCODE.Clear()
                    Exit Sub
                ElseIf DBT.Rows(0)("C_CHECK").Equals("출고") Then
                    MsgBox("이미 출고처리 되었습니다")
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

            For i = 0 To D.Rows.Count - 1
                If D.Item(0, i).Value.ToString.Equals(BARCODE.Text) Then
                    MsgBox("이미 같은 데이터가 들어가 있습니다.")
                    Exit Sub
                End If

            Next

            '그리드뷰에 옮기고
            StrSQL = "Select CODE, PL_NAME, QTY FROM LABEL WHERE CODE = '" & BARCODE.Text & "'"
            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                Exit Sub
            Else
                D.Rows.Add(DBT.Rows(0)("CODE"), DBT.Rows(0)("PL_NAME"), DBT.Rows(0)("QTY"))
            End If

            D.MultiSelect = False
            D.ClearSelection()
            BARCODE.Clear()
            ' Panel2.Visible = True
        End If
    End Sub
End Class