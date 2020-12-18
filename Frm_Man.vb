﻿Public Class Frm_Man
    Dim Grid_Display_Ch As Boolean
    Dim Man_Grid_Row As Integer
    Dim Man_Grid_Col As Integer
    Public hellocode As String
    Private Sub Frm_Man_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()
        Grid_Display_Ch = False

        Grid_Setup()
        Grid_Display()
        Man_Grid.ClearSelection()
    End Sub

    Public Function Grid_Setup() As Boolean
        Dim i As Integer
        Grid_Color(Man_Grid)
        Man_Grid.EnableHeadersVisualStyles = False

        Man_Grid.AllowUserToAddRows = False
        Man_Grid.RowTemplate.Height = 20
        Man_Grid.ColumnCount = 11
        Man_Grid.RowCount = 1


        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Man_Grid.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
        Man_Grid.RowHeadersWidth = 45
        Man_Grid.Columns(0).Width = 55
        For i = 1 To Man_Grid.ColumnCount - 1
            Man_Grid.Columns(i).Width = 100
            Man_Grid.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Next i

        Man_Grid.Columns(0).HeaderText = "코드"
        Man_Grid.Columns(1).HeaderText = "성명"
        'Man_Grid.Columns(2).HeaderText = "업체명"
        Man_Grid.Columns(2).HeaderText = "부서"
        Man_Grid.Columns(3).HeaderText = "직급"
        Man_Grid.Columns(4).HeaderText = "성별"
        'Man_Grid.Columns(6).HeaderText = "생년월일"
        Man_Grid.Columns(5).HeaderText = "내/외국인"
        Man_Grid.Columns(6).HeaderText = "전화번호"
        'Man_Grid.Columns(9).HeaderText = "차량번호"
        'Man_Grid.Columns(10).HeaderText = "주소"
        Man_Grid.Columns(7).HeaderText = "암호"
        Man_Grid.Columns(8).HeaderText = "사용Level"
        'Man_Grid.Columns(13).HeaderText = "사원구분"
        Man_Grid.Columns(9).HeaderText = "비고"
        Man_Grid.Columns(10).HeaderText = "보건증 갱신일자"


        Man_Grid.Columns(0).ReadOnly = True
        Grid_Setup = True
    End Function
    Public Function Grid_Display() As Boolean
        Grid_Display_Ch = False

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Man_Grid.RowCount = 1
        StrSQL = "Select * FROM SI_MAN with(NOLOCK) Where Man_Name Like '%" & Search_Text.Text & "%'  Or Man_Name is Null Order  By Man_Name"

        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Man_Grid.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To Man_Grid.ColumnCount - 1
                    Man_Grid.Item(j, i).Value = DBT.Rows(i).Item(j)
                Next j
            Next i
        Else
            Man_Grid.RowCount = 1
            For j = 0 To Man_Grid.ColumnCount - 1
                Man_Grid.Item(j, 0).Value = ""
            Next j
        End If
        Grid_Display = True
        Grid_Display_Ch = True

    End Function

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        Grid_Display()

    End Sub

    Private Sub Save_Com_Click(sender As Object, e As EventArgs) Handles Save_Com.Click
        Grid_Display_Ch = False

        Dim Table_Name As String
        Dim i As Integer
        Dim j As Long
        Dim k As Integer
        Dim Field_Name(500) As String
        Dim DBT As DataTable = Nothing
        Dim RRR As String
        Table_Name = "SI_Man"
        j = 0
        StrSQL = "Select sys.Columns.Name From sys.Tables with(NOLOCK) , sys.Columns with(NOLOCK) Where sys.Tables.name = '" & Table_Name & "' And sys.Tables.Object_id = sys.Columns.Object_id Order By sys.Columns.column_id"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            For i = 0 To DBT.Rows.Count - 1
                Field_Name(j) = DBT.Rows(i)("Name").ToString
                j = j + 1
            Next i
        End If
        j = j - 1

        For i = 0 To Man_Grid.Rows.Count - 1
            If Man_Grid.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE " & Table_Name & " SET "
                For k = 1 To j
                    'StrSQL = StrSQL & " " & Table_Name & "." & Field_Name(k) & " = '" & Replace(Man_Grid.Item(k, i).Value.ToString, "'", "''") & "'"
                    RRR = IIf(IsDBNull(Man_Grid.Item(k, i).Value), "", Man_Grid.Item(k, i).Value)

                    StrSQL = StrSQL & " " & Table_Name & "." & Field_Name(k) & " = '" & RRR & "'"

                    If k <> j Then
                        StrSQL = StrSQL & ","
                    End If
                Next k
                StrSQL = StrSQL & " WHERE " & Table_Name & "." & Field_Name(0) & " = '" & Man_Grid.Item(0, i).Value.ToString & "'"
                Re_Count = DB_Execute()
                Man_Grid.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display_Ch = True

        MsgBox("저장되었습니다")

    End Sub

    Private Sub Insert_Com_Click(sender As Object, e As EventArgs) Handles Insert_Com.Click
        '최종 값으로 추가 한다.
        '새로운 코드 추가
        Dim DBT As New DataTable
        Dim JP_Code As Long
        Dim check As Integer

        StrSQL = "Select Man_Code FROM SI_MAN with(NOLOCK)  Order By Convert(Decimal,Man_Code) Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            JP_Code = JP_Code & "100001"
        Else
            JP_Code = Val(DBT.Rows(0).Item(0)) + 1
        End If
        Manform.Info_Tx0.Text = JP_Code



        Manform.ShowDialog()
        If Manform.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If

    End Sub

    Private Sub Man_Grid_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Man_Grid.CellContentClick

    End Sub

    Private Sub Man_Grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Man_Grid.CellValueChanged
        If Grid_Display_Ch = True Then
            sender.CurrentRow.HeaderCell.Value = "*"
        End If

    End Sub

    Private Sub Delete_Com_Click(sender As Object, e As EventArgs) Handles Delete_Com.Click
        Grid_Display_Ch = False
        If Len(Man_Grid.Item(0, Man_Grid.CurrentCell.RowIndex).Value) < 1 Then
            Exit Sub
        End If

        If MsgBox("사원코드  '" & Man_Grid.Item(0, Man_Grid.CurrentCell.RowIndex).Value & "'를 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "사원코드 삭제") = vbNo Then
            Exit Sub
        End If


        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_MAN Where Man_Code = '" & Man_Grid.Item(0, Man_Grid.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        Grid_Display()
        Grid_Display_Ch = True

    End Sub

    Private Sub Man_Grid_MouseClick(sender As Object, e As MouseEventArgs) Handles Man_Grid.MouseClick
        Man_Grid_Row = Man_Grid.CurrentCell.RowIndex
        Man_Grid_Col = Man_Grid.CurrentCell.ColumnIndex
        Man_Combo.Visible = False
        If Man_Grid_Row < 0 Then
            Exit Sub
        End If
        If Man_Grid_Col < 0 Then
            Exit Sub
        End If
        Dim DBT As New DataTable
        If Man_Grid_Col < 1 Then
        Else
            Select Case Man_Grid_Col
                Case 2
                    StrSQL = "Select Base_Name  FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '회사' Order By Base_Name"
                    Re_Count = DB_Select(DBT)
                    Man_Combo.Items.Clear()
                    If Re_Count <> 0 Then
                        For i = 0 To Re_Count - 1
                            Man_Combo.Items.Add(DBT.Rows(i).Item(0))
                        Next i
                    End If
                Case 3
                    StrSQL = "Select Base_Name  FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '부서' Order By Base_Name"
                    Re_Count = DB_Select(DBT)
                    Man_Combo.Items.Clear()
                    If Re_Count <> 0 Then
                        For i = 0 To Re_Count - 1
                            Man_Combo.Items.Add(DBT.Rows(i).Item(0))
                        Next i
                    End If
                Case 4
                    StrSQL = "Select Base_Name  FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '직급' Order By Base_Name"
                    Re_Count = DB_Select(DBT)
                    Man_Combo.Items.Clear()
                    If Re_Count <> 0 Then
                        For i = 0 To Re_Count - 1
                            Man_Combo.Items.Add(DBT.Rows(i).Item(0))
                        Next i
                    End If
                Case 5
                    Man_Combo.Items.Clear()
                    Man_Combo.Items.Add("남")
                    Man_Combo.Items.Add("여")
                Case 7
                    Man_Combo.Items.Clear()
                    Man_Combo.Items.Add("내국인")
                    Man_Combo.Items.Add("외국인")
                Case 12
                    StrSQL = "Select Le_Code  FROM SI_MAN_Level with(NOLOCK) Order By Le_Code"
                    Re_Count = DB_Select(DBT)
                    Man_Combo.Items.Clear()
                    If Re_Count <> 0 Then
                        For i = 0 To Re_Count - 1
                            Man_Combo.Items.Add(DBT.Rows(i).Item(0))
                        Next i
                    End If

                Case Else

                    Exit Sub
            End Select
            Man_Combo.Size = Man_Grid.CurrentCell.Size
            Man_Combo.Top = Man_Grid.GetCellDisplayRectangle(Man_Grid_Col, Man_Grid_Row, True).Top + Man_Grid.Top
            Man_Combo.Left = Man_Grid.GetCellDisplayRectangle(Man_Grid_Col, Man_Grid_Row, True).Left + Man_Grid.Left
            Man_Combo.Text = Man_Grid.CurrentCell.Value.ToString
            Man_Combo.BackColor = Man_Grid.RowsDefaultCellStyle.BackColor
            Man_Combo.Visible = True
            Man_Combo.Focus.ToString()

        End If

    End Sub

    Private Sub Man_Grid_Scroll(sender As Object, e As ScrollEventArgs) Handles Man_Grid.Scroll
        Man_Combo.Visible = False

    End Sub

    Private Sub Man_Combo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Man_Combo.SelectedIndexChanged

    End Sub

    Private Sub Man_Combo_LostFocus(sender As Object, e As EventArgs) Handles Man_Combo.LostFocus
        Man_Combo.Visible = False

    End Sub

    Private Sub Man_Combo_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Man_Combo.SelectionChangeCommitted
        Man_Grid.Item(Man_Grid_Col, Man_Grid_Row).Value = Man_Combo.SelectedItem.ToString()

    End Sub



    Private Sub Man_Grid_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Man_Grid.CellClick
        Man_Grid_Row = Man_Grid.CurrentCell.RowIndex
        Man_Grid_Col = Man_Grid.CurrentCell.ColumnIndex

        Manform.Info_Tx0.Text = Man_Grid.Item(0, Man_Grid_Row).Value
        Manform.Info_Tx1.Text = Man_Grid.Item(1, Man_Grid_Row).Value

        Manform.ComboBox3.Text = Man_Grid.Item(2, Man_Grid_Row).Value
        Manform.ComboBox2.Text = Man_Grid.Item(3, Man_Grid_Row).Value
        Manform.ComboBox5.Text = Man_Grid.Item(4, Man_Grid_Row).Value

        Manform.ComboBox6.Text = Man_Grid.Item(5, Man_Grid_Row).Value
        Manform.Info_Tx12.Text = Man_Grid.Item(6, Man_Grid_Row).Value

        Manform.Info_Tx16.Text = Man_Grid.Item(7, Man_Grid_Row).Value
        Manform.ComboBox4.Text = Man_Grid.Item(8, Man_Grid_Row).Value

        Manform.Info_Tx21.Text = Man_Grid.Item(9, Man_Grid_Row).Value
        Manform.Info_Tx22.Text = Man_Grid.Item(10, Man_Grid_Row).Value

        Manform.ShowDialog()
        If Manform.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If
    End Sub
End Class