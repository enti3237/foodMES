﻿Public Class Frm_Device2
    Dim Grid_Display As Boolean
    Dim Grid_Display_Ch As Boolean

    Dim Grid_Bom_Row As Integer
    Dim Grid_Bom_Col As Integer

    Private Sub Frm_Device2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()
        Process_Line_Grid_Setup()
        Search_Combo.Items.Add("설비코드")
        Search_Combo.Items.Add("설비명")
        Search_Combo.Text = "설비코드"
    End Sub
    Public Function Process_Line_Grid_Setup() As Boolean

        Grid_Color(Process_Line_Grid)


        Process_Line_Grid.AllowUserToAddRows = False
        Process_Line_Grid.RowTemplate.Height = 20


        '열의 갯수는 하나 더 많아야 함.
        Process_Line_Grid.ColumnCount = 8
        Process_Line_Grid.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Process_Line_Grid.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Process_Line_Grid.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Process_Line_Grid.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Process_Line_Grid.RowHeadersWidth = 40
        Process_Line_Grid.Columns(0).Width = 60
        Process_Line_Grid.Columns(1).Width = 200
        Process_Line_Grid.Columns(2).Width = 100
        Process_Line_Grid.Columns(3).Width = 100
        Process_Line_Grid.Columns(4).Width = 100
        Process_Line_Grid.Columns(5).Width = 200
        Process_Line_Grid.Columns(0).HeaderText = "설비코드"
        Process_Line_Grid.Columns(1).HeaderText = "설비명"
        Process_Line_Grid.Columns(2).HeaderText = "설비그룹"

        Process_Line_Grid.Columns(3).HeaderText = "설비ID"
        ' Process_Line_Grid.Columns(3).Visible = False
        Process_Line_Grid.Columns(4).HeaderText = "PORT 번호"
        ' Process_Line_Grid.Columns(4).Visible = False
        Process_Line_Grid.Columns(5).HeaderText = "설비구분"
        ' Process_Line_Grid.Columns(5).Visible = False
        Process_Line_Grid.Columns(6).HeaderText = "현재사용횟수"
        ' Process_Line_Grid.Columns(6).Visible = False

        Process_Line_Grid.Columns(7).HeaderText = "비고"
        ' Process_Line_Grid.Columns(7).Visible = False

        Process_Line_Grid.Columns(0).ReadOnly = True
        Process_Line_Grid.Columns(1).ReadOnly = True
        Process_Line_Grid_Setup = True
    End Function

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        Process_Line_Grid_Display()

    End Sub
    Public Function Process_Line_Grid_Display() As Boolean

        Dim i As Integer
        Dim j As Integer
        Dim DBT As New DataTable
        Grid_Display = False
        Process_Line_Grid.RowCount = 1
        Select Case Search_Combo.Text
            Case "설비코드"
                StrSQL = "Select * FROM SI_MACHINE with(NOLOCK) Where PL_Code Like '%" & Search_Text.Text & "%' Or PL_Code Is Null Order By PL_Code"
            Case "설비명"
                StrSQL = "Select * FROM SI_MACHINE with(NOLOCK) Where PL_Name Like '%" & Search_Text.Text & "%' Or PL_Name Is Null Order By PL_Name"
        End Select
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Process_Line_Grid.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To DBT.Columns.Count - 1
                    Process_Line_Grid.Item(j, i).Value = DBT.Rows(i).Item(j)
                Next j
            Next i
        Else
            Process_Line_Grid.RowCount = 1
            For j = 0 To Process_Line_Grid.ColumnCount - 1
                Process_Line_Grid.Item(j, 0).Value = ""
            Next j

        End If
        Grid_Display = True
        Process_Line_Grid_Display = True
    End Function
    Private Sub Process_Line_Grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Process_Line_Grid.CellValueChanged
        If Grid_Display = True Then
            Process_Line_Grid.CurrentRow.HeaderCell.Value = "*"
        End If

    End Sub

    Private Sub Save_Com_Click(sender As Object, e As EventArgs) Handles Save_Com.Click
        Dim i As Integer
        Grid_Display = False
        For i = 0 To Process_Line_Grid.Rows.Count - 1
            If Process_Line_Grid.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SI_MACHINE Set
                                     PL_Group = '" & Process_Line_Grid.Item(2, i).Value & "'
                                      , PL_ID = '" & Process_Line_Grid.Item(3, i).Value & "'
                                      , PL_Port = '" & Process_Line_Grid.Item(4, i).Value & "'
                                      , PL_Sel = '" & Process_Line_Grid.Item(5, i).Value & "'
                                      , PL_Count = '" & Process_Line_Grid.Item(6, i).Value & "'
                                      , PL_Bigo = '" & Process_Line_Grid.Item(7, i).Value & "'
                                   where PL_Code = '" & Process_Line_Grid.Item(0, i).Value & "'"
                Re_Count = DB_Execute()
                Process_Line_Grid.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display = True
    End Sub

    Private Sub Insert_Com_Click(sender As Object, e As EventArgs) Handles Insert_Com.Click
        Frm_Insert.Text = "생산라인 코드 추가"
        Frm_Insert.Show()
        Frm_Main.Enabled = False
    End Sub

    Private Sub Delete_Com_Click(sender As Object, e As EventArgs) Handles Delete_Com.Click
        If Len(Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value) < 1 Then
            MsgBox("라인코드를 선택하세요.")
            Exit Sub
        End If
        If MsgBox("라인코드  '" & Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value & "'를 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "라인코드 삭제") = vbNo Then
            Exit Sub
        End If
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_MACHINE where PL_Code = '" & Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()
        Process_Line_Grid_Display()
        MsgBox("삭제 되었습니다.")
    End Sub
End Class
