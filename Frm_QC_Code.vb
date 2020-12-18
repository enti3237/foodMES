﻿Public Class Frm_QC_Code
    Dim Grid_Display_Ch As Boolean

    Dim Level_Grid_Row As Integer
    Dim Level_Grid_Col As Integer

    Private Sub Frm_QC_Code_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()
        '"수입검사", "공정검사", "출하검사", "Xbar-R 관리도", "부적합보고서"
        Grid_Setup(QC_Grid1)
        Grid_Display(QC_Grid1, "수입검사")

        Grid_Setup(QC_Grid2)
        Grid_Display(QC_Grid2, "공정검사")

        Grid_Setup(QC_Grid3)
        Grid_Display(QC_Grid3, "자주검사")

        Grid_Setup(QC_Grid4)
        Grid_Display(QC_Grid4, "출하검사")

        Grid_Setup(QC_Grid5)
        Grid_Display(QC_Grid5, "Xbar-R")


        Grid_Setup(QC_Grid6)
        Grid_Display(QC_Grid6, "부적합")

        Grid_Setup(QC_Grid7)
        Grid_Display(QC_Grid7, "일자")
        QC_Grid7.Columns(2).HeaderText = "일자"

        Panel3.Visible = False

    End Sub
    Public Function Grid_Setup(Grid As DataGridView) As Boolean
        Grid_Display_Ch = False
        Grid_Color(Grid)

        Grid.EnableHeadersVisualStyles = False

        Grid.AllowUserToAddRows = False
        Grid.RowTemplate.Height = 20
        Grid.ColumnCount = 4
        Grid.RowCount = 1


        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid.RowHeadersWidth = 45
        Grid.Columns(0).Width = 54
        Grid.Columns(1).Width = 115
        Grid.Columns(2).Width = 115
        Grid.Columns(3).Width = 105

        Grid.Columns(0).HeaderText = "순번"
        Grid.Columns(1).HeaderText = "구분"
        Grid.Columns(2).HeaderText = "File"
        Grid.Columns(3).HeaderText = "비고" '기초코드

        Grid.Columns(0).ReadOnly = True
        Grid_Display_Ch = True
        Grid_Setup = True

        Panel6.Visible = False

        Panel7.Visible = False

    End Function
    Public Function Grid_Display(Grid As DataGridView, Sel_Name As String) As Boolean
        Grid_Display_Ch = False
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid.RowCount = 1
        StrSQL = "Select * FROM SI_QC with(NOLOCK) Where QC_Sel = '" & Sel_Name & "'  Order  By Convert(Decimal,QC_Sun)"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                Grid.Item(0, i).Value = DBT.Rows(i)("QC_Sun")
                Grid.Item(1, i).Value = DBT.Rows(i)("QC_Name")
                Grid.Item(2, i).Value = DBT.Rows(i)("QC_File")
                Grid.Item(3, i).Value = DBT.Rows(i)("QC_Bigo")
            Next i
        Else
            Grid.RowCount = 1
            For j = 0 To Grid.ColumnCount - 1
                Grid.Item(j, 0).Value = ""
            Next j
        End If
        Grid_Display = True
        Grid_Display_Ch = True

    End Function
    Private Sub QC_Grid1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles QC_Grid1.CellValueChanged, QC_Grid2.CellValueChanged, QC_Grid3.CellValueChanged, QC_Grid4.CellValueChanged, QC_Grid5.CellValueChanged, QC_Grid6.CellValueChanged, QC_Grid7.CellValueChanged
        If Grid_Display_Ch = True Then
            sender.CurrentRow.HeaderCell.Value = "*"
        End If
    End Sub

    Private Sub QC_Save_Com1_Click(sender As Object, e As EventArgs) Handles QC_Save_Com1.Click, QC_Save_Com2.Click, QC_Save_Com3.Click, QC_Save_Com4.Click, QC_Save_Com5.Click, QC_Save_Com6.Click, QC_Save_Com7.Click
        Grid_Display_Ch = False
        Dim Grid As DataGridView
        Dim Sel_Name As String
        Select Case sender.name.ToString
            Case "QC_Save_Com1"
                Grid = QC_Grid1
                Sel_Name = "수입검사"
            Case "QC_Save_Com2"
                Grid = QC_Grid2
                Sel_Name = "공정검사"
            Case "QC_Save_Com3"
                Grid = QC_Grid3
                Sel_Name = "자주검사"
            Case "QC_Save_Com4"
                Grid = QC_Grid4
                Sel_Name = "출하검사"
            Case "QC_Save_Com5"
                Grid = QC_Grid5
                Sel_Name = "Xbar-R"
            Case "QC_Save_Com6"
                Grid = QC_Grid6
                Sel_Name = "부적합"
            Case "QC_Save_Com7"
                Grid = QC_Grid7
                Sel_Name = "일자"
            Case Else
                Exit Sub
        End Select
        Dim i As Integer

        For i = 0 To Grid.Rows.Count - 1
            If Grid.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SI_QC Set QC_Name = '" & Grid.Item(1, i).Value & "', QC_File = '" & Grid.Item(2, i).Value & "', QC_Bigo = '" & Grid.Item(3, i).Value & "' Where QC_Sel = '" & Sel_Name & "' And QC_Sun = '" & Grid.Item(0, i).Value & "'"
                Re_Count = DB_Execute()
                Grid.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display_Ch = True

    End Sub
    Private Sub QC_Insert_Com1_Click(sender As Object, e As EventArgs) Handles QC_Insert_Com1.Click, QC_Insert_Com2.Click, QC_Insert_Com3.Click, QC_Insert_Com4.Click, QC_Insert_Com5.Click, QC_Insert_Com6.Click
        Dim Grid As DataGridView
        Dim Sel_Name As String
        Select Case sender.name.ToString
            Case "QC_Insert_Com1"
                Grid = QC_Grid1
                Sel_Name = "수입검사"
            Case "QC_Insert_Com2"
                Grid = QC_Grid2
                Sel_Name = "공정검사"
            Case "QC_Insert_Com3"
                Grid = QC_Grid3
                Sel_Name = "자주검사"
            Case "QC_Insert_Com4"
                Grid = QC_Grid4
                Sel_Name = "출하검사"
            Case "QC_Insert_Com5"
                Grid = QC_Grid5
                Sel_Name = "Xbar-R"
            Case "QC_Insert_Com6"
                Grid = QC_Grid6
                Sel_Name = "부적합"
            Case Else
                Exit Sub
        End Select

        Dim DBT As New DataTable
        Dim Db_Sun As Long
        StrSQL = "Select QC_Sun FROM SI_QC with(NOLOCK) Where QC_Sel = '" & Sel_Name & "' Order By  Convert(Decimal,QC_Sun) Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            Db_Sun = 1
        Else
            Db_Sun = DBT.Rows(0).Item(0) + 1
        End If

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO SI_QC (QC_Sel, QC_Sun)  Values ('" & Sel_Name & "', '" & Db_Sun & "')"
        Re_Count = DB_Execute()
        Grid_Display(Grid, Sel_Name)

    End Sub
    Private Sub QC_Put_Com1_Click(sender As Object, e As EventArgs) Handles QC_Put_Com1.Click, QC_Put_Com2.Click, QC_Put_Com3.Click, QC_Put_Com4.Click, QC_Put_Com5.Click, QC_Put_Com6.Click
        Grid_Display_Ch = False
        Dim Grid As DataGridView
        Dim Sel_Name As String
        Select Case sender.name.ToString
            Case "QC_Put_Com1"
                Grid = QC_Grid1
                Sel_Name = "수입검사"
            Case "QC_Put_Com2"
                Grid = QC_Grid2
                Sel_Name = "공정검사"
            Case "QC_Put_Com3"
                Grid = QC_Grid3
                Sel_Name = "자주검사"
            Case "QC_Put_Com4"
                Grid = QC_Grid4
                Sel_Name = "출하검사"
            Case "QC_Put_Com5"
                Grid = QC_Grid5
                Sel_Name = "Xbar-R"
            Case "QC_Put_Com6"
                Grid = QC_Grid6
                Sel_Name = "부적합"
            Case Else
                Exit Sub
        End Select

        If Len(Grid.Item(0, Grid.CurrentCell.RowIndex).Value) < 1 Then
            Exit Sub
        End If
        Dim i As Integer
        For i = Grid.RowCount - 1 To Grid.CurrentCell.RowIndex Step -1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "UPDATE SI_QC Set QC_Sun = '" & i + 2 & "' Where QC_Sun = '" & i + 1 & "' And QC_Sel ='" & Sel_Name & "'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO SI_QC (QC_Sel, QC_Sun)  Values ('" & Sel_Name & "', '" & Grid.CurrentCell.RowIndex + 1 & "')"
        Re_Count = DB_Execute()
        Grid_Display(Grid, Sel_Name)
        Grid_Display_Ch = True

    End Sub
    Private Sub QC_Delete_Com1_Click(sender As Object, e As EventArgs) Handles QC_Delete_Com1.Click, QC_Delete_Com2.Click, QC_Delete_Com3.Click, QC_Delete_Com4.Click, QC_Delete_Com5.Click, QC_Delete_Com6.Click
        '해당 칼럼 삭제
        Grid_Display_Ch = False
        Dim Grid As DataGridView
        Dim Sel_Name As String
        Select Case sender.name.ToString
            Case "QC_Delete_Com1"
                Grid = QC_Grid1
                Sel_Name = "수입검사"
            Case "QC_Delete_Com2"
                Grid = QC_Grid2
                Sel_Name = "공정검사"
            Case "QC_Delete_Com3"
                Grid = QC_Grid3
                Sel_Name = "자주검사"
            Case "QC_Delete_Com4"
                Grid = QC_Grid4
                Sel_Name = "출하검사"
            Case "QC_Delete_Com5"
                Grid = QC_Grid5
                Sel_Name = "Xbar-R"
            Case "QC_Delete_Com6"
                Grid = QC_Grid6
                Sel_Name = "부적합"
            Case Else
                Exit Sub
        End Select
        If Len(Grid.Item(0, Grid.CurrentCell.RowIndex).Value) < 1 Then
            Exit Sub
        End If
        Dim i As Integer

        If MsgBox("'" & Grid.Item(1, Grid.CurrentCell.RowIndex).Value & "'을(를) 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "삭제") = vbNo Then
            Exit Sub
        End If

        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_QC Where QC_Sel = '" & Sel_Name & "' And   QC_Sun = '" & Grid.Item(0, Grid.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        For i = Grid.CurrentCell.RowIndex + 1 To Grid.RowCount - 1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "UPDATE SI_QC Set QC_Sun = '" & i & "' Where QC_Sel = '" & Sel_Name & "' And  QC_Sun = '" & i + 1 & "'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        Grid_Display(Grid, Sel_Name)
        Grid_Display_Ch = True

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

    End Sub
End Class
