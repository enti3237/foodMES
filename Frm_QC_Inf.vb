﻿Public Class Frm_QC_Inf
    Dim Grid_Display_Ch As Boolean

    Private Sub Frm_QC_Inf_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()

        Re_Str = Grid_Main_Setup(Grid_Main1)
        Re_Str = Grid_Main_Setup(Grid_Main2)
        Re_Str = Grid_Main_Setup(Grid_Main3)
        Re_Str = Grid_Main_Setup(Grid_Main4)
        Grid_Main_Display(Grid_Main1, "종합상황판 품질")
        Grid_Main1.Visible = True
        Grid_Main2.Visible = False
        Grid_Main3.Visible = False
        Grid_Main4.Visible = False


        Button1.Text = "종합상황판 품질"
        Button2.Text = "종합상황판 공지"
        Button3.Text = "현장모니터 공지"
        'Button4.Text = DBT.Rows(3)("QC_Name")


    End Sub
    Public Function Grid_Main_Setup(Grid_Main As Object) As Boolean
        Grid_Color(Grid_Main)

        Grid_Main.AllowUserToAddRows = False
        Grid_Main.RowTemplate.Height = 24
        Grid_Main.ColumnCount = 2
        Grid_Main.RowCount = 0

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Main.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.RowHeadersWidth = 10
        Grid_Main.Columns(0).Width = 60



        Grid_Main.Columns(0).HeaderText = "순번"
        Grid_Main.Columns(1).HeaderText = "내용"


        Grid_Main.Columns(0).Width = 60
        Grid_Main.Columns(1).Width = 300

        Grid_Main_Setup = True
    End Function
    Public Function Grid_Main_Display(Grid_Main As Object, Q_Gu As String) As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display_Ch = False
        Grid_Main.RowCount = 0
        StrSQL = "Select QC_CONTENT, QC_MAN  From QC_INFO with(NOLOCK)  Where QC_TITLE = '" & Q_Gu & "' Order By Convert(Decimal,QC_CONTENT)"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Main.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To Grid_Main.ColumnCount - 1
                    Grid_Main.Item(j, i).Value = DBT.Rows(i).Item(j).ToString
                Next j
            Next i
        Else
            Grid_Main.RowCount = 1
            For j = 0 To Grid_Main.ColumnCount - 1
                Grid_Main.Item(j, 0).Value = ""
            Next j
        End If
        Grid_Display_Ch = True
        Grid_Main_Display = True
        Grid_Main4.ClearSelection()

    End Function

    Private Sub Save_Main_Click(sender As Object, e As EventArgs) Handles Save_Main.Click
        Dim C_Sel As String
        Dim Grid_Main As Object

        C_Sel = ""

        If Grid_Main1.Visible = True Then
            C_Sel = "종합상황판 품질"
            Grid_Main = Grid_Main1
        End If
        If Grid_Main2.Visible = True Then
            C_Sel = "종합상황판 공지"
            Grid_Main = Grid_Main2
        End If
        If Grid_Main3.Visible = True Then
            C_Sel = "현장모니터 공지"
            Grid_Main = Grid_Main3
        End If
        If Grid_Main4.Visible = True Then
            C_Sel = "4"
            Grid_Main = Grid_Main4
        End If

        Dim i As Integer

        Grid_Display_Ch = False
        For i = 0 To Grid_Main.Rows.Count - 1
            If Grid_Main.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "update QC_INFO Set QC_MAN = '" & Grid_Main.Item(1, i).Value & "' where QC_TITLE = '" & C_Sel & "' And QC_CONTENT = '" & Grid_Main.Item(0, i).Value & "'"
                Re_Count = DB_Execute()
                Grid_Main.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display_Ch = True

        MsgBox("저장되었습니다")
    End Sub

    Private Sub Grid_Main4_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Main1.CellValueChanged, Grid_Main2.CellValueChanged, Grid_Main3.CellValueChanged, Grid_Main4.CellValueChanged
        If Grid_Display_Ch = True Then
            sender.CurrentRow.HeaderCell.Value = "*"
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Label_Color(sender, Color_Label, Di_Panel2)
        Grid_Main_Display(Grid_Main1, "종합상황판 품질")
        Grid_Main1.Visible = True
        Grid_Main2.Visible = False
        Grid_Main3.Visible = False
        Grid_Main4.Visible = False

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Label_Color(sender, Color_Label, Di_Panel2)
        Grid_Main_Display(Grid_Main2, "종합상황판 공지")
        Grid_Main1.Visible = False
        Grid_Main2.Visible = True
        Grid_Main3.Visible = False
        Grid_Main4.Visible = False
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Label_Color(sender, Color_Label, Di_Panel2)
        Grid_Main_Display(Grid_Main3, "현장모니터 공지")
        Grid_Main1.Visible = False
        Grid_Main2.Visible = False
        Grid_Main3.Visible = True
        Grid_Main4.Visible = False

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Label_Color(sender, Color_Label, Di_Panel2)
        Grid_Main_Display(Grid_Main4, "4")
        Grid_Main1.Visible = False
        Grid_Main2.Visible = False
        Grid_Main3.Visible = False
        Grid_Main4.Visible = True
    End Sub

    Private Sub Insert_Main_Click(sender As Object, e As EventArgs) Handles Insert_Main.Click
        Dim DBT As New DataTable
        Dim QC_CONTENT As Long
        Dim QC_TITLE As String
        QC_TITLE = ""

        If Grid_Main2.Visible = True Then
            QC_TITLE = "종합상황판 공지"
        ElseIf Grid_Main1.Visible = True Then
            QC_TITLE = "종합상황판 품질"
        End If

        StrSQL = "Select CONVERT(INT, QC_CONTENT) AS QC_CONTENT From QC_INFO with(NOLOCK) where QC_TITLE = '" & QC_TITLE & "' order by QC_CONTENT desc"
        Re_Count = DB_Select(DBT)

        If Re_Count = 0 Then
            QC_CONTENT = 1
        Else
            QC_CONTENT = DBT.Rows(0).Item(0) + 1
        End If

        Grid_Display_Ch = False
        '추가한다
        If Grid_Main2.Visible = True Then
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "Insert into QC_INFO (QC_CONTENT, QC_TITLE)  Values ('" & QC_CONTENT & "', '종합상황판 공지' )"
            Re_Count = DB_Execute()
        ElseIf Grid_Main1.Visible = True Then
            StrSQL = "Set TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "Insert into QC_INFO (QC_CONTENT, QC_TITLE)  Values ('" & QC_CONTENT & "', '종합상황판 품질' )"
            Re_Count = DB_Execute()
        End If

        Grid_Main_Display(Grid_Main1, QC_TITLE)
        Grid_Main_Display(Grid_Main2, QC_TITLE)
        Grid_Display_Ch = True

    End Sub

    Private Sub Delete_Main_Click(sender As Object, e As EventArgs) Handles Delete_Main.Click
        '해당 칼럼 삭제
        'If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        'Else
        '    Exit Sub
        'End If
        'If Grid_Info.Item(1, 10).Value = "처리" Then
        '    MsgBox("수량 처리가 되어서 수정 할 수 없습니다.")
        '    Exit Sub
        'End If

        Dim QC_TITLE As String
        Dim Grid_Main As Object
        QC_TITLE = ""
        Grid_Main = ""


        If Grid_Main2.Visible = True Then
            QC_TITLE = "종합상황판 공지"
            Grid_Main = Grid_Main2
        ElseIf Grid_Main1.Visible = True Then
            QC_TITLE = "종합상황판 품질"
            Grid_Main = Grid_Main1
        End If

        If MsgBox("순번 '" & Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value & "'을(를) 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "삭제") = vbNo Then
            Exit Sub
        End If



        Grid_Display_Ch = False

        Dim i As Integer
        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Delete QC_INFO Where QC_TITLE = '" & QC_TITLE & "' And  QC_CONTENT = '" & Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()


        Grid_Main_Display(Grid_Main, QC_TITLE)
        Grid_Display_Ch = True
    End Sub

End Class