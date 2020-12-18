﻿Imports Excel = Microsoft.Office.Interop.Excel

Public Class Frm_Haccp
    Dim Grid_Display As Boolean
    Dim Grid_Display_Ch As Boolean
    Dim Grid_Bom_Row As Integer
    Dim Grid_Bom_Col As Integer
    Dim Grid_Info_Col As Integer
    Dim Grid_Info_Row As Integer

    Dim ofd As New OpenFileDialog()

    Private Sub Frm_Device2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()
        Process_Line_Grid_Setup()
        Search_Combo.Items.Add("")
        '  Search_Combo.Items.Add("상태")
        Search_Combo.Text = ""

        ' Insert_Com.Visible = False

        Search_Combo.Visible = False
        Search_Text.Visible = False

        Search_Com.PerformClick()
        Process_Line_Grid.ClearSelection()
        Search_Com.Visible = False
    End Sub
    Public Function Process_Line_Grid_Setup() As Boolean

        Grid_Color(Process_Line_Grid)


        Process_Line_Grid.AllowUserToAddRows = False
        Process_Line_Grid.RowTemplate.Height = 20


        '열의 갯수는 하나 더 많아야 함.
        Process_Line_Grid.ColumnCount = 5
        Process_Line_Grid.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Process_Line_Grid.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Process_Line_Grid.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Process_Line_Grid.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Process_Line_Grid.RowHeadersWidth = 40
        Process_Line_Grid.Columns(0).Width = 100
        Process_Line_Grid.Columns(1).Width = 100
        Process_Line_Grid.Columns(2).Width = 100
        Process_Line_Grid.Columns(3).Width = 370
        Process_Line_Grid.Columns(4).Width = 700


        Process_Line_Grid.Columns(0).HeaderText = "순번"
        Process_Line_Grid.Columns(1).HeaderText = "재개정일자"
        Process_Line_Grid.Columns(1).Visible = False
        Process_Line_Grid.Columns(2).HeaderText = "기준"
        Process_Line_Grid.Columns(2).Visible = False
        Process_Line_Grid.Columns(3).HeaderText = "내용"
        Process_Line_Grid.Columns(3).Visible = False
        Process_Line_Grid.Columns(4).HeaderText = "파일위치"



        Process_Line_Grid_Setup = True
    End Function

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        Process_Line_Grid_Display()

    End Sub
    Public Function Process_Line_Grid_Display() As Boolean

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display = False
        Process_Line_Grid.RowCount = 1
        Select Case Search_Combo.Text
            Case ""
                StrSQL = "Select * FROM HP_STANDARD with(NOLOCK) Order By ST_CODE"

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
        Process_Line_Grid.ClearSelection()

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
                StrSQL = StrSQL & "UPDATE HP_STANDARD Set
                                     ST_DATE = '" & Process_Line_Grid.Item(1, i).Value & "',
                                     ST_STANDARD = '" & Process_Line_Grid.Item(2, i).Value & "',
                                     ST_CONTENT = '" & Process_Line_Grid.Item(3, i).Value & "',
                                     ST_BIGO = '" & Process_Line_Grid.Item(4, i).Value & "'
                                   where ST_CODE = '" & Process_Line_Grid.Item(0, i).Value & "'"
                Re_Count = DB_Execute()
                Process_Line_Grid.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display = True
        MsgBox("저장되었습니다")
    End Sub

    Private Sub Insert_Com_Click(sender As Object, e As EventArgs) Handles Insert_Com.Click
        'Frm_Insert.Text = "해썹 코드 추가"
        'Frm_Insert.Show()
        'Frm_Main.Enabled = False

        Grid_Display_Ch = False
        Dim DBT As New DataTable
        Dim Db_Sun As Long
        StrSQL = "Select ST_CODE FROM HP_STANDARD with(NOLOCK) Order By Convert(Decimal,ST_CODE) Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            Db_Sun = 1
        Else
            Db_Sun = DBT.Rows(0).Item(0) + 1
        End If

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO HP_STANDARD (ST_CODE)  Values ('" & Db_Sun & "')"
        Re_Count = DB_Execute()
        Process_Line_Grid_Display()
        Grid_Display_Ch = True
        Search_Com.PerformClick()

    End Sub

    Private Sub Delete_Com_Click(sender As Object, e As EventArgs) Handles Delete_Com.Click
        If Len(Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value) < 1 Then
            MsgBox("순번을 선택하세요.")
            Exit Sub
        End If
        If MsgBox("순번  '" & Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value & "'를 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "코드 삭제") = vbNo Then
            Exit Sub
        End If
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE HP_STANDARD where ST_CODE = '" & Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()
        Process_Line_Grid_Display()
        MsgBox("삭제 되었습니다.")
    End Sub

    Private Sub Panel_Menu_Paint(sender As Object, e As PaintEventArgs) Handles Panel_Menu.Paint

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Process_Line_Grid.Item(0, Process_Line_Grid.CurrentCell.RowIndex).Value Is Nothing Or Process_Line_Grid.Item(0, 0).Value Is Nothing Then
            MsgBox("추가 버튼을 눌러 순번을 먼저 추가해주세요")
            Exit Sub
        End If
        If ofd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            ' file_textbox.Text = ofd.FileName

            Process_Line_Grid.Item(4, Process_Line_Grid.CurrentCell.RowIndex).Value = ofd.FileName
            'MsgBox(Path.GetFileNameWithoutExtension(file_textbox.Text))
            'Me.Text = "메모장 " + ChkNewFile    '타이틀 셋팅
            'Dim openFile As StreamReader '파일읽기 변수
            'openFile = New StreamReader(OpenFileDialog.FileName, System.Text.Encoding.UTF8)
            'MainNote.Text = openFile.ReadToEnd ' 읽은 파일 RichTextBox에 모두 할당
            'openFile.Close()
        End If
    End Sub

    Private Sub Process_Line_Grid_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Process_Line_Grid.CellDoubleClick
        Grid_Info_Row = Process_Line_Grid.CurrentCell.RowIndex
        Grid_Info_Col = Process_Line_Grid.CurrentCell.ColumnIndex
        '  MsgBox(Grid_Info_Col)
        If Grid_Info_Col = 4 Then
            '  MsgBox("")

            Dim File_Name As String = Process_Line_Grid.Item(Grid_Info_Col, Grid_Info_Row).Value

            If System.IO.File.Exists(File_Name) = True Then
                Process.Start(File_Name)
            Else
                MsgBox("파일이 존재하지 않습니다")
            End If
        End If
    End Sub
End Class