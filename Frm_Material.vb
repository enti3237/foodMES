Public Class Frm_Material
    Dim Grid_Display_Ch As Boolean
    Dim Grid_Main_Row As Integer
    Dim Grid_Main_Col As Integer
    Private Sub Frm_Material_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()

        Grid_Code_Setup()
        Grid_Info_Setup()
        Grid_Main_Setup()

        Search_Combo.Items.Add("코드")
        Search_Combo.Items.Add("날짜")
        Search_Combo.Items.Add("담당")
        Search_Combo.Text = "코드"

        Main_Panel.Top = 138
        Main_Panel.Left = 386
    End Sub
    Public Function Grid_Code_Setup() As Boolean
        Grid_Color(Grid_Code)

        Grid_Code.AllowUserToAddRows = False
        Grid_Code.RowTemplate.Height = 20
        Grid_Code.ColumnCount = 3
        Grid_Code.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Code.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Code.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Code.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Code.RowHeadersWidth = 10
        Grid_Code.Columns(0).Width = 100
        Grid_Code.Columns(1).Width = 110
        Grid_Code.Columns(2).Width = 110

        Grid_Code.Columns(0).HeaderText = "코드"
        Grid_Code.Columns(1).HeaderText = "날짜"
        Grid_Code.Columns(2).HeaderText = "담당"

        Dim i As Integer
        For i = 0 To Grid_Code.ColumnCount - 1
            Grid_Code.Columns(i).ReadOnly = True
        Next i
        Grid_Code_Setup = True
    End Function

    Public Function Grid_Info_Setup() As Boolean

        Grid_Color(Grid_Info)

        Grid_Info.AllowUserToAddRows = False
        Grid_Info.RowTemplate.Height = 24
        Grid_Info.ColumnCount = 2
        Grid_Info.RowCount = 6

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Info.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Info.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Info.RowHeadersWidth = 10
        Grid_Info.Columns(0).Width = 100
        Grid_Info.Columns(1).Width = 240
        Grid_Info.Columns(0).HeaderText = "구분"
        Grid_Info.Columns(1).HeaderText = "내용"

        Grid_Info.Item(0, 0).Value = "코드"
        Grid_Info.Item(0, 1).Value = "날짜"
        Grid_Info.Item(0, 2).Value = "시간"
        Grid_Info.Item(0, 3).Value = "담당자"
        Grid_Info.Item(0, 4).Value = "수량"
        Grid_Info.Item(0, 5).Value = "비고"


        'Grid_Info.Columns(0).ReadOnly = True
        'Grid_Info.Columns(1).ReadOnly = True
        Grid_Info.Rows(0).ReadOnly = True
        Grid_Info.Rows(1).ReadOnly = True
        Grid_Info.Rows(2).ReadOnly = True
        Grid_Info.Rows(3).ReadOnly = True
        Grid_Info.Rows(4).ReadOnly = True


        Grid_Info_Setup = True
    End Function
    Public Function Grid_Main_Setup() As Boolean

        Grid_Color(Grid_Main)

        Grid_Main.AllowUserToAddRows = False
        Grid_Main.RowTemplate.Height = 24
        Grid_Main.ColumnCount = 7
        Grid_Main.RowCount = 0

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Main.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.RowHeadersWidth = 10
        Grid_Main.Columns(0).Width = 60



        Grid_Main.Columns(0).HeaderText = "순번"
        Grid_Main.Columns(1).HeaderText = "제품코드"
        Grid_Main.Columns(2).HeaderText = "제품명"
        Grid_Main.Columns(3).HeaderText = "공정코드"
        Grid_Main.Columns(4).HeaderText = "공정명"
        Grid_Main.Columns(5).HeaderText = "수량"
        Grid_Main.Columns(6).HeaderText = "비고"

        Dim i As Integer

        Grid_Main.Columns(0).Width = 60
        Grid_Main.Columns(1).Width = 150
        Grid_Main.Columns(2).Width = 150
        Grid_Main.Columns(3).Width = 100
        Grid_Main.Columns(4).Width = 150

        For i = 5 To Grid_Main.ColumnCount - 1
            Grid_Main.Columns(i).Width = 80
            'Grid_Main.Columns(i).ReadOnly = True
        Next i


        Grid_Main_Setup = True
    End Function

    Private Sub Insert_Com_Click(sender As Object, e As EventArgs) Handles Insert_Com.Click
        '새로운 코드 추가
        Dim DBT As New DataTable
        Dim JP_Code As Long
        Dim My_Date As String
        Dim My_Time As String

        StrSQL = "Select GETDATE() "
        Re_Count = DB_Select(DBT)

        If Re_Count = 0 Then
            Exit Sub
        Else
            My_Date = DBT.Rows(0).Item(0)
            JP_Code = Mid(My_Date, 1, 4) & Mid(My_Date, 6, 2) & Mid(My_Date, 9, 2)
            My_Time = DateTime.Now.ToString("HH:mm:ss")
            My_Date = Mid(My_Date, 1, 4) & "-" & Mid(My_Date, 6, 2) & "-" & Mid(My_Date, 9, 2)
        End If

        StrSQL = "Select MT_Code From MT_Material with(NOLOCK) Where MT_Date Like  '" & My_Date & "%' Order By MT_Code Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            JP_Code = JP_Code & "001"
        Else
            JP_Code = JP_Code & Format(Val(Mid(DBT.Rows(0).Item(0), 11, 3)) + 1, "00#")
        End If

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert into MT_Material (MT_Code, MT_Date, MT_Time, MT_Name)  Values ('MT" & JP_Code & "', '" & My_Date & "','" & My_Time & "','" & Frm_Main.Text_Man_Name.Text & "')"
        Re_Count = DB_Execute()
        Grid_Code_Display()
        Grid_Info_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)

    End Sub
    Public Function Grid_Code_Display() As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Code.RowCount = 1
        Select Case Search_Combo.Text
            Case "코드"
                StrSQL = "Select MT_Code, MT_Date, MT_Name From MT_Material with(NOLOCK) Where MT_Code Like '" & Search_Text.Text & "%'  Order  By MT_Code DESC"
            Case "날짜"
                StrSQL = "Select MT_Code, MT_Date, MT_Name From MT_Material with(NOLOCK) Where MT_Date Like '" & Search_Text.Text & "%'  Or MT_Date Is Null Order  By MT_Date Desc"
            Case "담당"
                StrSQL = "Select MT_Code, MT_Date, MT_Name From MT_Material with(NOLOCK) WhereMT_Name Like '" & Search_Text.Text & "%'   Order  By MT_Name Desc"
        End Select
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Code.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To Grid_Code.ColumnCount - 1
                    Grid_Code.Item(j, i).Value = DBT.Rows(i).Item(j)
                Next j
            Next i
        Else
            Grid_Code.RowCount = 1
            For j = 0 To Grid_Code.ColumnCount - 1
                Grid_Code.Item(j, 0).Value = ""
            Next j
        End If
        Grid_Code_Display = True
    End Function
    Public Function Grid_Info_Display(Code As String) As Boolean
        Grid_Display_Ch = False

        Grid_Info_Display = True
        Dim DBT As New DataTable
        Dim i As Integer
        StrSQL = "Select * From MT_Material With(NOLOCK) Where MT_Code = '" & Code & "'"
        Re_Count = DB_Select(DBT)
        If Re_Count > 0 Then
            For i = 0 To DBT.Columns.Count - 1
                Grid_Info.Item(1, i).Value = DBT.Rows(0).Item(i).ToString
            Next i
        Else
            For i = 0 To DBT.Columns.Count - 1
                Grid_Info.Item(1, i).Value = ""
            Next i
        End If
        Grid_Display_Ch = True

    End Function

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        Grid_Code_Display()
    End Sub
    Private Sub Grid_Code_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Code.CellEnter
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
            Grid_Info_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
            Grid_Main_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
        End If

    End Sub
    Public Function Grid_Main_Display(PL_Code As String) As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display_Ch = False
        Grid_Main.RowCount = 0
        StrSQL = "Select MT_Sun, MT_PL_Code, MT_PL_Name, MT_PC_Code, MT_PC_Name, MT_Total, MT_Bigo  From MT_Material_List with(NOLOCK)  Where MT_Code = '" & PL_Code & "' Order By Convert(Decimal,MT_Sun)"
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
    End Function
    Private Sub Grid_Info_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Info.CellValueChanged
        If Grid_Display_Ch = True Then
            Grid_Info.CurrentRow.HeaderCell.Value = "*"
        End If
    End Sub

    Private Sub Save_Com_Click(sender As Object, e As EventArgs) Handles Save_Com.Click
        'Grid_Info 를 저장 한다.
        'DataBase에서 필드 이름 가지고 오기


        Dim Table_Name As String
        Dim j As Long
        Dim DBT As DataTable = Nothing
        Dim Field_Name(500) As String
        Dim i As Integer
        j = 0
        For i = 1 To Grid_Info.RowCount - 1
            If Grid_Info.Rows(i).HeaderCell.Value = "*" Then
                j = 1
            End If
        Next i
        If j = 0 Then
        Else
            Table_Name = "MT_Material"
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
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "UPDATE " & Table_Name & " SET "
            Field_Name(500) = "Ok"
            For i = 1 To j
                If Grid_Info.Rows(i).HeaderCell.Value = "*" Then
                    If Field_Name(500) = "" Then
                        StrSQL = StrSQL & ","
                    End If
                    StrSQL = StrSQL & " " & Table_Name & "." & Field_Name(i) & " = '" & Replace(Grid_Info.Item(1, i).Value, "'", "''") & "'"
                    If Field_Name(500) = "Ok" Then
                        Field_Name(500) = ""
                    End If
                End If
            Next i
            StrSQL = StrSQL & " WHERE " & Table_Name & "." & Field_Name(0) & " = '" & Grid_Info.Item(1, 0).Value & "'"
            Re_Count = DB_Execute()
        End If

        If Grid_Main.Item(0, 0).Value = "1" Then
        Else
            Exit Sub

        End If


        Dim Val_Check As String



        Val_Check = ""
        StrSQL = "Select MT_Check From MT_Material With(NOLOCK) Where MT_Code = '" & Grid_Info.Item(1, 0).Value & "'"
        Re_Count = DB_Select(DBT)
        Grid_Info_Combo.Items.Clear()
        If Re_Count <> 0 Then
            If DBT.Rows(0)("MT_Check").ToString = "처리" Then
                Val_Check = "처리"
                MsgBox("이미 처리 되었습니다. 삭제 후 다시 처리 하세요")
                Exit Sub
            End If
        End If


        '수량을 변경한다.
        For i = 0 To Grid_Main.RowCount - 1
            If Len(Grid_Main.Item(5, i).Value) > 0 Then
                '변경
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SI_PROCESS_LIST Set PP_Amount =  Convert(Decimal, PP_Amount)  + " & Grid_Main.Item(5, i).Value & " Where PP_Code = '" & Grid_Main.Item(1, i).Value & "' And PP_PC_Code = '" & Grid_Main.Item(3, i).Value & "'"
                Re_Count = DB_Execute()
            End If
        Next i

        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Update MT_Material Set MT_Check = '처리' Where MT_Code = '" & Grid_Info.Item(1, 0).Value & "'"
        Re_Count = DB_Execute()
        Grid_Info.Item(1, 4).Value = "처리"

    End Sub

    Private Sub Insert__Main_Click(sender As Object, e As EventArgs) Handles Insert__Main.Click
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        If Grid_Info.Item(1, 4).Value = "처리" Then
            MsgBox("수량 처리가 되어서 수정 할 수 없습니다.")
            Exit Sub
        End If

        Grid_Display_Ch = False

        Dim DBT As New DataTable
        Dim Db_Sun As Long
        StrSQL = "Select MT_Sun From MT_Material_List with(NOLOCK) Where MT_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' Order By Convert(Decimal,MT_Sun) Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            Db_Sun = 1
        Else
            Db_Sun = DBT.Rows(0).Item(0) + 1
        End If
        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert into MT_Material_List (MT_Code, MT_Sun)  Values ('" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "', '" & Db_Sun & "')"
        Re_Count = DB_Execute()
        Grid_Main_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True
    End Sub

    Private Sub Save_Main_Click(sender As Object, e As EventArgs) Handles Save_Main.Click
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        If Grid_Info.Item(1, 4).Value = "처리" Then
            MsgBox("수량 처리가 되어서 수정 할 수 없습니다.")
            Exit Sub
        End If
        Dim i As Integer

        Grid_Display_Ch = False
        For i = 0 To Grid_Main.Rows.Count - 1
            If Grid_Main.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "update MT_Material_List Set MT_PL_Code = '" & Grid_Main.Item(1, i).Value & "', MT_PL_Name = '" & Grid_Main.Item(2, i).Value & "',  MT_PC_Code = '" & Grid_Main.Item(3, i).Value & "', MT_PC_Name = '" & Grid_Main.Item(4, i).Value & "',   MT_Total = '" & Grid_Main.Item(5, i).Value & "', MT_Bigo = '" & Grid_Main.Item(6, i).Value & "' where MT_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And MT_Sun = '" & Grid_Main.Item(0, i).Value & "'"
                Re_Count = DB_Execute()
                Grid_Main.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display_Ch = True
    End Sub
    Private Sub Grid_Main_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Main.CellValueChanged
        If Grid_Display_Ch = True Then
            Grid_Main.CurrentRow.HeaderCell.Value = "*"
        End If
    End Sub
    Private Sub Grid_Main_MouseClick(sender As Object, e As MouseEventArgs) Handles Grid_Main.MouseClick
        Grid_Main_Row = Grid_Main.CurrentCell.RowIndex
        Grid_Main_Col = Grid_Main.CurrentCell.ColumnIndex
        Grid_Main_Combo.Visible = False
        If Grid_Main_Row < 0 Then
            Exit Sub
        End If
        If Grid_Main_Col < 0 Then
            Exit Sub
        End If
        Dim DBT As New DataTable

        Select Case Grid_Main_Col
            Case 1
                StrSQL = "Select PL_Code, PL_Name  FROM SI_PRODUCT with(NOLOCK) Order By PL_Code"
                Re_Count = DB_Select(DBT)
                Grid_Main_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        Grid_Main_Combo.Items.Add(DBT.Rows(i).Item(0))
                    Next i
                End If
                Grid_Main_Combo.Top = Grid_Main.Top + Grid_Main.ColumnHeadersHeight + (Grid_Main_Row) * 24
                Grid_Main_Combo.Left = Grid_Main.Left + Grid_Main.RowHeadersWidth + Grid_Main.Columns(0).Width
                Grid_Main_Combo.Width = Grid_Main.Columns(Grid_Main_Col).Width
                Grid_Main_Combo.Text = Grid_Main.CurrentCell.Value.ToString
                Grid_Main_Combo.BackColor = Grid_Main.RowsDefaultCellStyle.BackColor
                Grid_Main_Combo.Visible = True
                Grid_Main_Combo.Focus.ToString()
            Case 2
                StrSQL = "Select PL_Code, PL_Name  FROM SI_PRODUCT with(NOLOCK) Order By PL_Name"
                Re_Count = DB_Select(DBT)
                Grid_Main_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        Grid_Main_Combo.Items.Add(DBT.Rows(i).Item(1))
                    Next i
                End If
                Grid_Main_Combo.Top = Grid_Main.Top + Grid_Main.ColumnHeadersHeight + (Grid_Main_Row) * 24
                Grid_Main_Combo.Left = Grid_Main.Left + Grid_Main.RowHeadersWidth + Grid_Main.Columns(0).Width + Grid_Main.Columns(1).Width
                Grid_Main_Combo.Width = Grid_Main.Columns(Grid_Main_Col).Width
                Grid_Main_Combo.Text = Grid_Main.CurrentCell.Value.ToString
                Grid_Main_Combo.BackColor = Grid_Main.RowsDefaultCellStyle.BackColor
                Grid_Main_Combo.Visible = True
                Grid_Main_Combo.Focus.ToString()
            Case 3
                StrSQL = "Select PC_Code, PC_Name  FROM SI_PROCESS with(NOLOCK), SI_Process_List with(NOLOCK)  Where PP_Code = '" & Grid_Main.Item(1, Grid_Main_Row).Value & "' And PP_PC_Code = PC_Code Order By PC_Code"
                Re_Count = DB_Select(DBT)
                Grid_Main_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        Grid_Main_Combo.Items.Add(DBT.Rows(i).Item(0))
                    Next i
                End If
                Grid_Main_Combo.Top = Grid_Main.Top + Grid_Main.ColumnHeadersHeight + (Grid_Main_Row) * 24
                Grid_Main_Combo.Left = Grid_Main.Left + Grid_Main.RowHeadersWidth + Grid_Main.Columns(0).Width + Grid_Main.Columns(1).Width + Grid_Main.Columns(2).Width
                Grid_Main_Combo.Width = Grid_Main.Columns(Grid_Main_Col).Width
                Grid_Main_Combo.Text = Grid_Main.CurrentCell.Value.ToString
                Grid_Main_Combo.BackColor = Grid_Main.RowsDefaultCellStyle.BackColor
                Grid_Main_Combo.Visible = True
                Grid_Main_Combo.Focus.ToString()
            Case 4
                StrSQL = "Select PC_Code, PC_Name  FROM SI_PROCESS with(NOLOCK), SI_Process_List with(NOLOCK)  Where PP_Code = '" & Grid_Main.Item(1, Grid_Main_Row).Value & "' And PP_PC_Code = PC_Code  Order By PC_Name"
                Re_Count = DB_Select(DBT)
                Grid_Main_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        Grid_Main_Combo.Items.Add(DBT.Rows(i).Item(1))
                    Next i
                End If
                Grid_Main_Combo.Top = Grid_Main.Top + Grid_Main.ColumnHeadersHeight + (Grid_Main_Row) * 24
                Grid_Main_Combo.Left = Grid_Main.Left + Grid_Main.RowHeadersWidth + Grid_Main.Columns(0).Width + Grid_Main.Columns(1).Width + Grid_Main.Columns(2).Width + Grid_Main.Columns(3).Width
                Grid_Main_Combo.Width = Grid_Main.Columns(Grid_Main_Col).Width
                Grid_Main_Combo.Text = Grid_Main.CurrentCell.Value.ToString
                Grid_Main_Combo.BackColor = Grid_Main.RowsDefaultCellStyle.BackColor
                Grid_Main_Combo.Visible = True
                Grid_Main_Combo.Focus.ToString()

        End Select

    End Sub
    Private Sub Grid_Main_Combo_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Grid_Main_Combo.SelectionChangeCommitted
        Dim DBT As New DataTable
        Grid_Main.Item(Grid_Main_Col, Grid_Main_Row).Value = Grid_Main_Combo.SelectedItem.ToString()
        Select Case Grid_Main_Col
            Case 1
                StrSQL = "Select PL_Code, PL_Name  FROM SI_PRODUCT with(NOLOCK) Where PL_Code = '" & Replace(Grid_Main_Combo.SelectedItem.ToString(), "'", "''") & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Main.Item(2, Grid_Main_Row).Value = DBT.Rows(0).Item(1)
                End If
            Case 2
                StrSQL = "Select PL_Code, PL_Name  FROM SI_PRODUCT with(NOLOCK) Where PL_Name = '" & Replace(Grid_Main_Combo.SelectedItem.ToString(), "'", "''") & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Main.Item(1, Grid_Main_Row).Value = DBT.Rows(0).Item(0)
                End If
            Case 3
                StrSQL = "Select PC_Code, PC_Name  FROM SI_PROCESS with(NOLOCK) Where PC_Code = '" & Replace(Grid_Main_Combo.SelectedItem.ToString(), "'", "''") & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Main.Item(4, Grid_Main_Row).Value = DBT.Rows(0).Item(1)
                End If
            Case 4
                StrSQL = "Select PC_Code, PC_Name  FROM SI_PROCESS with(NOLOCK) Where PC_Name = '" & Replace(Grid_Main_Combo.SelectedItem.ToString(), "'", "''") & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Main.Item(3, Grid_Main_Row).Value = DBT.Rows(0).Item(0)
                End If

        End Select

    End Sub

    Private Sub Grid_Main_Combo_LostFocus(sender As Object, e As EventArgs) Handles Grid_Main_Combo.LostFocus
        Grid_Main_Combo.Visible = False

    End Sub

    Private Sub Delete__Main_Click(sender As Object, e As EventArgs) Handles Delete__Main.Click
        '해당 칼럼 삭제
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        If Grid_Info.Item(1, 4).Value = "처리" Then
            MsgBox("수량 처리가 되어서 수정 할 수 없습니다.")
            Exit Sub
        End If

        Grid_Display_Ch = False
        Dim i As Integer

        Dim DBT As New DataTable


        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Delete MT_Material_List Where MT_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And MT_Sun = '" & Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        For i = Grid_Main.CurrentCell.RowIndex + 1 To Grid_Main.RowCount - 1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "Update  MT_Material_List Set MT_Sun = '" & i & "' Where MT_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And  MT_Sun = '" & i + 1 & "'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        Grid_Main_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True
    End Sub

    Private Sub Grid_Main_Combo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Grid_Main_Combo.SelectedIndexChanged

    End Sub

    Private Sub Delete_Com_Click(sender As Object, e As EventArgs) Handles Delete_Com.Click

        '해당 칼럼 삭제
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        If MsgBox("자재수량  '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "'를 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "자재수량전표 삭제") = vbNo Then
            Exit Sub
        End If


        '수량 변경
        '수량이 변경 되었는지 확인 한다.
        Dim Val_Check As String

        Dim DBT As DataTable = Nothing

        Val_Check = ""
        StrSQL = "Select MT_Check From MT_Material With(NOLOCK) Where MT_Code = '" & Grid_Info.Item(1, 0).Value & "'"
        Re_Count = DB_Select(DBT)
        Grid_Info_Combo.Items.Clear()
        If Re_Count <> 0 Then
            If DBT.Rows(0)("MT_Check").ToString = "처리" Then
                Val_Check = "처리"
            End If
        End If


        If Val_Check = "처리" Then
            '기존 Data를 삭제 한다.
            'Grid의 수량을 가지고 온다.
            For i = 0 To Grid_Main.RowCount - 1
                '변경
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SI_PROCESS_LIST Set PP_Amount =  Convert(Decimal, PP_Amount)  - " & Grid_Main.Item(5, i).Value & " Where PP_Code = '" & Grid_Main.Item(1, i).Value & "' And PP_PC_Code = '" & Grid_Main.Item(3, i).Value & "'"
                Re_Count = DB_Execute()
            Next i
        End If


        Grid_Display_Ch = False
        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Delete MT_Material Where MT_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' "
        Re_Count = DB_Execute()

        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Delete MT_Material_List Where MT_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' "
        Re_Count = DB_Execute()

        Grid_Code_Display()
        Grid_Info_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True
    End Sub
End Class
