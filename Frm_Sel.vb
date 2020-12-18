Public Class Frm_Sel
    Dim Grid_Display_Ch As Boolean
    Dim Level_Grid_Row As Integer
    Dim Level_Grid_Col As Integer

    Private Sub Frm_Base_Code_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()
        Grid_Display_Ch = False

        D_Grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        J_Grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        S_Grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        Grid_Setup(D_Grid)
        Grid_Display(D_Grid, "대")

        Grid_Display_Ch = False
        Grid_Setup(J_Grid)
        Grid_Display(J_Grid, "중")


        Grid_Display_Ch = False
        Grid_Setup(S_Grid)
        Grid_Display(S_Grid, "소")


        '사원정보

        Grid_Display_Ch = False




    End Sub




    Public Function Grid_Setup(Grid As DataGridView) As Boolean
        Grid_Color(Grid)

        Grid.EnableHeadersVisualStyles = False

        Grid.AllowUserToAddRows = False
        Grid.RowTemplate.Height = 20
        Grid.ColumnCount = 3
        Grid.RowCount = 1


        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid.RowHeadersWidth = 45
        Grid.Columns(0).Width = 55
        Grid.Columns(1).Width = 115
        Grid.Columns(2).Width = 115

        Grid.Columns(0).HeaderText = "코드"
        Grid.Columns(1).HeaderText = "이름"
        Grid.Columns(2).HeaderText = "비고" '기초코드

        Grid.Columns(0).ReadOnly = True
        Grid_Setup = True
    End Function
    Public Function Grid_Display(Grid As DataGridView, Sel_Name As String) As Boolean
        Grid_Display_Ch = False

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid.RowCount = 1
        StrSQL = "Select * FROM SI_BASE_CODE With(NOLOCK) Where Base_Sel = '" & Sel_Name & "'  Order  By Base_Code"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                Grid.Item(0, i).Value = DBT.Rows(i)("Base_Code")
                Grid.Item(1, i).Value = DBT.Rows(i)("Base_Name")
                Grid.Item(2, i).Value = DBT.Rows(i)("Base_Bigo")
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


    Private Sub Product_Grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles D_Grid.CellValueChanged, J_Grid.CellValueChanged, S_Grid.CellValueChanged
        '  If Grid_Display_Ch = True Then
        sender.CurrentRow.HeaderCell.Value = "*"
        '  End If
    End Sub

    Private Sub Product_Save_Com_Click(sender As Object, e As EventArgs) Handles J_Save_Com.Click, S_Save_Com.Click, D_Save_Com.Click


        '선택된 갈럼값을 가지고 온다

        Grid_Display_Ch = True

        Dim Grid As DataGridView
        Dim Sel_Name As String
        Select Case sender.name.ToString
            Case "D_Save_Com"
                Grid = D_Grid
                Sel_Name = "대"
            Case "J_Save_Com"
                Grid = J_Grid
                Sel_Name = "중"
            Case "S_Save_Com"
                Grid = S_Grid
                Sel_Name = "소"

            Case Else
                Exit Sub
        End Select
        Dim i As Integer
        For i = 0 To Grid.Rows.Count - 1
            If Grid.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SI_BASE_CODE Set Base_Name = '" & Grid.Item(1, i).Value & "', Base_Bigo = '" & Grid.Item(2, i).Value & "' Where Base_Sel = '" & Sel_Name & "' And Base_Code = '" & Grid.Item(0, i).Value & "'"
                Re_Count = DB_Execute()
                Grid.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display_Ch = True

    End Sub

    Private Sub Product_Insert_Com_Click(sender As Object, e As EventArgs) Handles D_Insert_Com.Click, J_Insert_Com.Click, S_Insert_Com.Click
        Dim Grid As DataGridView
        Dim Sel_Name As String
        Select Case sender.name.ToString
            Case "D_Insert_Com"
                Grid = D_Grid
                Sel_Name = "대"
            Case "J_Insert_Com"
                Grid = J_Grid
                Sel_Name = "중"
                'Case "Product_Insert_Com"
                '    Grid = Product_Grid
                '    Sel_Name = "자산"
            Case "S_Insert_Com"
                Grid = S_Grid
                Sel_Name = "소"

            Case Else
                Exit Sub
        End Select

        Dim DBT As New DataTable
        Dim Db_Sun As String
        StrSQL = "Select Base_Code FROM SI_BASE_CODE with(NOLOCK) Where base_Sel = '" & Sel_Name & "' Order By Base_Code Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            Db_Sun = "001"
        Else
            Db_Sun = Format(Val(DBT.Rows(0).Item(0)) + 1, "00#")
        End If

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert into si_Base_Code (Base_Sel, Base_Code)  Values ('" & Sel_Name & "', '" & Db_Sun & "')"
        Re_Count = DB_Execute()
        Grid_Display(Grid, Sel_Name)

    End Sub

    Private Sub Product_Put_Com_Click(sender As Object, e As EventArgs) Handles Product_Put_Com.Click, J_Put_Com.Click, S_Put_Com.Click
        Grid_Display_Ch = False
        Dim Grid As DataGridView
        Dim Sel_Name As String
        Select Case sender.name.ToString
            Case "D_Put_Com"
                Grid = D_Grid
                Sel_Name = "대"
            Case "J_Put_Com"
                Grid = J_Grid
                Sel_Name = "중"
                'Case "product_Put_Com"
                '    Grid = Product_Grid
                '    Sel_Name = "자산"
            Case "S_Put_Com"
                Grid = S_Grid
                Sel_Name = "소"

                'Case "Proc_QC_Put_Com"
                '    Grid = Proc_QC_Grid
                '    Sel_Name = "설비"
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
            StrSQL = StrSQL & "UPDATE SI_BASE_CODE Set Base_Code = '" & i + 2 & "' Where Base_Code = '" & i + 1 & "' And Base_Sel ='" & Sel_Name & "'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert into SI_Base_Code (Base_Sel, Base_Code)  Values ('" & Sel_Name & "', '" & Grid.CurrentCell.RowIndex + 1 & "')"
        Re_Count = DB_Execute()
        Grid_Display(Grid, Sel_Name)
        Grid_Display_Ch = True


    End Sub

    Private Sub Product_Delete_Com_Click(sender As Object, e As EventArgs) Handles D_Delete_Com.Click, J_Delete_Com.Click, S_Delete_Com.Click
        '해당 칼럼 삭제



        Grid_Display_Ch = False
        Dim Grid As DataGridView
        Dim Sel_Name As String
        Select Case sender.name.ToString
            Case "D_Delete_Com"
                Grid = D_Grid
                Sel_Name = "대"
            Case "J_Delete_Com"
                Grid = J_Grid
                Sel_Name = "중"
                'Case "Product_Delete_Com"
                '    Grid = Product_Grid
                '    Sel_Name = "자산"
            Case "S_Delete_Com"
                Grid = S_Grid
                Sel_Name = "소"

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
        StrSQL = StrSQL & "DELETE SI_BASE_CODE Where Base_Sel = '" & Sel_Name & "' And   Base_Code = '" & Grid.Item(0, Grid.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        For i = Grid.CurrentCell.RowIndex + 1 To Grid.RowCount - 1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "UPDATE SI_BASE_CODE Set Base_Code = '" & i & "' Where Base_Sel = '" & Sel_Name & "' And  Base_Code = '" & i + 1 & "'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        Grid_Display(Grid, Sel_Name)
        Grid_Display_Ch = True

    End Sub


End Class
