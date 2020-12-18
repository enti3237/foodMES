﻿Public Class Frm_Base_Code
    Dim Grid_Display_Ch As Boolean
    Dim Level_Grid_Row As Integer
    Dim Level_Grid_Col As Integer

    Private Sub Frm_Base_Code_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()
        Grid_Display_Ch = False
        Product_Sel2_Grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        Unit_Grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        Man_Grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        Level_Grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        stockgridview.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        Grid_Setup(Product_Grid)
        Grid_Display(Product_Grid, "제품구분")

        Grid_Display_Ch = False
        Grid_Setup(Product_Sel2_Grid)
        Grid_Display(Product_Sel2_Grid, "제품종류")


        Grid_Display_Ch = False
        Grid_Setup(Unit_Grid)
        Grid_Display(Unit_Grid, "단위")

        Grid_Display_Ch = False
        Grid_Setup(Condition_Grid)
        Grid_Display(Condition_Grid, "지불조건")


        Grid_Display_Ch = False
        Grid_Setup(Err_Grid)
        Grid_Display(Err_Grid, "불량유형")

        Grid_Display_Ch = False
        Grid_Setup(Dr_Grid)
        Grid_Display(Dr_Grid, "장비점검")


        Grid_Display_Ch = False
        Grid_Setup(Port_Grid)
        Grid_Display(Port_Grid, "Port")

        '사원정보
        Man_Combo.Items.Clear()
        '    Man_Combo.Items.Add("회사")
        Man_Combo.Items.Add("부서")
        Man_Combo.Items.Add("직급")
        Man_Combo.Text = "부서"
        Grid_Display_Ch = False
        Man_Grid_Setup()
        Man_Grid_Display()

        Grid_Display_Ch = False
        Level_Grid_Setup()
        Level_Grid_Display()


        STOCK_Grid_Setup()
        STOCK_Grid_Display()

        Panel7.Visible = False
        Panel8.Visible = False
    End Sub
    Public Function Level_Grid_Setup() As Boolean
        Grid_Color(Level_Grid)
        Level_Grid.EnableHeadersVisualStyles = False

        Level_Grid.AllowUserToAddRows = False
        Level_Grid.RowTemplate.Height = 24

        Dim i As Integer
        Dim j As Long
        For i = 0 To 9

            If Len(Menu_String(i, 0)) < 1 Then
                j = i + 1
                Exit For
            End If
        Next i

        Level_Grid.ColumnCount = j

        Level_Grid.RowCount = 1


        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Level_Grid.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Level_Grid.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Level_Grid.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Level_Grid.RowHeadersWidth = 45
        Level_Grid.Columns(0).Width = 55

        Level_Grid.Columns(0).HeaderText = "코드"
        For i = 1 To j - 1
            Level_Grid.Columns(i).HeaderText = Menu_String(i - 1, 0)
        Next i

        For i = 0 To Level_Grid.ColumnCount - 1
            Level_Grid.Columns(i).ReadOnly = True
        Next i

        For i = 1 To Level_Grid.ColumnCount - 1
            Level_Grid.Columns(i).Width = 80
        Next

        Level_Grid_Setup = True

    End Function

    Public Function Level_Grid_Display() As Boolean
        Grid_Display_Ch = False
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Level_Grid.RowCount = 1
        StrSQL = "Select * FROM SI_MAN_Level with(NOLOCK) Order  By Convert(Decimal,Le_Code)"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Level_Grid.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To Level_Grid.ColumnCount - 1
                    Level_Grid.Item(j, i).Value = DBT.Rows(i)(j)
                Next j
            Next i
        Else
        End If
        Grid_Display_Ch = True
        Level_Grid_Display = True
    End Function

    Public Function STOCK_Grid_Setup() As Boolean
        Grid_Color(stockgridview)

        stockgridview.EnableHeadersVisualStyles = False

        stockgridview.AllowUserToAddRows = False
        stockgridview.RowTemplate.Height = 20
        stockgridview.ColumnCount = 4
        stockgridview.RowCount = 1


        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        stockgridview.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        stockgridview.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        stockgridview.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        stockgridview.RowHeadersWidth = 45
        stockgridview.Columns(0).Width = 55
        stockgridview.Columns(1).Width = 115
        stockgridview.Columns(2).Width = 60

        stockgridview.Columns(0).HeaderText = "코드"
        stockgridview.Columns(1).HeaderText = "이름"
        stockgridview.Columns(2).HeaderText = "기준온도" '
        stockgridview.Columns(3).HeaderText = "IP" '

        stockgridview.Columns(0).ReadOnly = True
        Grid_Display_Ch = True
    End Function

    Public Function STOCK_Grid_Display() As Boolean
        Grid_Display_Ch = False
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        stockgridview.RowCount = 1
        StrSQL = "Select * FROM SI_STOCK with(NOLOCK) Order By Convert(Decimal,FL_CODE)"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            stockgridview.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To stockgridview.ColumnCount - 1
                    stockgridview.Item(j, i).Value = DBT.Rows(i)(j)
                Next j
            Next i
        Else
            stockgridview.RowCount = 1
            For j = 0 To stockgridview.ColumnCount - 1
                stockgridview.Item(j, 0).Value = ""
            Next j
        End If
        Grid_Display_Ch = True
        STOCK_Grid_Display = True
    End Function

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

        Grid.Columns(0).HeaderText = "순번"
        Grid.Columns(1).HeaderText = "내용"
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
        StrSQL = "Select * FROM SI_BASE_CODE With(NOLOCK) Where Base_Sel = '" & Sel_Name & "'  Order  By Base_Sun"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                Grid.Item(0, i).Value = DBT.Rows(i)("Base_Sun")
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


    Private Sub Product_Grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Product_Grid.CellValueChanged, Product_Sel2_Grid.CellValueChanged, Unit_Grid.CellValueChanged, Condition_Grid.CellValueChanged, Err_Grid.CellValueChanged, Dr_Grid.CellValueChanged
        If Grid_Display_Ch = True Then
            sender.CurrentRow.HeaderCell.Value = "*"
        End If
    End Sub

    Private Sub Product_Save_Com_Click(sender As Object, e As EventArgs) Handles Product_Save_Com.Click, Product_Sel2_Save_Com.Click, Unit_Save_Com.Click, Condition_Save_Com.Click, Err_Save_Com.Click, Port_Save_Com.Click, Dr_Save_Com.Click



        '선택된 갈럼값을 가지고 온다
        Man_Grid_Display()
        Grid_Display_Ch = True

        Grid_Display_Ch = False
        Dim Grid As DataGridView
        Dim Sel_Name As String
        Select Case sender.name.ToString
            Case "Product_Save_Com"
                Grid = Product_Grid
                Sel_Name = "제품구분"
            Case "Product_Sel2_Save_Com"
                Grid = Product_Sel2_Grid
                Sel_Name = "제품종류"
            Case "Unit_Save_Com"
                Grid = Unit_Grid
                Sel_Name = "단위"
            Case "Condition_Save_Com"
                Grid = Condition_Grid
                Sel_Name = "지불조건"
            Case "Err_Save_Com"
                Grid = Err_Grid
                Sel_Name = "불량유형"
            Case "Port_Save_Com"
                Grid = Port_Grid
                Sel_Name = "Port"
            Case "Dr_Save_Com"
                Grid = Dr_Grid
                Sel_Name = "장비점검"
            Case Else
                Exit Sub
        End Select
        Dim i As Integer
        For i = 0 To Grid.Rows.Count - 1
            If Grid.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SI_BASE_CODE Set Base_Name = '" & Grid.Item(1, i).Value & "', Base_Bigo = '" & Grid.Item(2, i).Value & "' Where Base_Sel = '" & Sel_Name & "' And Base_Sun = '" & Grid.Item(0, i).Value & "'"
                Re_Count = DB_Execute()
                Grid.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display_Ch = True

    End Sub

    Private Sub Product_Insert_Com_Click(sender As Object, e As EventArgs) Handles Product_Insert_Com.Click, Product_Sel2_Insert_Com.Click, Unit_Insert_Com.Click, Condition_Insert_Com.Click, Err_Insert_Com.Click, Dr_Insert_Com.Click
        Dim Grid As DataGridView
        Dim Sel_Name As String
        Select Case sender.name.ToString
            Case "Product_Insert_Com"
                Grid = Product_Grid
                Sel_Name = "제품구분"
            Case "Product_Sel2_Insert_Com"
                Grid = Product_Sel2_Grid
                Sel_Name = "제품종류"
                'Case "Product_Insert_Com"
                '    Grid = Product_Grid
                '    Sel_Name = "자산"
            Case "Unit_Insert_Com"
                Grid = Unit_Grid
                Sel_Name = "단위"
            Case "Condition_Insert_Com"
                Grid = Condition_Grid
                Sel_Name = "지불조건"
                'Case "Proc_Insert_Com"
                '    Grid = Proc_Grid
                '    Sel_Name = "중지"
            Case "Err_Insert_Com"
                Grid = Err_Grid
                Sel_Name = "불량유형"
            Case "Dr_Insert_Com"
                Grid = Dr_Grid
                Sel_Name = "장비점검"
                'Case "Proc_QC_Insert_Com"
                '    Grid = Proc_QC_Grid
                '    Sel_Name = "설비"
            Case Else
                Exit Sub
        End Select

        Dim DBT As New DataTable
        Dim Db_Sun As Long
        StrSQL = "Select Base_Sun FROM SI_BASE_CODE with(NOLOCK) Where base_Sel = '" & Sel_Name & "' Order By Base_Sun Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            Db_Sun = 1
        Else
            Db_Sun = DBT.Rows(0).Item(0) + 1
        End If

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert into si_Base_Code (Base_Sel, Base_Sun)  Values ('" & Sel_Name & "', '" & Db_Sun & "')"
        Re_Count = DB_Execute()
        Grid_Display(Grid, Sel_Name)

    End Sub

    Private Sub Product_Put_Com_Click(sender As Object, e As EventArgs) Handles Product_Put_Com.Click, Product_Sel2_Put_Com.Click, Unit_Put_Com.Click, Condition_Put_Com.Click, Err_Put_Com.Click, Dr_Put_Com.Click
        Grid_Display_Ch = False
        Dim Grid As DataGridView
        Dim Sel_Name As String
        Select Case sender.name.ToString
            Case "Product_Put_Com"
                Grid = Product_Grid
                Sel_Name = "제품구분"
            Case "Product_Sel2_Put_Com"
                Grid = Product_Sel2_Grid
                Sel_Name = "제품종류"
                'Case "product_Put_Com"
                '    Grid = Product_Grid
                '    Sel_Name = "자산"
            Case "Unit_Put_Com"
                Grid = Unit_Grid
                Sel_Name = "단위"
            Case "Condition_Put_Com"
                Grid = Condition_Grid
                Sel_Name = "지불조건"
                'Case "Proc_Put_Com"
                '    Grid = Proc_Grid
                '    Sel_Name = "중지"
            Case "Err_Put_Com"
                Grid = Err_Grid
                Sel_Name = "불량유형"
            Case "Dr_Put_Com"
                Grid = Dr_Grid
                Sel_Name = "장비점검"
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
            StrSQL = StrSQL & "UPDATE SI_BASE_CODE Set Base_Sun = '" & i + 2 & "' Where Base_Sun = '" & i + 1 & "' And Base_Sel ='" & Sel_Name & "'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert into SI_Base_Code (Base_Sel, Base_Sun)  Values ('" & Sel_Name & "', '" & Grid.CurrentCell.RowIndex + 1 & "')"
        Re_Count = DB_Execute()
        Grid_Display(Grid, Sel_Name)
        Grid_Display_Ch = True


    End Sub

    Private Sub Product_Delete_Com_Click(sender As Object, e As EventArgs) Handles Product_Delete_Com.Click, Product_Sel2_Delete_Com.Click, Unit_Delete_Com.Click, Condition_Delete_Com.Click, Err_Delete_Com.Click, Dr_Delete_Com.Click
        '해당 칼럼 삭제



        Grid_Display_Ch = False
        Dim Grid As DataGridView
        Dim Sel_Name As String
        Select Case sender.name.ToString
            Case "Product_Delete_Com"
                Grid = Product_Grid
                Sel_Name = "제품구분"
            Case "Product_Sel2_Delete_Com"
                Grid = Product_Sel2_Grid
                Sel_Name = "제품종류"
                'Case "Product_Delete_Com"
                '    Grid = Product_Grid
                '    Sel_Name = "자산"
            Case "Unit_Delete_Com"
                Grid = Unit_Grid
                Sel_Name = "단위"
            Case "Condition_Delete_Com"
                Grid = Condition_Grid
                Sel_Name = "지불조건"
                'Case "Proc_Delete_Com"
                '    Grid = Proc_Grid
                '    Sel_Name = "중지"
            Case "Err_Delete_Com"
                Grid = Err_Grid
                Sel_Name = "불량유형"
            Case "Dr_Delete_Com"
                Grid = Dr_Grid
                Sel_Name = "장비점검"
                'Case "Proc_QC_Delete_Com"
                '    Grid = Proc_QC_Grid
                '    Sel_Name = "설비"
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
        StrSQL = StrSQL & "DELETE SI_BASE_CODE Where Base_Sel = '" & Sel_Name & "' And   Base_Sun = '" & Grid.Item(0, Grid.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        For i = Grid.CurrentCell.RowIndex + 1 To Grid.RowCount - 1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "UPDATE SI_BASE_CODE Set Base_Sun = '" & i & "' Where Base_Sel = '" & Sel_Name & "' And  Base_Sun = '" & i + 1 & "'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        Grid_Display(Grid, Sel_Name)
        Grid_Display_Ch = True

    End Sub

    Private Sub Man_Combo_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Man_Combo.SelectionChangeCommitted
        Man_Grid_Display()

    End Sub

    Private Sub Man_Delete_Com_Click(sender As Object, e As EventArgs) Handles Man_Delete_Com.Click
        Grid_Display_Ch = False
        If Len(Man_Grid.Item(0, Man_Grid.CurrentCell.RowIndex).Value) < 1 Then
            Exit Sub
        End If

        If MsgBox("'" & Man_Grid.Item(1, Man_Grid.CurrentCell.RowIndex).Value & "' 을(를) 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "삭제") = vbNo Then
            Exit Sub
        End If
        Dim i As Integer

        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_BASE_CODE Where Base_Sel = '" & Man_Combo.Text & "' And   Base_Sun = '" & Man_Grid.Item(0, Man_Grid.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        For i = Man_Grid.CurrentCell.RowIndex + 1 To Man_Grid.RowCount - 1
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "UPDATE SI_BASE_CODE Set Base_Sun = '" & i & "' Where Base_Sel = '" & Man_Combo.Text & "' And  Base_Sun = '" & i + 1 & "'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        Man_Grid_Display()
        Grid_Display_Ch = True


    End Sub
    Private Sub Man_Insert_Com_Click(sender As Object, e As EventArgs) Handles Man_Insert_Com.Click
        Dim DBT As New DataTable
        Dim Db_Sun As Long
        StrSQL = "Select Base_Sun FROM SI_BASE_CODE with(NOLOCK) Where base_Sel = '" & Man_Combo.Text & "' Order By Base_Sun Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            Db_Sun = 1
        Else
            Db_Sun = DBT.Rows(0).Item(0) + 1
        End If

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert into SI_Base_Code (Base_Sel, Base_Sun)  Values ('" & Man_Combo.Text & "', '" & Db_Sun & "')"
        Re_Count = DB_Execute()
        Man_Grid_Display()

    End Sub

    Private Sub Man_Grid_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Man_Grid.CellContentClick

    End Sub

    Private Sub Man_Grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Man_Grid.CellValueChanged
        If Grid_Display_Ch = True Then
            sender.CurrentRow.HeaderCell.Value = "*"
        End If

    End Sub

    Private Sub Man_Save_Com_Click(sender As Object, e As EventArgs) Handles Man_Save_Com.Click
        Grid_Display_Ch = False
        Dim i As Integer
        For i = 0 To Man_Grid.Rows.Count - 1
            If Man_Grid.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SI_BASE_CODE Set Base_Name = '" & Man_Grid.Item(1, i).Value & "' Where Base_Sel = '" & Man_Combo.Text & "' And Base_Sun = '" & Man_Grid.Item(0, i).Value & "'"
                Re_Count = DB_Execute()
                Man_Grid.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display_Ch = True


    End Sub
    Public Function Man_Grid_Setup() As Boolean
        Grid_Color(Man_Grid)
        Man_Grid.EnableHeadersVisualStyles = False

        Man_Grid.AllowUserToAddRows = False
        Man_Grid.RowTemplate.Height = 20
        Man_Grid.ColumnCount = 2
        Man_Grid.RowCount = 1


        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Man_Grid.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Man_Grid.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Man_Grid.RowHeadersWidth = 45
        Man_Grid.Columns(0).Width = 55
        Man_Grid.Columns(1).Width = 145

        Man_Grid.Columns(0).HeaderText = "코드"
        Man_Grid.Columns(1).HeaderText = "명칭"

        Man_Grid.Columns(0).ReadOnly = True

        Man_Grid_Setup = True

    End Function

    Public Function Man_Grid_Display() As Boolean
        Grid_Display_Ch = False

        Dim DBT As New DataTable
        Dim i As Integer
        Man_Grid.RowCount = 0
        StrSQL = "Select * FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '" & Man_Combo.SelectedItem.ToString() & "'  Order  By Base_Sun"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Man_Grid.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                Man_Grid.Item(0, i).Value = DBT.Rows(i)("Base_Sun")
                Man_Grid.Item(1, i).Value = DBT.Rows(i)("Base_Name")
            Next i
        Else
            Man_Grid.RowCount = 1
        End If
        Grid_Display_Ch = True
        Man_Grid_Display = True
    End Function
    Private Sub Level_Grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Level_Grid.CellValueChanged
        If Grid_Display_Ch = True Then
            Level_Grid.CurrentRow.HeaderCell.Value = "*"
        End If

    End Sub

    Private Sub Level_Grid_MouseClick(sender As Object, e As MouseEventArgs) Handles Level_Grid.MouseClick
        Level_Grid_Row = Level_Grid.CurrentCell.RowIndex
        Level_Grid_Col = Level_Grid.CurrentCell.ColumnIndex
        Level_Grid_Combo.Visible = False
        If Level_Grid_Row < 0 Then
            Exit Sub
        End If
        If Level_Grid_Col < 0 Then
            Exit Sub
        End If
        Dim DBT As New DataTable
        If Level_Grid_Col < 1 Then
        Else
            Level_Grid_Combo.Items.Clear()
            Level_Grid_Combo.Items.Add("True")
            Level_Grid_Combo.Items.Add("False")

            Level_Grid_Combo.Size = Level_Grid.CurrentCell.Size
            Level_Grid_Combo.Top = Level_Grid.GetCellDisplayRectangle(Level_Grid_Col, Level_Grid_Row, True).Top + Level_Grid.Top
            Level_Grid_Combo.Left = Level_Grid.GetCellDisplayRectangle(Level_Grid_Col, Level_Grid_Row, True).Left + Level_Grid.Left
            Level_Grid_Combo.Text = Level_Grid.CurrentCell.Value.ToString


            Level_Grid_Combo.BackColor = Level_Grid.RowsDefaultCellStyle.BackColor
            Level_Grid_Combo.Visible = True
            Level_Grid_Combo.Focus.ToString()
        End If

    End Sub

    Private Sub Level_Grid_Scroll(sender As Object, e As ScrollEventArgs) Handles Level_Grid.Scroll
        Level_Grid_Combo.Visible = False

    End Sub

    Private Sub Level_Grid_Combo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Level_Grid_Combo.SelectedIndexChanged

    End Sub

    Private Sub Level_Grid_Combo_LostFocus(sender As Object, e As EventArgs) Handles Level_Grid_Combo.LostFocus
        Level_Grid_Combo.Visible = False

    End Sub

    Private Sub Level_Grid_Combo_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Level_Grid_Combo.SelectionChangeCommitted
        Level_Grid.Item(Level_Grid_Col, Level_Grid_Row).Value = Level_Grid_Combo.SelectedItem.ToString()

    End Sub

    Private Sub Level_Save_Com_Click(sender As Object, e As EventArgs) Handles Level_Save_Com.Click
        Grid_Display_Ch = False
        Dim Table_Name As String
        Dim DBT As DataTable = Nothing
        Dim Field_Name(500) As String
        Dim i As Integer
        Dim j As Long
        Dim k As Integer

        Table_Name = "SI_Man_Level"
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
        For k = 0 To Level_Grid.RowCount - 1
            If Level_Grid.Rows(k).HeaderCell.Value = "*" Then
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE " & Table_Name & " SET "
                For i = 1 To Level_Grid.ColumnCount - 1
                    StrSQL = StrSQL & " " & Table_Name & "." & Field_Name(i) & " = '" & Trim(Level_Grid.Item(i, k).Value.ToString) & "'"
                    If i = Level_Grid.ColumnCount - 1 Then
                    Else
                        StrSQL = StrSQL & ","
                    End If
                Next i
                StrSQL = StrSQL & " WHERE " & Table_Name & "." & Field_Name(0) & " = '" & Level_Grid.Item(0, k).Value & "'"
                Re_Count = DB_Execute()
            End If
            Level_Grid.Rows(k).HeaderCell.Value = ""
        Next k
        Grid_Display_Ch = True
        MsgBox("저장되었습니다")
    End Sub

    Private Sub Level_Insert_Com_Click(sender As Object, e As EventArgs) Handles Level_Insert_Com.Click
        Dim DBT As New DataTable
        Dim Db_Sun As Long
        StrSQL = "Select Le_Code FROM SI_MAN_Level with(NOLOCK)  Order By Convert(Decimal,Le_Code) Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            Db_Sun = 1
        Else
            Db_Sun = DBT.Rows(0).Item(0) + 1
        End If
        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO SI_MAN_Level ( Le_Code)  Values ('" & Format("0#", Db_Sun) & "')"
        Re_Count = DB_Execute()
        Level_Grid_Display()
    End Sub

    Private Sub Product_Sel2_Save_Com_Click(sender As Object, e As EventArgs) Handles Product_Sel2_Save_Com.Click

    End Sub

    Private Sub Product_Grid_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Product_Grid.CellContentClick

    End Sub

    Private Sub Product_Sel2_Grid_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Product_Sel2_Grid.CellContentClick

    End Sub

    Private Sub Unit_Grid_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Unit_Grid.CellContentClick

    End Sub

    Private Sub Err_Grid_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Err_Grid.CellContentClick

    End Sub

    Private Sub Port_Grid_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Port_Grid.CellContentClick

    End Sub

    Private Sub Port_Grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Port_Grid.CellValueChanged
        If Grid_Display_Ch = True Then
            sender.CurrentRow.HeaderCell.Value = "*"
        End If
    End Sub

    Private Sub Stock_Insert_Com_Click(sender As Object, e As EventArgs) Handles Stock_Insert_Com.Click
        Dim DBT As New DataTable
        Dim Db_Sun As Long
        StrSQL = "Select FL_CODE FROM SI_STOCK with(NOLOCK)  Order By Convert(Decimal,FL_Code) Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            Db_Sun = 1
        Else
            Db_Sun = DBT.Rows(0).Item(0) + 1
        End If
        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO SI_STOCK ( FL_CODE)  Values ('" & Format("0#", Db_Sun) & "')"
        Re_Count = DB_Execute()
        STOCK_Grid_Display()

    End Sub

    Private Sub Stock_Save_Com_Click(sender As Object, e As EventArgs) Handles Stock_Save_Com.Click
        Grid_Display_Ch = False
        Dim i As Integer
        For i = 0 To stockgridview.Rows.Count - 1
            If stockgridview.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SI_STOCK Set FL_NAME = '" & stockgridview.Item(1, i).Value & "', FL_ST = '" & stockgridview.Item(2, i).Value & "', FL_IP = '" & stockgridview.Item(3, i).Value & "' Where FL_CODE = '" & stockgridview.Item(0, i).Value & "' "
                Re_Count = DB_Execute()
                stockgridview.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display_Ch = True
    End Sub

    Private Sub Stock_Delete_Com_Click(sender As Object, e As EventArgs) Handles Stock_Delete_Com.Click

        If MsgBox("'" & stockgridview.Item(1, stockgridview.CurrentCell.RowIndex).Value & "' 을(를) 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "삭제") = vbNo Then
            Exit Sub
        End If

        Grid_Display_Ch = False
        If Len(stockgridview.Item(0, stockgridview.CurrentCell.RowIndex).Value) < 1 Then
            Exit Sub
        End If
        Dim i As Integer

        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_STOCK Where FL_CODE = '" & stockgridview.Item(0, stockgridview.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        'For i = stockgridview.CurrentCell.RowIndex + 1 To stockgridview.RowCount - 1
        '    '수정한다
        '    StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        '    StrSQL = StrSQL & "UPDATE SI_STOCK Set FL_CODE = '" & i & "' Where FL_CODE = '" & stockgridview.Item(0, stockgridview.CurrentCell.RowIndex).Value & "'"
        '    Re_Count = DB_Execute()
        'Next i

        '선택된 갈럼값을 가지고 온다
        STOCK_Grid_Display()
        Grid_Display_Ch = True
    End Sub

    Private Sub stockgridview_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles stockgridview.CellValueChanged
        If Grid_Display_Ch = True Then
            sender.CurrentRow.HeaderCell.Value = "*"
        End If
    End Sub
End Class
