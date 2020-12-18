Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
'작성자: 태승업
Public Class Frm_Product
    Dim Grid_Display_Ch As Boolean

    Dim Grid_Info_Row As Integer
    Dim Grid_Info_Col As Integer

    Dim Grid_Bom_Row As Integer
    Dim Grid_Bom_Col As Integer

    Dim Grid_Pro_Row As Integer
    Dim Grid_Pro_Col As Integer

    Dim Info_Lab() As Button
    Dim Cellclick As Boolean = False

    Private Sub Frm_Product_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Me.Dock = DockStyle.Fill
        Me.BringToFront()
        Me.BackColor = Color.White

        Grid_Info.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        Grid_Bom.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        Grid_Pro.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        '그리드뷰 설정
        Grid_Info_Setup()
        Grid_Bom_Setup()
        Grid_Pro_Setup()

        Grid_Info.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Search_Com.PerformClick()

    End Sub


    Public Sub Grid_Bom_Setup()
        Grid_Color(Grid_Bom)

        Grid_Bom.AllowUserToAddRows = False
        Grid_Bom.RowTemplate.Height = 24
        Grid_Bom.ColumnCount = 5
        Grid_Bom.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Bom.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Bom.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Bom.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Bom.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Grid_Bom.RowHeadersWidth = 40
        Grid_Bom.Columns(0).Width = 40
        Grid_Bom.Columns(1).Width = 100
        Grid_Bom.Columns(2).Width = 150
        Grid_Bom.Columns(3).Width = 50
        Grid_Bom.Columns(4).Width = 90

        Grid_Bom.Columns(0).HeaderText = "순번"
        Grid_Bom.Columns(1).HeaderText = "제품코드"
        Grid_Bom.Columns(2).HeaderText = "제품명"
        Grid_Bom.Columns(3).HeaderText = "Qty"
        Grid_Bom.Columns(4).HeaderText = "단위"


        Grid_Bom.Item(0, 0).Value = ""
        Grid_Bom.Item(1, 0).Value = ""
        Grid_Bom.Item(2, 0).Value = ""
        Grid_Bom.Item(3, 0).Value = ""
        Grid_Bom.Item(4, 0).Value = ""

        Grid_Bom.Columns(0).ReadOnly = True
        Grid_Bom.ClearSelection()
    End Sub
    Public Sub Grid_Pro_Setup()

        Grid_Color(Grid_Pro)



        Grid_Pro.AllowUserToAddRows = False
        Grid_Pro.RowTemplate.Height = 24

        'Grid_Pro.ColumnHeadersHeight = 10
        Grid_Pro.ColumnCount = 3
        Grid_Pro.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Pro.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Pro.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Pro.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft


        Grid_Pro.RowHeadersWidth = 40
        Grid_Pro.Columns(0).Width = 40
        Grid_Pro.Columns(1).Width = 60
        Grid_Pro.Columns(2).Width = 300


        Grid_Pro.Columns(0).HeaderText = "순번"
        Grid_Pro.Columns(1).HeaderText = "공정코드"
        Grid_Pro.Columns(2).HeaderText = "공정명"


        'Grid_Pro.Columns(4).HeaderText = "공정전이미지"
        'Grid_Pro.Columns(5).HeaderText = "공정후이미지"


        Grid_Pro.Item(0, 0).Value = ""
        Grid_Pro.Item(1, 0).Value = ""
        Grid_Pro.Item(2, 0).Value = ""

        'Grid_Pro.Item(5, 0).Value = ""

        Grid_Pro.Columns(0).ReadOnly = True
        Grid_Pro.ClearSelection()
    End Sub

    Public Function Grid_Info_Setup() As Boolean
        Grid_Color(Grid_Info)

        '제목 폭
        Grid_Info.RowHeadersWidth = 10
        Grid_Info.ColumnHeadersHeight = 24

        'Data 부분 높이
        Grid_Info.RowTemplate.Height = 24
        Grid_Info.ColumnCount = 6
        Grid_Info.RowCount = 1






        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Info.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Info.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Info.Columns(0).Width = 100
        Grid_Info.Columns(1).Width = 240

        Grid_Info.Columns(0).HeaderText = "코드"
        Grid_Info.Columns(1).HeaderText = "품명"
        Grid_Info.Columns(2).HeaderText = "종류"
        Grid_Info.Columns(3).HeaderText = "규격"
        Grid_Info.Columns(4).HeaderText = "유통기한"
        Grid_Info.Columns(5).HeaderText = "비고"


        Grid_Info.Columns(0).ReadOnly = True
        Grid_Info.Columns(0).DefaultCellStyle.BackColor = Color.White



    End Function


    Private Sub Insert_Com_Click(sender As Object, e As EventArgs) Handles Insert_Com.Click
        '  Product_Insert.Label1.Text = "insert"
        '새 폼에서 제품정보 추가
        '    Product_Insert.Show()
        Dim form As New Product_Insert
        form.Label1.Text = "insert"
        form.ShowDialog()

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If


    End Sub


    Public Function Grid_Bom_Display(PL_Code As String) As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display_Ch = False
        Grid_Bom.RowCount = 1
        StrSQL = "Select PL_Sun, PL_Sub_Code, PL_Qty, PL_Unit  FROM SI_PRODUCT_RECIPE with(NOLOCK)  Where PL_Code = '" & PL_Code & "' Order By Convert(Decimal,PL_Sun)"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Bom.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                Grid_Bom.Item(0, i).Value = DBT.Rows(i).Item(0).ToString
                Grid_Bom.Item(1, i).Value = DBT.Rows(i).Item(1).ToString
                Grid_Bom.Item(2, i).Value = ""
                Grid_Bom.Item(3, i).Value = DBT.Rows(i).Item(2).ToString
                Grid_Bom.Item(4, i).Value = DBT.Rows(i).Item(3).ToString
            Next i

        Else
            Grid_Bom.RowCount = 1
            For j = 0 To Grid_Bom.ColumnCount - 1
                Grid_Bom.Item(j, 0).Value = ""
            Next j

        End If

        If Re_Count <> 0 Then
            For i = 0 To Grid_Bom.RowCount - 1
                StrSQL = "Select PL_Name  FROM SI_PRODUCT with(NOLOCK)  Where PL_Code = '" & Grid_Bom.Item(1, i).Value & "' "
                Re_Count2 = DB_Select(DBT)
                If Re_Count2 <> 0 Then
                    Grid_Bom.Item(2, i).Value = DBT.Rows(0).Item(0).ToString
                End If
            Next
        End If




        Grid_Display_Ch = True
        Grid_Bom_Display = True
    End Function
    Public Function Grid_Pro_Display(PL_Code As String) As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display_Ch = False
        Grid_Pro.RowCount = 1
        StrSQL = "Select PP_Sun, PP_Code FROM SI_Process_List with(NOLOCK)  Where PP_PC_Code = '" & PL_Code & "' Order By Convert(Decimal,PP_Sun)"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Pro.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                Grid_Pro.Item(0, i).Value = DBT.Rows(i).Item(0).ToString
                Grid_Pro.Item(1, i).Value = DBT.Rows(i).Item(1).ToString
            Next i
        Else
            Grid_Pro.RowCount = 1
            For j = 0 To Grid_Pro.ColumnCount - 1
                Grid_Pro.Item(j, 0).Value = ""
            Next j

        End If

        If Re_Count <> 0 Then
            For i = 0 To Grid_Pro.RowCount - 1
                StrSQL = "Select PC_Name  FROM SI_PROCESS with(NOLOCK)  Where PC_Code = '" & Grid_Pro.Item(1, i).Value & "' "
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Pro.Item(2, i).Value = DBT.Rows(0).Item(0).ToString
                End If
            Next

        End If


        Grid_Display_Ch = True
        Grid_Pro_Display = True
    End Function

    Public Function Grid_Info_Display_Pro(PL_Code As String) As Boolean
        Grid_Display_Ch = False
        'Try
        Dim DBT As New DataTable
        Dim i As Integer
        StrSQL = "Select * FROM SI_PRODUCT With(NOLOCK) Where PL_Code = '" & PL_Code & "' AND PL_SEL != '원부재료'"
        Re_Count = DB_Select(DBT)
        If Re_Count > 0 Then
            For i = 0 To DBT.Columns.Count - 1
                Grid_Info.Item(1, i).Value = DBT.Rows(0).Item(i).ToString
                '이미지 처리 한다.
                If i = 7 Then
                    If Len(Grid_Info.Item(1, i).Value) < 100 Then
                        '   Grid_Info_Picture1.Image = Nothing
                    Else
                        Dim b() As Byte = HexStr2Array(Grid_Info.Item(1, i).Value)
                        '  Grid_Info_Picture1.Image = Image.FromStream(New MemoryStream(b, 0, b.Length))
                    End If
                End If
                If i = 6 Then
                    Grid_Info.Item(1, i).Value = Format(Val(Grid_Info.Item(1, i).Value), "###,###,###,##0.00")
                End If
            Next i
        Else
            For i = 0 To DBT.Columns.Count - 1
                Grid_Info.Item(1, i).Value = ""
            Next i
        End If
        Grid_Display_Ch = True
        Grid_Info_Display_Pro = True

    End Function



    Private Sub Grid_Info_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Info.CellEnter
        '선택 시
        Cellclick = True
        Grid_Bom_Display(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value)
        Grid_Pro_Display(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value)
    End Sub


    Private Sub Save_Com_Click(sender As Object, e As EventArgs)
        'Grid_Info 를 저장 한다.
        'DataBase에서 필드 이름 가지고 오기

        If Grid_Info.Item(1, 1).Value = "" Then
            MsgBox("품명은 공백으로 둘 수 없습니다.")
            Exit Sub
        End If

        If Grid_Info.Item(1, 2).Value = "" Then
            MsgBox("구분은 공백으로 둘 수 없습니다.")
            Exit Sub
        End If

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
            Exit Sub
        End If


        Table_Name = "SI_Product"
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

        If Grid_Info.Item(1, 2).Value = "" Then
            MsgBox("구분은 공백으로 둘 수 없습니다.")
            Exit Sub
        End If


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
        StrSQL = StrSQL & " WHERE " & Table_Name & "." & Field_Name(0) & " = '" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        Search_Com.PerformClick()

    End Sub

    Private Sub Grid_Info_Picture1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Grid_Info_Picture1_DoubleClick(sender As Object, e As EventArgs)
        '이미지 파일 선택
        Pic_File.InitialDirectory = Application.StartupPath
        Pic_File.Filter = "All files (*.*)|*.*"
        Pic_File.FilterIndex = 1
        Pic_File.RestoreDirectory = True
        Grid_Display_Ch = True
        If Pic_File.ShowDialog() = DialogResult.OK Then
            'Try
            Dim fs As Stream = Pic_File.OpenFile()
            If (fs IsNot Nothing) Then
                Dim b(fs.Length() - 1) As Byte
                fs.Read(b, 0, b.Length)
                Grid_Info.Item(1, 7).Value = Array2HexStr(b)
                '   Grid_Info_Picture1.Image = Image.FromStream(fs)
                Grid_Info.Rows(7).HeaderCell.Value = "*"
                fs.Close()
            End If
            'Catch Ex As Exception
            'MsgBox(Ex.Message)
            'End Try
        Else
            '   Grid_Info.Item(1, 7).Value = ""
            '     Grid_Info_Picture1.Image = Nothing
            Grid_Info.Rows(7).HeaderCell.Value = "*"
        End If

    End Sub
    Private Sub Grid_Bom_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Bom.CellValueChanged
        If Grid_Display_Ch = True Then
            Grid_Bom.CurrentRow.HeaderCell.Value = "*"
        End If
    End Sub

    Private Sub Grid_Bom_Scroll(sender As Object, e As ScrollEventArgs) Handles Grid_Bom.Scroll
        Grid_Bom_Combo.Visible = False

    End Sub
    Private Sub Grid_Bom_Combo_LostFocus(sender As Object, e As EventArgs) Handles Grid_Bom_Combo.LostFocus
        Grid_Bom_Combo.Visible = False

    End Sub

    Private Sub Grid_Bom_Combo_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Grid_Bom_Combo.SelectionChangeCommitted
        Dim DBT As New DataTable
        Grid_Bom.Item(Grid_Bom_Col, Grid_Bom_Row).Value = Grid_Bom_Combo.SelectedItem.ToString()
        Select Case Grid_Bom_Col
            Case 1
                StrSQL = "Select PL_Code, PL_Name  FROM SI_PRODUCT with(NOLOCK) Where PL_Code = '" & Grid_Bom_Combo.SelectedItem.ToString() & "' and pl_sel != '제품'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Bom.Item(2, Grid_Bom_Row).Value = DBT.Rows(0).Item(1)
                End If
            Case 2
                StrSQL = "Select PL_Code, PL_Name  FROM SI_PRODUCT with(NOLOCK) Where PL_Name= '" & Replace(Grid_Bom_Combo.SelectedItem.ToString(), "'", "''") & "' and pl_sel != '제품'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Bom.Item(1, Grid_Bom_Row).Value = DBT.Rows(0).Item(0)
                End If
        End Select

    End Sub

    Private Sub Grid_Pro_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Pro.CellValueChanged
        If Grid_Display_Ch = True Then
            Grid_Pro.CurrentRow.HeaderCell.Value = "*"
        End If

    End Sub



    Private Sub Save_Pro_Click(sender As Object, e As EventArgs)
        If Len(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If


        Grid_Display_Ch = False
        Dim i As Integer

        For i = 0 To Grid_Pro.Rows.Count - 1
            If Grid_Pro.Rows(i).HeaderCell.Value = "*" Then
                If Grid_Pro.Item(3, i).Value = "" Then
                    MsgBox("값은 공백일 수 없습니다")
                    Exit Sub
                End If
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SI_PROCESS_LIST Set PP_PC_Code = '" & Grid_Pro.Item(1, i).Value.ToString & "', PP_Amount = '" & Grid_Pro.Item(3, i).Value.ToString & "' , PP_GOLD = '" & Grid_Pro.Item(4, i).Value.ToString & "' where PP_Code = '" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value.ToString & "' And PP_Sun = '" & Grid_Pro.Item(0, i).Value.ToString & "'"
                Re_Count = DB_Execute()
                Grid_Pro.Rows(i).HeaderCell.Value = ""
            End If
        Next i

        Grid_Display_Ch = True
        MsgBox("저장되었습니다")
    End Sub
    Private Sub Insert__Pro_Click(sender As Object, e As EventArgs)
        If Len(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        Grid_Display_Ch = False
        Dim DBT As New DataTable
        Dim Db_Sun As Long
        StrSQL = "Select PP_Sun FROM SI_Process_List with(NOLOCK) Where PP_Code = '" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value & "' Order By PP_Sun Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            Db_Sun = 1
        Else
            Db_Sun = DBT.Rows(0).Item(0) + 1
            If Db_Sun = 10 Then
                MsgBox("최대수 도달")
                Exit Sub
            End If
            For i = 0 To Grid_Pro.Rows.Count - 1
                If Grid_Pro.Rows(i).HeaderCell.Value = "*" Then
                    If Grid_Pro.Item(3, i).Value = "" Or Grid_Pro.Item(4, i).Value = "" Then
                        MsgBox("값은 공백일 수 없습니다")
                        Exit Sub
                    End If
                    'Update
                    StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                    StrSQL = StrSQL & "UPDATE SI_PROCESS_LIST Set PP_PC_Code = '" & Grid_Pro.Item(1, i).Value.ToString & "', PP_Amount = '" & Grid_Pro.Item(3, i).Value.ToString & "' , PP_GOLD = '" & Grid_Pro.Item(4, i).Value.ToString & "' where PP_Code = '" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value.ToString & "' And PP_Sun = '" & Grid_Pro.Item(0, i).Value.ToString & "'"
                    Re_Count = DB_Execute()
                    Grid_Pro.Rows(i).HeaderCell.Value = ""
                End If
            Next i
        End If

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO SI_PROCESS_LIST (PP_Code, PP_Sun)  Values ('" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value & "', '" & Db_Sun & "')"
        Re_Count = DB_Execute()
        Grid_Pro_Display(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True



    End Sub

    Private Sub Put__Pro_Click(sender As Object, e As EventArgs)
        If Len(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If


        If Len(Grid_Pro.Item(0, Grid_Pro.CurrentCell.RowIndex).Value) < 1 Then
            Exit Sub
        End If
        Grid_Display_Ch = False
        Dim i As Integer
        For i = Grid_Pro.RowCount - 1 To Grid_Pro.CurrentCell.RowIndex Step -1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "UPDATE SI_PROCESS_LIST Set PP_Sun = '" & i + 2 & "' Where PP_Sun = '" & i + 1 & "'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO SI_PROCESS_LIST (PP_Code, PP_Sun)  Values ('" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value & "', '" & Grid_Pro.CurrentCell.RowIndex + 1 & "')"
        Re_Count = DB_Execute()

        Grid_Pro_Display(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True

    End Sub

    Private Sub Delete__Pro_Click(sender As Object, e As EventArgs)
        If Len(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If


        If Len(Grid_Pro.Item(0, Grid_Pro.CurrentCell.RowIndex).Value) < 1 Then
            Exit Sub
        End If

        If MsgBox("'" & Grid_Pro.Item(2, Grid_Pro.CurrentCell.RowIndex).Value & "'을(를) 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "삭제") = vbNo Then
            Exit Sub
        End If

        Grid_Display_Ch = False
        Dim i As Integer
        Dim DBT As New DataTable
        Dim count As Integer = 0

        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_PROCESS_LIST Where PP_Code = '" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value & "' And   PP_Sun = '" & Grid_Pro.Item(0, Grid_Pro.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        For i = Grid_Pro.CurrentCell.RowIndex + 1 To Grid_Pro.RowCount - 1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "UPDATE SI_PROCESS_LIST Set PP_Sun = '" & i & "' Where PP_Code = '" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value & "' And  PP_Sun = '" & i + 1 & "'"
            Re_Count = DB_Execute()
        Next i

        '수정후

        'StrSQL = "Select PP_Sun FROM SI_Process_List with(NOLOCK) Where PP_Code = '" & Grid_Info.Item(0, Grid_info.currentcell.rowindex).Value & "' Order By PP_Sun Desc "
        'Re_Count = DB_Select(DBT)

        'If Re_Count <> 0 Then
        '    count = DBT.Rows(0)("PP_Sun")
        'Else
        '    Exit Sub
        'End If

        'StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        'StrSQL = StrSQL & "DELETE SI_PROCESS_LIST Where PP_Code = '" & Grid_Info.Item(0, Grid_info.currentcell.rowindex).Value & "' And   PP_Sun = '" & count & "'"
        'Re_Count = DB_Execute()

        Grid_Pro_Display(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True

    End Sub
    Private Sub Grid_Pro_Combo_LostFocus(sender As Object, e As EventArgs) Handles Grid_Pro_Combo.LostFocus
        Grid_Pro_Combo.Visible = False

    End Sub

    Private Sub Grid_Pro_Combo_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Grid_Pro_Combo.SelectionChangeCommitted
        Dim DBT As New DataTable
        Grid_Pro.Item(Grid_Pro_Col, Grid_Pro_Row).Value = Grid_Pro_Combo.SelectedItem.ToString()
        Select Case Grid_Pro_Col
            Case 1
                StrSQL = "Select PC_Code, PC_Name  FROM SI_PROCESS with(NOLOCK) Where PC_Code = '" & Grid_Pro_Combo.SelectedItem.ToString() & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Pro.Item(2, Grid_Pro_Row).Value = DBT.Rows(0).Item(1)
                End If
            Case 2
                StrSQL = "Select PC_Code, PC_Name  FROM SI_PROCESS with(NOLOCK) Where PC_Name = '" & Grid_Pro_Combo.SelectedItem.ToString() & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Pro.Item(1, Grid_Pro_Row).Value = DBT.Rows(0).Item(0)
                End If
        End Select


    End Sub

    Private Sub Delete_Com_Click(sender As Object, e As EventArgs) Handles Delete_Com.Click
        '제품 코드 삭제
        '다른 제품에 BOM에 있으면 안내 메세지를 표시해 준다.
        If Len(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value) < 1 Then
            Exit Sub
        End If


        If MsgBox("제품 코드 '" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value & "' 제품명'" & Grid_Info.Item(1, Grid_Info.CurrentCell.RowIndex).Value & "' 를 삭제 하면 다른 제품에 영향을 줄 수 있습니다. 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "제품 삭제") = vbNo Then
            Exit Sub
        End If
        Grid_Display_Ch = False

        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_PRODUCT Where PL_Code = '" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value & "' "
        Re_Count = DB_Execute()

        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_PRODUCT_RECIPE Where PL_Code = '" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value & "' "
        Re_Count = DB_Execute()

        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_PRODUCT_QC Where PQ_Code = '" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value & "' "
        Re_Count = DB_Execute()

        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_PROCESS_LIST Where PP_Code = '" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value & "' "

        Re_Count = DB_Execute()

        '재고테이블에서 삭제

        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE STOCK Where PL_Code = '" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value & "' "
        Re_Count = DB_Execute()


        Info_Display()
        Grid_Bom.RowCount = 0
        Grid_Bom.RowCount = 0

        Grid_Display_Ch = True

    End Sub


    Private Sub Btn_Form_Click(sender As Object, e As EventArgs) Handles Btn_Form.Click
        '파일을 불러온다.
        OpenFileDialog1.InitialDirectory = Application.StartupPath
        OpenFileDialog1.Filter = "xlsx file(*.xlsx)|*.xlsx|All files(*,*)|*.*"
        OpenFileDialog1.FilterIndex = 0
        OpenFileDialog1.RestoreDirectory = True
        Dim i As Integer
        Dim j As Integer
        Dim Ex_Count As Integer

        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            Try
                Dim xlapp As Excel.Application
                Dim xlbook As Excel.Workbook
                Dim xlsheet As Excel.Worksheet
                xlapp = New Excel.Application
                xlbook = xlapp.Workbooks.Open(OpenFileDialog1.FileName)
                'xlsheet = xlbook.ActiveSheet
                xlsheet = xlbook.Sheets(1)
                xlapp.DisplayAlerts = False
                xlapp.Visible = False
                Dim popup_excel As New popup_excel
                popup_excel.Pro_Info.RowCount = xlsheet.UsedRange.Rows.Count - 1
                ' MsgBox(xlsheet.UsedRange.Rows.Count)
                ' MsgBox(xlsheet.UsedRange.Columns.Count)
                For i = 0 To xlsheet.UsedRange.Rows.Count - 1
                    For j = 0 To xlsheet.UsedRange.Columns.Count - 1
                        ' MsgBox(xlsheet.Cells(i + 1, j).Value)
                        ' popup_excel.Pro_Info.Item(j, i).Value = xlsheet.Cells(i + 1, j).Value
                        '   MsgBox(xlsheet.Cells(i + 1, j).Value)
                    Next j
                Next i

                xlsheet = Nothing
                xlbook.Close()
                xlbook = Nothing
                xlapp.Quit()
                xlapp = Nothing
                '   releaseObject(xlsheet)
                '  releaseObject(xlbook)
                '  releaseObject(xlapp)
                Dim xlp() As Process = Process.GetProcessesByName("EXCEL")
                For Each Process As Process In xlp
                    Process.Kill()
                Next

                MsgBox("제품 정보가 추가되었습니다")
                '  popup_excel.Show()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        Else
            MsgBox("취소하셨습니다.")
            Exit Sub
        End If
        popup_excel.Show()

    End Sub

    Private Sub Insert__Bom_Click(sender As Object, e As EventArgs) Handles Insert__Bom.Click

        If CellClick = False Then
            MsgBox("제품을 선택해주세요")
            Exit Sub
        End If

        Dim form As New Bom_Insert
        form.Label1.Text = "insert2"
        form.process_code.Text = Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value
        form.process_name.Text = Grid_Info.Item(1, Grid_Info.CurrentCell.RowIndex).Value
        form.Process_bigo.Text = Grid_Info.Item(5, Grid_Info.CurrentCell.RowIndex).Value

        form.ShowDialog()

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If

        'If Len(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value) > 0 Then
        'Else
        '    MsgBox("제품코드를 선택해주세요")
        '    Exit Sub
        'End If

        'Grid_Display_Ch = False

        'Dim DBT As New DataTable
        'Dim Db_Sun As Long
        'StrSQL = "Select PL_Sun FROM SI_PRODUCT_RECIPE with(NOLOCK) Where PL_Code = '" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value & "' Order By Convert(Decimal,PL_Sun) Desc "
        'Re_Count = DB_Select(DBT)
        'If Re_Count = 0 Then
        '    Db_Sun = 1
        'Else
        '    Db_Sun = DBT.Rows(0).Item(0) + 1

        '    For i = 0 To Grid_Bom.Rows.Count - 1
        '        If Grid_Bom.Rows(i).HeaderCell.Value = "*" Then
        '            'Update
        '            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        '            StrSQL = StrSQL & "UPDATE SI_PRODUCT_RECIPE Set PL_Sub_Code = '" & Grid_Bom.Item(1, i).Value & "', PL_Qty = '" & Grid_Bom.Item(3, i).Value & "', PL_UNIT = '" & Grid_Bom.Item(4, i).Value & "' where PL_Code = '" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value & "' And PL_Sun = '" & Grid_Bom.Item(0, i).Value & "'"
        '            Re_Count = DB_Execute()
        '            Grid_Bom.Rows(i).HeaderCell.Value = ""
        '        End If
        '    Next i
        'End If

        ''추가한다
        'StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        'StrSQL = StrSQL & "Insert INTO SI_PRODUCT_RECIPE (PL_Code, PL_Sun)  Values ('" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value & "', '" & Db_Sun & "')"
        'Re_Count = DB_Execute()
        'Grid_Bom_Display(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value)
        'Grid_Display_Ch = True
    End Sub

    Private Sub Save_Bom_Click(sender As Object, e As EventArgs)
        If Len(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        Dim i As Integer


        Grid_Display_Ch = False
        For i = 0 To Grid_Bom.Rows.Count - 1
            If Grid_Bom.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SI_PRODUCT_RECIPE 
                            Set PL_Sub_Code = '" & Grid_Bom.Item(1, i).Value & "', PL_Qty = '" & Grid_Bom.Item(3, i).Value & "', PL_UNIT = '" & Grid_Bom.Item(4, i).Value & "' 
                            where PL_Code = '" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value & "' And PL_Sun = '" & Grid_Bom.Item(0, i).Value & "'"
                Re_Count = DB_Execute()
                Grid_Bom.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display_Ch = True
        MsgBox("저장되었습니다")
    End Sub

    Private Sub Delete__Bom_Click(sender As Object, e As EventArgs) Handles Delete__Bom.Click
        '해당 칼럼 삭제
        If Len(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        If MsgBox("'" & Grid_Bom.Item(2, Grid_Bom.CurrentCell.RowIndex).Value & "'을(를) 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "삭제") = vbNo Then
            Exit Sub
        End If

        Grid_Display_Ch = False
        Dim i As Integer
        Dim DBT As New DataTable
        Dim count As Integer = 0
        '수정전
        ''삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_PRODUCT_RECIPE Where PL_Code = '" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value & "' And   PL_Sun = '" & Grid_Bom.Item(0, Grid_Bom.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        For i = Grid_Bom.CurrentCell.RowIndex + 1 To Grid_Bom.RowCount - 1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "UPDATE SI_PRODUCT_RECIPE Set PL_Sun = '" & i & "' Where PL_Code = '" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value & "' And  PL_Sun = '" & i + 1 & "'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다

        '수정후
        'StrSQL = "Select PL_Sun FROM SI_PRODUCT_RECIPE with(NOLOCK) Where PL_Code = '" & Grid_Info.Item(0, Grid_info.currentcell.rowindex).Value & "' Order By Convert(Decimal,PL_Sun) Desc "
        'Re_Count = DB_Select(DBT)
        'If Re_Count <> 0 Then
        '    count = DBT.Rows(0)("PL_Sun")
        'Else
        '    Exit Sub
        'End If

        'StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        'StrSQL = StrSQL & "DELETE SI_PRODUCT_RECIPE Where PL_Code = '" & Grid_Info.Item(0, Grid_info.currentcell.rowindex).Value & "' And   PL_Sun = '" & count & "'"
        'Re_Count = DB_Execute()

        Grid_Bom_Display(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True
    End Sub

    Private Sub Insert_Pro_Click(sender As Object, e As EventArgs) Handles Insert_Pro.Click

        If Cellclick = False Then
            MsgBox("제품을 선택해주세요")
            Exit Sub
        End If

        Dim form As New PProcess_Insert
        form.Label1.Text = "insert2"
        form.process_code.Text = Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value
        form.process_name.Text = Grid_Info.Item(1, Grid_Info.CurrentCell.RowIndex).Value
        form.Process_bigo.Text = Grid_Info.Item(5, Grid_Info.CurrentCell.RowIndex).Value

        form.ShowDialog()

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If

        ' Grid_Pro.Rows.Add()
        'If Len(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value) > 0 Then
        'Else
        '    Exit Sub
        'End If

        'Grid_Display_Ch = False
        'Dim DBT As New DataTable
        'Dim Db_Sun As Long
        'StrSQL = "Select PP_Sun FROM SI_Process_List with(NOLOCK) Where PP_Code = '" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value & "' Order By PP_Sun Desc "
        'Re_Count = DB_Select(DBT)
        'If Re_Count = 0 Then
        '    Db_Sun = 1
        'Else
        '    Db_Sun = DBT.Rows(0).Item(0) + 1
        'End If
        ''추가한다
        'StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        'StrSQL = StrSQL & "Insert INTO SI_PROCESS_LIST (PP_Code, PP_Sun)  Values ('" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value & "', '" & Db_Sun & "')"
        'Re_Count = DB_Execute()

        'Grid_Pro_Display(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value)
        'Grid_Display_Ch = True
    End Sub

    Private Sub Save_Pro_Click_1(sender As Object, e As EventArgs)
        If Len(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If


        Grid_Display_Ch = False
        Dim i As Integer

        For i = 0 To Grid_Pro.Rows.Count - 1
            If Grid_Pro.Rows(i).HeaderCell.Value = "*" Then

                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SI_PROCESS_LIST Set PP_PC_Code = '" & Grid_Pro.Item(1, i).Value.ToString & "' where PP_Code = '" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value.ToString & "' And PP_Sun = '" & Grid_Pro.Item(0, i).Value.ToString & "'"
                Re_Count = DB_Execute()
                Grid_Pro.Rows(i).HeaderCell.Value = ""
            End If
        Next i

        Grid_Display_Ch = True
        MsgBox("저장되었습니다")
    End Sub

    Private Sub Save_delete_Click(sender As Object, e As EventArgs) Handles Save_delete.Click
        If Len(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If


        If Len(Grid_Pro.Item(0, Grid_Pro.CurrentCell.RowIndex).Value) < 1 Then
            Exit Sub
        End If

        If MsgBox("'" & Grid_Pro.Item(2, Grid_Pro.CurrentCell.RowIndex).Value & "'을(를) 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "삭제") = vbNo Then
            Exit Sub
        End If

        Grid_Display_Ch = False
        Dim i As Integer
        Dim DBT As New DataTable
        Dim count As Integer = 0

        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_PROCESS_LIST Where PP_Code = '" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value & "' And   PP_Sun = '" & Grid_Pro.Item(0, Grid_Pro.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        For i = Grid_Pro.CurrentCell.RowIndex + 1 To Grid_Pro.RowCount - 1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "UPDATE SI_PROCESS_LIST Set PP_Sun = '" & i & "' Where PP_Code = '" & Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value & "' And  PP_Sun = '" & i + 1 & "'"
            Re_Count = DB_Execute()
        Next i

        '수정후

        'StrSQL = "Select PP_Sun FROM SI_Process_List with(NOLOCK) Where PP_Code = '" & Grid_Info.Item(0, Grid_info.currentcell.rowindex).Value & "' Order By PP_Sun Desc "
        'Re_Count = DB_Select(DBT)

        'If Re_Count <> 0 Then
        '    count = DBT.Rows(0)("PP_Sun")
        'Else
        '    Exit Sub
        'End If

        'StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        'StrSQL = StrSQL & "DELETE SI_PROCESS_LIST Where PP_Code = '" & Grid_Info.Item(0, Grid_info.currentcell.rowindex).Value & "' And   PP_Sun = '" & count & "'"
        'Re_Count = DB_Execute()

        Grid_Pro_Display(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True
    End Sub

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        Info_Display()
    End Sub

    Public Sub Info_Display()
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer

        StrSQL = "Select * FROM SI_PRODUCT With(NOLOCK) Where PL_Name Like '%" & Search_Text.Text & "%' and PL_SEL !='원부재료' Order By PL_Name"

        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Info.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To Grid_Info.ColumnCount - 1
                    Grid_Info.Item(j, i).Value = DBT.Rows(i).Item(j)
                Next j
            Next i
        Else
            Grid_Info.RowCount = 1
            For j = 0 To Grid_Info.ColumnCount - 1
                Grid_Info.Item(j, 0).Value = ""
            Next j
        End If
        Grid_Info.ClearSelection()
    End Sub

    Private Sub Grid_Info_DoubleClick(sender As Object, e As EventArgs) Handles Grid_Info.DoubleClick
        '유효성 검사
        If Len(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value) < 0 Then
            MsgBox("해당 행은 공백이므로 처리할 수 없습니다")
            Exit Sub
        End If

        '해당 행 내용 수정
        Product_Insert.Label1.Text = "update"
        Product_Insert.product_code.Text = Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value
        ' Product_Insert.Show()
        Dim form As New Product_Insert
        form.Label1.Text = "update"
        form.product_code.Text = Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value
        form.ShowDialog()

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If

    End Sub

    Private Sub Grid_Bom_DoubleClick(sender As Object, e As EventArgs) Handles Grid_Bom.DoubleClick
        If Len(Grid_Bom.Item(0, Grid_Bom.CurrentCell.RowIndex).Value) < 0 Then
            MsgBox("해당 행은 공백이므로 처리할 수 없습니다")
            Exit Sub
        End If

        '해당 행 내용 수정
        Dim form As New Bom_Insert
        form.Label1.Text = "update2"
        form.process_code.Text = Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value
        form.sun.Text = Grid_Bom.Item(0, Grid_Bom.CurrentCell.RowIndex).Value
        form.ShowDialog()

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If

    End Sub

    Private Sub Grid_Pro_DoubleClick(sender As Object, e As EventArgs) Handles Grid_Pro.DoubleClick
        If Len(Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value) < 0 Then
            MsgBox("해당 행은 공백이므로 처리할 수 없습니다")
            Exit Sub
        End If

        '해당 행 내용 수정
        Dim form As New PProcess_Insert
        form.Label1.Text = "update2"
        form.process_code.Text = Grid_Info.Item(0, Grid_Info.CurrentCell.RowIndex).Value
        form.TextBox3.Text = Grid_Pro.Item(0, Grid_Pro.CurrentCell.RowIndex).Value
        form.ShowDialog()

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If
    End Sub
End Class
