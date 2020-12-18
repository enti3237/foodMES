Imports System.IO

Public Class Frm_Product4
    Dim Grid_Display_Ch As Boolean

    Dim Grid_Info_Row As Integer
    Dim Grid_Info_Col As Integer


    Dim Grid_Bom_Row As Integer
    Dim Grid_Bom_Col As Integer

    Dim Grid_Pro_Row As Integer
    Dim Grid_Pro_Col As Integer

    Dim Grid_Cus_In_Row As Integer
    Dim Grid_Cus_In_Col As Integer
    Dim Grid_Cus_Out_Row As Integer
    Dim Grid_Cus_Out_Col As Integer


    Dim Tree_Nodes_Count As Long

    Private Sub Frm_Product_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()
        Me.BackColor = Color.White

        Grid_Code_Setup()
        Search_Com.PerformClick()

        Grid_Code.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

    End Sub

    Public Function Grid_Code_Setup() As Boolean
        Grid_Color(Grid_Code)
        Grid_Code.EnableHeadersVisualStyles = False
        Grid_Code.AllowUserToAddRows = False

        Grid_Code.RowTemplate.Height = 24
        Grid_Code.ColumnCount = 5
        Grid_Code.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Code.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Code.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Code.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Code.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Code.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        '지정하지 않은 헤더는 전체 사이즈에서 균일하게 나뉜다
        'Grid_Code.RowHeadersWidth = 100
        Grid_Code.Columns(0).Width = 100
        Grid_Code.Columns(1).Width = 250
        Grid_Code.Columns(2).Width = 120
        Grid_Code.Columns(3).Width = 120
        'Grid_Code.Columns(4).Width = 70
        Grid_Code.Columns(0).HeaderText = "코드"
        Grid_Code.Columns(1).HeaderText = "원부재료품명"
        Grid_Code.Columns(2).HeaderText = "구분"
        Grid_Code.Columns(3).HeaderText = "규격"
        Grid_Code.Columns(4).HeaderText = "비고"
        Dim i As Integer
        For i = 0 To Grid_Code.ColumnCount - 1
            Grid_Code.Columns(i).ReadOnly = True
        Next i
    End Function

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click

        Grid_Code_Display(True)

    End Sub

    Public Function Grid_Code_Display(Display_Sel As Boolean) As Boolean

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer

        StrSQL = "Select PL_Code, PL_Name, PL_Sel, PL_Standard, PL_Bigo FROM SI_PRODUCT With(NOLOCK) Where PL_Name Like '%" & Search_Text.Text & "%' AND PL_SEL='원부재료'  Order By PL_Code DESC"

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

    Private Sub Save_Com_Click_1(sender As Object, e As EventArgs)

        Dim Table_Name As String
        Dim j As Long
        Dim DBT As DataTable = Nothing
        Dim Field_Name(500) As String
        Dim i As Integer

        Table_Name = "SI_Product"
        j = 0
        StrSQL = "Select sys.Columns.Name From sys.Tables With(NOLOCK) , sys.Columns With(NOLOCK) Where sys.Tables.name = '" & Table_Name & "' And sys.Tables.Object_id = sys.Columns.Object_id Order By sys.Columns.column_id"
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

        '리프레쉬가 안되는데 어떻게 하지
        Search_Com.PerformClick()
        Grid_Code.ClearSelection()
        MsgBox("저장되었습니다")

    End Sub

    Private Sub Insert_Com_Click(sender As Object, e As EventArgs) Handles Insert_Com.Click

        '새로운 폼에서 입력 받는다.
        'Marterial_Insert.Show()
        Dim form As New Material_Insert
        form.Label1.Text = "insert"
        form.ShowDialog()

    End Sub


    Private Sub Grid_Code_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Code.CellValueChanged
        If Grid_Display_Ch = True Then
            sender.CurrentRow.HeaderCell.Value = "*"
        End If

    End Sub

    Private Sub Delete_Com_Click_1(sender As Object, e As EventArgs) Handles Delete_Com.Click

        '활성화
        '해당 칼럼 삭제
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        If MsgBox("원부재료 '" & Grid_Code.Item(1, Grid_Code.CurrentCell.RowIndex).Value & "'을(를) 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "삭제") = vbNo Then
            Exit Sub
        End If

        Grid_Display_Ch = False
        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_PRODUCT Where PL_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' "
        Re_Count = DB_Execute()

        '재고테이블에서 삭제

        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE STOCK Where PL_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' "
        Re_Count = DB_Execute()

        Search_Com.PerformClick()
        Grid_Code.ClearSelection()
        MsgBox("삭제되었습니다")

    End Sub

    Private Sub Search_Text_TextChanged(sender As Object, e As EventArgs) Handles Search_Text.TextChanged

    End Sub



    Private Sub Grid_Code_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles Grid_Code.MouseDoubleClick
        '유효성 검사
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) < 0 Then
            MsgBox("해당 행은 공백이므로 처리할 수 없습니다")
            Exit Sub
        End If

        '해당 행 내용 수정
        'Material_Insert.Label1.Text = "update"
        'Material_Insert.txt_Matecode.Text = Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value
        ' Product_Insert.Show()
        Dim form As New Material_Insert
        form.Label1.Text = "update"
        form.txt_Matecode.Text = Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value
        form.ShowDialog()

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If

    End Sub
End Class
