Public Class Frm_adjust

    Dim Grid_Display_Ch As Boolean
    Dim Man_Grid_Row As Integer
    Dim Man_Grid_Col As Integer

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        Grid_Display()
    End Sub

    Private Sub Man_Grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Main.CellValueChanged
        If Grid_Display_Ch = True Then
            sender.CurrentRow.HeaderCell.Value = "*"
        End If

    End Sub

    Public Function Grid_Setup() As Boolean
        Dim i As Integer
        Grid_Color(Grid_Main)
        Grid_Main.EnableHeadersVisualStyles = False

        Grid_Main.AllowUserToAddRows = False
        Grid_Main.RowTemplate.Height = 20
        Grid_Main.ColumnCount = 6
        Grid_Main.RowCount = 1


        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Main.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Main.RowHeadersWidth = 45
        Grid_Main.Columns(0).Width = 70
        For i = 1 To Grid_Main.ColumnCount - 1
            Grid_Main.Columns(i).Width = 100
            Grid_Main.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Next i
        Grid_Main.Columns(2).Width = 250
        Grid_Main.Columns(5).Width = 150

        Grid_Main.Columns(0).HeaderText = "구분"
        Grid_Main.Columns(1).HeaderText = "제품코드"
        Grid_Main.Columns(2).HeaderText = "제품명"
        Grid_Main.Columns(3).HeaderText = "현재고"
        Grid_Main.Columns(4).HeaderText = "실재고"
        Grid_Main.Columns(5).HeaderText = "사유"

        Grid_Main.Columns(0).ReadOnly = True
        Grid_Main.Columns(1).ReadOnly = True
        Grid_Main.Columns(2).ReadOnly = True
        Grid_Main.Columns(3).ReadOnly = True
        Grid_Setup = True
    End Function
    Public Function Grid_Display() As Boolean
        Grid_Display_Ch = False
        Dim DBT As New DataTable
        Dim DBT2 As New DataTable

        Dim Re_Count2 As Integer = 0
        Dim i As Integer
        Dim j As Integer
        Grid_Main.RowCount = 0

        If Search_Combo.Text = "전체" Then
            StrSQL = "Select * FROM STOCK with(NOLOCK) Where PL_Name Like '%" & Search_Text.Text & "%' Order By PL_Code "
        Else
            StrSQL = "Select * FROM STOCK with(NOLOCK) Where PL_Name Like '%" & Search_Text.Text & "%' And PL_SEL = '" & Search_Combo.Text & "' Order By PL_Code "
        End If
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Main.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To 3
                    Grid_Main.Item(j, i).Value = DBT.Rows(i).Item(j).ToString
                Next j
            Next i
        Else
            Grid_Main.RowCount = 1
            For j = 0 To Grid_Main.ColumnCount - 1
                Grid_Main.Item(j, 0).Value = ""
            Next j
        End If
        Grid_Display = True
        Grid_Display_Ch = True

    End Function

    Private Sub Save_Com_Click(sender As Object, e As EventArgs) Handles Save_Com.Click
        Grid_Display_Ch = False

        Dim Table_Name As String
        Dim i As Integer
        Dim j As Long
        Dim k As Integer
        Dim Field_Name(500) As String
        Dim DBT As DataTable = Nothing
        Dim RRR As String
        j = 0

        '현재 일자
        Dim tDate As Date
        Dim tDateTime As Date
        tDate = Now()
        tDate = Format(Now(), "yyyy-MM-dd")
        tDateTime = Format(Now(), "yyyy-MM-dd HH:mm:ss")

        For i = 0 To Grid_Main.Rows.Count - 1
            If Grid_Main.Rows(i).HeaderCell.Value = "*" Then
                '재고조정
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE STOCK SET PL_QTY='" & Grid_Main.Item(4, i).Value & "' "
                StrSQL = StrSQL & "WHERE PL_CODE='" & Grid_Main.Item(1, i).Value & "'"
                Re_Count = DB_Execute()

                Dim result As String
                Dim res As Integer
                result = ""
                If Val(Grid_Main.Item(3, i).Value) - Val(Grid_Main.Item(4, i).Value) > 0 Then
                    res = Val(Grid_Main.Item(3, i).Value) - Val(Grid_Main.Item(4, i).Value)
                    result = "+" & res
                Else
                    res = Val(Grid_Main.Item(3, i).Value) - Val(Grid_Main.Item(4, i).Value)
                    result = "-" & res
                End If


                '조정내역추가
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "INSERT INTO STOCK_ADJUST "
                StrSQL = StrSQL & "values('" & DateTime.Now & "', '" & Grid_Main.Item(1, i).Value & "', '" & Grid_Main.Item(2, i).Value & "', '" & Grid_Main.Item(4, i).Value & "' ,'" & Grid_Main.Item(5, i).Value & "', '" & result & "' , '" & Frm_Main.Text_Man_Name.Text & "')"
                Re_Count = DB_Execute()
                Grid_Main.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        MsgBox("저장되었습니다")
        Grid_Display_Ch = True
    End Sub

    Private Sub Frm_Inven_Adj_Reg_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()
        Dim DBT As New DataTable

        Search_Combo.Items.Clear()
        Search_Combo.Items.Add("전체")

        StrSQL = "Select Base_Name  From SI_BASE_CODE with(NOLOCK) Where Base_Sel = '제품종류' Order By Base_Name"
        Re_Count = DB_Select(DBT)

        If Re_Count <> 0 Then
            For i = 0 To Re_Count - 1
                Search_Combo.Items.Add(DBT.Rows(i)("Base_Name"))
            Next i 'DBT.Rows(i)("Base_Bigo")
        End If
        Search_Combo.Text = "전체"

        Grid_Display_Ch = False
        Grid_Main.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        Grid_Setup()
        Grid_Display()

    End Sub
End Class
