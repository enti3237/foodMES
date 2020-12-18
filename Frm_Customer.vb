Public Class Frm_Customer
    Dim Grid_Display_Ch As Boolean
    Dim Info_Lab() As Button
    Dim Info_Txt() As TextBox
    Dim Info_Com() As ComboBox

    Dim Dis_Check As String


    Private Sub Frm_Customer_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()

        'Info_Lab = New Button() {Info_La0, Info_La1, Info_La2, Info_La3, Info_La4, Info_La5, Info_La6, Info_La7, Info_La8, Info_La9, Info_La10, Info_La11, Info_La12, Info_La13, Info_La14, Info_La15, Info_La16, Info_La17, Info_La18, Info_La19, Info_La20, Info_La21, Info_La22, Info_La23}
        'Info_Txt = New TextBox() {Info_Tx0, Info_Tx1, Info_Tx2, Info_Tx3, Info_Tx4, Info_Tx5, Info_Tx6, Info_Tx7, Info_Tx8, Info_Tx9, Info_Tx10, Info_Tx11, Info_Tx12, Info_Tx13, Info_Tx14, Info_Tx15, Info_Tx16, Info_Tx17, Info_Tx18, Info_Tx19, Info_Tx20, Info_Tx21, Info_Tx22, Info_Tx23}

        Grid_Display_Ch = False
        ' Search_Combo.Items.Add("코드")
        ' Search_Combo.Items.Add("업체명")
        ' Search_Combo.Text = "업체명"

        Grid_Display_Ch = False

        Grid_Setup()
        Grid_Display()
        Customer_Grid.ClearSelection()

    End Sub
    Public Function Grid_Setup() As Boolean
        Dim i As Integer
        Grid_Color(Customer_Grid)
        Customer_Grid.EnableHeadersVisualStyles = False
        Customer_Grid.AllowUserToAddRows = False
        Customer_Grid.RowTemplate.Height = 20
        Customer_Grid.ColumnCount = 14
        Customer_Grid.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Customer_Grid.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
        Customer_Grid.RowHeadersWidth = 45
        Customer_Grid.Columns(0).Width = 55

        For i = 1 To Customer_Grid.ColumnCount - 1
            Customer_Grid.Columns(i).Width = 100
            Customer_Grid.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Next i
        Customer_Grid.Columns(13).Width = 450
        Customer_Grid.Columns(0).HeaderText = "코드"
        Customer_Grid.Columns(1).HeaderText = "업체명"
        Customer_Grid.Columns(2).HeaderText = "구분(매입/매출)"
        Customer_Grid.Columns(3).HeaderText = "사업자번호"
        Customer_Grid.Columns(4).HeaderText = "법인번호"
        Customer_Grid.Columns(5).HeaderText = "대표자"
        Customer_Grid.Columns(6).HeaderText = "우편번호"
        Customer_Grid.Columns(7).HeaderText = "주소"
        Customer_Grid.Columns(8).HeaderText = "전화번호"
        Customer_Grid.Columns(9).HeaderText = "팩스번호"
        Customer_Grid.Columns(10).HeaderText = "담당자이름"
        Customer_Grid.Columns(11).HeaderText = "담당자전화"
        Customer_Grid.Columns(12).HeaderText = "담당자메일"
        Customer_Grid.Columns(13).HeaderText = "비고"

        Dim k As Integer
        For k = 0 To Customer_Grid.ColumnCount - 1
            ' Info_Lab(k).Text = Customer_Grid.Item(0, k).Value
            'Info_Lab(k).Text = Customer_Grid.Columns(k).HeaderText
        Next k


        Customer_Grid.Columns(0).ReadOnly = True
        Grid_Setup = True
    End Function
    Public Function Grid_Display() As Boolean
        Grid_Display_Ch = False

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Customer_Grid.RowCount = 1
        StrSQL = "Select * FROM SI_CUSTOMER with(NOLOCK) Where CM_Name Like '%" & Search_Text.Text & "%'  Or CM_Name is Null Order  By CM_Name"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Customer_Grid.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To Customer_Grid.ColumnCount - 1
                    Customer_Grid.Item(j, i).Value = DBT.Rows(i).Item(j)
                Next j
            Next i
        Else
            Customer_Grid.RowCount = 1
            For j = 0 To Customer_Grid.ColumnCount - 1
                Customer_Grid.Item(j, 0).Value = ""
            Next j
        End If
        Grid_Display = True
        Grid_Display_Ch = True

    End Function

    Private Sub Insert_Com_Click(sender As Object, e As EventArgs) Handles Insert_Com.Click
        Dim form As New customer
        form.Label1.Text = "insert"
        form.ShowDialog()

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If
        'customer.Show()
        'Search_Com.PerformClick()
    End Sub

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        Grid_Display()

    End Sub

    Public Function Customer_Grid_Display(Grid As DataGridView, SQL_String As String, Info_Text() As TextBox) As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        'StrSQL = SQL_String

        '''MessageBox.Show(StrSQL)
        ''If Grid_Display_Ch = False Then
        ''    Exit Function


        'Re_Count = DB_Select(DBT)
        'If Re_Count > 0 Then
        '    For i = 0 To DBT.Columns.Count - 1
        '        Grid.Item(i, 0).Value = DBT.Rows(0).Item(i).ToString
        '    Next i
        'Else
        '    For i = 0 To DBT.Columns.Count - 1
        '        Grid.Item(i, 0).Value = ""
        '    Next i
        'End If


        For i = 0 To 23
            'If Len(Grid.Item(i, Grid.CurrentCell.RowIndex).Value) < 0 Then
            '    Grid.Item(i, Grid.CurrentCell.RowIndex).Value = ""
            'End If
            'Info_Text(i).Text = Grid.Item(i, Grid.CurrentCell.RowIndex).Value


            'Info_Combo(i).Text = Grid.Item(1, i).Value
        Next i

        Customer_Grid_Display = True
    End Function


    Private Sub Customer_Grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Customer_Grid.CellValueChanged
        If Grid_Display_Ch = True Then
            sender.CurrentRow.HeaderCell.Value = "*"
        End If

    End Sub

    Private Sub Save_Com_Click(sender As Object, e As EventArgs) Handles Save_Com.Click
        Grid_Display_Ch = False

        Dim Table_Name As String
        Dim i As Integer
        Dim j As Long
        Dim k As Integer
        Dim Field_Name(500) As String
        Dim DBT As DataTable = Nothing

        Table_Name = "SI_Customer"
        j = 0
        StrSQL = "Select sys.Columns.Name From sys.Tables with(NOLOCK) , sys.Columns With(NOLOCK) Where sys.Tables.name = '" & Table_Name & "' And sys.Tables.Object_id = sys.Columns.Object_id Order By sys.Columns.column_id"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            For i = 0 To DBT.Rows.Count - 1
                Field_Name(j) = DBT.Rows(i)("Name").ToString
                j = j + 1
            Next i
        End If
        j = j - 1
        Dim RRR As String

        For i = 0 To Customer_Grid.Rows.Count - 1
            If Customer_Grid.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE " & Table_Name & " SET "
                For k = 1 To j
                    'StrSQL = StrSQL & " " & Table_Name & "." & Field_Name(k) & " = '" & Replace(Customer_Grid.Item(k, i).Value.ToString, "'", "''") & "'"
                    'StrSQL = StrSQL & " " & Table_Name & "." & Field_Name(k) & " = '" & Trim(Customer_Grid.Item(k, i).Value.ToString) & "'" 'IIf(IsNull(rs!user_tel), "", rs!user_tel)
                    'Trim(IIf(Len(DBRs![Go_Total]) <> 0, DBRs![Go_Total], "0"))
                    'StrSQL = StrSQL & " " & Table_Name & "." & Field_Name(k) & " = '" & Replace(IIf(Len(Customer_Grid.Item(k, i).Value.ToString) <> 0, Customer_Grid.Item(k, i).Value.ToString, ""), "'", "''") & "'"
                    'RRR = IIf(IsDBNull(Customer_Grid.Item(k, i).Value.ToString), "", Customer_Grid.Item(k, i).Value.ToString)
                    RRR = IIf(IsDBNull(Customer_Grid.Item(k, i).Value), "", Customer_Grid.Item(k, i).Value)
                    StrSQL = StrSQL & " " & Table_Name & "." & Field_Name(k) & " = '" & Replace(RRR, "'", "''") & "'"

                    If k <> j Then
                        StrSQL = StrSQL & ","
                    End If
                Next k
                StrSQL = StrSQL & " WHERE " & Table_Name & "." & Field_Name(0) & " = '" & Customer_Grid.Item(0, i).Value.ToString & "'"
                Re_Count = DB_Execute()
                Customer_Grid.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display_Ch = True

        MsgBox("저장되었습니다")


    End Sub

    Private Sub Delete_Com_Click(sender As Object, e As EventArgs) Handles Delete_Com.Click
        '활성화
        '해당 칼럼 삭제
        If Len(Customer_Grid.Item(0, Customer_Grid.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        If MsgBox("거래처 '" & Customer_Grid.Item(1, Customer_Grid.CurrentCell.RowIndex).Value & "'을(를) 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "삭제") = vbNo Then
            Exit Sub
        End If


        Grid_Display_Ch = False
        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_CUSTOMER Where CM_Code = '" & Customer_Grid.Item(0, Customer_Grid.CurrentCell.RowIndex).Value & "' "
        Re_Count = DB_Execute()

        '선택된 갈럼값을 가지고 온다
        Grid_Display()
        Grid_Display_Ch = True


    End Sub

    'Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    '    Dim Li_Data As String
    '    Dim D_Col As Integer
    '    Dim NullStr(1000) As String
    '    Dim AD1 As String
    '    Dim AD2 As String
    '    Dim i As Integer
    '    Dim j As Integer

    '    AD1 = ""
    '    AD2 = ""

    '    FileOpen(1, Application.StartupPath + "\거래처등록.csv", OpenMode.Input)
    '    For i = 1 To 40
    '        Li_Data = LineInput(1)
    '        D_Col = Data_Col_Count_Str(Li_Data, NullStr)
    '        For j = 20 To Len(NullStr(4))
    '            If Mid(NullStr(4), j, 1) = " " Or Mid(NullStr(4), j, 1) = "(" Then
    '                AD1 = Mid(NullStr(4), 1, j - 1)
    '                AD2 = Mid(NullStr(4), j, 250)
    '                Exit For

    '            End If
    '        Next j
    '        If Len(AD1) < 1 Then
    '            AD1 = NullStr(4)
    '            AD2 = ""
    '        End If
    '        '추가한다.
    '        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
    '        StrSQL = StrSQL & "Insert INTO SI_CUSTOMER (CM_Code, CM_Op_Number, CM_Name, CM_Leader, CM_Add1, CM_Add2, CM_Conditions, CM_Item,CM_Man_Name,CM_Tel,CM_Man_Tel,CM_Fax,CM_Man_Email )  Values ('" & 10000 + i & "','" & NullStr(1) & "','" & NullStr(2) & "','" & NullStr(3) & "','" & AD1 & "','" & AD2 & "','" & NullStr(5) & "','" & NullStr(6) & "','" & NullStr(8) & "','" & NullStr(9) & "','" & NullStr(10) & "' ,'" & NullStr(11) & "' ,'" & NullStr(12) & "' )"
    '        Re_Count = DB_Execute()
    '    Next i


    '    FileClose(1)
    'End Sub



    Private Sub Info_Tx0_TextChanged(sender As Object, e As EventArgs)
        Dim index As Integer

        If Grid_Display_Ch = False Then
            Exit Sub

        Else
            'MsgBox(Mid(sender.name.ToString, 8, 2))
        End If

        index = Val(Mid(sender.name.ToString, 8, 2))

        Customer_Grid.Item(index, Customer_Grid.CurrentCell.RowIndex).Value = sender.text

        ' Info_Com(index).Text = sender.text
    End Sub

    Private Sub Customer_Grid_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles Customer_Grid.CellEnter


        Customer_Grid_Display(Customer_Grid, "", Info_Txt)


    End Sub

    Private Sub Customer_Grid_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Customer_Grid.CellDoubleClick
        '유효성 검사
        If Len(Customer_Grid.Item(0, Customer_Grid.CurrentCell.RowIndex).Value) < 0 Then
            MsgBox("해당 행은 공백이므로 처리할 수 없습니다")
            Exit Sub
        End If

        '해당 행 내용 수정
        Product_Insert.Label1.Text = "update"
        Product_Insert.product_code.Text = Customer_Grid.Item(0, Customer_Grid.CurrentCell.RowIndex).Value
        ' Product_Insert.Show()
        Dim form As New customer
        form.Label1.Text = "update"
        form.Info_Tx0.Text = Customer_Grid.Item(0, Customer_Grid.CurrentCell.RowIndex).Value
        form.ShowDialog()

        '재검색
        If form.DialogResult = DialogResult.OK Then
            Search_Com.PerformClick()
        Else
            Search_Com.PerformClick()
        End If
    End Sub
End Class
