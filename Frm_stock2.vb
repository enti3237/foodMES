Public Class Frm_stock2

    Dim Grid_Display_Ch As Boolean
    Dim Grid_Info_Row As Integer
    Dim Grid_Info_Col As Integer

    Dim Grid_Main_Row As Integer
    Dim Grid_Main_Col As Integer
    Dim Grid_Str(1000, 1000) As String
    Dim Je_ToT(1000, 1000) As String

    Dim Grid_Str_CM(1000, 1000) As String


    Dim Grid_Row As Integer
    Dim Grid_Col As Integer
    Dim Grid_Col2 As Integer

    Dim Product_Count_Cnt As Long

    Dim SEL_MC_CODE As String = "" '선택한 바코드 저장

    Public Function Grid_Code_Setup() As Boolean
        Grid_Color(Grid_Code)


        Grid_Code.ColumnHeadersHeight = 40
        Grid_Code.AllowUserToAddRows = False
        Grid_Code.RowTemplate.Height = 50
        Grid_Code.ColumnCount = 3
        Grid_Code.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Code.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Code.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Code.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Code.RowHeadersWidth = 10
        Grid_Code.Columns(0).Width = 100
        Grid_Code.Columns(1).Width = 120
        Grid_Code.Columns(2).Width = 120

        Grid_Code.Columns(0).HeaderText = "코드"
        Grid_Code.Columns(1).HeaderText = "일자"
        Grid_Code.Columns(2).HeaderText = "담당자"

        Dim i As Integer
        For i = 0 To Grid_Code.ColumnCount - 1
            Grid_Code.Columns(i).ReadOnly = True
        Next i
        Grid_Code_Setup = True
    End Function

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        Grid_Code_Display()
    End Sub

    Public Function Grid_Code_Display() As Boolean

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Code.RowCount = 1

        StrSQL = "SELECT [MC_CODE], [MC_DATE], [MC_MAN]
                          FROM [MC_STOCK_MOVE] with(NOLOCK)
                          Where [MC_CODE] Like '%" & Search_Text.Text & "%'  Order  By [MC_CODE] DESC"

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
        Grid_Code.ClearSelection()
    End Function

    Private Sub Frm_stock2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()

        Barcode_input_txtbox.TabStop = True
        Barcode_input_txtbox.TabIndex = 0
        Barcode_input_txtbox.Select()


        Button14.Visible = True
        Barcode_input_txtbox.Visible = True
        Search_Text.Visible = False
        Save_Com.Visible = False
        Insert_Com.Visible = False
        Grid_Code_Setup()
        Grid_Info_Setup()
        Grid_Main_Setup()
    End Sub

    Public Function Grid_Info_Setup() As Boolean
        Grid_Color(Grid_Info)

        Grid_Info.AllowUserToAddRows = False
        Grid_Info.RowTemplate.Height = 50
        Grid_Info.ColumnCount = 2
        Grid_Info.RowCount = 5
        Grid_Info.ColumnHeadersHeight = 40
        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Info.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Info.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Info.RowHeadersWidth = 10
        Grid_Info.Columns(0).Width = 100
        Grid_Info.Columns(1).Width = 240
        Grid_Info.Columns(0).HeaderText = "구분"
        Grid_Info.Columns(1).HeaderText = "내용"

        Grid_Info.Item(0, 0).Value = "코드"
        Grid_Info.Item(0, 1).Value = "일자"
        Grid_Info.Item(0, 2).Value = "시간"
        Grid_Info.Item(0, 3).Value = "담당자"
        Grid_Info.Item(0, 4).Value = "비고"

        Dim i As Integer

        '모든 행 일기전용
        For i = 0 To 3
            Grid_Info.Rows(i).ReadOnly = True
        Next i

        Grid_Info_Setup = True
    End Function

    Public Function Grid_Main_Setup() As Boolean
        Grid_Color(Grid_Main)

        Grid_Main.AllowUserToAddRows = False
        Grid_Main.RowTemplate.Height = 50
        Grid_Main.ColumnCount = 10
        Grid_Main.RowCount = 0
        Grid_Main.ColumnHeadersHeight = 40
        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Main.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Grid_Main.RowHeadersWidth = 10


        Grid_Main.Columns(0).HeaderText = "순번"
        Grid_Main.Columns(1).HeaderText = "입고코드"
        Grid_Main.Columns(2).HeaderText = "제품코드"
        Grid_Main.Columns(2).Visible = False
        Grid_Main.Columns(3).HeaderText = "제품명"
        Grid_Main.Columns(3).Visible = False
        Grid_Main.Columns(4).HeaderText = "규격"
        Grid_Main.Columns(4).Visible = False
        Grid_Main.Columns(5).HeaderText = "단위"
        Grid_Main.Columns(5).Visible = False
        Grid_Main.Columns(6).HeaderText = "입고수량"
        Grid_Main.Columns(6).Visible = False
        Grid_Main.Columns(7).HeaderText = "실제출고수량"
        Grid_Main.Columns(8).HeaderText = "출고장소"
        Grid_Main.Columns(9).HeaderText = "비고"


        Dim i As Integer

        Grid_Main.Columns(0).Width = 80
        Grid_Main.Columns(1).Width = 250
        Grid_Main.Columns(2).Width = 200

        For i = 3 To 9
            Grid_Main.Columns(i).Width = 150
        Next

        Grid_Main.Columns(7).Width = 200

        Grid_Main_Setup = True
    End Function

    Private Sub Save_Com_Click(sender As Object, e As EventArgs) Handles Save_Com.Click

    End Sub

    Private Sub Insert_Com_Click(sender As Object, e As EventArgs) Handles Insert_Com.Click

    End Sub

    Private Sub Delete_Com_Click(sender As Object, e As EventArgs) Handles Delete_Com.Click

    End Sub

    Public Function Grid_Info_Display(Code As String) As Boolean
        Grid_Display_Ch = False
        Grid_Info_Display = True
        Dim DBT As New DataTable
        Dim i As Integer
        StrSQL = "Select * From MC_STOCK_MOVE With(NOLOCK) Where MC_Code = '" & Code & "'"
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

    Public Function Grid_Main_Display(MC_CODE As String) As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Dim cs_code As String
        Dim cs_sun As String
        Dim m_code As String
        Grid_Display_Ch = False

        Grid_Main.RowCount = 0
        m_code = MC_CODE


        StrSQL = "SELECT MC_SUN, MC_IN_CODE, MC_NUM, MC_LOCATE, MC_BIGO
                    FROM MC_STOCK_MOVE_LIST with(NOLOCK)
                    WHERE MC_CODE = '" & MC_CODE & "'
                    ORDER BY MC_CODE DESC"

        'StrSQL = "Select CL_Data From MATERIAL_IN_Plan_List with(NOLOCK), Product_Customer with(NOLOCK)  Where CL_Code = '" & PL_Code & "' Order By Convert(Decimal,CL_Sun)"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then

            Grid_Main.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To 1
                    Grid_Main.Item(j, i).Value = DBT.Rows(i).Item(j).ToString
                    Grid_Main.Item(7, i).Value = DBT.Rows(i).Item(2).ToString
                    Grid_Main.Item(8, i).Value = DBT.Rows(i).Item(3).ToString
                    Grid_Main.Item(9, i).Value = DBT.Rows(i).Item(4).ToString
                Next j
            Next i
        Else
            Grid_Main.RowCount = 0
            'For j = 0 To Grid_Main.ColumnCount - 1
            '    Grid_Main.Item(j, 0).Value = ""
            'Next j
        End If

        'cs_code = "CS" & Mid(MC_CODE, 3, 11)
        'cs_sun = Mid(MC_CODE, 15)
        '' MsgBox(cs_code)
        '' MsgBox(cs_sun)

        'StrSQL = "SELECT CS_PL_CODE, CS_PL_NAME, CS_STANDARD, CS_UNIT, CS_TOTAL FROM MC_STOCK_IN_LIST with(NOLOCK) WHERE CS_CODE = '" & cs_code & "' AND CS_SUN = '" & cs_sun  &"' ORDER BY CS_CODE DESC"

        '' StrSQL = "Select CL_Data From MATERIAL_IN_Plan_List with(NOLOCK), Product_Customer with(NOLOCK)  Where CL_Code = '" & PL_Code & "' Order By Convert(Decimal,CL_Sun)"
        'Re_Count = DB_Select(DBT)
        'If Re_Count <> 0 Then

        '    ' Grid_Main.RowCount = Re_Count
        '    Dim k As Integer


        '    For i = 0 To Grid_Main.RowCount - 1
        '            For j = 0 To DBT.Columns.Count - 1
        '                Grid_Main.Item(j + 2, i).Value = DBT.Rows(i).Item(j).ToString
        '            Next j
        '        Next i



        'Else
        '    ' Grid_Main.RowCount = 0
        '    'For j = 0 To Grid_Main.ColumnCount - 1
        '    '    Grid_Main.Item(j, 0).Value = ""
        '    'Next j
        'End If


        '그리드 읽기전용으로 설정
        For i = 0 To 6
            Grid_Main.Columns(i).ReadOnly = True
        Next i

        Grid_Display_Ch = True
        Grid_Main_Display = True

    End Function

    Private Sub Barcode_input_txtbox_KeyDown(sender As Object, e As KeyEventArgs) Handles Barcode_input_txtbox.KeyDown
        If e.KeyCode = Keys.Return Then
            Dim DBT As New DataTable
            If Barcode_input_txtbox.Text = "" Then
                Exit Sub
            End If

            Dim barcode As String
            barcode = Mid(Barcode_input_txtbox.Text, 1, 13)

            ' 입고코드에서 확인
            StrSQL = "Select CS_CODE From MC_STOCK_IN with(NOLOCK) WHERE CS_CODE = '" & barcode & "'Order  By [CS_CODE] DESC"
            Re_Count = DB_Select(DBT)

            If Re_Count <> 1 Then
                MsgBox("존재하지 않는 바코드 입니다.")
                Barcode_input_txtbox.Text = ""
                Exit Sub
                ' MsgBox(barcode)
            End If

            Dim B As String
            B = "MC" & Mid(barcode, 3, 15) & Mid(Barcode_input_txtbox.Text, 14)
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



            StrSQL = "Select MC_CODE From MC_STOCK_MOVE with(NOLOCK) WHERE MC_CODE = '" & B & "'Order  By [MC_CODE] DESC"
            Re_Count = DB_Select(DBT)

            If Re_Count = 0 Then
                '입고코드가 출고된 전적이 없음

                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert into MC_STOCK_MOVE (MC_Code, MC_DATE, MC_TIME, MC_MAN)  Values ('" & B & "', '" & My_Date & "','" & My_Time & "', '태블릿')"
                Re_Count = DB_Execute()

                StrSQL = "Set TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert into MC_STOCK_MOVE_LIST (MC_Code, MC_SUN, MC_IN_CODE, MC_DATE, MC_TIME)  Values ('" & B & "', '1', '" & Barcode_input_txtbox.Text & "', ' " & My_Date & "','" & My_Time & "')"
                Re_Count = DB_Execute()
                Dim mc As String = "MC" & Barcode_input_txtbox.Text

                Grid_Code_Display()
                Grid_Info_Display(mc)
                Grid_Main_Display(mc)

                MsgBox("바코드번호가 " & Barcode_input_txtbox.Text & " :  출고처리 되었습니다. 출고수량과 위치를 확인해주세요")
                Barcode_input_txtbox.Text = ""
            Else


                'Dim DBT As New DataTable
                Dim Db_Sun As Long
                StrSQL = "Select MC_SUN From MC_STOCK_MOVE_LIST with(NOLOCK) Where MC_Code = '" & B & "' Order By Convert(Decimal,MC_SUN) Desc "
                Re_Count = DB_Select(DBT)
                If Re_Count = 0 Then
                    Db_Sun = 1
                Else
                    Db_Sun = DBT.Rows(0).Item(0) + 1
                End If
                'MsgBox("이미 출고처리 되었습니다.")
                'Barcode_input_txtbox.Text = ""
                'Exit Sub

                StrSQL = "Set TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert into MC_STOCK_MOVE_LIST (MC_Code, MC_SUN, MC_IN_CODE, MC_DATE, MC_TIME)  Values ('" & B & "', '" & Db_Sun & "', '" & Barcode_input_txtbox.Text & "', ' " & My_Date & "','" & My_Time & "')"
                Re_Count = DB_Execute()
                Dim mc As String = "MC" & Barcode_input_txtbox.Text

                Grid_Code_Display()
                Grid_Info_Display(mc)
                Grid_Main_Display(mc)

                MsgBox("바코드번호가 " & Barcode_input_txtbox.Text & " :  출고처리 되었습니다. 출고수량과 위치를 확인해주세요")
                Barcode_input_txtbox.Text = ""
            End If

            ' 추가한다

            '    Grid_Main_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)


        End If

        '  입력한 바코드가 창고 바코드인지 확인

        'For i = 0 To Re_Count - 1
        '    If Barcode_input_txtbox.Text = DBT.Rows(i).Item(0) Then
        '        Grid_Code_Display()
        '        Grid_Code.Rows(i).Selected = True
        '        Grid_Main_Display(Grid_Code.Item(0, i).Value)
        '        SEL_MC_CODE = Barcode_input_txtbox.Text
        '        Barcode_input_txtbox.Text = ""
        '        Exit Sub
        '    End If
        'Next i

        'Dim code_sun_ARR() As String = Split(Barcode_input_txtbox.Text, "-")
        ''   제품 바코드 검증
        'If code_sun_ARR.Length < 2 Then
        '    MsgBox("입력한 " & Barcode_input_txtbox.Text & "정상적인 바코드 값이 아닙니다.")
        'End If

        ''입력한 바코드가 현재선택된 위치의 보관내역에 포함되어 있는지 확인
        'For i = 0 To Grid_Main.Rows.Count - 1
        '    If Barcode_input_txtbox.Text = Grid_Main.Item(1, i).Value Then
        '        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED 
        '                      Update MC_STOCK_IN SET MOLD_LOCATION = NULL 
        '                      WHERE CS_CODE = '" & Grid_Main.Item(2, i).Value & "'"
        '        Re_Count = DB_Execute()


        '        MsgBox("바코드번호가 " & Barcode_input_txtbox.Text & "인 금형이 입고처리 되었습니다.")


        '        Barcode_input_txtbox.Text = ""
        '        Grid_Main_Display(SEL_MC_CODE)
        '        Exit Sub
        '    End If
        'Next

        ''입력한 바코드가 새로운 내용일 경우
        'StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED 
        '              Update MOLD SET MOLD_LOCATION = '" & SEL_MC_CODE & "'
        '              WHERE MOLD_LOT = '" & Barcode_input_txtbox.Text & "'"
        'Re_Count = DB_Execute()
        'If Re_Count <> 0 Then '변경처리가 되었으면 -1오류코드가 아니면 완료 메세지 출력 함수 종료 



        '    MsgBox("바코드번호가 " & Barcode_input_txtbox.Text & "인 제품이 출고처리 되었습니다.")



        '    Barcode_input_txtbox.Text = ""
        '    Grid_Main_Display(SEL_MC_CODE)
        '    Exit Sub
        'End If


        '  MsgBox("위치나 설비를 선택한 후 바코드를 입력해서 처리 해주십시오. 존재하지 않는 바코드 입니다.")



    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Grid_Code_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Code.CellEnter
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
            Panel_Main.Visible = True
            Dim DBT As New DataTable
            Dim i As Integer
            StrSQL = "Select * FROM MC_STOCK_MOVE With(NOLOCK) Where MC_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "'"
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
            Label_Color(Com_Main, Color_Label, Di_Panel2)
            Grid_Main_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
            Panel_Main.Visible = True

        End If
    End Sub

    Private Sub SAVE_BTN_Click(sender As Object, e As EventArgs) Handles SAVE_BTN.Click
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        Dim i As Integer

        Grid_Display_Ch = False
        For i = 0 To Grid_Main.Rows.Count - 1
            If Grid_Main.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE MC_STOCK_MOVE_LIST Set MC_NUM = '" & Grid_Main.Item(7, i).Value & "', MC_LOCATE = '" & Grid_Main.Item(8, i).Value & "',  MC_BIGO = '" & Grid_Main.Item(9, i).Value & "' 
                                        where MC_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And MC_Sun = '" & Grid_Main.Item(0, i).Value & "'"
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

    Private Sub DEL_BTN_Click(sender As Object, e As EventArgs) Handles DEL_BTN.Click
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        Grid_Display_Ch = False
        Dim i As Integer
        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE MC_STOCK_MOVE_LIST Where MC_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And  MC_Sun = '" & Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        For i = Grid_Main.CurrentCell.RowIndex + 1 To Grid_Main.RowCount - 1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "Update  MC_STOCK_MOVE_LIST Set MC_Sun = '" & i & "' Where MC_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And  MC_Sun = '" & i + 1 & "'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        Grid_Main_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True
    End Sub

    Private Sub Grid_Main_Combo_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Grid_Main_Combo.SelectionChangeCommitted
        Dim DBT As New DataTable
        Grid_Main.Item(Grid_Main_Col, Grid_Main_Row).Value = Grid_Main_Combo.SelectedItem.ToString()
        Select Case Grid_Main_Col
            Case 8
                StrSQL = "Select FL_NAME FROM SI_STOCK with(NOLOCK) Where FL_NAME = '" & Grid_Main_Combo.SelectedItem.ToString() & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Main.Item(8, Grid_Main_Row).Value = DBT.Rows(0).Item(0)
                End If
        End Select

        Grid_Main_Combo.Visible = False
    End Sub

    Private Sub Grid_Main_Click(sender As Object, e As EventArgs) Handles Grid_Main.Click
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
            Case 8
                StrSQL = "Select FL_NAME FROM SI_STOCK with(NOLOCK) Order By FL_Code"
                Re_Count = DB_Select(DBT)
                Grid_Main_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        Grid_Main_Combo.Items.Add(DBT.Rows(i).Item(0))
                    Next i
                End If
                Grid_Main_Combo.Top = Grid_Main.ColumnHeadersHeight + (Grid_Main_Row) * 50 + 250
                Grid_Main_Combo.Left = Grid_Main.Left + Grid_Main.RowHeadersWidth + Grid_Main.Columns(0).Width + Grid_Main.Columns(1).Width + Grid_Main.Columns(2).Width + 390
                Grid_Main_Combo.Width = Grid_Main.Columns(Grid_Main_Col).Width
                ' Grid_Main_Combo.Text = Grid_Main.CurrentCell.Value.ToString
                Grid_Main_Combo.Text = ""
                Grid_Main_Combo.BackColor = Grid_Main.RowsDefaultCellStyle.BackColor
                Grid_Main_Combo.Visible = True
                Grid_Main_Combo.Focus.ToString()

        End Select
    End Sub
End Class