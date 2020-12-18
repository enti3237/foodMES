Public Class Frm_stock1

    Dim Grid_Info_Row As Integer
    Dim Grid_Info_Col As Integer
    Dim Grid_Main_Row As Integer
    Dim Grid_Main_Col As Integer
    Dim lot_num As String
    Dim Grid_Display_Ch As Boolean

    Public Function Grid_Info_Display(Code As String) As Boolean
        Grid_Display_Ch = False
        Grid_Info_Display = True
        Dim DBT As New DataTable
        Dim i As Integer
        StrSQL = "Select * From MC_STOCK_IN With(NOLOCK) Where CS_Code = '" & Code & "'"
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

        Grid_Info_Setup()
        Grid_Main_Setup()
        Grid_Code_Display()
    End Sub

    Private Sub Frm_stock1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        print_tag.Visible = True
        print_tag.Enabled = True
        Search_Text.Visible = False
        Grid_Code_Setup()
        Grid_Code_Display()

        Grid_Info_Setup()
        Grid_Info_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)

        Grid_Main_Setup()

        Grid_Main_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)

    End Sub

    Public Function Grid_Code_Setup() As Boolean
        Grid_Color(Grid_Code)

        Grid_Code.AllowUserToAddRows = False
        Grid_Code.RowTemplate.Height = 50

        Grid_Code.ColumnCount = 3
        Grid_Code.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Code.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Code.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Code.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Code.ColumnHeadersHeight = 40
        Grid_Code.RowHeadersWidth = 10
        Grid_Code.Columns(0).Width = 100
        Grid_Code.Columns(1).Width = 100
        Grid_Code.Columns(2).Width = 100

        Grid_Code.Columns(0).HeaderText = "코드"
        Grid_Code.Columns(1).HeaderText = "날짜"
        Grid_Code.Columns(2).HeaderText = "거래처"

        Dim i As Integer
        For i = 0 To Grid_Code.ColumnCount - 1
            Grid_Code.Columns(i).ReadOnly = True
        Next i
        Grid_Code_Setup = True
    End Function

    Public Function Grid_Info_Setup() As Boolean

        Grid_Color(Grid_Info)


        Grid_Info.AllowUserToAddRows = False
        Grid_Info.RowTemplate.Height = 50
        Grid_Info.ColumnCount = 2
        Grid_Info.RowCount = 11

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Info.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Info.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Info.RowHeadersWidth = 10
        Grid_Info.Columns(0).Width = 100
        Grid_Info.Columns(1).Width = 240
        Grid_Info.Columns(0).HeaderText = "구분"
        Grid_Info.Columns(1).HeaderText = "내용"
        Grid_Info.ColumnHeadersHeight = 40

        Grid_Info.Item(0, 0).Value = "코드"
        Grid_Info.Item(0, 1).Value = "담당자"
        Grid_Info.Item(0, 2).Value = "날짜"
        Grid_Info.Item(0, 3).Value = "시간"

        Grid_Info.Item(0, 4).Value = "거래처코드"
        Grid_Info.Item(0, 5).Value = "거래처명"
        Grid_Info.Item(0, 6).Value = "담당자"
        Grid_Info.Item(0, 7).Value = "연락처"
        Grid_Info.Item(0, 8).Value = "입고일자"
        Grid_Info.Item(0, 9).Value = "수량"
        Grid_Info.Item(0, 10).Value = "비고"

        'Grid_Info.Columns(0).ReadOnly = True
        'Grid_Info.Columns(1).ReadOnly = True

        Grid_Info.Rows(0).ReadOnly = True
        Grid_Info.Rows(1).ReadOnly = True
        Grid_Info.Rows(2).ReadOnly = True
        Grid_Info.Rows(3).ReadOnly = True
        Grid_Info.Rows(9).ReadOnly = True
        Grid_Info.Rows(10).ReadOnly = True
        Grid_Info_Setup = True



    End Function
    Public Function Grid_Main_Setup() As Boolean


        Grid_Color(Grid_Main)
        Grid_Main.AllowUserToAddRows = False

        Grid_Main.RowTemplate.Height = 50
        Grid_Main.ColumnCount = 11
        Grid_Main.RowCount = 0

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Main.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Main.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Main.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Grid_Main.RowHeadersWidth = 10

        Grid_Main.ColumnHeadersHeight = 40
        Grid_Main.Columns(0).HeaderText = "순번"
        Grid_Main.Columns(1).HeaderText = "제품코드"
        Grid_Main.Columns(2).HeaderText = "제품명"
        Grid_Main.Columns(3).HeaderText = "규격"
        Grid_Main.Columns(4).HeaderText = "단위"
        Grid_Main.Columns(5).HeaderText = "단가"
        Grid_Main.Columns(6).HeaderText = "수량"
        Grid_Main.Columns(7).HeaderText = "금액"
        Grid_Main.Columns(8).HeaderText = "비고"
        Grid_Main.Columns(9).HeaderText = "입고장소"
        Grid_Main.Columns(10).HeaderText = "입고장소온도"

        Dim i As Integer

        Grid_Main.Columns(0).Width = 80
        Grid_Main.Columns(1).Width = 150
        Grid_Main.Columns(2).Width = 150

        For i = 3 To Grid_Main.ColumnCount - 1
            Grid_Main.Columns(i).Width = 180
        Next i


        Grid_Main_Setup = True
    End Function

    Public Function Grid_Code_Display() As Boolean

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Code.RowCount = 1


        StrSQL = "Select CS_Code, CS_Date, CS_Customer_Name From MC_STOCK_IN with(NOLOCK) ORDER BY CS_CODE DESC"

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
        Panel_Main.Visible = True



        Grid_Code_Display = True

        Grid_Code.MultiSelect = False
        Grid_Code.ClearSelection()



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



        StrSQL = "Select CS_Code From MC_STOCK_IN with(NOLOCK) Where CS_Date Like  '" & My_Date & "%' Order By CS_Code Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            JP_Code = JP_Code & "001"
        Else
            JP_Code = JP_Code & Format(Val(Mid(DBT.Rows(0).Item(0), 11, 3)) + 1, "00#")
        End If

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert into MC_STOCK_IN (CS_Code, CS_Date, CS_Time, CS_Name, CS_Day, CS_Check)  Values ('CS" & JP_Code & "', '" & My_Date & "','" & My_Time & "','태블릿','" & My_Date & "','' )"
        Re_Count = DB_Execute()
        Grid_Code_Display()
    End Sub

    Private Sub Delete_Com_Click(sender As Object, e As EventArgs) Handles Delete_Com.Click
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If



        If MsgBox("입고전표  '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "'를 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "입고전표 삭제") = vbNo Then
                Exit Sub
            End If


            '수량 변경
            '수량이 변경 되었는지 확인 한다.
            Dim Val_Check As String
            Dim PP_G_Sun As String
            Dim PP_G_Amount As String
            Dim DBT As DataTable = Nothing

            Val_Check = ""
            StrSQL = "Select CS_Check From MC_STOCK_IN With(NOLOCK) Where CS_Code = '" & Grid_Info.Item(1, 0).Value & "'"
            Re_Count = DB_Select(DBT)
            Grid_Info_Combo.Items.Clear()
            If Re_Count <> 0 Then
                If DBT.Rows(0)("CS_Check") = "처리" Then
                    Val_Check = "처리"
                End If
            End If


            If Val_Check = "처리" Then
                '기존 Data를 삭제 한다.
                'Grid의 수량을 가지고 온다.
                For i = 0 To Grid_Main.RowCount - 1
                    PP_G_Sun = "0"
                    PP_G_Amount = "0"
                    StrSQL = "Select PP_Sun, PP_Amount From SI_PROCESS_LIST With(NOLOCK) Where PP_Code = '" & Grid_Main.Item(1, i).Value & "' Order By Convert(Decimal,PP_Sun) "
                    Re_Count = DB_Select(DBT)
                    If Re_Count > 0 Then
                        PP_G_Sun = DBT.Rows(0)("PP_Sun")
                        PP_G_Amount = DBT.Rows(0)("PP_Amount")
                    End If
                    If PP_G_Sun = "0" Then
                    Else
                        '변경
                        PP_G_Amount = Val(PP_G_Amount) - Val(Grid_Main.Item(6, i).Value)
                        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                        StrSQL = StrSQL & "Update SI_PROCESS_LIST Set PP_Amount = '" & PP_G_Amount & "' Where PP_Code = '" & Grid_Main.Item(1, i).Value & "' And PP_Sun = '" & PP_G_Sun & "'"
                        Re_Count = DB_Execute()
                    End If
                Next i

            End If





            Grid_Display_Ch = False
            '삭제한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "Delete MC_STOCK_IN Where CS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' "
            Re_Count = DB_Execute()

            '삭제한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "Delete MC_STOCK_IN_List Where CS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' "
            Re_Count = DB_Execute()

            Grid_Code_Display()
            Grid_Info_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
    End Sub

    Private Sub Insert__Main_Click(sender As Object, e As EventArgs) Handles Insert__Main.Click
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        If Grid_Info.Item(1, 9).Value = "처리" Then
            MsgBox("수량 처리가 되어서 수정 할 수 없습니다.")
            Exit Sub
        End If

        Grid_Display_Ch = False

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



        'Dim DBT As New DataTable
        Dim Db_Sun As Long
        StrSQL = "Select CS_Sun From MC_STOCK_IN_List with(NOLOCK) Where CS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' Order By Convert(Decimal,CS_Sun) Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            Db_Sun = 1
        Else
            Db_Sun = DBT.Rows(0).Item(0) + 1
        End If


        Dim lot As String
        lot = Mid(JP_Code, 3)

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert into MC_STOCK_IN_List (CS_Code, CS_Sun)  Values ('" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "', '" & Db_Sun & "')"
        Re_Count = DB_Execute()
        Grid_Main_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)

        Grid_Display_Ch = True


    End Sub

    Private Sub Save_Main_Click(sender As Object, e As EventArgs) Handles Save_Main.Click
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        If Grid_Info.Item(1, 9).Value = "처리" Then
            MsgBox("수량 처리가 되어서 수정 할 수 없습니다.")
            Exit Sub
        End If

        If Grid_Info.Item(1, 10).Value = "처리" Then
            MsgBox("수량 처리가 되어서 수정 할 수 없습니다.")
            Exit Sub
        End If

        Dim i As Integer

        Grid_Display_Ch = False
        For i = 0 To Grid_Main.Rows.Count - 1
                If Grid_Main.Rows(i).HeaderCell.Value = "*" Then
                    'Update
                    StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "update MC_STOCK_IN_List Set CS_PL_Code = '" & Grid_Main.Item(1, i).Value & "', CS_PL_Name = '" & Grid_Main.Item(2, i).Value & "', CS_Standard = '" & Grid_Main.Item(3, i).Value & "', CS_Unit = '" & Grid_Main.Item(4, i).Value & "', CS_Unit_Amount = '" & Grid_Main.Item(5, i).Value & "', CS_Total = '" & Grid_Main.Item(6, i).Value & "', CS_Gold = '" & Grid_Main.Item(7, i).Value & "', CS_Bigo = '" & Grid_Main.Item(8, i).Value & "', CS_LOCATE = '" & Grid_Main.Item(9, i).Value & "', CS_Temp = '" & Grid_Main.Item(10, i).Value & "', LOT_Num = '" & lot_num & Grid_Main.Item(1, i).Value & "-' where CS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And CS_Sun = '" & Grid_Main.Item(0, i).Value & "'"
                Re_Count = DB_Execute()
                Grid_Main.Rows(i).HeaderCell.Value = ""
                End If
            Next i

        Grid_Display_Ch = True

    End Sub

    Private Sub Delete__Main_Click(sender As Object, e As EventArgs) Handles Delete__Main.Click
        '해당 칼럼 삭제
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        If Grid_Info.Item(1, 9).Value = "처리" Then
            MsgBox("수량 처리가 되어서 수정 할 수 없습니다.")
            Exit Sub
        End If

        Grid_Display_Ch = False


        Dim i As Integer



        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Delete MC_STOCK_IN_List Where CS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And  CS_Sun = '" & Grid_Main.Item(0, Grid_Main.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        For i = Grid_Main.CurrentCell.RowIndex + 1 To Grid_Main.RowCount - 1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "Update  MC_STOCK_IN_List Set CS_Sun = '" & i & "' Where CS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And  CS_Sun = '" & i + 1 & "'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        Grid_Main_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)

        Grid_Display_Ch = True

    End Sub

    Public Function Grid_Main_Display(PL_Code As String) As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display_Ch = False
        Grid_Main.RowCount = 0



        StrSQL = "Select CS_Sun, CS_PL_Code, CS_PL_Name, CS_Standard, CS_Unit, CS_Unit_Amount, CS_Total, CS_Gold, CS_Bigo, CS_Locate, CS_Temp  From MC_STOCK_IN_List with(NOLOCK)  Where CS_Code = '" & PL_Code & "' Order By Convert(Decimal,CS_Sun)"
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




        Grid_Info_Combo.Visible = False
        Grid_Display_Ch = True
        Grid_Main_Display = True

    End Function

    Private Sub Grid_Code_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Code.CellEnter
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
            Panel_Main.Visible = True
            Dim DBT As New DataTable
            Dim i As Integer
            StrSQL = "Select * FROM MC_STOCK_IN With(NOLOCK) Where CS_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "'"
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
            If Grid_Info.Item(1, 10).Value = "Y" Then
                print_tag.Enabled = False
            Else
                print_tag.Enabled = True
            End If
        End If
    End Sub

    Private Sub Grid_Info_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Info.CellEnter
        '선택되었을때 처리
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) < 1 Then
            Exit Sub
        End If
        'Car_Code 선택 처리
        Grid_Info_Row = Grid_Info.CurrentCell.RowIndex
        Grid_Info_Col = Grid_Info.CurrentCell.ColumnIndex
        Grid_Info_Combo.Visible = False
        If Grid_Info_Row < 0 Then
            Exit Sub
        End If
        If Grid_Info_Col < 0 Then
            Exit Sub
        End If

        Dim i As Integer
        Dim WW As Long
        Dim DBT As New DataTable



        If Grid_Info_Col = 1 Then
            Select Case Grid_Info_Row
                Case 4
                    StrSQL = "Select CM_Code  From SI_Customer With(NOLOCK) Order By CM_Code"
                    Re_Count = DB_Select(DBT)
                    Grid_Info_Combo.Items.Clear()
                    If Re_Count <> 0 Then
                        For i = 0 To Re_Count - 1
                            Grid_Info_Combo.Items.Add(DBT.Rows(i)("CM_Code"))
                        Next i
                    End If
                    Grid_Info_Combo.Top = Grid_Info.Top + Grid_Info.ColumnHeadersHeight + (Grid_Info_Row) * 50
                    WW = 0
                    For i = 0 To Grid_Info_Col - 1
                        WW = WW + Grid_Info.Columns(i).Width
                    Next
                    Grid_Info_Combo.Left = Grid_Info.Left + Grid_Info.RowHeadersWidth + WW
                    Grid_Info_Combo.Width = Grid_Info.Columns(Grid_Info_Col).Width
                    Grid_Info_Combo.Text = Grid_Info.CurrentCell.Value.ToString
                    Grid_Info_Combo.BackColor = Grid_Info.RowsDefaultCellStyle.BackColor
                    Grid_Info_Combo.Visible = True
                    Grid_Info_Combo.Focus.ToString()
                Case 5
                    StrSQL = "Select CM_Name  From SI_Customer With(NOLOCK) Order By CM_Name"
                    Re_Count = DB_Select(DBT)
                    Grid_Info_Combo.Items.Clear()
                    If Re_Count <> 0 Then
                        For i = 0 To Re_Count - 1
                            Grid_Info_Combo.Items.Add(DBT.Rows(i)("CM_Name"))
                        Next i
                    End If
                    Grid_Info_Combo.Top = Grid_Info.Top + Grid_Info.ColumnHeadersHeight + (Grid_Info_Row) * 50
                    WW = 0
                    For i = 0 To Grid_Info_Col - 1
                        WW = WW + Grid_Info.Columns(i).Width
                    Next
                    Grid_Info_Combo.Left = Grid_Info.Left + Grid_Info.RowHeadersWidth + WW
                    Grid_Info_Combo.Width = Grid_Info.Columns(Grid_Info_Col).Width
                    Grid_Info_Combo.Text = Grid_Info.CurrentCell.Value.ToString
                    Grid_Info_Combo.BackColor = Grid_Info.RowsDefaultCellStyle.BackColor
                    Grid_Info_Combo.Visible = True
                    Grid_Info_Combo.Focus.ToString()
                Case Else
                    Exit Sub
            End Select
        End If
    End Sub


    Private Sub print_tag_Click(sender As Object, e As EventArgs) Handles print_tag.Click
        Dim pc_stock_print_data As Pc_stock_print_data '출력할 자료를 저장할 구조체 배열
        Dim barcode_ArrList As New ArrayList()

        '체크된 그리드내용 ArrList에 저장하기
        Dim i As Integer
        For i = 0 To Grid_Main.RowCount - 1
            ' If Grid_Main.Item(9, i).Value = True Then
            pc_stock_print_data.psCode_sunCode = Grid_Info.Item(1, 0).Value & "-" & Grid_Main.Item(0, i).Value
                pc_stock_print_data.PS_PL_Name = Grid_Main.Item(2, i).Value
                pc_stock_print_data.PS_Standard = Grid_Main.Item(3, i).Value
                pc_stock_print_data.PS_Unit = Grid_Main.Item(4, i).Value
                pc_stock_print_data.PS_PO_Total = Grid_Main.Item(6, i).Value
                pc_stock_print_data.PS_DAY = Grid_Info.Item(1, 2).Value
                pc_stock_print_data.PS_Bigo = Grid_Main.Item(8, i).Value
                pc_stock_print_data.PS_CM = Grid_Info.Item(1, 5).Value
                pc_stock_print_data.PS_CODE = Grid_Info.Item(1, 0).Value
                barcode_ArrList.Add(pc_stock_print_data)
            ' End If
        Next

        If barcode_ArrList.Count = 0 Then
            MsgBox("출력할 행을 체크하신 후 출력버튼을 눌려주세요. 출력 할 행이 선택되지 않으면 실행되지 않습니다.", 0, "안내메세지")
        Else
            '출력할 바코드 list형태로 전달
            '  Dia_Stock_barcode_print.SetBarcode_arr(barcode_ArrList)
            '  Dia_Stock_barcode_print.Show()

        End If
    End Sub

    Public Structure Pc_stock_print_data
        Dim psCode_sunCode As String   '제품바코드 일련번호
        Dim PS_PL_Name As String        '제품명
        Dim PS_Standard As String       '규격
        Dim PS_Unit As String           '단위
        Dim PS_PO_Total As String       '생산량
        Dim PS_DAY As String            '생산일자
        Dim PS_Bigo As String           '비고
        Dim PS_CM As String             '고객
        Dim PS_CODE As String             '고객
    End Structure

    Private Sub Grid_Info_Combo_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Grid_Info_Combo.SelectionChangeCommitted
        Dim DBT As New DataTable
        Grid_Info.Item(Grid_Info_Col, Grid_Info_Row).Value = Grid_Info_Combo.SelectedItem.ToString()


        If Grid_Info_Row = 4 Then
            StrSQL = "Select CM_Code, CM_Name, CM_Man_Name, CM_Man_Tel  From SI_Customer With(NOLOCK) Where CM_Code = '" & Grid_Info_Combo.SelectedItem.ToString() & "'"
            Re_Count = DB_Select(DBT)
            Grid_Info_Combo.Items.Clear()
            If Re_Count <> 0 Then
                Grid_Info.Item(Grid_Info_Col, 5).Value = DBT.Rows(0)("CM_Name")
                Grid_Info.Item(Grid_Info_Col, 6).Value = DBT.Rows(0)("CM_Man_Name")
                Grid_Info.Item(Grid_Info_Col, 7).Value = DBT.Rows(0)("CM_Man_Tel")
                Grid_Info.Rows(4).HeaderCell.Value = "*"
                Grid_Info.Rows(5).HeaderCell.Value = "*"
                Grid_Info.Rows(6).HeaderCell.Value = "*"
                Grid_Info.Rows(7).HeaderCell.Value = "*"
            End If
        End If
        If Grid_Info_Row = 5 Then
            StrSQL = "Select CM_Code, CM_Name, CM_Man_Name, CM_Man_Tel  From SI_Customer With(NOLOCK) Where CM_Name = '" & Grid_Info_Combo.SelectedItem.ToString() & "'"
            Re_Count = DB_Select(DBT)
            Grid_Info_Combo.Items.Clear()
            If Re_Count <> 0 Then
                Grid_Info.Item(Grid_Info_Col, 4).Value = DBT.Rows(0)("CM_Code")
                Grid_Info.Item(Grid_Info_Col, 6).Value = DBT.Rows(0)("CM_Man_Name")
                Grid_Info.Item(Grid_Info_Col, 7).Value = DBT.Rows(0)("CM_Man_Tel")
                Grid_Info.Rows(4).HeaderCell.Value = "*"
                Grid_Info.Rows(5).HeaderCell.Value = "*"
                Grid_Info.Rows(6).HeaderCell.Value = "*"
                Grid_Info.Rows(7).HeaderCell.Value = "*"
            End If
        End If


        Grid_Info.Rows(Grid_Info_Row).HeaderCell.Value = "*"
        Grid_Info_Combo.Visible = False

    End Sub

    Private Sub Grid_Main_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Main.CellEnter
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
                StrSQL = "Select SI_Product.PL_Code, SI_Product.PL_Name  From SI_Product with(NOLOCK), SI_Product_Customer with(NOLOCK) Where SI_Product_Customer.PL_CM_Code = '" & Grid_Info.Item(1, 4).Value & "' And SI_Product_Customer.PL_Code = SI_Product.PL_Code And SI_Product_Customer.PL_Sel = '매입' Order By PL_Code"
                Re_Count = DB_Select(DBT)
                Grid_Main_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        Grid_Main_Combo.Items.Add(DBT.Rows(i).Item(0))
                    Next i
                End If
                Grid_Main_Combo.Top = Grid_Main.Top + Grid_Main.ColumnHeadersHeight + (Grid_Main_Row) * 50
                Grid_Main_Combo.Left = Grid_Main.Left + Grid_Main.RowHeadersWidth + Grid_Main.Columns(0).Width
                Grid_Main_Combo.Width = Grid_Main.Columns(Grid_Main_Col).Width
                Grid_Main_Combo.Text = Grid_Main.CurrentCell.Value.ToString
                Grid_Main_Combo.BackColor = Grid_Main.RowsDefaultCellStyle.BackColor
                Grid_Main_Combo.Visible = True
                Grid_Main_Combo.Focus.ToString()
            Case 2
                StrSQL = "Select SI_Product.PL_Code, SI_Product.PL_Name  From SI_Product with(NOLOCK), SI_Product_Customer with(NOLOCK) Where SI_Product_Customer.PL_CM_Code = '" & Grid_Info.Item(1, 4).Value & "' And SI_Product_Customer.PL_Code = SI_Product.PL_Code And SI_Product_Customer.PL_Sel = '매입' Order By PL_Name"
                Re_Count = DB_Select(DBT)
                Grid_Main_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        Grid_Main_Combo.Items.Add(DBT.Rows(i).Item(1))
                    Next i
                End If
                Grid_Main_Combo.Top = Grid_Main.Top + Grid_Main.ColumnHeadersHeight + (Grid_Main_Row) * 50
                Grid_Main_Combo.Left = Grid_Main.Left + Grid_Main.RowHeadersWidth + Grid_Main.Columns(0).Width + Grid_Main.Columns(1).Width
                Grid_Main_Combo.Width = Grid_Main.Columns(Grid_Main_Col).Width
                Grid_Main_Combo.Text = Grid_Main.CurrentCell.Value.ToString
                Grid_Main_Combo.BackColor = Grid_Main.RowsDefaultCellStyle.BackColor
                Grid_Main_Combo.Visible = True
                Grid_Main_Combo.Focus.ToString()

            Case 9
                StrSQL = "Select FL_Name  From SI_Stock with(NOLOCK)"
                Re_Count = DB_Select(DBT)
                Grid_Main_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        Grid_Main_Combo.Items.Add(DBT.Rows(i).Item(0))
                    Next i
                End If
                Grid_Main_Combo.Top = Grid_Main.Top + Grid_Main.ColumnHeadersHeight + (Grid_Main_Row) * 50
                Grid_Main_Combo.Left = Grid_Main.Left + Grid_Main.RowHeadersWidth + Grid_Main.Columns(0).Width + Grid_Main.Columns(1).Width + Grid_Main.Columns(2).Width + Grid_Main.Columns(3).Width + Grid_Main.Columns(4).Width + Grid_Main.Columns(5).Width + Grid_Main.Columns(6).Width + 15
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
                StrSQL = "Select PL_Code, PL_Name, PL_Standard ,PL_Unit, PL_Unit_Gold From SI_Product with(NOLOCK) Where PL_Code = '" & Grid_Main_Combo.SelectedItem.ToString() & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Main.Item(2, Grid_Main_Row).Value = DBT.Rows(0).Item(1)
                    Grid_Main.Item(3, Grid_Main_Row).Value = DBT.Rows(0).Item(2)
                    Grid_Main.Item(4, Grid_Main_Row).Value = DBT.Rows(0).Item(3)
                    Grid_Main.Item(5, Grid_Main_Row).Value = DBT.Rows(0).Item(4)

                End If
            Case 2
                StrSQL = "Select PL_Code, PL_Name, PL_Standard ,PL_Unit, PL_Unit_Gold  From SI_Product with(NOLOCK) Where PL_Name= '" & Replace(Grid_Main_Combo.SelectedItem.ToString(), "'", "''") & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Main.Item(1, Grid_Main_Row).Value = DBT.Rows(0).Item(0)
                    Grid_Main.Item(3, Grid_Main_Row).Value = DBT.Rows(0).Item(2)
                    Grid_Main.Item(4, Grid_Main_Row).Value = DBT.Rows(0).Item(3)
                    Grid_Main.Item(5, Grid_Main_Row).Value = DBT.Rows(0).Item(4)
                End If

            Case 9
                '창고위치
                StrSQL = "Select FL_Name  From SI_Stock with(NOLOCK) WHERE FL_NAME ='" & Grid_Main_Combo.SelectedItem.ToString() & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Main.Item(9, Grid_Main_Row).Value = DBT.Rows(0).Item(0)
                End If
        End Select
        Grid_Main_Combo.Visible = False
    End Sub

    Private Sub Save_Com_Click(sender As Object, e As EventArgs) Handles Save_Com.Click
        Dim Table_Name As String
        Dim j As Long
        Dim DBT As DataTable = Nothing
        Dim Field_Name(500) As String
        Dim i As Integer
        j = 0

        If Grid_Code.Item(0, 0).Value = "" Then
            MsgBox("공백은 저장할 수 없습니다")
            Exit Sub
        End If

        For i = 1 To Grid_Info.RowCount - 1
                If Grid_Info.Rows(i).HeaderCell.Value = "*" Then
                    j = 1
                End If
            Next i
            If j = 0 Then
            Else
                Table_Name = "MC_STOCK_IN"
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


            '수량 변경
            '수량이 변경 되었는지 확인 한다.
            Dim Val_Check As String
            Dim PP_G_Sun As String
            Dim PP_G_Amount As String

            Val_Check = ""
            StrSQL = "Select CS_Check From MC_STOCK_IN With(NOLOCK) Where CS_Code = '" & Grid_Info.Item(1, 0).Value & "'"
            Re_Count = DB_Select(DBT)
            Grid_Info_Combo.Items.Clear()
            If Re_Count <> 0 Then
                If DBT.Rows(0)("CS_Check") = "처리" Then
                    Val_Check = "처리"
                    MsgBox("이미 처리 되었습니다. 삭제 후 다시 처리 하세요")
                    Exit Sub
                End If
            End If


            '수량을 변경한다.
            For i = 0 To Grid_Main.RowCount - 1
                PP_G_Sun = "0"
                PP_G_Amount = "0"
                StrSQL = "Select PP_Sun, PP_Amount From SI_PROCESS_LIST With(NOLOCK) Where PP_Code = '" & Grid_Main.Item(1, i).Value & "' Order By Convert(Decimal,PP_Sun)"
                Re_Count = DB_Select(DBT)
                If Re_Count > 0 Then
                    PP_G_Sun = DBT.Rows(0)("PP_Sun")
                    PP_G_Amount = DBT.Rows(0)("PP_Amount")
                End If
                If PP_G_Sun = "0" Then
                Else
                    '변경
                    PP_G_Amount = Val(PP_G_Amount) + Val(Grid_Main.Item(6, i).Value)
                    StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                    StrSQL = StrSQL & "Update SI_PROCESS_LIST Set PP_Amount = '" & PP_G_Amount & "' Where PP_Code = '" & Grid_Main.Item(1, i).Value & "' And PP_Sun = '" & PP_G_Sun & "'"
                    Re_Count = DB_Execute()
                End If
            Next i

            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "Update MC_STOCK_IN Set CS_Check = '처리' Where CS_Code = '" & Grid_Info.Item(1, 0).Value & "'"
            Re_Count = DB_Execute()
            Grid_Info.Item(1, 9).Value = "처리"
    End Sub

    Private Sub Grid_Main_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Main.CellValueChanged

        'If Grid_Display_Ch = True Then
        '    Grid_Main.CurrentRow.HeaderCell.Value = "*"
        '    Grid_Main.Item(7, Grid_Main.CurrentCell.RowIndex).Value = Val(Grid_Main.Item(5, Grid_Main.CurrentCell.RowIndex).Value.ToString) * Val(Grid_Main.Item(6, Grid_Main.CurrentCell.RowIndex).Value.ToString)
        'End If
    End Sub

    Private Sub Grid_Info_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Info.CellValueChanged
        If Grid_Display_Ch = True Then
            Grid_Info.CurrentRow.HeaderCell.Value = "*"
        End If
    End Sub
End Class