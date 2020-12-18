﻿Imports System.IO

Public Class Frm_Product2
    Dim Grid_Display_Ch As Boolean

    Dim Grid_Info_Row As Integer
    Dim Grid_Info_Col As Integer


    Dim Grid_Bom_Row As Integer
    Dim Grid_Bom_Col As Integer

    Dim Grid_Pro_Row As Integer
    Dim Grid_Pro_Col As Integer

    Dim Product_Line_QC_Row As Integer
    Dim Product_Line_QC_Col As Integer

    Dim Grid_Cus_In_Row As Integer
    Dim Grid_Cus_In_Col As Integer
    Dim Grid_Cus_Out_Row As Integer
    Dim Grid_Cus_Out_Col As Integer


    Dim Tree_Nodes_Count As Long

    Private Sub Frm_Product_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()
        Me.BackColor = Color.White


        Grid_Info_Picture1.Visible = False
        Grid_Code_Setup()
        Grid_Info_Setup()


        Search_Combo.Items.Add("코드")
        Search_Combo.Items.Add("품명")
        Search_Combo.Text = "코드"

        Dim DBT As New DataTable
        Search_Sel_Combo.Items.Clear()
        Search_Sel_Combo.Items.Add("전체")
        StrSQL = "Select Base_Name  FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '제품구분' Order By Base_Name"
        Re_Count = DB_Select(DBT)
        'Grid_Info_Combo.Items.Clear()
        If Re_Count <> 0 Then
            For i = 0 To Re_Count - 1
                Search_Sel_Combo.Items.Add(DBT.Rows(i)("Base_Name"))
            Next i
        End If
        Search_Sel_Combo.Text = "전체"

        Grid_Bom_Setup()
        Grid_Pro_Setup()

        Product_Line_QC_Setup()



        Bom_Panel.Left = 370
        Bom_Panel.Top = 127
        Bom_Panel.Visible = True

        'Product_Line_QC.Visible = False

        Grid_Code_Display(True)
        Grid_Code.ClearSelection()

        'Me.Com_Customer.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange

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
        Grid_Bom.Columns(3).Width = 80
        Grid_Bom.Columns(4).Width = 50

        Grid_Bom.Columns(0).HeaderText = "순번"
        Grid_Bom.Columns(1).HeaderText = "제품코드"
        Grid_Bom.Columns(2).HeaderText = "제품명"
        Grid_Bom.Columns(3).HeaderText = "Qty"
        Grid_Bom.Columns(4).HeaderText = "비고"


        Grid_Bom.Item(0, 0).Value = ""
        Grid_Bom.Item(1, 0).Value = ""
        Grid_Bom.Item(2, 0).Value = ""
        Grid_Bom.Item(3, 0).Value = ""
        Grid_Bom.Item(4, 0).Value = ""

        Grid_Bom.Columns(0).ReadOnly = True

    End Sub
    Public Sub Grid_Pro_Setup()

        Grid_Color(Grid_Pro)



        Grid_Pro.AllowUserToAddRows = False
        Grid_Pro.RowTemplate.Height = 24

        'Grid_Pro.ColumnHeadersHeight = 10
        Grid_Pro.ColumnCount = 5
        Grid_Pro.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Pro.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_Pro.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Pro.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Pro.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Grid_Pro.RowHeadersWidth = 40
        Grid_Pro.Columns(0).Width = 40
        Grid_Pro.Columns(1).Width = 60
        Grid_Pro.Columns(2).Width = 100
        Grid_Pro.Columns(3).Width = 80
        Grid_Pro.Columns(4).Width = 90
        '        Grid_Pro.Columns(4).Width = 2601
        '        Grid_Pro.Columns(5).Width = 260

        Grid_Pro.Columns(0).HeaderText = "순번"
        Grid_Pro.Columns(1).HeaderText = "공정코드"
        Grid_Pro.Columns(2).HeaderText = "공정명"
        Grid_Pro.Columns(3).HeaderText = "수량"
        Grid_Pro.Columns(4).HeaderText = "공정단가"
        'Grid_Pro.Columns(4).HeaderText = "공정전이미지"
        'Grid_Pro.Columns(5).HeaderText = "공정후이미지"


        Grid_Pro.Item(0, 0).Value = ""
        Grid_Pro.Item(1, 0).Value = ""
        Grid_Pro.Item(2, 0).Value = ""
        Grid_Pro.Item(3, 0).Value = ""
        'Grid_Pro.Item(4, 0).Value = ""
        'Grid_Pro.Item(5, 0).Value = ""

        Grid_Pro.Columns(0).ReadOnly = True

    End Sub
    Public Function Product_Line_QC_Setup() As Boolean
        Grid_Color(Product_Line_QC)



        Product_Line_QC.AllowUserToAddRows = False
        Product_Line_QC.RowTemplate.Height = 24


        '열의 갯수는 하나 더 많아야 함.
        Product_Line_QC.ColumnCount = 4
        Product_Line_QC.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Product_Line_QC.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Product_Line_QC.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Product_Line_QC.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Product_Line_QC.RowHeadersWidth = 40
        Product_Line_QC.Columns(0).Width = 40
        Product_Line_QC.Columns(1).Width = 100
        Product_Line_QC.Columns(2).Width = 115
        Product_Line_QC.Columns(0).HeaderText = "순번"
        Product_Line_QC.Columns(1).HeaderText = "점검항목"
        Product_Line_QC.Columns(2).HeaderText = "기준값"
        Product_Line_QC.Columns(3).HeaderText = "비고"


        Product_Line_QC.Columns(0).ReadOnly = True
        Product_Line_QC_Setup = True
    End Function

    Public Function Grid_Code_Setup() As Boolean
        Grid_Color(Grid_Code)
        Grid_Code.EnableHeadersVisualStyles = False
        Grid_Code.AllowUserToAddRows = False
        Grid_Info.ColumnHeadersHeight = 24

        Grid_Code.RowTemplate.Height = 24
        Grid_Code.ColumnCount = 3
        Grid_Code.RowCount = 1
        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Code.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Code.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Code.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Code.RowHeadersWidth = 10
        Grid_Code.Columns(0).Width = 100
        Grid_Code.Columns(1).Width = 150
        Grid_Code.Columns(2).Width = 70
        Grid_Code.Columns(0).HeaderText = "코드"
        Grid_Code.Columns(1).HeaderText = "품명"
        Grid_Code.Columns(2).HeaderText = "구분"
        Dim i As Integer
        For i = 0 To Grid_Code.ColumnCount - 1
            Grid_Code.Columns(i).ReadOnly = True
        Next i
    End Function

    Public Function Grid_Info_Setup() As Boolean
        Grid_Color(Grid_Info)

        '제목 폭
        Grid_Info.RowHeadersWidth = 10
        Grid_Info.ColumnHeadersHeight = 24

        'Data 부분 높이
        Grid_Info.RowTemplate.Height = 24
        Grid_Info.ColumnCount = 2
        Grid_Info.RowCount = 13






        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Info.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Info.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Info.Columns(0).Width = 100
        Grid_Info.Columns(1).Width = 240
        Grid_Info.Columns(0).HeaderText = "구분"
        Grid_Info.Columns(1).HeaderText = "내용"

        Grid_Info.Item(0, 0).Value = "코드"
        Grid_Info.Item(0, 1).Value = "품명"
        Grid_Info.Item(0, 2).Value = "구분"
        Grid_Info.Item(0, 3).Value = "종류"
        Grid_Info.Item(0, 4).Value = "규격"
        Grid_Info.Item(0, 5).Value = "단위"
        Grid_Info.Item(0, 6).Value = "단가"
        ''Grid_Info.Item(0, 7).Value = "이미지"
        ''Grid_Info.Rows(7).Height = 110
        ''Grid_Info_Picture_Setup()

        Grid_Info.Item(0, 7).Value = "수입검사"
        Grid_Info.Item(0, 8).Value = "공정검사"
        Grid_Info.Item(0, 9).Value = "출하검사"
        Grid_Info.Item(0, 10).Value = "유통기한"
        Grid_Info.Item(0, 11).Value = "안전재고수량"
        Grid_Info.Item(0, 12).Value = "비고"

        Grid_Info.Columns(0).ReadOnly = True



    End Function
    Public Function Grid_Info_Picture_Setup() As Boolean
        Grid_Info_Picture1.Top = Grid_Info.Top + Grid_Info.ColumnHeadersHeight + (7) * 24 + 1
        Grid_Info_Picture1.Left = Grid_Info.Left + Grid_Info.RowHeadersWidth + Grid_Info.Columns(0).Width + 1
        Grid_Info_Picture1.Width = Grid_Info.Columns(1).Width - 1
        Grid_Info_Picture1.Height = Grid_Info.Rows(7).Height - 1
        Grid_Info_Picture1.BackColor = Grid_Info.RowsDefaultCellStyle.BackColor
        Grid_Info_Picture1.Visible = True
        Grid_Info_Picture_Setup = True
    End Function

    Private Sub Insert_Com_Click(sender As Object, e As EventArgs)
        'DB를 추가 한다.
        '코드 번호 입력
        '새로운 푬에서 입력 받는다.

        Dim RR As String
        Dim DBT As New DataTable

        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Delete Temp "
        Re_Count = DB_Execute()



        Frm_Insert.Text = "제품 코드 추가"
        Frm_Insert.Show()
        Frm_Main.Enabled = False

        Do
            Application.DoEvents()

            StrSQL = "Select Temp From Temp With(NOLOCK)"
            Re_Count = DB_Select(DBT)

            If Re_Count <> 0 Then
                RR = DBT.Rows(0)("Temp")
                Exit Do
            Else
            End If
        Loop
        If RR = "취소" Then
            Exit Sub
        End If

        Grid_Info_Display_Pro(RR)
        If Len(Grid_Info.Item(1, 0).Value) > 0 Then

        Else

            Exit Sub
        End If

        Label_Color(Com_BOM, Color_Label, Di_Panel2)
        'BOM 입력 페널
        Grid_Bom_Display(Grid_Info.Item(1, 0).Value)
        Grid_Pro_Display(Grid_Info.Item(1, 0).Value)
        Bom_Panel.Visible = True


    End Sub

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        Grid_Code_Display(True)

    End Sub
    Public Function Grid_Code_Display(Display_Sel As Boolean) As Boolean

        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Dim Je_Gu As String
        If Display_Sel Then
            Grid_Code.RowCount = 1
            If Search_Sel_Combo.Text = "전체" Then
                Je_Gu = ""
            Else
                Je_Gu = " And PL_Sel = '" & Search_Sel_Combo.Text & "'"
            End If
            Select Case Search_Combo.Text
                Case "코드"
                    StrSQL = "Select PL_Code, PL_Name, PL_Sel FROM SI_PRODUCT with(NOLOCK) Where PL_Code Like '%" & Search_Text.Text & "%' " & Je_Gu & "  Order  By PL_Code"
                Case "품명"
                    StrSQL = "Select PL_Code, PL_Name, PL_Sel FROM SI_PRODUCT With(NOLOCK) Where PL_Name Like '%" & Search_Text.Text & "%' Or PL_Name Is Null " & Je_Gu & " Order By PL_Name"
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
        End If
        Grid_Code_Display = True
    End Function
    Private Sub Grid_Code_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Code.CellEnter
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
            Grid_Info_Display_Pro(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
            If Len(Grid_Info.Item(1, 0).Value) > 0 Then
            Else
                Exit Sub
            End If

            Label_Color(Com_BOM, Color_Label, Di_Panel2)
            'BOM 입력 페널
            Grid_Bom_Display(Grid_Info.Item(1, 0).Value)
            Grid_Pro_Display(Grid_Info.Item(1, 0).Value)
            Product_Line_QC_Display(Grid_Info.Item(1, 0).Value, Grid_Pro.Item(1, 0).Value)
            Bom_Panel.Visible = True

        End If
    End Sub
    Public Function Grid_Bom_Display(PL_Code As String) As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display_Ch = False
        Grid_Bom.RowCount = 1
        StrSQL = "Select PL_Sun, PL_Sub_Code, PL_Qty, PL_Bigo   FROM SI_PRODUCT_RECIPE with(NOLOCK)  Where PL_Code = '" & PL_Code & "' Order By Convert(Decimal,PL_Sun)"
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
        For i = 0 To Grid_Bom.RowCount - 1
            StrSQL = "Select PL_Name  FROM SI_PRODUCT with(NOLOCK)  Where PL_Code = '" & Grid_Bom.Item(1, i).Value & "' "
            Re_Count = DB_Select(DBT)
            If Re_Count <> 0 Then
                Grid_Bom.Item(2, i).Value = DBT.Rows(0).Item(0).ToString
            End If
        Next




        Grid_Display_Ch = True
        Grid_Bom_Display = True
    End Function
    Public Function Grid_Pro_Display(PL_Code As String) As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display_Ch = False
        Grid_Pro.RowCount = 1
        StrSQL = "Select PP_Sun, PP_PC_Code, pp_amount    FROM SI_Process_List with(NOLOCK)  Where PP_Code = '" & PL_Code & "' Order By Convert(Decimal,PP_Sun)"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Pro.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                Grid_Pro.Item(0, i).Value = DBT.Rows(i).Item(0).ToString
                Grid_Pro.Item(1, i).Value = DBT.Rows(i).Item(1).ToString
                Grid_Pro.Item(2, i).Value = ""
                Grid_Pro.Item(3, i).Value = DBT.Rows(i).Item(2).ToString
                '  Grid_Pro.Item(4, i).Value = DBT.Rows(i).Item(5).ToString
                'Grid_Pro.Item(4, i).Value = DBT.Rows(i).Item(3).ToString
                'Grid_Pro.Item(5, i).Value = DBT.Rows(i).Item(4).ToString
            Next i
        Else
            Grid_Pro.RowCount = 1
            For j = 0 To Grid_Pro.ColumnCount - 1
                Grid_Pro.Item(j, 0).Value = ""
            Next j

        End If

        For i = 0 To Grid_Pro.RowCount - 1
            StrSQL = "Select PC_Name  FROM SI_PROCESS with(NOLOCK)  Where PC_Code = '" & Grid_Pro.Item(1, i).Value & "' "
            Re_Count = DB_Select(DBT)
            If Re_Count <> 0 Then
                Grid_Pro.Item(2, i).Value = DBT.Rows(0).Item(0).ToString
            End If
        Next


        Grid_Display_Ch = True
        Grid_Pro_Display = True
    End Function

    Public Function Grid_Info_Display_Pro(PL_Code As String) As Boolean
        Grid_Display_Ch = False
        'Try
        Dim DBT As New DataTable
        Dim i As Integer
        StrSQL = "Select * FROM SI_PRODUCT With(NOLOCK) Where PL_Code = '" & PL_Code & "'"
        Re_Count = DB_Select(DBT)
        If Re_Count > 0 Then
            For i = 0 To DBT.Columns.Count - 1
                Grid_Info.Item(1, i).Value = DBT.Rows(0).Item(i).ToString
                '이미지 처리 한다.
                If i = 7 Then
                    If Len(Grid_Info.Item(1, i).Value) < 100 Then
                        Grid_Info_Picture1.Image = Nothing
                    Else
                        Dim b() As Byte = HexStr2Array(Grid_Info.Item(1, i).Value)
                        Grid_Info_Picture1.Image = Image.FromStream(New MemoryStream(b, 0, b.Length))
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
        '선택되었을때 처리
        If Len(Grid_Info.Item(1, 0).Value) < 1 Then
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
                Case 2
                    StrSQL = "Select Base_Name  FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '제품구분' Order By Base_Name"
                    Re_Count = DB_Select(DBT)
                    Grid_Info_Combo.Items.Clear()
                    If Re_Count <> 0 Then
                        For i = 0 To Re_Count - 1
                            Grid_Info_Combo.Items.Add(DBT.Rows(i)("Base_Name"))
                        Next i 'DBT.Rows(i)("Base_Bigo")
                    End If
                    Grid_Info_Combo.Top = Grid_Info.Top + Grid_Info.ColumnHeadersHeight + (Grid_Info_Row) * 24
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
                Case 3
                    StrSQL = "Select Base_Name  FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '제품종류'  Order By Base_Name"
                    Re_Count = DB_Select(DBT)
                    Grid_Info_Combo.Items.Clear()
                    If Re_Count <> 0 Then
                        For i = 0 To Re_Count - 1
                            Grid_Info_Combo.Items.Add(DBT.Rows(i)("Base_Name"))
                        Next i 'DBT.Rows(i)("Base_Bigo")
                    End If
                    Grid_Info_Combo.Top = Grid_Info.Top + Grid_Info.ColumnHeadersHeight + (Grid_Info_Row) * 24
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
                    StrSQL = "Select Base_Name  FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '단위'  Order By Base_Name"
                    Re_Count = DB_Select(DBT)
                    Grid_Info_Combo.Items.Clear()
                    If Re_Count <> 0 Then
                        For i = 0 To Re_Count - 1
                            Grid_Info_Combo.Items.Add(DBT.Rows(i)("Base_Name"))
                        Next i 'DBT.Rows(i)("Base_Bigo")
                    End If
                    Grid_Info_Combo.Top = Grid_Info.Top + Grid_Info.ColumnHeadersHeight + (Grid_Info_Row) * 24
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
                Case 8
                    StrSQL = "Select QC_Name  FROM SI_QC with(NOLOCK) Where QC_Sel = '수입검사' Order By QC_Name"
                    Re_Count = DB_Select(DBT)
                    Grid_Info_Combo.Items.Clear()
                    If Re_Count <> 0 Then
                        For i = 0 To Re_Count - 1
                            Grid_Info_Combo.Items.Add(DBT.Rows(i)("QC_Name"))
                        Next i
                    End If
                    Grid_Info_Combo.Top = Grid_Info.Top + Grid_Info.ColumnHeadersHeight + (Grid_Info_Row - 1) * 24 + Grid_Info.Rows(7).Height
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
                Case 9
                    ''"수입검사", "순회검사", "출하검사", "Xbar-R 관리도", "부적합보고서"
                    StrSQL = "Select QC_Name  FROM SI_QC with(NOLOCK) Where QC_Sel = '순회검사' Order By QC_Name"
                    Re_Count = DB_Select(DBT)
                    Grid_Info_Combo.Items.Clear()
                    If Re_Count <> 0 Then
                        For i = 0 To Re_Count - 1
                            Grid_Info_Combo.Items.Add(DBT.Rows(i)("QC_Name"))
                        Next i
                    End If
                    Grid_Info_Combo.Top = Grid_Info.Top + Grid_Info.ColumnHeadersHeight + (Grid_Info_Row - 1) * 24 + Grid_Info.Rows(7).Height
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
                Case 10
                    StrSQL = "Select QC_Name  FROM SI_QC with(NOLOCK) Where QC_Sel = '자주검사' Order By QC_Name"
                    Re_Count = DB_Select(DBT)
                    Grid_Info_Combo.Items.Clear()
                    If Re_Count <> 0 Then
                        For i = 0 To Re_Count - 1
                            Grid_Info_Combo.Items.Add(DBT.Rows(i)("QC_Name"))
                        Next i
                    End If
                    Grid_Info_Combo.Top = Grid_Info.Top + Grid_Info.ColumnHeadersHeight + (Grid_Info_Row - 1) * 24 + Grid_Info.Rows(7).Height
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
                Case 11
                    StrSQL = "Select QC_Name  FROM SI_QC with(NOLOCK) Where QC_Sel = '출하검사' Order By QC_Name"
                    Re_Count = DB_Select(DBT)
                    Grid_Info_Combo.Items.Clear()
                    If Re_Count <> 0 Then
                        For i = 0 To Re_Count - 1
                            Grid_Info_Combo.Items.Add(DBT.Rows(i)("QC_Name"))
                        Next i
                    End If
                    Grid_Info_Combo.Top = Grid_Info.Top + Grid_Info.ColumnHeadersHeight + (Grid_Info_Row - 1) * 24 + Grid_Info.Rows(7).Height
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
                Case 12
                    StrSQL = "Select QC_Name  FROM SI_QC with(NOLOCK) Where QC_Sel = 'Xbar-R' Order By QC_Name"
                    Re_Count = DB_Select(DBT)
                    Grid_Info_Combo.Items.Clear()
                    If Re_Count <> 0 Then
                        For i = 0 To Re_Count - 1
                            Grid_Info_Combo.Items.Add(DBT.Rows(i)("QC_Name"))
                        Next i
                    End If
                    Grid_Info_Combo.Top = Grid_Info.Top + Grid_Info.ColumnHeadersHeight + (Grid_Info_Row - 1) * 24 + Grid_Info.Rows(7).Height
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
                Case 13
                    StrSQL = "Select QC_Name  FROM SI_QC with(NOLOCK) Where QC_Sel = '부적합' Order By QC_Name"
                    Re_Count = DB_Select(DBT)
                    Grid_Info_Combo.Items.Clear()
                    If Re_Count <> 0 Then
                        For i = 0 To Re_Count - 1
                            Grid_Info_Combo.Items.Add(DBT.Rows(i)("QC_Name"))
                        Next i
                    End If
                    Grid_Info_Combo.Top = Grid_Info.Top + Grid_Info.ColumnHeadersHeight + (Grid_Info_Row - 1) * 24 + Grid_Info.Rows(7).Height
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


    Private Sub Grid_Code_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Code.CellContentClick

    End Sub

    Private Sub Save_Com_Click(sender As Object, e As EventArgs)
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



    End Sub

    Private Sub Grid_Info_Combo_LostFocus(sender As Object, e As EventArgs) Handles Grid_Info_Combo.LostFocus
        Grid_Info_Combo.Visible = False
    End Sub

    Private Sub Grid_Info_Combo_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Grid_Info_Combo.SelectionChangeCommitted
        Dim DBT As New DataTable
        Grid_Info.Item(Grid_Info_Col, Grid_Info_Row).Value = Grid_Info_Combo.SelectedItem.ToString()

    End Sub
    Private Sub Grid_Info_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Info.CellValueChanged
        If Grid_Display_Ch = True Then
            Grid_Info.CurrentRow.HeaderCell.Value = "*"
        End If
    End Sub

    Private Sub Grid_Info_Picture1_Click(sender As Object, e As EventArgs) Handles Grid_Info_Picture1.Click

    End Sub

    Private Sub Grid_Info_Picture1_DoubleClick(sender As Object, e As EventArgs) Handles Grid_Info_Picture1.DoubleClick
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
                Grid_Info_Picture1.Image = Image.FromStream(fs)
                Grid_Info.Rows(7).HeaderCell.Value = "*"
                fs.Close()
            End If
            'Catch Ex As Exception
            'MsgBox(Ex.Message)
            'End Try
        Else
            Grid_Info.Item(1, 7).Value = ""
            Grid_Info_Picture1.Image = Nothing
            Grid_Info.Rows(7).HeaderCell.Value = "*"
        End If

    End Sub

    Private Sub Save_Bom_Click(sender As Object, e As EventArgs)
        If Len(Grid_Info.Item(1, 0).Value) > 0 Then
        Else
            Exit Sub
        End If

        Dim i As Integer


        Grid_Display_Ch = False
        For i = 0 To Grid_Bom.Rows.Count - 1
            If Grid_Bom.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SI_PRODUCT_RECIPE Set PL_Sub_Code = '" & Grid_Bom.Item(1, i).Value & "', PL_Qty = '" & Grid_Bom.Item(3, i).Value & "', PL_Bigo = '" & Grid_Bom.Item(4, i).Value & "' where PL_Code = '" & Grid_Info.Item(1, 0).Value & "' And PL_Sun = '" & Grid_Bom.Item(0, i).Value & "'"
                Re_Count = DB_Execute()
                Grid_Bom.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display_Ch = True

    End Sub
    Private Sub Grid_Bom_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Bom.CellValueChanged
        If Grid_Display_Ch = True Then
            Grid_Bom.CurrentRow.HeaderCell.Value = "*"
        End If
    End Sub

    Private Sub Grid_Bom_MouseClick(sender As Object, e As MouseEventArgs) Handles Grid_Bom.MouseClick
        Grid_Bom_Row = Grid_Bom.CurrentCell.RowIndex
        Grid_Bom_Col = Grid_Bom.CurrentCell.ColumnIndex
        Grid_Bom_Combo.Visible = False
        If Grid_Bom_Row < 0 Then
            Exit Sub
        End If
        If Grid_Bom_Col < 0 Then
            Exit Sub
        End If
        Dim DBT As New DataTable

        Select Case Grid_Bom_Col
            Case 1
                StrSQL = "Select PL_Code, PL_Name  FROM SI_PRODUCT with(NOLOCK)  Order By PL_Code"
                Re_Count = DB_Select(DBT)
                Grid_Bom_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        Grid_Bom_Combo.Items.Add(DBT.Rows(i).Item(0))
                    Next i
                End If

                Grid_Bom_Combo.Top = Grid_Bom.Top + Grid_Bom.ColumnHeadersHeight + (Grid_Bom_Row) * 24
                Grid_Bom_Combo.Left = Grid_Bom.Left + Grid_Bom.RowHeadersWidth + Grid_Bom.Columns(0).Width
                Grid_Bom_Combo.Width = Grid_Bom.Columns(Grid_Bom_Col).Width
                Grid_Bom_Combo.Text = Grid_Bom.CurrentCell.Value.ToString
                Grid_Bom_Combo.BackColor = Grid_Bom.RowsDefaultCellStyle.BackColor
                Grid_Bom_Combo.Visible = True
                Grid_Bom_Combo.Focus.ToString()
            Case 2
                StrSQL = "Select PL_Code, PL_Name  FROM SI_PRODUCT with(NOLOCK)  Order By PL_Code"

                Re_Count = DB_Select(DBT)
                Grid_Bom_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        Grid_Bom_Combo.Items.Add(DBT.Rows(i).Item(1))
                    Next i
                End If
                Grid_Bom_Combo.Top = Grid_Bom.Top + Grid_Bom.ColumnHeadersHeight + (Grid_Bom_Row) * 24
                Grid_Bom_Combo.Left = Grid_Bom.Left + Grid_Bom.RowHeadersWidth + Grid_Bom.Columns(0).Width + Grid_Bom.Columns(1).Width
                Grid_Bom_Combo.Width = Grid_Bom.Columns(Grid_Bom_Col).Width
                Grid_Bom_Combo.Text = Grid_Bom.CurrentCell.Value.ToString
                Grid_Bom_Combo.BackColor = Grid_Bom.RowsDefaultCellStyle.BackColor
                Grid_Bom_Combo.Visible = True
                Grid_Bom_Combo.Focus.ToString()
        End Select

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
                StrSQL = "Select PL_Code, PL_Name  FROM SI_PRODUCT with(NOLOCK) Where PL_Code = '" & Grid_Bom_Combo.SelectedItem.ToString() & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Bom.Item(2, Grid_Bom_Row).Value = DBT.Rows(0).Item(1)
                End If
            Case 2
                StrSQL = "Select PL_Code, PL_Name  FROM SI_PRODUCT with(NOLOCK) Where PL_Name= '" & Replace(Grid_Bom_Combo.SelectedItem.ToString(), "'", "''") & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Bom.Item(1, Grid_Bom_Row).Value = DBT.Rows(0).Item(0)
                End If
        End Select

    End Sub

    Private Sub Put__Bom_Click(sender As Object, e As EventArgs)
        If Len(Grid_Info.Item(1, 0).Value) > 0 Then
        Else
            Exit Sub
        End If

        Grid_Display_Ch = False
        Dim i As Integer
        For i = Grid_Bom.RowCount - 1 To Grid_Bom.CurrentCell.RowIndex Step -1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "UPDATE SI_PRODUCT_RECIPE Set PL_Sun = '" & i + 2 & "' Where PL_Sun = '" & i + 1 & "'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO SI_PRODUCT_RECIPE (PL_Code, PL_Sun)  Values ('" & Grid_Info.Item(1, 0).Value & "', '" & Grid_Bom.CurrentCell.RowIndex + 1 & "')"
        Re_Count = DB_Execute()
        Grid_Bom_Display(Grid_Info.Item(1, 0).Value)
        Grid_Display_Ch = True


    End Sub

    Private Sub Delete__Bom_Click(sender As Object, e As EventArgs)
        '해당 칼럼 삭제
        If Len(Grid_Info.Item(1, 0).Value) > 0 Then
        Else
            Exit Sub
        End If

        Grid_Display_Ch = False
        Dim i As Integer
        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_PRODUCT_RECIPE Where PL_Code = '" & Grid_Info.Item(1, 0).Value & "' And   PL_Sun = '" & Grid_Bom.Item(0, Grid_Bom.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        For i = Grid_Bom.CurrentCell.RowIndex + 1 To Grid_Bom.RowCount - 1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "UPDATE SI_PRODUCT_RECIPE Set PL_Sun = '" & i & "' Where PL_Code = '" & Grid_Info.Item(1, 0).Value & "' And  PL_Sun = '" & i + 1 & "'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        Grid_Bom_Display(Grid_Info.Item(1, 0).Value)
        Grid_Display_Ch = True

    End Sub


    Private Sub Insert__Bom_Click(sender As Object, e As EventArgs)
        If Len(Grid_Info.Item(1, 0).Value) > 0 Then
        Else
            Exit Sub
        End If

        Grid_Display_Ch = False

        Dim DBT As New DataTable
        Dim Db_Sun As Long
        StrSQL = "Select PL_Sun FROM SI_PRODUCT_RECIPE with(NOLOCK) Where PL_Code = '" & Grid_Info.Item(1, 0).Value & "' Order By Convert(Decimal,PL_Sun) Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            Db_Sun = 1
        Else
            Db_Sun = DBT.Rows(0).Item(0) + 1
        End If

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO SI_PRODUCT_RECIPE (PL_Code, PL_Sun)  Values ('" & Grid_Info.Item(1, 0).Value & "', '" & Db_Sun & "')"
        Re_Count = DB_Execute()
        Grid_Bom_Display(Grid_Info.Item(1, 0).Value)
        Grid_Display_Ch = True

    End Sub
    Private Sub Grid_Pro_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Pro.CellValueChanged
        If Grid_Display_Ch = True Then
            Grid_Pro.CurrentRow.HeaderCell.Value = "*"
        End If

    End Sub

    Private Sub Grid_Pro_MouseClick(sender As Object, e As MouseEventArgs) Handles Grid_Pro.MouseClick
        Grid_Pro_Row = Grid_Pro.CurrentCell.RowIndex
        Grid_Pro_Col = Grid_Pro.CurrentCell.ColumnIndex
        Grid_Pro_Combo.Visible = False
        If Grid_Pro_Row < 0 Then
            Exit Sub
        End If
        If Grid_Pro_Col < 0 Then
            Exit Sub
        End If
        Dim DBT As New DataTable

        Select Case Grid_Pro_Col
            Case 1
                StrSQL = "Select PC_Code, PC_Name  FROM SI_PROCESS with(NOLOCK)  Order By PC_Code"
                Re_Count = DB_Select(DBT)
                Grid_Pro_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        Grid_Pro_Combo.Items.Add(DBT.Rows(i).Item(0))
                    Next i
                End If

                Grid_Pro_Combo.Top = Grid_Pro.Top + Grid_Pro.ColumnHeadersHeight + (Grid_Pro_Row) * 24
                Grid_Pro_Combo.Left = Grid_Pro.Left + Grid_Pro.RowHeadersWidth + Grid_Pro.Columns(0).Width
                Grid_Pro_Combo.Width = Grid_Pro.Columns(Grid_Pro_Col).Width
                Grid_Pro_Combo.Text = Grid_Pro.CurrentCell.Value.ToString
                Grid_Pro_Combo.BackColor = Grid_Pro.RowsDefaultCellStyle.BackColor
                Grid_Pro_Combo.Visible = True
                Grid_Pro_Combo.Focus.ToString()
            Case 2
                StrSQL = "Select PC_Code, PC_Name  FROM SI_PROCESS with(NOLOCK)  Order By PC_Code"

                Re_Count = DB_Select(DBT)
                Grid_Pro_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        Grid_Pro_Combo.Items.Add(DBT.Rows(i).Item(1))
                    Next i
                End If
                Grid_Pro_Combo.Top = Grid_Pro.Top + Grid_Pro.ColumnHeadersHeight + (Grid_Pro_Row) * 24
                Grid_Pro_Combo.Left = Grid_Pro.Left + Grid_Pro.RowHeadersWidth + Grid_Pro.Columns(0).Width + Grid_Pro.Columns(1).Width
                Grid_Pro_Combo.Width = Grid_Pro.Columns(Grid_Pro_Col).Width
                Grid_Pro_Combo.Text = Grid_Pro.CurrentCell.Value.ToString
                Grid_Pro_Combo.BackColor = Grid_Pro.RowsDefaultCellStyle.BackColor
                Grid_Pro_Combo.Visible = True
                Grid_Pro_Combo.Focus.ToString()
        End Select
        Product_Line_QC_Display(Grid_Info.Item(1, 0).Value, Grid_Pro.Item(1, Grid_Pro.CurrentCell.RowIndex).Value)


    End Sub
    Public Function Product_Line_QC_Display(PL_Code As String, PL_PC_Code As String) As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display_Ch = False
        Product_Line_QC.RowCount = 1
        StrSQL = "Select PQ_Sun, PQ_History1, PQ_History2, PQ_Bigo   FROM SI_PRODUCT_QC with(NOLOCK)  Where PQ_Code = '" & PL_Code & "' And PQ_PC_Code = '" & PL_PC_Code & "' Order By Convert(Decimal,PQ_Sun)"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Product_Line_QC.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                Product_Line_QC.Item(0, i).Value = DBT.Rows(i).Item(0).ToString
                Product_Line_QC.Item(1, i).Value = DBT.Rows(i).Item(1).ToString
                Product_Line_QC.Item(2, i).Value = DBT.Rows(i).Item(2).ToString
                Product_Line_QC.Item(3, i).Value = DBT.Rows(i).Item(3).ToString
            Next i
        Else
            Product_Line_QC.RowCount = 1
            For j = 0 To Product_Line_QC.ColumnCount - 1
                Product_Line_QC.Item(j, 0).Value = ""
            Next j

        End If

        Grid_Display_Ch = True
        Product_Line_QC_Display = True
    End Function
    Private Sub Product_Line_QC_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Product_Line_QC.CellValueChanged
        If Grid_Display_Ch = True Then
            Product_Line_QC.CurrentRow.HeaderCell.Value = "*"
        End If

    End Sub

    Private Sub Save_Pro_Click(sender As Object, e As EventArgs)
        If Len(Grid_Info.Item(1, 0).Value) > 0 Then
        Else
            Exit Sub
        End If

        Grid_Display_Ch = False
        Dim i As Integer

        For i = 0 To Grid_Pro.Rows.Count - 1
            If Grid_Pro.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SI_PROCESS_LIST Set PP_PC_Code = '" & Grid_Pro.Item(1, i).Value.ToString & "', PP_Amount = '" & Grid_Pro.Item(3, i).Value.ToString & "', PP_Gold = '" & Grid_Pro.Item(4, i).Value.ToString & "' where PP_Code = '" & Grid_Info.Item(1, 0).Value.ToString & "' And PP_Sun = '" & Grid_Pro.Item(0, i).Value.ToString & "'"
                Re_Count = DB_Execute()
                Grid_Pro.Rows(i).HeaderCell.Value = ""
            End If
        Next i

        Grid_Display_Ch = True
    End Sub
    Private Sub Insert__Pro_Click(sender As Object, e As EventArgs)
        If Len(Grid_Info.Item(1, 0).Value) > 0 Then
        Else
            Exit Sub
        End If

        Grid_Display_Ch = False
        Dim DBT As New DataTable
        Dim Db_Sun As Long
        StrSQL = "Select PP_Sun FROM SI_Process_List with(NOLOCK) Where PP_Code = '" & Grid_Info.Item(1, 0).Value & "' Order By PP_Sun Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            Db_Sun = 1
        Else
            Db_Sun = DBT.Rows(0).Item(0) + 1
        End If

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO SI_PROCESS_LIST (PP_Code, PP_Sun)  Values ('" & Grid_Info.Item(1, 0).Value & "', '" & Db_Sun & "')"
        Re_Count = DB_Execute()
        Grid_Pro_Display(Grid_Info.Item(1, 0).Value)
        Grid_Display_Ch = True



    End Sub

    Private Sub Put__Pro_Click(sender As Object, e As EventArgs)
        If Len(Grid_Info.Item(1, 0).Value) > 0 Then
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
        StrSQL = StrSQL & "Insert INTO SI_PROCESS_LIST (PP_Code, PP_Sun)  Values ('" & Grid_Info.Item(1, 0).Value & "', '" & Grid_Pro.CurrentCell.RowIndex + 1 & "')"
        Re_Count = DB_Execute()

        Grid_Pro_Display(Grid_Info.Item(1, 0).Value)
        Grid_Display_Ch = True

    End Sub

    Private Sub Delete__Pro_Click(sender As Object, e As EventArgs)
        If Len(Grid_Info.Item(1, 0).Value) > 0 Then
        Else
            Exit Sub
        End If


        If Len(Grid_Pro.Item(0, Grid_Pro.CurrentCell.RowIndex).Value) < 1 Then
            Exit Sub
        End If
        Grid_Display_Ch = False
        Dim i As Integer

        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_PROCESS_LIST Where PP_Code = '" & Grid_Info.Item(1, 0).Value & "' And   PP_Sun = '" & Grid_Pro.Item(0, Grid_Pro.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        For i = Grid_Pro.CurrentCell.RowIndex + 1 To Grid_Pro.RowCount - 1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "UPDATE SI_PROCESS_LIST Set PP_Sun = '" & i & "' Where PP_Code = '" & Grid_Info.Item(1, 0).Value & "' And  PP_Sun = '" & i + 1 & "'"
            Re_Count = DB_Execute()
        Next i
        Grid_Pro_Display(Grid_Info.Item(1, 0).Value)
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


    Private Sub Button6_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Com_BOM_Click(sender As Object, e As EventArgs) Handles Com_BOM.Click

        If Len(Grid_Info.Item(1, 0).Value) > 0 Then
        Else
            Exit Sub
        End If
        Label_Color(sender, Color_Label, Di_Panel2)
        'BOM 입력 페널

        Grid_Bom_Display(Grid_Info.Item(1, 0).Value)
        Grid_Pro_Display(Grid_Info.Item(1, 0).Value)
        Bom_Panel.Visible = True

    End Sub

    Private Sub Com_Tree_Click(sender As Object, e As EventArgs)
        If Len(Grid_Info.Item(1, 0).Value) > 0 Then
        Else
            Exit Sub
        End If
        If Len(Len(Grid_Info.Item(1, 1).Value)) > 0 Then
        Else
            MsgBox("제품명을 입력하세요")
            Exit Sub
        End If

        Label_Color(sender, Color_Label, Di_Panel2)

        Bom_Panel.Visible = False

        Dim PD_Code As String
        Dim Re_Text As String
        Dim My_Nod As TreeNode
        Dim Nod_Count As Long

        PD_Code = Grid_Info.Item(1, 0).Value
        If Len(PD_Code) < 1 Then
            Exit Sub
        End If
    End Sub
    Public Function Product_Tree_Search(Pr_Nod As TreeNode, PD_Code As String, Nod_Count As Long) As String
        Dim My_Nod As TreeNode
        Dim i As Long
        Dim Re_Text As String
        Nod_Count = Nod_Count + 1
        If Nod_Count > 20 Then
            MsgBox("제품에 무한으로 생성 됩니다. BOM 구조를 확인 부탁 드립니다.")
            Exit Function
        End If
        If Len(PD_Code) < 1 Then
            Product_Tree_Search = "BOM에 코드가 없는 자료가 있습니다. 확인 부탁 드립니다."
            Exit Function
        End If
        'Data를 가지고 온다.
        Dim DBT As New DataTable
        StrSQL = "Select PL_Sun, PL_Sub_Code, SI_Product.PL_Name, PL_Qty,SI_Product.PL_Sel   FROM SI_PRODUCT_RECIPE With(NOLOCK), Product With(NOLOCK) Where Product_BOM.PL_Code = '" & PD_Code & "' And  Product_BOM.PL_Sub_Code = SI_Product.PL_Code Order By PL_Sun"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            For i = 0 To Re_Count - 1
                Tree_Nodes_Count = Tree_Nodes_Count + 1
                My_Nod = Pr_Nod.Nodes.Add(DBT.Rows(i).Item(1).ToString & "-" & Tree_Nodes_Count, DBT.Rows(i).Item(1).ToString & " - " & DBT.Rows(i).Item(2).ToString, "단품.png")
                '상위노드검색
                If My_Nod.Parent IsNot Nothing Then
                    Re_Text = Product_Tree_Parant_Search(My_Nod.Parent, DBT.Rows(i).Item(1).ToString)
                    If Re_Text = "Ok" Then
                    Else
                        Product_Tree_Search = Re_Text
                        Exit Function
                    End If
                End If

                Re_Text = Product_Tree_Search(My_Nod, DBT.Rows(i).Item(1).ToString, Nod_Count)
                If Re_Text = "Ok" Then
                Else
                    Product_Tree_Search = Re_Text
                    Exit Function
                End If

                '공정을 삽입한다.
                Re_Text = Product_Tree_Process(My_Nod, DBT.Rows(i).Item(1).ToString, Nod_Count)
                If Re_Text = "Ok" Then
                Else
                    Product_Tree_Search = Re_Text
                    Exit Function
                End If
            Next i
        Else
        End If



        Product_Tree_Search = "Ok"


    End Function
    Public Function Product_Tree_Parant_Search(Pr_Nod As TreeNode, PD_Code As String) As String
        'Dim Re_Text As String
        'Dim i As Long
        'Dim PPD_Code As String
        'PPD_Code = Pr_Nod.Text.ToString

        'For i = 1 To Len(PPD_Code)
        '    If Mid(PPD_Code, i, 1) = "-" Then
        '        PPD_Code = Mid(PPD_Code, 1, i - 2)
        '        Exit For
        '    End If
        'Next i
        'If Len(PPD_Code) < 2 Then
        '    Product_Tree_Parant_Search = PD_Code & " 등록 오류 입니다."
        '    Exit Function
        'End If

        'If PPD_Code = PD_Code Then
        '    Product_Tree_Parant_Search = "상위에" & PD_Code & "가 중복입니다."
        '    Exit Function
        'Else
        '    If Pr_Nod.Parent IsNot Nothing Then
        '        Re_Text = Product_Tree_Parant_Search(Pr_Nod.Parent, PD_Code)
        '        If Re_Text = "Ok" Then
        '        Else
        '            Product_Tree_Parant_Search = Re_Text
        '            Exit Function
        '        End If

        '    End If
        'End If
        Product_Tree_Parant_Search = "Ok"
    End Function
    Public Function Product_Tree_Process(Pr_Nod As TreeNode, PD_Code As String, Nod_Count As Long) As String
        Dim My_Nod As TreeNode
        Dim i As Long
        Dim Re_Text As String
        'If Tree_Nodes_Count > 10 Then
        '    Product_Tree_Search = "Error"
        '    Exit Function
        'End If
        If Len(PD_Code) < 1 Then
            Product_Tree_Process = "BOM에 코드가 없는 자료가 있습니다. 확인 부탁 드립니다."
            Exit Function
        End If
        'Data를 가지고 온다.
        Dim DBT As New DataTable
        StrSQL = "Select PP_Sun, PP_PC_Code, Process.PC_Name, PP_Amount  FROM SI_Process_List with(NOLOCK), Process With(NOLOCK) Where SI_Process_List.PP_Code = '" & PD_Code & "' And SI_Process_List.PP_PC_Code = Process.PC_Code Order By PP_Sun"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            For i = 0 To Re_Count - 1
                Tree_Nodes_Count = Tree_Nodes_Count + 1
                My_Nod = Pr_Nod.Nodes.Add(DBT.Rows(i).Item(1).ToString & "-" & Tree_Nodes_Count, DBT.Rows(i).Item(1).ToString & " - " & DBT.Rows(i).Item(2).ToString & " 수량 : " & DBT.Rows(i)("PP_Amount").ToString, "공정.Png")
                My_Nod.BackColor = Color.FromArgb(248, 203, 173)
                Re_Text = Product_Tree_Search(My_Nod, DBT.Rows(i).Item(1).ToString, Nod_Count)
                If Re_Text = "Ok" Then
                Else
                    Product_Tree_Process = Re_Text
                    Exit Function
                End If
            Next i
        Else
        End If
        Product_Tree_Process = "Ok"
    End Function

    Private Sub Com_Customer_Click(sender As Object, e As EventArgs)
        If Len(Grid_Info.Item(1, 0).Value) > 0 Then
        Else
            Exit Sub
        End If
        Label_Color(sender, Color_Label, Di_Panel2)
        'BOM 입력 페널

        Bom_Panel.Visible = False
    End Sub


    Private Sub Insert__Cus_In_Click(sender As Object, e As EventArgs)
        If Len(Grid_Info.Item(1, 0).Value) > 0 Then
        Else
            Exit Sub
        End If

        Grid_Display_Ch = False

        Dim DBT As New DataTable
        Dim Db_Sun As Long
        StrSQL = "Select PL_Sun FROM SI_PRODUCT_Customer with(NOLOCK) Where PL_Code = '" & Grid_Info.Item(1, 0).Value & "' And PL_Sel = '매입' Order By Convert(Decimal,PL_Sun) Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            Db_Sun = 1
        Else
            Db_Sun = DBT.Rows(0).Item(0) + 1
        End If

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO SI_PRODUCT_Customer (PL_Sel, PL_Code, PL_Sun)  Values ('매입','" & Grid_Info.Item(1, 0).Value & "', '" & Db_Sun & "')"
        Re_Count = DB_Execute()

        Grid_Display_Ch = True

    End Sub




    Private Sub Grid_Cus_In_Combo_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub



    Private Sub Grid_Cus_Out_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub



    Private Sub Grid_Cus_Out_Combo_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub


    Private Sub Delete_Com_Click(sender As Object, e As EventArgs)
        '제품 코드 삭제
        '다른 제품에 BOM에 있으면 안내 메세지를 표시해 준다.
        If Len(Grid_Info.Item(1, 0).Value) < 1 Then
            Exit Sub
        End If


        If MsgBox("제품 코드 '" & Grid_Info.Item(1, 0).Value & "' 제품명'" & Grid_Info.Item(1, 1).Value & "' 를 삭제 하면 다른 제품에 영향을 줄 수 있습니다. 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "제품 삭제") = vbNo Then
            Exit Sub
        End If
        Grid_Display_Ch = False

        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_PRODUCT Where PL_Code = '" & Grid_Info.Item(1, 0).Value & "' "
        Re_Count = DB_Execute()

        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_PRODUCT_RECIPE Where PL_Code = '" & Grid_Info.Item(1, 0).Value & "' "
        Re_Count = DB_Execute()

        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_PRODUCT_Customer Where PL_Code = '" & Grid_Info.Item(1, 0).Value & "' "
        Re_Count = DB_Execute()

        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_PRODUCT_QC Where PQ_Code = '" & Grid_Info.Item(1, 0).Value & "' "
        Re_Count = DB_Execute()

        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_PROCESS_LIST Where PP_Code = '" & Grid_Info.Item(1, 0).Value & "' "

        Re_Count = DB_Execute()
        Grid_Code_Display(True)
        Grid_Display_Ch = True

    End Sub

    Private Sub Grid_Cus_In_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub Grid_Info_CellContentClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Info.CellContentClick

    End Sub

    Private Sub Panel_Menu_Paint(sender As Object, e As PaintEventArgs) Handles Panel_Menu.Paint

    End Sub

    Private Sub Btn_Save_Click(sender As Object, e As EventArgs) Handles Btn_Save.Click
        If Len(Grid_Info.Item(1, 0).Value) > 0 Then
        Else
            Exit Sub
        End If
        If Len(Grid_Pro.Item(1, Grid_Pro.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If


        Dim i As Integer

        Grid_Display_Ch = False
        For i = 0 To Product_Line_QC.Rows.Count - 1
            If Product_Line_QC.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED " 'Replace(Grid_Info.Item(1, i).Value, "'", "''")
                StrSQL = StrSQL & "UPDATE SI_PRODUCT_QC Set PQ_History1 = '" & Replace(Product_Line_QC.Item(1, i).Value, "'", "''") & "', PQ_History2 = '" & Replace(Product_Line_QC.Item(2, i).Value, "'", "''") & "',PQ_Bigo = '" & Product_Line_QC.Item(3, i).Value & "' where  PQ_Code = '" & Grid_Info.Item(1, 0).Value & "' And PQ_PC_Code = '" & Grid_Pro.Item(1, Grid_Pro.CurrentCell.RowIndex).Value & "' And PQ_Sun = '" & Product_Line_QC.Item(0, i).Value & "'"
                Re_Count = DB_Execute()
                Product_Line_QC.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display_Ch = True
        MsgBox("저장되었습니다")
    End Sub

    Private Sub Btn_Add_Click(sender As Object, e As EventArgs) Handles Btn_Add.Click
        If Len(Grid_Info.Item(1, 0).Value) > 0 Then
        Else
            Exit Sub
        End If
        If Len(Grid_Pro.Item(1, Grid_Pro.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If
        '
        Grid_Display_Ch = False

        Dim DBT As New DataTable
        Dim Db_Sun As Long
        StrSQL = "Select PQ_Sun FROM SI_PRODUCT_QC with(NOLOCK) Where PQ_Code = '" & Grid_Info.Item(1, 0).Value & "' And PQ_PC_Code = '" & Grid_Pro.Item(1, Grid_Pro.CurrentCell.RowIndex).Value & "' Order By Convert(Decimal,PQ_Sun) Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            Db_Sun = 1
        Else
            Db_Sun = DBT.Rows(0).Item(0) + 1

            '이전 데이터 저장

            '수정한다
            Dim i As Integer


            For i = 0 To Product_Line_QC.Rows.Count - 1
                If Product_Line_QC.Rows(i).HeaderCell.Value = "*" Then
                    'Update
                    StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED " 'Replace(Grid_Info.Item(1, i).Value, "'", "''")
                    StrSQL = StrSQL & "UPDATE SI_PRODUCT_QC Set PQ_History1 = '" & Replace(Product_Line_QC.Item(1, i).Value, "'", "''") & "', PQ_History2 = '" & Replace(Product_Line_QC.Item(2, i).Value, "'", "''") & "',PQ_Bigo = '" & Product_Line_QC.Item(3, i).Value & "' where  PQ_Code = '" & Grid_Info.Item(1, 0).Value & "' And PQ_PC_Code = '" & Grid_Pro.Item(1, Grid_Pro.CurrentCell.RowIndex).Value & "' And PQ_Sun = '" & Product_Line_QC.Item(0, i).Value & "'"
                    Re_Count = DB_Execute()
                    Product_Line_QC.Rows(i).HeaderCell.Value = ""
                End If
            Next i
        End If

        ''추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO SI_PRODUCT_QC (PQ_Code, PQ_PC_Code, PQ_Sun)  Values ('" & Grid_Info.Item(1, 0).Value & "', '" & Grid_Pro.Item(1, Grid_Pro.CurrentCell.RowIndex).Value & "', '" & Db_Sun & "')"
        Re_Count = DB_Execute()
        Product_Line_QC_Display(Grid_Info.Item(1, 0).Value, Grid_Pro.Item(1, Grid_Pro.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True
    End Sub

    Private Sub Btn_Insert_Click(sender As Object, e As EventArgs) Handles Btn_Insert.Click
        If Len(Grid_Info.Item(1, 0).Value) > 0 Then
        Else
            Exit Sub
        End If
        If Len(Grid_Pro.Item(0, Grid_Pro.CurrentCell.RowIndex).Value) < 1 Then
            Exit Sub
        End If
        Grid_Display_Ch = False
        Dim i As Integer
        For i = Product_Line_QC.RowCount - 1 To Product_Line_QC.CurrentCell.RowIndex Step -1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "UPDATE SI_PRODUCT_QC Set PQ_Sun = '" & i + 2 & "' Where PQ_Sun= '" & i + 1 & "' And PQ_Code = '" & Grid_Info.Item(1, 0).Value & "' And PQ_PC_Code = '" & Grid_Pro.Item(1, Grid_Pro.CurrentCell.RowIndex).Value & "'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO SI_PRODUCT_QC (PQ_Code, PQ_PC_Code, PQ_Sun) Values ('" & Grid_Info.Item(1, 0).Value & "', '" & Grid_Pro.Item(1, Grid_Pro.CurrentCell.RowIndex).Value & "', '" & Product_Line_QC.CurrentCell.RowIndex + 1 & "')"
        Re_Count = DB_Execute()
        Product_Line_QC_Display(Grid_Info.Item(1, 0).Value, Grid_Pro.Item(1, Grid_Pro.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True
    End Sub

    Private Sub Btn_delete_Click(sender As Object, e As EventArgs) Handles Btn_delete.Click
        If Len(Grid_Info.Item(1, 0).Value) > 0 Then
        Else
            Exit Sub
        End If
        If Len(Grid_Pro.Item(0, Grid_Pro.CurrentCell.RowIndex).Value) < 1 Then
            Exit Sub
        End If

        If MsgBox("순번 '" & Product_Line_QC.Item(0, Product_Line_QC.CurrentCell.RowIndex).Value & "'을(를) 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "삭제") = vbNo Then
            Exit Sub
        End If

        Grid_Display_Ch = False
        Dim i As Integer

        ''삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_PRODUCT_QC Where PQ_Code = '" & Grid_Info.Item(1, 0).Value & "' And PQ_PC_Code = '" & Grid_Pro.Item(1, Grid_Pro.CurrentCell.RowIndex).Value & "' And   PQ_Sun = '" & Product_Line_QC.Item(0, Product_Line_QC.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        For i = Product_Line_QC.CurrentCell.RowIndex + 1 To Product_Line_QC.RowCount - 1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "UPDATE SI_PRODUCT_QC Set PQ_Sun = '" & i & "' Where  PQ_Code = '" & Grid_Info.Item(1, 0).Value & "' And PQ_PC_Code = '" & Grid_Pro.Item(1, Grid_Pro.CurrentCell.RowIndex).Value & "' And  PQ_Sun = '" & i + 1 & "'"
            Re_Count = DB_Execute()
        Next i

        '수정후

        'Dim count As Integer = 0


        'Dim DBT As New DataTable
        'Dim Db_Sun As Long
        'StrSQL = "Select PQ_Sun FROM SI_PRODUCT_QC with(NOLOCK) Where PQ_Code = '" & Grid_Info.Item(1, 0).Value & "' And PQ_PC_Code = '" & Grid_Pro.Item(1, Grid_Pro.CurrentCell.RowIndex).Value & "' Order By Convert(Decimal,PQ_Sun) Desc "
        'Re_Count = DB_Select(DBT)
        'If Re_Count <> 0 Then
        '    count = DBT.Rows(0)("PQ_Sun")
        'Else
        '    Exit Sub
        'End If


        'StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        'StrSQL = StrSQL & "DELETE SI_PRODUCT_QC Where PQ_Code = '" & Grid_Info.Item(1, 0).Value & "' And PQ_PC_Code = '" & Grid_Pro.Item(1, Grid_Pro.CurrentCell.RowIndex).Value & "' And PQ_Sun = '" & count & "'"
        'Re_Count = DB_Execute()


        Product_Line_QC_Display(Grid_Info.Item(1, 0).Value, Grid_Pro.Item(1, Grid_Pro.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True
    End Sub

    Private Sub Product_Line_QC_Combo_SelectionChangeCommitted(sender As Object, e As EventArgs)

    End Sub

    Private Sub Product_Line_QC_MouseClick(sender As Object, e As MouseEventArgs) Handles Product_Line_QC.MouseClick
        'Product_Line_QC_Row = Product_Line_QC.CurrentCell.RowIndex
        'Product_Line_QC_Col = Product_Line_QC.CurrentCell.ColumnIndex
        'Product_Line_QC_Combo.Visible = False
        'If Product_Line_QC_Row < 0 Then
        '    Exit Sub
        'End If
        'If Product_Line_QC_Col < 0 Then
        '    Exit Sub
        'End If
        'Dim DBT As New DataTable

        'Select Case Product_Line_QC_Col
        '    Case 1
        '        StrSQL = "Select PQ_History1, PQ_History2  FROM SI_PRODUCT_QC with(NOLOCK)  Order By PQ_History1"
        '        Re_Count = DB_Select(DBT)
        '        Product_Line_QC_Combo.Items.Clear()
        '        If Re_Count <> 0 Then
        '            For i = 0 To Re_Count - 1
        '                Product_Line_QC_Combo.Items.Add(DBT.Rows(i).Item(0))
        '            Next i
        '        End If

        '        Product_Line_QC_Combo.Top = Product_Line_QC.Top + Product_Line_QC.ColumnHeadersHeight + (Product_Line_QC_Row) * 24
        '        Product_Line_QC_Combo.Left = Product_Line_QC.Left + Grid_Pro.RowHeadersWidth + Grid_Pro.Columns(0).Width
        '        Product_Line_QC_Combo.Width = Product_Line_QC.Columns(Product_Line_QC_Col).Width
        '        Product_Line_QC_Combo.Text = Product_Line_QC.CurrentCell.Value.ToString
        '        Product_Line_QC_Combo.BackColor = Product_Line_QC.RowsDefaultCellStyle.BackColor
        '        Product_Line_QC_Combo.Visible = True
        '        Product_Line_QC_Combo.Focus.ToString()
        '    Case 2
        '        StrSQL = "Select PQ_History1, PQ_History2  FROM SI_PRODUCT_QC with(NOLOCK)  Order By PQ_History2"

        '        Re_Count = DB_Select(DBT)
        '        Product_Line_QC_Combo.Items.Clear()
        '        If Re_Count <> 0 Then
        '            For i = 0 To Re_Count - 1
        '                Product_Line_QC_Combo.Items.Add(DBT.Rows(i).Item(1))
        '            Next i
        '        End If
        '        Product_Line_QC_Combo.Top = Product_Line_QC.Top + Product_Line_QC.ColumnHeadersHeight + (Product_Line_QC_Row) * 24
        '        Product_Line_QC_Combo.Left = Product_Line_QC.Left + Product_Line_QC.RowHeadersWidth + Product_Line_QC.Columns(0).Width + Product_Line_QC.Columns(1).Width
        '        Product_Line_QC_Combo.Width = Product_Line_QC.Columns(Product_Line_QC_Col).Width
        '        Product_Line_QC_Combo.Text = Product_Line_QC.CurrentCell.Value.ToString
        '        Product_Line_QC_Combo.BackColor = Product_Line_QC.RowsDefaultCellStyle.BackColor
        '        Product_Line_QC_Combo.Visible = True
        '        Product_Line_QC_Combo.Focus.ToString()
        'End Select
    End Sub
End Class