﻿Imports System.IO

Public Class Frm_Product3
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

    '체크박스 셀 용

    Private Sub Frm_Product_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()
        Me.BackColor = Color.White



        Grid_Code_Setup()
        Grid_Info_Setup()
        ' Grid_LG_Setup()

        Search_Combo.Items.Add("코드")
        Search_Combo.Items.Add("품명")
        Search_Combo.Text = "코드"

        Dim DBT As New DataTable
        Search_Sel_Combo.Items.Clear()
        Search_Sel_Combo.Items.Add("전체")
        'StrSQL = "Select Base_Name  FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '제품구분' Order By Base_Name"
        'Re_Count = DB_Select(DBT)
        ''Grid_Info_Combo.Items.Clear()
        'If Re_Count <> 0 Then
        '    For i = 0 To Re_Count - 1
        '        Search_Sel_Combo.Items.Add(DBT.Rows(i)("Base_Name"))
        '    Next i
        'End If
        Search_Sel_Combo.Text = "전체"


        '  Grid_Cus_In_Setup()
        Grid_Cus_Out_Setup()



        Customer_Panel.Left = 391
        Customer_Panel.Top = 127
        Customer_Panel.Visible = True

        Grid_LG.Visible = False
        Panel2.Visible = False
        Grid_Info_Picture1.Visible = False
        'Me.Com_Customer.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange


        Grid_Code_Display(True)
        Grid_Code.ClearSelection()

        Button14.Visible = False
        Grid_Cus_In.Visible = False
        Grid_Cus_In_Combo.Visible = False
        Panel7.Visible = False
        Panel1.Visible = False
    End Sub
    Public Function Grid_Cus_In_Setup() As Boolean


        Grid_Color(Grid_Cus_In)

        Grid_Cus_In.EnableHeadersVisualStyles = False

        Grid_Cus_In.AllowUserToAddRows = False




        'Grid_Cus_In.RowTemplate.Height = 20
        Grid_Cus_In.RowTemplate.Height = 24
        '순번, 거래처코드, 거래처명,  단위, 단가, 기초수량, 비고
        Grid_Cus_In.ColumnCount = 7
        Grid_Cus_In.RowCount = 1


        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Cus_In.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Cus_In.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Cus_In.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Cus_In.RowHeadersWidth = 10
        Grid_Cus_In.Columns(0).Width = 55
        Grid_Cus_In.Columns(1).Width = 55
        Grid_Cus_In.Columns(2).Width = 110
        Grid_Cus_In.Columns(3).Width = 60
        Grid_Cus_In.Columns(4).Width = 60
        Grid_Cus_In.Columns(5).Width = 60
        Grid_Cus_In.Columns(6).Width = 100

        Grid_Cus_In.Columns(0).HeaderText = "순번"
        Grid_Cus_In.Columns(1).HeaderText = "코드"
        Grid_Cus_In.Columns(2).HeaderText = "거래처상호"
        Grid_Cus_In.Columns(3).HeaderText = "단위"
        Grid_Cus_In.Columns(4).HeaderText = "단가"
        Grid_Cus_In.Columns(5).HeaderText = "기초수량"
        Grid_Cus_In.Columns(6).HeaderText = "비고"

    End Function
    Public Function Grid_Cus_Out_Setup() As Boolean
        Grid_Color(Grid_Cus_Out)
        Grid_Cus_Out.EnableHeadersVisualStyles = False

        Grid_Cus_Out.AllowUserToAddRows = False




        'Grid_Cus_Out.RowTemplate.Height = 20
        Grid_Cus_Out.RowTemplate.Height = 24
        '순번, 거래처코드, 거래처명,  단위, 단가, 기초수량, 비고
        Grid_Cus_Out.ColumnCount = 7
        Grid_Cus_Out.RowCount = 1


        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Cus_Out.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Cus_Out.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Cus_Out.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Cus_Out.RowHeadersWidth = 10

        Grid_Cus_Out.Columns(0).Width = 55
        Grid_Cus_Out.Columns(1).Width = 55
        Grid_Cus_Out.Columns(2).Width = 110
        Grid_Cus_Out.Columns(3).Width = 50
        Grid_Cus_Out.Columns(4).Width = 60
        Grid_Cus_Out.Columns(5).Width = 60
        Grid_Cus_Out.Columns(6).Width = 108
        '   Grid_Cus_Out.Columns(7).Width = 50

        '이전
        Grid_Cus_Out.Columns(0).HeaderText = "순번"
        Grid_Cus_Out.Columns(1).HeaderText = "코드"
        Grid_Cus_Out.Columns(2).HeaderText = "거래처상호"
        Grid_Cus_Out.Columns(3).HeaderText = "단위"
        Grid_Cus_Out.Columns(4).HeaderText = "단가"
        Grid_Cus_Out.Columns(5).HeaderText = "기초수량"
        Grid_Cus_Out.Columns(6).HeaderText = "비고"

        '체크박스
        '   Dim check As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()


        ' Grid_Cus_Out.Columns.Add(check)



        Grid_Cus_Out_Setup = True

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


    Public Sub Grid_LG_Setup()
        Grid_LG.GridColor = Color.White
        Grid_LG.BackgroundColor = Color.White
        Grid_LG.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        Grid_LG.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(91, 155, 213)
        Grid_LG.RowsDefaultCellStyle.BackColor = Color.FromArgb(210, 222, 239)
        Grid_LG.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(234, 239, 247)
        Grid_LG.AllowUserToAddRows = False
        Grid_LG.RowTemplate.Height = 24

        Grid_LG.ColumnCount = 2
        Grid_LG.RowCount = 1

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_LG.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Grid_LG.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Grid_LG.RowHeadersWidth = 40
        Grid_LG.Columns(0).Width = 40
        Grid_LG.Columns(1).Width = 150

        Grid_LG.Columns(0).HeaderText = "순번"
        Grid_LG.Columns(1).HeaderText = "LG 코드"


        Grid_LG.Item(0, 0).Value = ""
        Grid_LG.Item(1, 0).Value = ""

        Grid_LG.Columns(0).ReadOnly = True

    End Sub

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
        ' Grid_Info.Item(0, 7).Value = "이미지"
        '   Grid_Info.Rows(7).Height = 110
        ' Grid_Info_Picture_Setup()

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
                    StrSQL = "Select PL_Code, PL_Name, PL_Sel FROM SI_PRODUCT with(NOLOCK) Where PL_Code Like '%" & Search_Text.Text & "%' " & Je_Gu & " and PL_SEL!='단품'  Order  By PL_Code"
                Case "품명"
                    StrSQL = "Select PL_Code, PL_Name, PL_Sel FROM SI_PRODUCT With(NOLOCK) Where PL_Name Like '%" & Search_Text.Text & "%' and PL_Name Is not Null " & Je_Gu & " Order By PL_Name"
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

            Label_Color(Com_Customer, Color_Label, Di_Panel2)
            'BOM 입력 페널
            Grid_Cus_In_Display(Grid_Info.Item(1, 0).Value)
            Grid_Cus_Out_Display(Grid_Info.Item(1, 0).Value)
            ' Grid_LG_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
            'Customer_Panel.Visible = False


        End If
    End Sub


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
                Case 7
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
                Case 8
                    ''"수입검사", "순회검사", "출하검사", "Xbar-R 관리도", "부적합보고서"
                    StrSQL = "Select QC_Name  FROM SI_QC with(NOLOCK) Where QC_Sel = '공정검사' Order By QC_Name"
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


                Case Else
                    Exit Sub
            End Select
        End If


    End Sub


    Private Sub Grid_Code_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Code.CellContentClick

    End Sub

    Private Sub Save_Com_Click(sender As Object, e As EventArgs)

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


    Public Function Grid_Cus_In_Display(PL_Code As String) As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display_Ch = False
        Grid_Cus_In.RowCount = 1
        StrSQL = "Select PL_Sun, PL_CM_Code, PL_CM_Name, PL_Unit, PL_Unit_Gold, PL_Base_Amount, PL_Bigo   FROM SI_PRODUCT_Customer with(NOLOCK)  Where PL_Code = '" & PL_Code & "' And PL_Sel = '매입'  Order By Convert(Decimal,PL_Sun)"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Cus_In.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                Grid_Cus_In.Item(0, i).Value = DBT.Rows(i).Item(0).ToString
                Grid_Cus_In.Item(1, i).Value = DBT.Rows(i).Item(1).ToString
                Grid_Cus_In.Item(2, i).Value = DBT.Rows(i).Item(2).ToString
                Grid_Cus_In.Item(3, i).Value = DBT.Rows(i).Item(3).ToString
                Grid_Cus_In.Item(4, i).Value = DBT.Rows(i).Item(4).ToString
                Grid_Cus_In.Item(5, i).Value = DBT.Rows(i).Item(5).ToString
                Grid_Cus_In.Item(6, i).Value = DBT.Rows(i).Item(6).ToString
            Next i
        Else
            Grid_Cus_In.RowCount = 1
            For j = 0 To Grid_Cus_In.ColumnCount - 1
                Grid_Cus_In.Item(j, 0).Value = ""
            Next j

        End If

        Grid_Display_Ch = True
        Grid_Cus_In_Display = True
    End Function

    Private Sub Insert__Cus_In_Click(sender As Object, e As EventArgs) Handles Insert__Cus_In.Click
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
            For i = 0 To Grid_Cus_In.Rows.Count - 1
                If Grid_Cus_In.Rows(i).HeaderCell.Value = "*" Then
                    'Update
                    StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                    StrSQL = StrSQL & "UPDATE SI_PRODUCT_Customer Set PL_Cm_Code = '" & Grid_Cus_In.Item(1, i).Value & "', PL_CM_Name = '" & Grid_Cus_In.Item(2, i).Value & "', PL_Unit = '" & Grid_Cus_In.Item(3, i).Value & "', PL_Unit_Gold = '" & Grid_Cus_In.Item(4, i).Value & "', PL_Base_Amount = '" & Grid_Cus_In.Item(5, i).Value & "', PL_Bigo = '" & Grid_Cus_In.Item(6, i).Value & "' where PL_Code = '" & Grid_Info.Item(1, 0).Value & "' And PL_Sun = '" & Grid_Cus_In.Item(0, i).Value & "' And PL_Sel = '매입'"
                    Re_Count = DB_Execute()
                    Grid_Cus_In.Rows(i).HeaderCell.Value = ""

                End If
            Next i
        End If

        'MessageBox.Show(Grid_Info.Item(1, 0).Value)

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO SI_PRODUCT_Customer (PL_Sel, PL_Code, PL_Sun)  Values ('매입','" & Grid_Info.Item(1, 0).Value & "', '" & Db_Sun & "')"
        Re_Count = DB_Execute()
        Grid_Cus_In_Display(Grid_Info.Item(1, 0).Value)
        Grid_Display_Ch = True

    End Sub



    Private Sub Grid_Cus_In_MouseClick(sender As Object, e As MouseEventArgs) Handles Grid_Cus_In.MouseClick
        Grid_Cus_In_Row = Grid_Cus_In.CurrentCell.RowIndex
        Grid_Cus_In_Col = Grid_Cus_In.CurrentCell.ColumnIndex
        Grid_Cus_In_Combo.Visible = False
        If Grid_Cus_In_Row < 0 Then
            Exit Sub
        End If
        If Grid_Cus_In_Col < 0 Then
            Exit Sub
        End If
        Dim DBT As New DataTable

        Select Case Grid_Cus_In_Col
            Case 1
                StrSQL = "Select CM_Code, CM_Name  FROM SI_CUSTOMER with(NOLOCK)  Order By CM_Code"
                Re_Count = DB_Select(DBT)
                Grid_Cus_In_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        Grid_Cus_In_Combo.Items.Add(DBT.Rows(i).Item(0))
                    Next i
                End If

                Grid_Cus_In_Combo.Top = Grid_Cus_In.Top + Grid_Cus_In.ColumnHeadersHeight + (Grid_Cus_In_Row) * 24
                Grid_Cus_In_Combo.Left = Grid_Cus_In.Left + Grid_Cus_In.RowHeadersWidth + Grid_Cus_In.Columns(0).Width
                Grid_Cus_In_Combo.Width = Grid_Cus_In.Columns(Grid_Cus_In_Col).Width
                If Len(Grid_Cus_In.CurrentCell.Value) < 1 Then
                    Grid_Cus_In_Combo.Text = ""
                Else
                    Grid_Cus_In_Combo.Text = Grid_Cus_In.CurrentCell.Value.ToString

                End If
                Grid_Cus_In_Combo.BackColor = Grid_Cus_In.RowsDefaultCellStyle.BackColor
                Grid_Cus_In_Combo.Visible = True
                Grid_Cus_In_Combo.Focus.ToString()
            Case 2
                StrSQL = "Select CM_Code, CM_Name  FROM SI_CUSTOMER with(NOLOCK)  Order By CM_Name"

                Re_Count = DB_Select(DBT)
                Grid_Cus_In_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        Grid_Cus_In_Combo.Items.Add(DBT.Rows(i).Item(1))
                    Next i
                End If
                Grid_Cus_In_Combo.Top = Grid_Cus_In.Top + Grid_Cus_In.ColumnHeadersHeight + (Grid_Cus_In_Row) * 24
                Grid_Cus_In_Combo.Left = Grid_Cus_In.Left + Grid_Cus_In.RowHeadersWidth + Grid_Cus_In.Columns(0).Width + Grid_Cus_In.Columns(1).Width
                Grid_Cus_In_Combo.Width = Grid_Cus_In.Columns(Grid_Cus_In_Col).Width
                Grid_Cus_In_Combo.Text = Grid_Cus_In.CurrentCell.Value
                Grid_Cus_In_Combo.BackColor = Grid_Cus_In.RowsDefaultCellStyle.BackColor
                Grid_Cus_In_Combo.Visible = True
                Grid_Cus_In_Combo.Focus.ToString()
            Case 3
                StrSQL = "Select Base_Name  FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '단위'  Order By Base_Name"
                Re_Count = DB_Select(DBT)
                Grid_Cus_In_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        Grid_Cus_In_Combo.Items.Add(DBT.Rows(i).Item(0))
                    Next i
                End If
                Grid_Cus_In_Combo.Top = Grid_Cus_In.Top + Grid_Cus_In.ColumnHeadersHeight + (Grid_Cus_In_Row) * 24
                Grid_Cus_In_Combo.Left = Grid_Cus_In.Left + Grid_Cus_In.RowHeadersWidth + Grid_Cus_In.Columns(0).Width + Grid_Cus_In.Columns(1).Width + Grid_Cus_In.Columns(2).Width
                Grid_Cus_In_Combo.Width = Grid_Cus_In.Columns(Grid_Cus_In_Col).Width
                Grid_Cus_In_Combo.Text = Grid_Cus_In.CurrentCell.Value
                Grid_Cus_In_Combo.BackColor = Grid_Cus_In.RowsDefaultCellStyle.BackColor
                Grid_Cus_In_Combo.Visible = True
                Grid_Cus_In_Combo.Focus.ToString()
        End Select
    End Sub

    Private Sub Grid_Cus_In_Combo_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Grid_Cus_In_Combo_LostFocus(sender As Object, e As EventArgs)
        Grid_Cus_In_Combo.Visible = False

    End Sub

    Private Sub Grid_Cus_In_Combo_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Grid_Cus_In_Combo.SelectionChangeCommitted
        Dim DBT As New DataTable
        Grid_Cus_In.Item(Grid_Cus_In_Col, Grid_Cus_In_Row).Value = Grid_Cus_In_Combo.SelectedItem.ToString()
        Select Case Grid_Cus_In_Col
            Case 1
                StrSQL = "Select CM_Code, CM_Name  FROM SI_CUSTOMER with(NOLOCK) Where CM_Code = '" & Grid_Cus_In_Combo.SelectedItem.ToString() & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Cus_In.Item(2, Grid_Cus_In_Row).Value = DBT.Rows(0).Item(1)
                End If
            Case 2
                StrSQL = "Select CM_Code, CM_Name  FROM SI_CUSTOMER with(NOLOCK) Where CM_Name = '" & Grid_Cus_In_Combo.SelectedItem.ToString() & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Cus_In.Item(1, Grid_Cus_In_Row).Value = DBT.Rows(0).Item(0)
                End If
        End Select

    End Sub



    Private Sub Grid_Cus_In_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Cus_In.CellValueChanged
        If Grid_Display_Ch = True Then
            Grid_Cus_In.CurrentRow.HeaderCell.Value = "*"
        End If

    End Sub


    Private Sub Save_Cus_In_Click(sender As Object, e As EventArgs) Handles Save_Cus_In.Click
        If Len(Grid_Info.Item(1, 0).Value) > 0 Then
        Else
            Exit Sub
        End If

        Dim i As Integer
        Grid_Display_Ch = False
        For i = 0 To Grid_Cus_In.Rows.Count - 1
            If Grid_Cus_In.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SI_PRODUCT_Customer Set PL_Cm_Code = '" & Grid_Cus_In.Item(1, i).Value & "', PL_CM_Name = '" & Grid_Cus_In.Item(2, i).Value & "', PL_Unit = '" & Grid_Cus_In.Item(3, i).Value & "', PL_Unit_Gold = '" & Grid_Cus_In.Item(4, i).Value & "', PL_Base_Amount = '" & Grid_Cus_In.Item(5, i).Value & "', PL_Bigo = '" & Grid_Cus_In.Item(6, i).Value & "' where PL_Code = '" & Grid_Info.Item(1, 0).Value & "' And PL_Sun = '" & Grid_Cus_In.Item(0, i).Value & "' And PL_Sel = '매입'"
                Re_Count = DB_Execute()
                Grid_Cus_In.Rows(i).HeaderCell.Value = ""

            End If
        Next i
        Grid_Display_Ch = True
        MessageBox.Show("저장되었습니다")
    End Sub

    Private Sub Put__Cus_In_Click(sender As Object, e As EventArgs) Handles Put__Cus_In.Click
        If Len(Grid_Info.Item(1, 0).Value) > 0 Then
        Else
            Exit Sub
        End If

        Grid_Display_Ch = False
        Dim i As Integer
        For i = Grid_Cus_In.RowCount - 1 To Grid_Cus_In.CurrentCell.RowIndex Step -1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "UPDATE SI_PRODUCT_Customer Set PL_Sun = '" & i + 2 & "' Where PL_Sun = '" & i + 1 & "' And PL_Code = '" & Grid_Info.Item(1, 0).Value & "' And PL_Sel = '매입'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO SI_PRODUCT_Customer (PL_Sel, PL_Code, PL_Sun)  Values ('매입', '" & Grid_Info.Item(1, 0).Value & "', '" & Grid_Cus_In.CurrentCell.RowIndex + 1 & "')"
        Re_Count = DB_Execute()
        Grid_Cus_In_Display(Grid_Info.Item(1, 0).Value)
        Grid_Display_Ch = True

    End Sub

    Private Sub Delete__Cus_In_Click(sender As Object, e As EventArgs) Handles Delete__Cus_In.Click
        '해당 칼럼 삭제
        If Len(Grid_Info.Item(1, 0).Value) > 0 Then
        Else
            Exit Sub
        End If

        Grid_Display_Ch = False
        Dim i As Integer
        '수정전
        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_PRODUCT_Customer Where PL_Sel = '매입' And PL_Code = '" & Grid_Info.Item(1, 0).Value & "' And   PL_Sun = '" & Grid_Cus_In.Item(0, Grid_Cus_In.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        For i = Grid_Cus_In.CurrentCell.RowIndex + 1 To Grid_Cus_In.RowCount - 1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "UPDATE SI_PRODUCT_Customer Set PL_Sun = '" & i & "' Where  PL_Sel = '매입' And PL_Code = '" & Grid_Info.Item(1, 0).Value & "' And  PL_Sun = '" & i + 1 & "'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        Grid_Cus_In_Display(Grid_Info.Item(1, 0).Value)
        Grid_Display_Ch = True

        '수정후
        'Dim DBT As New DataTable
        'Dim count As Integer = 0

        'StrSQL = "Select PL_Sun FROM SI_PRODUCT_Customer with(NOLOCK) Where PL_Code = '" & Grid_Info.Item(1, 0).Value & "' And PL_Sel = '매입' Order By Convert(Decimal,PL_Sun) Desc "
        'Re_Count = DB_Select(DBT)
        'If Re_Count <> 0 Then
        '    count = DBT.Rows(0)("PL_Sun")
        'Else
        '    Exit Sub
        'End If

        'StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        'StrSQL = StrSQL & "DELETE SI_PRODUCT_Customer Where PL_Sel = '매입' And PL_Code = '" & Grid_Info.Item(1, 0).Value & "' And   PL_Sun = '" & count & "'"
        'Re_Count = DB_Execute()

        ''For i = Grid_Cus_In.CurrentCell.RowIndex + 1 To Grid_Cus_In.RowCount - 1
        ''    '수정한다
        ''    StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        ''    StrSQL = StrSQL & "UPDATE SI_PRODUCT_Customer Set PL_Sun = '" & i & "' Where  PL_Sel = '매입' And PL_Code = '" & Grid_Info.Item(1, 0).Value & "' And  PL_Sun = '" & i + 1 & "'"
        ''    Re_Count = DB_Execute()
        ''Next i   '선택된 갈럼값을 가지고 온다

        'Grid_Cus_In_Display(Grid_Info.Item(1, 0).Value)
        'Grid_Display_Ch = True


    End Sub

    Private Sub Grid_Cus_Out_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub Grid_Cus_Out_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Cus_Out.CellValueChanged
        If Grid_Display_Ch = True Then
            Grid_Cus_Out.CurrentRow.HeaderCell.Value = "*"
        End If

    End Sub

    Private Sub Grid_Cus_Out_MouseClick(sender As Object, e As MouseEventArgs) Handles Grid_Cus_Out.MouseClick
        Grid_Cus_Out_Row = Grid_Cus_Out.CurrentCell.RowIndex
        Grid_Cus_Out_Col = Grid_Cus_Out.CurrentCell.ColumnIndex
        Grid_Cus_Out_Combo.Visible = False
        If Grid_Cus_Out_Row < 0 Then
            Exit Sub
        End If
        If Grid_Cus_Out_Col < 0 Then
            Exit Sub
        End If
        Dim DBT As New DataTable

        Select Case Grid_Cus_Out_Col
            Case 1
                StrSQL = "Select CM_Code, CM_Name  FROM SI_CUSTOMER with(NOLOCK)  Order By CM_Code"
                Re_Count = DB_Select(DBT)
                Grid_Cus_Out_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        Grid_Cus_Out_Combo.Items.Add(DBT.Rows(i).Item(0))
                    Next i
                End If

                Grid_Cus_Out_Combo.Top = Grid_Cus_Out.Top + Grid_Cus_Out.ColumnHeadersHeight + (Grid_Cus_Out_Row) * 24
                Grid_Cus_Out_Combo.Left = Grid_Cus_Out.Left + Grid_Cus_Out.RowHeadersWidth + Grid_Cus_Out.Columns(0).Width
                Grid_Cus_Out_Combo.Width = Grid_Cus_Out.Columns(Grid_Cus_Out_Col).Width
                Grid_Cus_Out_Combo.Text = Grid_Cus_Out.CurrentCell.Value
                Grid_Cus_Out_Combo.BackColor = Grid_Cus_Out.RowsDefaultCellStyle.BackColor
                Grid_Cus_Out_Combo.Visible = True
                Grid_Cus_Out_Combo.Focus.ToString()
            Case 2
                StrSQL = "Select CM_Code, CM_Name  FROM SI_CUSTOMER with(NOLOCK)  Order By CM_Name"

                Re_Count = DB_Select(DBT)
                Grid_Cus_Out_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        Grid_Cus_Out_Combo.Items.Add(DBT.Rows(i).Item(1))
                    Next i
                End If
                Grid_Cus_Out_Combo.Top = Grid_Cus_Out.Top + Grid_Cus_Out.ColumnHeadersHeight + (Grid_Cus_Out_Row) * 24
                Grid_Cus_Out_Combo.Left = Grid_Cus_Out.Left + Grid_Cus_Out.RowHeadersWidth + Grid_Cus_Out.Columns(0).Width + Grid_Cus_Out.Columns(1).Width
                Grid_Cus_Out_Combo.Width = Grid_Cus_Out.Columns(Grid_Cus_Out_Col).Width
                Grid_Cus_Out_Combo.Text = Grid_Cus_Out.CurrentCell.Value
                Grid_Cus_Out_Combo.BackColor = Grid_Cus_Out.RowsDefaultCellStyle.BackColor
                Grid_Cus_Out_Combo.Visible = True
                Grid_Cus_Out_Combo.Focus.ToString()
            Case 3
                StrSQL = "Select Base_Name  FROM SI_BASE_CODE with(NOLOCK) Where Base_Sel = '단위'  Order By Base_Name"
                Re_Count = DB_Select(DBT)
                Grid_Cus_Out_Combo.Items.Clear()
                If Re_Count <> 0 Then
                    For i = 0 To Re_Count - 1
                        Grid_Cus_Out_Combo.Items.Add(DBT.Rows(i).Item(0))
                    Next i
                End If
                Grid_Cus_Out_Combo.Top = Grid_Cus_Out.Top + Grid_Cus_Out.ColumnHeadersHeight + (Grid_Cus_Out_Row) * 24
                Grid_Cus_Out_Combo.Left = Grid_Cus_Out.Left + Grid_Cus_Out.RowHeadersWidth + Grid_Cus_Out.Columns(0).Width + Grid_Cus_Out.Columns(1).Width + Grid_Cus_Out.Columns(2).Width
                Grid_Cus_Out_Combo.Width = Grid_Cus_Out.Columns(Grid_Cus_Out_Col).Width
                Grid_Cus_Out_Combo.Text = Grid_Cus_Out.CurrentCell.Value
                Grid_Cus_Out_Combo.BackColor = Grid_Cus_Out.RowsDefaultCellStyle.BackColor
                Grid_Cus_Out_Combo.Visible = True
                Grid_Cus_Out_Combo.Focus.ToString()
        End Select

    End Sub

    Private Sub Grid_Cus_Out_Combo_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Grid_Cus_Out_Combo_LostFocus(sender As Object, e As EventArgs)
        Grid_Cus_Out_Combo.Visible = False

    End Sub

    Private Sub Grid_Cus_Out_Combo_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Grid_Cus_Out_Combo.SelectionChangeCommitted
        Dim DBT As New DataTable
        Grid_Cus_Out.Item(Grid_Cus_Out_Col, Grid_Cus_Out_Row).Value = Grid_Cus_Out_Combo.SelectedItem.ToString()
        Select Case Grid_Cus_Out_Col
            Case 1
                StrSQL = "Select CM_Code, CM_Name  FROM SI_CUSTOMER with(NOLOCK) Where CM_Code = '" & Grid_Cus_Out_Combo.SelectedItem.ToString() & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Cus_Out.Item(2, Grid_Cus_Out_Row).Value = DBT.Rows(0).Item(1)
                End If
            Case 2
                StrSQL = "Select CM_Code, CM_Name  FROM SI_CUSTOMER with(NOLOCK) Where CM_Name = '" & Grid_Cus_Out_Combo.SelectedItem.ToString() & "'"
                Re_Count = DB_Select(DBT)
                If Re_Count <> 0 Then
                    Grid_Cus_Out.Item(1, Grid_Cus_Out_Row).Value = DBT.Rows(0).Item(0)
                End If
        End Select

    End Sub

    Private Sub Save_Cus_Out_Click(sender As Object, e As EventArgs) Handles Save_Cus_Out.Click
        If Len(Grid_Info.Item(1, 0).Value) > 0 Then
        Else
            Exit Sub
        End If

        Dim i As Integer
        Grid_Display_Ch = False
        For i = 0 To Grid_Cus_Out.Rows.Count - 1
            If Grid_Cus_Out.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SI_PRODUCT_Customer Set PL_Cm_Code = '" & Grid_Cus_Out.Item(1, i).Value & "', PL_CM_Name = '" & Grid_Cus_Out.Item(2, i).Value & "', PL_Unit = '" & Grid_Cus_Out.Item(3, i).Value & "', PL_Unit_Gold = '" & Grid_Cus_Out.Item(4, i).Value & "', PL_Base_Amount = '" & Grid_Cus_Out.Item(5, i).Value & "', PL_Bigo = '" & Grid_Cus_Out.Item(6, i).Value & "' where PL_Code = '" & Grid_Info.Item(1, 0).Value & "' And PL_Sun = '" & Grid_Cus_Out.Item(0, i).Value & "' And PL_Sel = '매출'"
                Re_Count = DB_Execute()
                Grid_Cus_Out.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        Grid_Display_Ch = True
        MsgBox("저장되었습니다")
    End Sub

    Private Sub Insert__Cus_Out_Click(sender As Object, e As EventArgs) Handles Insert__Cus_Out.Click
        If Len(Grid_Info.Item(1, 0).Value) > 0 Then
        Else
            Exit Sub
        End If

        Grid_Display_Ch = False

        Dim DBT As New DataTable
        Dim Db_Sun As Long
        StrSQL = "Select PL_Sun FROM SI_PRODUCT_Customer with(NOLOCK) Where PL_Code = '" & Grid_Info.Item(1, 0).Value & "' And PL_Sel = '매출' Order By Convert(Decimal,PL_Sun) Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            Db_Sun = 1
        Else
            Db_Sun = DBT.Rows(0).Item(0) + 1

            For i = 0 To Grid_Cus_Out.Rows.Count - 1
                If Grid_Cus_Out.Rows(i).HeaderCell.Value = "*" Then
                    'Update
                    StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                    StrSQL = StrSQL & "UPDATE SI_PRODUCT_Customer Set PL_Cm_Code = '" & Grid_Cus_Out.Item(1, i).Value & "', PL_CM_Name = '" & Grid_Cus_Out.Item(2, i).Value & "', PL_Unit = '" & Grid_Cus_Out.Item(3, i).Value & "', PL_Unit_Gold = '" & Grid_Cus_Out.Item(4, i).Value & "', PL_Base_Amount = '" & Grid_Cus_Out.Item(5, i).Value & "', PL_Bigo = '" & Grid_Cus_Out.Item(6, i).Value & "' where PL_Code = '" & Grid_Info.Item(1, 0).Value & "' And PL_Sun = '" & Grid_Cus_Out.Item(0, i).Value & "' And PL_Sel = '매출'"
                    Re_Count = DB_Execute()
                    Grid_Cus_Out.Rows(i).HeaderCell.Value = ""
                End If
            Next i
        End If

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO SI_PRODUCT_Customer (PL_Sel, PL_Code, PL_Sun)  Values ('매출','" & Grid_Info.Item(1, 0).Value & "', '" & Db_Sun & "')"
        Re_Count = DB_Execute()
        Grid_Cus_Out_Display(Grid_Info.Item(1, 0).Value)
        Grid_Display_Ch = True

    End Sub

    Private Sub Put__Cus_Out_Click(sender As Object, e As EventArgs) Handles Put__Cus_Out.Click
        If Len(Grid_Info.Item(1, 0).Value) > 0 Then
        Else
            Exit Sub
        End If

        Grid_Display_Ch = False
        Dim i As Integer
        For i = Grid_Cus_Out.RowCount - 1 To Grid_Cus_Out.CurrentCell.RowIndex Step -1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "UPDATE SI_PRODUCT_Customer Set PL_Sun = '" & i + 2 & "' Where PL_Sun = '" & i + 1 & "' And PL_Code = '" & Grid_Info.Item(1, 0).Value & "' And PL_Sel = '매출'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO SI_PRODUCT_Customer (PL_Sel, PL_Code, PL_Sun)  Values ('매출', '" & Grid_Info.Item(1, 0).Value & "', '" & Grid_Cus_Out.CurrentCell.RowIndex + 1 & "')"
        Re_Count = DB_Execute()
        Grid_Cus_Out_Display(Grid_Info.Item(1, 0).Value)
        Grid_Display_Ch = True

    End Sub

    Private Sub Delete__Cus_Out_Click(sender As Object, e As EventArgs) Handles Delete__Cus_Out.Click
        '해당 칼럼 삭제
        If Len(Grid_Info.Item(1, 0).Value) > 0 Then
        Else
            Exit Sub
        End If

        If MsgBox("'" & Grid_Cus_Out.Item(2, Grid_Cus_Out.CurrentCell.RowIndex).Value & "'을(를) 삭제 하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1, "삭제") = vbNo Then
            Exit Sub
        End If

        Grid_Display_Ch = False
        Dim i As Integer

        '수정전

        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_PRODUCT_Customer Where PL_Sel = '매출' And PL_Code = '" & Grid_Info.Item(1, 0).Value & "' And   PL_Sun = '" & Grid_Cus_Out.Item(0, Grid_Cus_Out.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()

        For i = Grid_Cus_Out.CurrentCell.RowIndex + 1 To Grid_Cus_Out.RowCount - 1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "UPDATE SI_PRODUCT_Customer Set PL_Sun = '" & i & "' Where  PL_Sel = '매출' And PL_Code = '" & Grid_Info.Item(1, 0).Value & "' And  PL_Sun = '" & i + 1 & "'"
            Re_Count = DB_Execute()
        Next i
        '선택된 갈럼값을 가지고 온다
        Grid_Cus_Out_Display(Grid_Info.Item(1, 0).Value)
        Grid_Display_Ch = True

        'Dim DBT As New DataTable
        'Dim count As Integer = 0

        'StrSQL = "Select PL_Sun FROM SI_PRODUCT_Customer with(NOLOCK) Where PL_Code = '" & Grid_Info.Item(1, 0).Value & "' And PL_Sel = '매출' Order By Convert(Decimal,PL_Sun) Desc "
        'Re_Count = DB_Select(DBT)
        'If Re_Count <> 0 Then
        '    count = DBT.Rows(0)("PL_Sun")
        'Else
        '    Exit Sub
        'End If
        '' MsgBox(count)
        ''수정후
        'StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        'StrSQL = StrSQL & "DELETE SI_PRODUCT_Customer Where PL_Sel = '매출' And PL_Code = '" & Grid_Info.Item(1, 0).Value & "' And   PL_Sun = '" & count & "'"
        'Re_Count = DB_Execute()

        'For i = Grid_Cus_Out.CurrentCell.RowIndex + 1 To Grid_Cus_Out.RowCount - 1
        '    '수정한다
        '    StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        '    StrSQL = StrSQL & "UPDATE SI_PRODUCT_Customer Set PL_Sun = '" & i & "' Where  PL_Sel = '매출' And PL_Code = '" & Grid_Info.Item(1, 0).Value & "' And  PL_Sun = '" & i + 1 & "'"
        '    Re_Count = DB_Execute()
        'Next i
        '선택된 갈럼값을 가지고 온다
        'Grid_Cus_Out_Display(Grid_Info.Item(1, 0).Value)
        'Grid_Display_Ch = True

    End Sub
    Public Function Grid_Cus_Out_Display(PL_Code As String) As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display_Ch = False
        Grid_Cus_Out.RowCount = 1
        StrSQL = "Select PL_Sun, PL_CM_Code, PL_CM_Name, PL_Unit, PL_Unit_Gold, PL_Base_Amount, PL_Bigo   FROM SI_PRODUCT_Customer with(NOLOCK)  Where PL_Code = '" & PL_Code & "' And PL_Sel = '매출'  Order By Convert(Decimal,PL_Sun)"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_Cus_Out.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                Grid_Cus_Out.Item(0, i).Value = DBT.Rows(i).Item(0).ToString
                Grid_Cus_Out.Item(1, i).Value = DBT.Rows(i).Item(1).ToString
                Grid_Cus_Out.Item(2, i).Value = DBT.Rows(i).Item(2).ToString
                Grid_Cus_Out.Item(3, i).Value = DBT.Rows(i).Item(3).ToString
                Grid_Cus_Out.Item(4, i).Value = DBT.Rows(i).Item(4).ToString
                Grid_Cus_Out.Item(5, i).Value = DBT.Rows(i).Item(5).ToString
                Grid_Cus_Out.Item(6, i).Value = DBT.Rows(i).Item(6).ToString
            Next i
        Else
            Grid_Cus_Out.RowCount = 1
            For j = 0 To Grid_Cus_Out.ColumnCount - 1
                Grid_Cus_Out.Item(j, 0).Value = ""
            Next j
        End If

        Grid_Display_Ch = True
        Grid_Cus_Out_Display = True
    End Function

    Private Sub Delete_Com_Click(sender As Object, e As EventArgs)

    End Sub


    Public Function Grid_LG_Display(PL_Code As String) As Boolean
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        Grid_Display_Ch = False
        Grid_LG.RowCount = 1
        StrSQL = "Select PL_Sun, PL_LG_Code FROM SI_PRODUCT_LG with(NOLOCK)  Where PL_Code = '" & PL_Code & "' Order By Convert(Decimal,PL_Sun)"
        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            Grid_LG.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                Grid_LG.Item(0, i).Value = DBT.Rows(i).Item(0).ToString
                Grid_LG.Item(1, i).Value = DBT.Rows(i).Item(1).ToString
            Next i
        Else
            Grid_LG.RowCount = 1
            For j = 0 To Grid_LG.ColumnCount - 1
                Grid_LG.Item(j, 0).Value = ""
            Next j
        End If

        Grid_Display_Ch = True
        Grid_LG_Display = True
    End Function



    Private Sub Com_Customer_Click(sender As Object, e As EventArgs) Handles Com_Customer.Click
        Customer_Panel.Visible = True
    End Sub

    Private Sub Save_LG_Click(sender As Object, e As EventArgs) Handles Save_LG.Click
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        Grid_Display_Ch = False
        Dim i As Integer

        For i = 0 To Grid_LG.Rows.Count - 1
            If Grid_LG.Rows(i).HeaderCell.Value = "*" Then
                'Update
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SI_PRODUCT_LG Set PL_LG_Code = '" & Grid_LG.Item(1, i).Value & "' where PL_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And PL_Sun = '" & Grid_LG.Item(0, i).Value & "'"
                Re_Count = DB_Execute()
                Grid_LG.Rows(i).HeaderCell.Value = ""
            End If
        Next i
        MessageBox.Show("저장되었습니다")

        Grid_Display_Ch = True
    End Sub

    Private Sub Insert__LG_Click(sender As Object, e As EventArgs) Handles Insert__LG.Click
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        Grid_Display_Ch = False
        Dim DBT As New DataTable
        Dim Db_Sun As Long
        StrSQL = "Select PL_Sun FROM SI_PRODUCT_LG with(NOLOCK) Where PL_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' Order By Convert(Decimal,PL_Sun) Desc "
        Re_Count = DB_Select(DBT)
        If Re_Count = 0 Then
            Db_Sun = 1
        Else
            Db_Sun = DBT.Rows(0).Item(0) + 1
        End If

        '추가한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Insert INTO SI_PRODUCT_LG (PL_Code, PL_Sun)  Values ('" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "', '" & Db_Sun & "')"
        Re_Count = DB_Execute()
        Grid_LG_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True
    End Sub

    Private Sub Delete__LG_Click(sender As Object, e As EventArgs) Handles Delete__LG.Click
        If Len(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value) > 0 Then
        Else
            Exit Sub
        End If

        If Len(Grid_LG.Item(0, Grid_LG.CurrentCell.RowIndex).Value) < 1 Then
            Exit Sub
        End If
        Grid_Display_Ch = False
        Dim i As Integer
        '삭제한다
        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_PRODUCT_LG Where PL_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And   PL_Sun = '" & Grid_LG.Item(0, Grid_LG.CurrentCell.RowIndex).Value & "'"
        Re_Count = DB_Execute()
        For i = Grid_LG.CurrentCell.RowIndex + 1 To Grid_LG.RowCount - 1
            '수정한다
            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "UPDATE SI_PRODUCT_LG Set PL_Sun = '" & i & "' Where PL_Code = '" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' And  PL_Sun = '" & i + 1 & "'"
            Re_Count = DB_Execute()
        Next i
        Grid_LG_Display(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
        Grid_Display_Ch = True
    End Sub

    Private Sub Grid_LG_MouseClick(sender As Object, e As MouseEventArgs) Handles Grid_LG.MouseClick



    End Sub

    Private Sub Grid_LG_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_LG.CellValueChanged
        If Grid_Display_Ch = True Then
            Grid_LG.CurrentRow.HeaderCell.Value = "*"
        End If
    End Sub

    Private Sub Save_Com_Click_1(sender As Object, e As EventArgs) Handles Save_Com.Click
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
        StrSQL = StrSQL & " WHERE " & Table_Name & "." & Field_Name(0) & " = '" & Grid_Info.Item(1, 0).Value & "'"
        Re_Count = DB_Execute()


        Search_Com.PerformClick()
        Grid_Code.ClearSelection()
    End Sub

    Private Sub Insert_Com_Click(sender As Object, e As EventArgs) Handles Insert_Com.Click
        'DB를 추가 한다.
        '코드 번호 입력
        '새로운 푬에서 입력 받는다.

        Dim RR As String
        Dim DBT As New DataTable


        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "Delete Temp "
        Re_Count = DB_Execute()



        Frm_Insert.Text = "가공품/제품 코드 추가"
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
    End Sub

    Private Sub Delete_Com_Click_1(sender As Object, e As EventArgs) Handles Delete_Com.Click
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
        StrSQL = StrSQL & "DELETE SI_PRODUCT_Customer Where PL_Code = '" & Grid_Info.Item(1, 0).Value & "' "
        Re_Count = DB_Execute()

        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_PRODUCT_QC Where PQ_Code = '" & Grid_Info.Item(1, 0).Value & "' "
        Re_Count = DB_Execute()

        StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        StrSQL = StrSQL & "DELETE SI_PROCESS_LIST Where PP_Code = '" & Grid_Info.Item(1, 0).Value & "' "

        'StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
        'StrSQL = StrSQL & "DELETE SI_PRODUCT_LG Where PL_Code = '" & Grid_Info.Item(1, 0).Value & "' "

        Re_Count = DB_Execute()
        Grid_Code_Display(True)
        Grid_Display_Ch = True
    End Sub

    Private Sub Panel_Menu_Paint(sender As Object, e As PaintEventArgs) Handles Panel_Menu.Paint

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim RR As String
        Dim DBT As New DataTable
        Dim Db_Sun As String

        If Grid_Info.Item(1, 0).Value = "" Then
            MsgBox("대상이 될 제품을 선택해주세요")
            Exit Sub
        End If

        Panel1.Visible = True
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        '기존에 있는 데이터인지 확인
        Dim DBT As New DataTable
        Dim Db_Sun As String

        StrSQL = "Select PL_CODE From SI_PRODUCT with(NOLOCK) WHERE PL_CODE ='" & TextBox1.Text & "' Order By PL_CODE Desc "
        Re_Count = DB_Select(DBT)

        If Re_Count = 0 Then
            MsgBox("입력하신 코드가 유효하지 않습니다")
            Exit Sub
        Else
        End If


        '매출처가 존재하는 지 확인
        StrSQL = "Select PL_SUN, PL_CM_CODE, PL_UNIT, PL_UNIT_GOLD, PL_BASE_AMOUNT, PL_BIGO
                    From SI_PRODUCT_CUSTOMER with(NOLOCK) 
                    WHERE PL_CODE ='" & Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value & "' and PL_SEL ='매출' Order By PL_CODE Desc "
        Re_Count = DB_Select(DBT)

        If Re_Count = 0 Then
            MsgBox("매출처를 먼저 입력해주세요")
            Exit Sub
        Else
            ' MsgBox(Grid_Cus_Out.Item(1, 3).Value)
            ' MsgBox(Grid_Cus_Out.Item(3, 1).Value)
            Dim i As Integer = 0
            For i = 0 To Re_Count - 1
                '매출처 복사
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert into SI_PRODUCT_CUSTOMER (PL_SEL, PL_CODE, PL_SUN, PL_CM_CODE, PL_CM_NAME, PL_UNIT, PL_UNIT_GOLD, PL_BASE_AMOUNT, PL_BIGO)  Values 
                                ('매출', '" & TextBox1.Text & "', '" & Grid_Cus_Out.Item(0, i).Value & "',
                 '" & Grid_Cus_Out.Item(1, i).Value & "', '" & Grid_Cus_Out.Item(2, i).Value & "', '" & Grid_Cus_Out.Item(3, i).Value & "', 
                 '" & Grid_Cus_Out.Item(4, i).Value & "', '" & Grid_Cus_Out.Item(5, i).Value & "', '" & Grid_Cus_Out.Item(6, i).Value & "')"

                Re_Count = DB_Execute()
            Next


            Grid_Info_Display_Pro(Grid_Code.Item(0, Grid_Code.CurrentCell.RowIndex).Value)
            'BOM 입력 페널
            Grid_Code_Display(True)
            Customer_Panel.Visible = True
            Panel1.Visible = False
            MsgBox("완료되었습니다")
        End If
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        Panel1.Visible = False
    End Sub
End Class