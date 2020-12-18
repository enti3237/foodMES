Public Class Chul
    Private Sub D_DoubleClick(sender As Object, e As EventArgs) Handles D.DoubleClick

    End Sub

    Private Sub Chul_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TopMost = True
        setup()
        display()
        '  setup2()
        '  Panel1.Visible = False
        D.ReadOnly = True
    End Sub

    Public Sub display()
        Dim DBT As New DataTable
        Dim i As Integer
        Dim j As Integer
        D.RowCount = 1
        '  Select Case Search_Combo.Text
        ' Case "코드"
        StrSQL = "Select CODE, PL_NAME, QTY, OUT_DATE, OUT_TIME, CM_NAME
                    FROM LABEL with(NOLOCK) 
                   Where C_CHECK='출고' ORDER BY OUT_DATE DESC, OUT_TIME DESC"
        '    Case "날짜"
        '  StrSQL = "Select CR_Code, CR_Date, CR_Customer_Name FROM SC_SALES with(NOLOCK) Where CR_Date Like '%" & Search_Text.Text & "%' Or CR_Date Is Null  Order By CR_Date"
        '     Case "거래처"
        '  StrSQL = "Select CR_Code, CR_Date, CR_Customer_Name FROM SC_SALES with(NOLOCK) Where CR_Customer_Name Like '%" & Search_Text.Text & "%' Or CR_Customer_Name Is Null  Order By CR_Customer_Name"
        '  End Select

        Re_Count = DB_Select(DBT)
        If Re_Count <> 0 Then
            D.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                For j = 0 To D.ColumnCount - 1
                    D.Item(j, i).Value = DBT.Rows(i).Item(j)
                Next j
            Next i
        Else
            D.RowCount = 1
            For j = 0 To D.ColumnCount - 1
                D.Item(j, 0).Value = ""
            Next j
        End If
        '   D_Display = True

        D.MultiSelect = False
        D.ClearSelection()

        'StrSQL = "Select GETDATE()"
        'Re_Count = DB_Select(DBT)


        'StrSQL = "Select GETDATE()"
        'Re_Count = DB_Select(DBT)
        '    If Re_Count = 0 Then
        '        Exit Sub
        '    Else
        '        My_Date = DBT.Rows(0).Item(0)
        '        My_Time = DateTime.Now.ToString("HH:mm")
        '        JP_Code = Mid(My_Date, 1, 4) & Mid(My_Date, 6, 2) & Mid(My_Date, 9, 2)
        '        My_Time = Replace(My_Time, ":", "-")
        '        My_Date = Mid(My_Date, 1, 4) & "-" & Mid(My_Date, 6, 2) & "-" & Mid(My_Date, 9, 2)
        '    End If



        'StrSQL = "UPDATE LABEL SET C_CHECK = '출고', OUT_DATE = '" & My_Date & "', OUT_TIME = '" & My_Time & "', 
        '                CM_CODE = '" & D.Item(5, i).Value & "', CM_NAME = '" & D.Item(6, i).Value & "'
        '       WHERE CODE ='" & D.Item(0, i).Value & "'"

        'Re_Count = DB_Execute()

    End Sub

    Public Sub setup()
        Grid_Color(D)

        D.AllowUserToAddRows = False
        D.RowTemplate.Height = 70
        D.ColumnCount = 6
        D.RowCount = 0
        ' D.RowCount = 7

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        D.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        D.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        D.Columns(0).Width = 240
        D.Columns(1).Width = 300
        D.RowHeadersWidth = 10
        D.Columns(0).HeaderText = "라벨코드"
        D.Columns(1).HeaderText = "제품명"
        D.Columns(2).HeaderText = "수량"
        D.Columns(3).HeaderText = "출고일자"
        D.Columns(4).HeaderText = "출고시간"
        D.Columns(5).HeaderText = "거래처명"
        ' D.Columns(6).HeaderText = "거래처명"

        D.ColumnHeadersHeight = 40


        D.ColumnHeadersHeight = 40
    End Sub
End Class