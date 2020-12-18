Public Class Frm_Company
    Dim Grid_Display_Ch As Boolean

    Private Sub Frm_Company_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        Me.BringToFront()
        Grid_Info_Setup()
        Grid_Info_Display(Grid_Info, "Select * From Company With(NOLOCK) Where CM_Code = '10000'")
        '10000



        'Search_Combo.Items.Add("코드")
        'Search_Combo.Items.Add("날짜")
        'Search_Combo.Items.Add("거래처")
        'Search_Combo.Text = "코드"


        'Panel_Main.Top = 136
        'Panel_Main.Left = 389
        'Panel_Excel.Top = 136
        'Panel_Excel.Left = 389

        'Panel_Main.Visible = True
        'Panel_Excel.Visible = False

        'Com_Excel_Print.Visible = False
    End Sub
    Public Function Grid_Info_Setup() As Boolean

        Grid_Color(Grid_Info)


        Grid_Info.AllowUserToAddRows = False
        Grid_Info.RowTemplate.Height = 24
        Grid_Info.ColumnCount = 2
        Grid_Info.RowCount = 24

        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Grid_Info.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid_Info.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid_Info.RowHeadersWidth = 10
        Grid_Info.Columns(0).Width = 100
        Grid_Info.Columns(1).Width = 350
        Grid_Info.Columns(0).HeaderText = "구분"
        Grid_Info.Columns(1).HeaderText = "내용"


        Grid_Info.Item(0, 0).Value = "코드"
        Grid_Info.Item(0, 1).Value = "업체명"
        Grid_Info.Item(0, 2).Value = "사업자번호"
        Grid_Info.Item(0, 3).Value = "법인번호"
        Grid_Info.Item(0, 4).Value = "대표자"
        Grid_Info.Item(0, 5).Value = "업태"
        Grid_Info.Item(0, 6).Value = "종목"
        Grid_Info.Item(0, 7).Value = "우편번호"
        Grid_Info.Item(0, 8).Value = "주소1"
        Grid_Info.Item(0, 9).Value = "주소2"
        Grid_Info.Item(0, 10).Value = "전화번호"
        Grid_Info.Item(0, 11).Value = "팩스번호"
        Grid_Info.Item(0, 12).Value = "거래시작일"
        Grid_Info.Item(0, 13).Value = "거래종료일"
        Grid_Info.Item(0, 14).Value = "담당자"
        Grid_Info.Item(0, 15).Value = "담당자전화"
        Grid_Info.Item(0, 16).Value = "담당자메일"
        Grid_Info.Item(0, 17).Value = "은행명"
        Grid_Info.Item(0, 18).Value = "계좌번호"
        Grid_Info.Item(0, 19).Value = "생성담당"
        Grid_Info.Item(0, 20).Value = "생성일자"
        Grid_Info.Item(0, 21).Value = "수정담당"
        Grid_Info.Item(0, 22).Value = "수정일"
        Grid_Info.Item(0, 23).Value = "비고"

        Grid_Info.Rows(0).ReadOnly = True
        Grid_Info_Setup = True
    End Function

    Private Sub Grid_Info_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid_Info.CellValueChanged
        If Grid_Display_Ch = True Then
            Grid_Info.CurrentRow.HeaderCell.Value = "*"
        End If
    End Sub

    Private Sub Save_Main_Click(sender As Object, e As EventArgs) Handles Save_Main.Click
        Dim i As Integer


        '  For i = 0 To Grid_Info.Rows.Count - 1
        'CM_Code = '" & Grid_Info.Item(1, 0).Value & "'
        StrSQL = "update Company Set 
                        CM_Name = '" & Grid_Info.Item(1, 1).Value & "'
                        , CM_Op_Number = '" & Grid_Info.Item(1, 2).Value & "' 
                        , CM_Co_Number = '" & Grid_Info.Item(1, 3).Value & "'
                        , CM_Leader = '" & Grid_Info.Item(1, 4).Value & "'
                        , CM_Conditions = '" & Grid_Info.Item(1, 5).Value & "'
                        , CM_Item = '" & Grid_Info.Item(1, 6).Value & "'
                        , CM_Add_Number = '" & Grid_Info.Item(1, 7).Value & "'
                        , CM_Add1 = '" & Grid_Info.Item(1, 8).Value & "'
                        , CM_Add2 = '" & Grid_Info.Item(1, 9).Value & "'
                        , CM_Tel = '" & Grid_Info.Item(1, 10).Value & "'
                        , CM_Fax = '" & Grid_Info.Item(1, 11).Value & "'
                        , CM_ST_Date = '" & Grid_Info.Item(1, 12).Value & "'
                        , CM_EN_Date = '" & Grid_Info.Item(1, 13).Value & "'
                        , CM_Man_Name = '" & Grid_Info.Item(1, 14).Value & "'
                        , CM_Man_Tel = '" & Grid_Info.Item(1, 15).Value & "'
                        , CM_Man_Email = '" & Grid_Info.Item(1, 16).Value & "'
                        , CM_Bank = '" & Grid_Info.Item(1, 17).Value & "'
                        , CM_Bank_Number = '" & Grid_Info.Item(1, 18).Value & "'
                        , CM_Create = '" & Grid_Info.Item(1, 19).Value & "'
                        , CM_Create_Date = '" & Grid_Info.Item(1, 20).Value & "'
                        , CM_Update = '" & Grid_Info.Item(1, 21).Value & "'
                        , CM_Update_Date = '" & Grid_Info.Item(1, 22).Value & "'
                        , CM_Bigo = '" & Grid_Info.Item(1, 23).Value & "'
                         where CM_Code = '" & Grid_Info.Item(1, 0).Value & "'"

        Re_Count = DB_Execute()

        MessageBox.Show("저장했습니다")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        If Grid_Info.Item(1, 0).Value = "" Then
            Exit Sub
        End If

        '삭제

        StrSQL = "DELETE COMPNAY where CM_Code = '" & Grid_Info.Item(1, 0).Value & "'"
        Re_Count = DB_Execute()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        '추가

        If Grid_Info.Item(1, 0).Value = "10000" Then
            MsgBox("기존 코드가 있어 추가할 수 없습니다")
            Exit Sub
        Else
            StrSQL = "INSERT INTO COMPNAY (CM_CODE) VALUES('10000')"
            Re_Count = DB_Execute()
        End If



    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '검색
        StrSQL = "SELECT * FROM COMPANY"
        Re_Count = DB_Execute()

    End Sub
End Class
