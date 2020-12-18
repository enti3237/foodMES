Imports Excel = Microsoft.Office.Interop.Excel

Public Class Frm_HpSchedule
    Dim xlapp As Excel.Application
    Dim xlbook As Excel.Workbook
    Dim xlsheet As Excel.Worksheet
    Dim Excel_Check As Boolean

    Private Sub Frm_CC_Report1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Dock = DockStyle.Fill
        Me.BringToFront()
        'Search_Combo.Items.Add("코드")
        'Search_Combo.Items.Add("날짜")
        'Search_Combo.Items.Add("거래처")
        'Search_Combo.Text = "코드"
        '12, 135

        '  ListBox1.Items.Add(foundFile)
        ' ListBox1.Items.Add(Application.StartupPath & "/Excel/HACCP")


        ListBox1.Items.Clear()
        Dim dir As New IO.DirectoryInfo(Application.StartupPath & "/Excel/HACCP")
        Dim fname As IO.FileInfo
        For Each fname In dir.GetFiles()
            If fname.Extension.Equals(".xlsx") Then 'Or fname.Extension.Equals(".png") Or fname.Extension.Equals(".gif") Or fname.Extension.Equals(".txt") Then
                ListBox1.Items.Add(fname)
                '  ListView1.Items.Add(fname)
            End If
        Next


    End Sub


    Private Sub Frm_CC_Report1_VisibleChanged(sender As Object, e As EventArgs) Handles Me.VisibleChanged
        If Excel_Check = True Then
            xlbook.Close()
            xlapp.Quit()
            xlsheet = Nothing
            xlbook = Nothing
            xlapp = Nothing
            releaseObject(xlsheet)
            releaseObject(xlbook)
            releaseObject(xlapp)
            Excel_Check = False
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'ccp 모니터링 점검표
        'Dim Frm_Excel1 As New Excel_Panel(Panel1, "개선조치요구서_작성본 2019.xlsx")
        Dim Frm_Excel1 As New Excel_Panel(Panel1, "CCP-1P모니터링_윤푸드.xlsx")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '중요관리점 
        Dim Frm_Excel2 As New Excel_Panel(Panel1, "중요관리점(CCP)_윤푸드.xlsx")

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        '냉장 냉동고
        Dim Frm_Excel3 As New Excel_Panel(Panel1, "냉장냉동고온도_윤푸드.xlsx")
    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        '교육훈련 
        Dim Frm_Excel4 As New Excel_Panel(Panel1, "교육훈련일지_윤푸드.xlsx")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        '위생점검 일
        Dim Frm_Excel5 As New Excel_Panel(Panel1, "위생일지_윤푸드.xlsx")
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        '위생점검 정기
        Dim Frm_Excel6 As New Excel_Panel(Panel1, "정기위생일지_윤푸드.xlsx")
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        '이물혼입
        Dim Frm_Excel7 As New Excel_Panel(Panel1, "이물발생관리대장_윤푸드.xlsx")
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        '부적합품
        Dim Frm_Excel8 As New Excel_Panel(Panel1, "부적합품발생및처리_윤푸드.xlsx")
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        '출하검사서
        Dim Frm_Excel9 As New Excel_Panel(Panel1, "출하검사서_윤푸드.xlsx")
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        '자체검교정
        Dim Frm_Excel10 As New Excel_Panel(Panel1, "자체검교정일지_윤푸드.xlsx")
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        '검증계획
        Dim Frm_Excel11 As New Excel_Panel(Panel1, "검증결과보고및개선조치_윤푸드.xlsx")
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        '협력업체
        Dim Frm_Excel12 As New Excel_Panel(Panel1, "협력업체평가표_윤푸드.xlsx")
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        '탈의실
        Dim Frm_Excel13 As New Excel_Panel(Panel1, "탈의실 관리_윤푸드.xlsx")
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        '화장실
        Dim Frm_Excel14 As New Excel_Panel(Panel1, "화장실관리점검표_윤푸드.xlsx")
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        '위생전실
        Dim Frm_Excel15 As New Excel_Panel(Panel1, "위생전실관리_윤푸드.xlsx")
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        'Dim index As Integer
        'Dim search, text As String
        'Dim i As Integer

        'search = TextBox1.Text.Trim

        'If search = "" Then
        '    Exit Sub
        'End If

        '' ListBox1.Items.Clear()

        'For i = 0 To ListBox1.Items.Count - 1
        '    If InStr(ListBox1.Items(i).ToString, search) Then '리스트박스 행에 검색하려는 글자가 있을때
        '        '  ListBox1.Items.Add(ListBox1.Items(i).ToString)

        '    Else
        '        ListBox1.Items.Remove(ListBox1.Items(i).ToString)
        '        Continue For
        '    End If
        'Next

        ' ListBox1.Items.Clear()

        'MsgBox(ListBox1.Items(0).ToString)
        'MsgBox(ListBox1.Items(1).ToString)
        'If InStr(ListBox1.Items()) Then

        'End If

        'index = ListBox1.FindString(search)
        'text = ListBox1.Items(0).ToString


        'ListBox1.TopIndex = index
        'ListBox1.SelectedIndex = index

        'MsgBox(text)

        'For i = 0 To ListBox1.Items.Count - 1
        '    ListBox1.Items.
        'Next

    End Sub

    Private Sub Search_Com_Click(sender As Object, e As EventArgs) Handles Search_Com.Click
        '리스트 박스에 있는 데이터 검색하기

        Dim index As Integer
        Dim search, text As String
        Dim i As Integer

        search = TextBox1.Text.Trim

        If search = "" Then
            Exit Sub
        End If

        ListBox2.Items.Clear()

        For i = 0 To ListBox1.Items.Count - 1
            If InStr(ListBox1.Items(i).ToString, search) Then '리스트박스 행에 검색하려는 글자가 있을때
                ListBox2.Items.Add(ListBox1.Items(i).ToString)
                ' MsgBox(True)
            Else
                ' ListBox1.Items.Remove(i)
                ' MsgBox(False)
                '  Continue For
            End If
        Next
    End Sub
End Class
