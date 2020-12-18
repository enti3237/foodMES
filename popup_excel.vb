Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel


Public Class popup_excel

    Private Sub popup_excel_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Pro_Info_Setup()

    End Sub

    Public Function Pro_Info_Setup() As Boolean
        Grid_Color(Pro_Info)
        Pro_Info.EnableHeadersVisualStyles = False
        Pro_Info.AllowUserToAddRows = False
        Pro_Info.ColumnHeadersHeight = 24

        Pro_Info.RowTemplate.Height = 24
        Pro_Info.ColumnCount = 6
        'Pro_Info.RowCount = 1
        '만약 헤더의 스타일을 조정하려면 아래 코드에 헤더를 추가한다.
        Pro_Info.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Pro_Info.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Pro_Info.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft



        Pro_Info.RowHeadersWidth = 10
        Pro_Info.Columns(0).Width = 100
        Pro_Info.Columns(1).Width = 150
        Pro_Info.Columns(2).Width = 70
        Pro_Info.Columns(3).Width = 70
        Pro_Info.Columns(4).Width = 70
        Pro_Info.Columns(5).Width = 70

        '   Pro_Info.Columns(13).Width = 70

        Pro_Info.Columns(0).HeaderText = "코드"
        Pro_Info.Columns(1).HeaderText = "품명"
        Pro_Info.Columns(2).HeaderText = "종류"
        Pro_Info.Columns(3).HeaderText = "규격"
        Pro_Info.Columns(4).HeaderText = "유통기한"
        Pro_Info.Columns(5).HeaderText = "비고"


    End Function

    Private Sub Btn_exSave_Click(sender As Object, e As EventArgs) Handles Btn_exSave.Click
        Dim DBT As DataTable
        StrSQL = "SELECT PL_Code
                    FROM SI_PRODUCT
                   WHERE PL_Code ='" & Pro_Info.Item(0, 0).Value & "'"
        Re_Count = DB_Select(DBT)

        If Re_Count <> 0 Then
            MsgBox("기존 코드가 존재합니다!")
            Exit Sub
        End If

        For i = 0 To Pro_Info.Rows.Count - 1

            If Pro_Info.Item(0, i).Value = "" Then
                Exit For
            End If

            StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
            StrSQL = StrSQL & "Insert INTO SI_PRODUCT (PL_Code, PL_Name, PL_Sel, PL_Standard, PL_DATE, PL_BIGO) 
          Values ('" & Pro_Info.Item(0, i).Value & "', '" & Pro_Info.Item(1, i).Value & "', '" & Pro_Info.Item(2, i).Value & "', '" & Pro_Info.Item(3, i).Value & "',
                '" & Pro_Info.Item(4, i).Value & "', '" & Pro_Info.Item(5, i).Value & "')"


            Re_Count = DB_Execute()
        Next i

        MsgBox("저장되었습니다!")
    End Sub
End Class