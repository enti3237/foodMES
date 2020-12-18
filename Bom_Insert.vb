Public Class Bom_Insert
    Dim JP_Code As String
    Private Sub Process_Insert_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim DBT As New DataTable
        Label1.Visible = False
        Sub_Code.Visible = False
        If Label1.Text = "insert2" Then
            'BOM추가
            process_code.Enabled = False
            process_name.Enabled = False
            Process_bigo.Enabled = False
            Insert_Com.Enabled = False

            Dim Db_Sun As Long
            StrSQL = "Select PL_Sun FROM SI_PRODUCT_RECIPE with(NOLOCK) Where PL_Code = '" & process_code.Text & "' Order By Convert(Decimal,PL_Sun) Desc "
            Re_Count = DB_Select(DBT)
            If Re_Count = 0 Then
                Db_Sun = 1
            Else
                Db_Sun = DBT.Rows(0).Item(0) + 1

            End If

            sun.Text = Db_Sun
            Unit.Text = "KG"
        ElseIf Label1.Text = "update2" Then
            process_code.Enabled = False
            process_name.Enabled = False
            Process_bigo.Enabled = False
            Insert_Com.Enabled = False
            Unit.Text = "KG"
        End If

    End Sub

    Private Sub Btn_Save_Click(sender As Object, e As EventArgs) Handles Btn_Save.Click
        If Label1.Text = "insert2" Then
            Dim DBT As New DataTable
            Try
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "Insert INTO SI_PRODUCT_RECIPE Values ('" & process_code.Text & "', '" & sun.Text & "', '" & Sub_Code.Text & "', '" & Sub_Name.Text & "', '" & Qty.Text & "', '" & Unit.Text & "')"
                Re_Count = DB_Execute()
                MsgBox("저장되었습니다")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Me.Close()

        ElseIf Label1.Text = "update2" Then

            'UPDATE문 실행
            Try
                StrSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED "
                StrSQL = StrSQL & "UPDATE SI_PRODUCT_RECIPE
                                          SET PL_SUB_CODE = '" & Sub_Code.Text & "',
                                              PL_QTY = '" & Qty.Text & "',
                                              PL_UNIT = '" & Unit.Text & "'
                                        WHERE PL_CODE ='" & process_code.Text & "' AND PL_SUN='" & sun.Text & "'"
                Re_Count = DB_Execute()
                MsgBox("수정되었습니다")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Me.Close()

        End If
    End Sub

    Private Sub btnJobSearch_Click(sender As Object, e As EventArgs) Handles btnJobSearch.Click
        Dim popup_product As New popup_product
        popup_product.ShowDialog()

        If popup_product.DialogResult = DialogResult.OK Then
            Sub_Name.Text = popup_product.Product_Name
            Sub_Code.Text = popup_product.Product_Code
            Qty.Text = popup_product.Product_Standard
        End If
    End Sub
End Class