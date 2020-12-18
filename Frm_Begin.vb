Public Class Frm_Begin
    Private Sub Frm_Begin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Haccp_Display()
        Main_Display()
        Man_Display()
        SALES_Display()
    End Sub

    Private Sub Haccp_Display()
        Dim i As Integer

        Dim DBT As DataTable
        DBT = Nothing
        '문서명 작성여부 일자
        StrSQL = "SELECT ST_DOC
                    FROM HP_STANDARD_LIST AS A"

        Re_Count = DB_Select(DBT)
        If Re_Count > 0 Then
            Main.DataSource = DBT
            Main.ClearSelection()
        End If
    End Sub

    Private Sub Main_Display()
        Dim i As Integer

        Dim DBT As DataTable
        DBT = Nothing

        StrSQL = "SELECT CR_CODE, CR_DATE, CR_CM_NAME, CL_PL_NAME, CL_QTY 
                    FROM SC_SALES AS A
                    JOIN SC_SALES_LIST AS B ON A.CR_CODE = B.CL_CR_CODE 
                    ORDER BY CR_CODE DESC"

        Re_Count = DB_Select(DBT)
        If Re_Count > 0 Then
            Main.DataSource = DBT
            Main.ClearSelection()
        End If
    End Sub

    Private Sub Man_Display()
        Dim i As Integer

        Dim DBT As DataTable
        DBT = Nothing

        StrSQL = "SELECT Man_Name, Man_Update 
                    FROM SI_MAN 
                    ORDER BY MAN_UPDATE"

        Re_Count = DB_Select(DBT)
        If Re_Count > 0 Then
            MAN.DataSource = DBT
            MAN.ClearSelection()
        End If

        For i = 0 To MAN.RowCount - 1
            If MAN.Item(1, i).Value.ToString < DateTime.Today.ToString("yyyy-MM-dd") Then
                MAN.Rows.Item(i).DefaultCellStyle.BackColor = Color.Yellow
            Else
                MAN.Rows.Item(i).DefaultCellStyle.BackColor = Color.Green
            End If
        Next

        MAN.ClearSelection()

    End Sub

    Private Sub SALES_Display()
        Dim i As Integer

        Dim DBT As DataTable
        DBT = Nothing

        StrSQL = "SELECT CR_DAY, CR_CM_NAME, CL_PL_NAME, SUM(CONVERT(DECIMAL,CL_QTY)) AS CL_QTY
                    FROM SC_SALES AS A
                    JOIN SC_SALES_LIST AS B ON A.CR_CODE = B.CL_CR_CODE 
                    GROUP BY CR_DAY, CR_CM_NAME, CL_PL_NAME 
                    ORDER BY CR_DAY"

        Re_Count = DB_Select(DBT)

        If Re_Count <> 0 Then
            SALES.RowCount = Re_Count
            For i = 0 To Re_Count - 1
                '   SALES.Item(0, i).Value = i + 1
                For j = 0 To SALES.ColumnCount - 1
                    SALES.Item(j, i).Value = DBT.Rows(i).Item(j).ToString
                    If SALES.Item(0, i).Value.ToString < DateTime.Today.ToString("yyyy-MM-dd") Then
                        SALES.Rows(i).DefaultCellStyle.BackColor = Color.OrangeRed
                    Else
                        SALES.Rows(i).DefaultCellStyle.BackColor = Color.White
                    End If
                Next j
            Next i
        Else
            SALES.RowCount = 1
            For j = 0 To SALES.ColumnCount - 1
                SALES.Item(j, 0).Value = ""
            Next j
        End If
        SALES.ClearSelection()

    End Sub


End Class
