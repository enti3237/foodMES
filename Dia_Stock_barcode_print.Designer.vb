<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Dia_Stock_barcode_print
    Inherits System.Windows.Forms.Form

    'Form은 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows Form 디자이너에 필요합니다.
    Private components As System.ComponentModel.IContainer

    '참고: 다음 프로시저는 Windows Form 디자이너에 필요합니다.
    '수정하려면 Windows Form 디자이너를 사용하십시오.  
    '코드 편집기에서는 수정하지 마세요.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.PrintDocument1 = New System.Drawing.Printing.PrintDocument()
        Me.Print_bnt = New System.Windows.Forms.Button()
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog()
        Me.next_page = New System.Windows.Forms.Button()
        Me.back_page = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Location = New System.Drawing.Point(12, 44)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(446, 337)
        Me.PictureBox1.TabIndex = 7
        Me.PictureBox1.TabStop = False
        '
        'Print_bnt
        '
        Me.Print_bnt.Location = New System.Drawing.Point(10, 8)
        Me.Print_bnt.Name = "Print_bnt"
        Me.Print_bnt.Size = New System.Drawing.Size(93, 30)
        Me.Print_bnt.TabIndex = 6
        Me.Print_bnt.Text = "인쇄하기"
        Me.Print_bnt.UseVisualStyleBackColor = True
        '
        'PrintDialog1
        '
        Me.PrintDialog1.UseEXDialog = True
        '
        'next_page
        '
        Me.next_page.Location = New System.Drawing.Point(381, 12)
        Me.next_page.Name = "next_page"
        Me.next_page.Size = New System.Drawing.Size(79, 23)
        Me.next_page.TabIndex = 10
        Me.next_page.Text = "다음 페이지"
        Me.next_page.UseVisualStyleBackColor = True
        '
        'back_page
        '
        Me.back_page.Location = New System.Drawing.Point(289, 12)
        Me.back_page.Name = "back_page"
        Me.back_page.Size = New System.Drawing.Size(86, 23)
        Me.back_page.TabIndex = 9
        Me.back_page.Text = "이전 페이지"
        Me.back_page.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(359, 395)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(42, 12)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Label1"
        '
        'Dia_Stock_barcode_print
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(470, 414)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Print_bnt)
        Me.Controls.Add(Me.next_page)
        Me.Controls.Add(Me.back_page)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Dia_Stock_barcode_print"
        Me.Text = "라벨 출력"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents PrintDocument1 As Printing.PrintDocument
    Friend WithEvents Print_bnt As Button
    Friend WithEvents PrintDialog1 As PrintDialog
    Friend WithEvents next_page As Button
    Friend WithEvents back_page As Button
    Friend WithEvents Label1 As Label
End Class
