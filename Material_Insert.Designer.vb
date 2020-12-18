<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Material_Insert
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
        Me.Btn_Save = New System.Windows.Forms.Button()
        Me.txt_Matebigo = New System.Windows.Forms.TextBox()
        Me.Info_La9 = New System.Windows.Forms.Button()
        Me.Info_La4 = New System.Windows.Forms.Button()
        Me.Info_La3 = New System.Windows.Forms.Button()
        Me.txt_Matename = New System.Windows.Forms.TextBox()
        Me.Info_La1 = New System.Windows.Forms.Button()
        Me.txt_Matecode = New System.Windows.Forms.TextBox()
        Me.Info_La0 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txt_Standard = New System.Windows.Forms.TextBox()
        Me.txt_Sel = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'Btn_Save
        '
        Me.Btn_Save.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Btn_Save.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Btn_Save.ForeColor = System.Drawing.Color.Black
        Me.Btn_Save.Location = New System.Drawing.Point(14, 195)
        Me.Btn_Save.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Btn_Save.Name = "Btn_Save"
        Me.Btn_Save.Size = New System.Drawing.Size(546, 34)
        Me.Btn_Save.TabIndex = 600
        Me.Btn_Save.Text = "저장"
        Me.Btn_Save.UseVisualStyleBackColor = False
        '
        'txt_Matebigo
        '
        Me.txt_Matebigo.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.txt_Matebigo.Location = New System.Drawing.Point(135, 142)
        Me.txt_Matebigo.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_Matebigo.Name = "txt_Matebigo"
        Me.txt_Matebigo.Size = New System.Drawing.Size(425, 29)
        Me.txt_Matebigo.TabIndex = 500
        '
        'Info_La9
        '
        Me.Info_La9.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Info_La9.Enabled = False
        Me.Info_La9.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Info_La9.ForeColor = System.Drawing.Color.Black
        Me.Info_La9.Location = New System.Drawing.Point(14, 141)
        Me.Info_La9.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Info_La9.Name = "Info_La9"
        Me.Info_La9.Size = New System.Drawing.Size(114, 34)
        Me.Info_La9.TabIndex = 999
        Me.Info_La9.Text = "비고"
        Me.Info_La9.UseVisualStyleBackColor = False
        '
        'Info_La4
        '
        Me.Info_La4.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Info_La4.Enabled = False
        Me.Info_La4.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Info_La4.ForeColor = System.Drawing.Color.Black
        Me.Info_La4.Location = New System.Drawing.Point(14, 98)
        Me.Info_La4.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Info_La4.Name = "Info_La4"
        Me.Info_La4.Size = New System.Drawing.Size(114, 34)
        Me.Info_La4.TabIndex = 999
        Me.Info_La4.Text = "규격"
        Me.Info_La4.UseVisualStyleBackColor = False
        '
        'Info_La3
        '
        Me.Info_La3.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Info_La3.Enabled = False
        Me.Info_La3.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Info_La3.ForeColor = System.Drawing.Color.Black
        Me.Info_La3.Location = New System.Drawing.Point(349, 98)
        Me.Info_La3.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Info_La3.Name = "Info_La3"
        Me.Info_La3.Size = New System.Drawing.Size(114, 34)
        Me.Info_La3.TabIndex = 999
        Me.Info_La3.Text = "구분"
        Me.Info_La3.UseVisualStyleBackColor = False
        '
        'txt_Matename
        '
        Me.txt_Matename.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.txt_Matename.Location = New System.Drawing.Point(135, 54)
        Me.txt_Matename.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_Matename.Name = "txt_Matename"
        Me.txt_Matename.Size = New System.Drawing.Size(269, 29)
        Me.txt_Matename.TabIndex = 200
        '
        'Info_La1
        '
        Me.Info_La1.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Info_La1.Enabled = False
        Me.Info_La1.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Info_La1.ForeColor = System.Drawing.Color.Black
        Me.Info_La1.Location = New System.Drawing.Point(14, 56)
        Me.Info_La1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Info_La1.Name = "Info_La1"
        Me.Info_La1.Size = New System.Drawing.Size(114, 34)
        Me.Info_La1.TabIndex = 999
        Me.Info_La1.Text = "원부재료명"
        Me.Info_La1.UseVisualStyleBackColor = False
        '
        'txt_Matecode
        '
        Me.txt_Matecode.Enabled = False
        Me.txt_Matecode.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.txt_Matecode.Location = New System.Drawing.Point(162, 16)
        Me.txt_Matecode.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_Matecode.Name = "txt_Matecode"
        Me.txt_Matecode.Size = New System.Drawing.Size(242, 29)
        Me.txt_Matecode.TabIndex = 999
        '
        'Info_La0
        '
        Me.Info_La0.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Info_La0.Enabled = False
        Me.Info_La0.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Info_La0.ForeColor = System.Drawing.Color.Black
        Me.Info_La0.Location = New System.Drawing.Point(14, 15)
        Me.Info_La0.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Info_La0.Name = "Info_La0"
        Me.Info_La0.Size = New System.Drawing.Size(142, 34)
        Me.Info_La0.TabIndex = 999
        Me.Info_La0.Text = "원부재료코드"
        Me.Info_La0.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(512, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(50, 15)
        Me.Label1.TabIndex = 241
        Me.Label1.Text = "Label1"
        '
        'txt_Standard
        '
        Me.txt_Standard.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.txt_Standard.Location = New System.Drawing.Point(135, 100)
        Me.txt_Standard.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_Standard.Name = "txt_Standard"
        Me.txt_Standard.Size = New System.Drawing.Size(182, 29)
        Me.txt_Standard.TabIndex = 300
        '
        'txt_Sel
        '
        Me.txt_Sel.Enabled = False
        Me.txt_Sel.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.txt_Sel.Location = New System.Drawing.Point(470, 100)
        Me.txt_Sel.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_Sel.Name = "txt_Sel"
        Me.txt_Sel.Size = New System.Drawing.Size(90, 29)
        Me.txt_Sel.TabIndex = 400
        '
        'Material_Insert
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(579, 249)
        Me.Controls.Add(Me.txt_Sel)
        Me.Controls.Add(Me.txt_Standard)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Btn_Save)
        Me.Controls.Add(Me.txt_Matebigo)
        Me.Controls.Add(Me.Info_La9)
        Me.Controls.Add(Me.Info_La4)
        Me.Controls.Add(Me.Info_La3)
        Me.Controls.Add(Me.txt_Matename)
        Me.Controls.Add(Me.Info_La1)
        Me.Controls.Add(Me.txt_Matecode)
        Me.Controls.Add(Me.Info_La0)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "Material_Insert"
        Me.Text = "원부재료 추가"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Btn_Save As Button
    Friend WithEvents txt_Matebigo As TextBox
    Friend WithEvents Info_La9 As Button
    Friend WithEvents Info_La4 As Button
    Friend WithEvents Info_La3 As Button
    Friend WithEvents txt_Matename As TextBox
    Friend WithEvents Info_La1 As Button
    Friend WithEvents txt_Matecode As TextBox
    Friend WithEvents Info_La0 As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents txt_Standard As TextBox
    Friend WithEvents txt_Sel As TextBox
End Class
