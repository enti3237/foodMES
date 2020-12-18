<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Claim_Insert
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button14 = New System.Windows.Forms.Button()
        Me.Button13 = New System.Windows.Forms.Button()
        Me.QC_DATE = New System.Windows.Forms.TextBox()
        Me.QC_BIGO = New System.Windows.Forms.TextBox()
        Me.QC_CM_NAME = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.QC_MAN = New System.Windows.Forms.TextBox()
        Me.QC_CONTENT = New System.Windows.Forms.TextBox()
        Me.QC_CODE = New System.Windows.Forms.TextBox()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Btn_Save = New System.Windows.Forms.Button()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(584, 211)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(42, 12)
        Me.Label1.TabIndex = 277
        Me.Label1.Text = "Label1"
        Me.Label1.Visible = False
        '
        'Button14
        '
        Me.Button14.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button14.Enabled = False
        Me.Button14.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button14.ForeColor = System.Drawing.Color.Black
        Me.Button14.Location = New System.Drawing.Point(361, 78)
        Me.Button14.Name = "Button14"
        Me.Button14.Size = New System.Drawing.Size(128, 27)
        Me.Button14.TabIndex = 261
        Me.Button14.Text = "일자"
        Me.Button14.UseVisualStyleBackColor = False
        '
        'Button13
        '
        Me.Button13.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button13.Enabled = False
        Me.Button13.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button13.ForeColor = System.Drawing.Color.Black
        Me.Button13.Location = New System.Drawing.Point(361, 111)
        Me.Button13.Name = "Button13"
        Me.Button13.Size = New System.Drawing.Size(128, 27)
        Me.Button13.TabIndex = 259
        Me.Button13.Text = "비고"
        Me.Button13.UseVisualStyleBackColor = False
        '
        'QC_DATE
        '
        Me.QC_DATE.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.QC_DATE.Location = New System.Drawing.Point(495, 80)
        Me.QC_DATE.Name = "QC_DATE"
        Me.QC_DATE.ReadOnly = True
        Me.QC_DATE.Size = New System.Drawing.Size(156, 25)
        Me.QC_DATE.TabIndex = 253
        '
        'QC_BIGO
        '
        Me.QC_BIGO.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.QC_BIGO.Location = New System.Drawing.Point(495, 111)
        Me.QC_BIGO.Name = "QC_BIGO"
        Me.QC_BIGO.Size = New System.Drawing.Size(156, 25)
        Me.QC_BIGO.TabIndex = 250
        '
        'QC_CM_NAME
        '
        Me.QC_CM_NAME.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.QC_CM_NAME.Location = New System.Drawing.Point(494, 14)
        Me.QC_CM_NAME.Name = "QC_CM_NAME"
        Me.QC_CM_NAME.Size = New System.Drawing.Size(123, 25)
        Me.QC_CM_NAME.TabIndex = 248
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button1.Enabled = False
        Me.Button1.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.Black
        Me.Button1.Location = New System.Drawing.Point(361, 14)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(128, 27)
        Me.Button1.TabIndex = 247
        Me.Button1.Text = "거래처명"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'QC_MAN
        '
        Me.QC_MAN.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.QC_MAN.Location = New System.Drawing.Point(495, 47)
        Me.QC_MAN.Name = "QC_MAN"
        Me.QC_MAN.Size = New System.Drawing.Size(156, 25)
        Me.QC_MAN.TabIndex = 243
        '
        'QC_CONTENT
        '
        Me.QC_CONTENT.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.QC_CONTENT.Location = New System.Drawing.Point(147, 47)
        Me.QC_CONTENT.Multiline = True
        Me.QC_CONTENT.Name = "QC_CONTENT"
        Me.QC_CONTENT.Size = New System.Drawing.Size(208, 91)
        Me.QC_CONTENT.TabIndex = 242
        '
        'QC_CODE
        '
        Me.QC_CODE.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.QC_CODE.Location = New System.Drawing.Point(146, 14)
        Me.QC_CODE.Name = "QC_CODE"
        Me.QC_CODE.ReadOnly = True
        Me.QC_CODE.Size = New System.Drawing.Size(209, 25)
        Me.QC_CODE.TabIndex = 241
        '
        'Button6
        '
        Me.Button6.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button6.Enabled = False
        Me.Button6.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button6.ForeColor = System.Drawing.Color.Black
        Me.Button6.Location = New System.Drawing.Point(13, 45)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(128, 27)
        Me.Button6.TabIndex = 238
        Me.Button6.Text = "내용"
        Me.Button6.UseVisualStyleBackColor = False
        '
        'Button5
        '
        Me.Button5.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button5.Enabled = False
        Me.Button5.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button5.ForeColor = System.Drawing.Color.Black
        Me.Button5.Location = New System.Drawing.Point(361, 47)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(128, 27)
        Me.Button5.TabIndex = 237
        Me.Button5.Text = "담당자"
        Me.Button5.UseVisualStyleBackColor = False
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button4.Enabled = False
        Me.Button4.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button4.ForeColor = System.Drawing.Color.Black
        Me.Button4.Location = New System.Drawing.Point(12, 12)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(128, 27)
        Me.Button4.TabIndex = 236
        Me.Button4.Text = "코드"
        Me.Button4.UseVisualStyleBackColor = False
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.White
        Me.Panel2.Controls.Add(Me.Button2)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.Button14)
        Me.Panel2.Controls.Add(Me.Button13)
        Me.Panel2.Controls.Add(Me.QC_DATE)
        Me.Panel2.Controls.Add(Me.QC_BIGO)
        Me.Panel2.Controls.Add(Me.QC_CM_NAME)
        Me.Panel2.Controls.Add(Me.Button1)
        Me.Panel2.Controls.Add(Me.QC_MAN)
        Me.Panel2.Controls.Add(Me.QC_CONTENT)
        Me.Panel2.Controls.Add(Me.QC_CODE)
        Me.Panel2.Controls.Add(Me.Button6)
        Me.Panel2.Controls.Add(Me.Button5)
        Me.Panel2.Controls.Add(Me.Button4)
        Me.Panel2.Controls.Add(Me.Btn_Save)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(658, 205)
        Me.Panel2.TabIndex = 4
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.White
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Font = New System.Drawing.Font("맑은 고딕", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button2.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.Button2.Location = New System.Drawing.Point(619, 13)
        Me.Button2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(32, 26)
        Me.Button2.TabIndex = 278
        Me.Button2.Text = "⌕"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Btn_Save
        '
        Me.Btn_Save.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Btn_Save.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Btn_Save.ForeColor = System.Drawing.Color.Black
        Me.Btn_Save.Location = New System.Drawing.Point(17, 148)
        Me.Btn_Save.Name = "Btn_Save"
        Me.Btn_Save.Size = New System.Drawing.Size(638, 50)
        Me.Btn_Save.TabIndex = 233
        Me.Btn_Save.Text = "저장"
        Me.Btn_Save.UseVisualStyleBackColor = False
        '
        'Claim_Insert
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(658, 205)
        Me.Controls.Add(Me.Panel2)
        Me.Name = "Claim_Insert"
        Me.Text = "Claim 추가/수정"
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents Button14 As Button
    Friend WithEvents Button13 As Button
    Friend WithEvents QC_DATE As TextBox
    Friend WithEvents QC_BIGO As TextBox
    Friend WithEvents QC_CM_NAME As TextBox
    Friend WithEvents Button1 As Button
    Friend WithEvents QC_MAN As TextBox
    Friend WithEvents QC_CONTENT As TextBox
    Friend WithEvents QC_CODE As TextBox
    Friend WithEvents Button6 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Btn_Save As Button
    Friend WithEvents Button2 As Button
End Class
