<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Bom_Insert
    Inherits System.Windows.Forms.Form

    'Form은 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Process_bigo = New System.Windows.Forms.TextBox()
        Me.process_name = New System.Windows.Forms.TextBox()
        Me.process_code = New System.Windows.Forms.TextBox()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Insert_Com = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Sub_Code = New System.Windows.Forms.TextBox()
        Me.Button9 = New System.Windows.Forms.Button()
        Me.btnJobSearch = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Unit = New System.Windows.Forms.TextBox()
        Me.Qty = New System.Windows.Forms.TextBox()
        Me.Sub_Name = New System.Windows.Forms.TextBox()
        Me.sun = New System.Windows.Forms.TextBox()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Btn_Save = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(91, Byte), Integer), CType(CType(155, Byte), Integer), CType(CType(213, Byte), Integer))
        Me.Panel1.Controls.Add(Me.Process_bigo)
        Me.Panel1.Controls.Add(Me.process_name)
        Me.Panel1.Controls.Add(Me.process_code)
        Me.Panel1.Controls.Add(Me.Button3)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.Button2)
        Me.Panel1.Controls.Add(Me.Insert_Com)
        Me.Panel1.Location = New System.Drawing.Point(14, 26)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(887, 179)
        Me.Panel1.TabIndex = 0
        '
        'Process_bigo
        '
        Me.Process_bigo.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Process_bigo.Location = New System.Drawing.Point(159, 68)
        Me.Process_bigo.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Process_bigo.Name = "Process_bigo"
        Me.Process_bigo.Size = New System.Drawing.Size(686, 29)
        Me.Process_bigo.TabIndex = 240
        '
        'process_name
        '
        Me.process_name.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.process_name.Location = New System.Drawing.Point(631, 18)
        Me.process_name.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.process_name.Name = "process_name"
        Me.process_name.Size = New System.Drawing.Size(214, 29)
        Me.process_name.TabIndex = 239
        '
        'process_code
        '
        Me.process_code.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.process_code.Location = New System.Drawing.Point(159, 18)
        Me.process_code.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.process_code.Name = "process_code"
        Me.process_code.ReadOnly = True
        Me.process_code.Size = New System.Drawing.Size(283, 29)
        Me.process_code.TabIndex = 238
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button3.Enabled = False
        Me.Button3.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button3.ForeColor = System.Drawing.Color.Black
        Me.Button3.Location = New System.Drawing.Point(6, 64)
        Me.Button3.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(146, 34)
        Me.Button3.TabIndex = 237
        Me.Button3.Text = "비고"
        Me.Button3.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button1.Enabled = False
        Me.Button1.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.Black
        Me.Button1.Location = New System.Drawing.Point(470, 14)
        Me.Button1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(155, 34)
        Me.Button1.TabIndex = 236
        Me.Button1.Text = "제품명"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button2.Enabled = False
        Me.Button2.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.Black
        Me.Button2.Location = New System.Drawing.Point(9, 9)
        Me.Button2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(146, 34)
        Me.Button2.TabIndex = 235
        Me.Button2.Text = "제품코드"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Insert_Com
        '
        Me.Insert_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Insert_Com.Location = New System.Drawing.Point(3, 116)
        Me.Insert_Com.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Insert_Com.Name = "Insert_Com"
        Me.Insert_Com.Size = New System.Drawing.Size(880, 59)
        Me.Insert_Com.TabIndex = 6
        Me.Insert_Com.Text = "저장"
        Me.Insert_Com.UseVisualStyleBackColor = False
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.White
        Me.Panel2.Controls.Add(Me.Sub_Code)
        Me.Panel2.Controls.Add(Me.Button9)
        Me.Panel2.Controls.Add(Me.btnJobSearch)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.Unit)
        Me.Panel2.Controls.Add(Me.Qty)
        Me.Panel2.Controls.Add(Me.Sub_Name)
        Me.Panel2.Controls.Add(Me.sun)
        Me.Panel2.Controls.Add(Me.Button7)
        Me.Panel2.Controls.Add(Me.Button6)
        Me.Panel2.Controls.Add(Me.Button4)
        Me.Panel2.Controls.Add(Me.Btn_Save)
        Me.Panel2.Location = New System.Drawing.Point(17, 212)
        Me.Panel2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(887, 209)
        Me.Panel2.TabIndex = 1
        '
        'Sub_Code
        '
        Me.Sub_Code.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Sub_Code.Location = New System.Drawing.Point(743, 18)
        Me.Sub_Code.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Sub_Code.Name = "Sub_Code"
        Me.Sub_Code.Size = New System.Drawing.Size(111, 29)
        Me.Sub_Code.TabIndex = 248
        Me.Sub_Code.Text = "원부재료코드"
        '
        'Button9
        '
        Me.Button9.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button9.Enabled = False
        Me.Button9.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button9.ForeColor = System.Drawing.Color.Black
        Me.Button9.Location = New System.Drawing.Point(6, 63)
        Me.Button9.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(146, 34)
        Me.Button9.TabIndex = 247
        Me.Button9.Text = "규격"
        Me.Button9.UseVisualStyleBackColor = False
        '
        'btnJobSearch
        '
        Me.btnJobSearch.BackColor = System.Drawing.Color.White
        Me.btnJobSearch.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnJobSearch.Font = New System.Drawing.Font("맑은 고딕", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.btnJobSearch.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.btnJobSearch.Location = New System.Drawing.Point(639, 18)
        Me.btnJobSearch.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnJobSearch.Name = "btnJobSearch"
        Me.btnJobSearch.Size = New System.Drawing.Size(37, 32)
        Me.btnJobSearch.TabIndex = 245
        Me.btnJobSearch.Text = "⌕"
        Me.btnJobSearch.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(821, 150)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(50, 15)
        Me.Label1.TabIndex = 241
        Me.Label1.Text = "Label1"
        Me.Label1.Visible = False
        '
        'Unit
        '
        Me.Unit.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Unit.Location = New System.Drawing.Point(504, 67)
        Me.Unit.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Unit.Name = "Unit"
        Me.Unit.ReadOnly = True
        Me.Unit.Size = New System.Drawing.Size(139, 29)
        Me.Unit.TabIndex = 244
        '
        'Qty
        '
        Me.Qty.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Qty.Location = New System.Drawing.Point(159, 68)
        Me.Qty.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Qty.Name = "Qty"
        Me.Qty.Size = New System.Drawing.Size(175, 29)
        Me.Qty.TabIndex = 243
        '
        'Sub_Name
        '
        Me.Sub_Name.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Sub_Name.Location = New System.Drawing.Point(421, 24)
        Me.Sub_Name.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Sub_Name.Name = "Sub_Name"
        Me.Sub_Name.ReadOnly = True
        Me.Sub_Name.Size = New System.Drawing.Size(212, 29)
        Me.Sub_Name.TabIndex = 242
        '
        'sun
        '
        Me.sun.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.sun.Location = New System.Drawing.Point(159, 23)
        Me.sun.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.sun.Name = "sun"
        Me.sun.ReadOnly = True
        Me.sun.Size = New System.Drawing.Size(89, 29)
        Me.sun.TabIndex = 241
        '
        'Button7
        '
        Me.Button7.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button7.Enabled = False
        Me.Button7.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button7.ForeColor = System.Drawing.Color.Black
        Me.Button7.Location = New System.Drawing.Point(352, 63)
        Me.Button7.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(146, 34)
        Me.Button7.TabIndex = 239
        Me.Button7.Text = "단위"
        Me.Button7.UseVisualStyleBackColor = False
        '
        'Button6
        '
        Me.Button6.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button6.Enabled = False
        Me.Button6.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button6.ForeColor = System.Drawing.Color.Black
        Me.Button6.Location = New System.Drawing.Point(269, 21)
        Me.Button6.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(146, 34)
        Me.Button6.TabIndex = 238
        Me.Button6.Text = "원부재료명"
        Me.Button6.UseVisualStyleBackColor = False
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button4.Enabled = False
        Me.Button4.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button4.ForeColor = System.Drawing.Color.Black
        Me.Button4.Location = New System.Drawing.Point(6, 19)
        Me.Button4.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(146, 34)
        Me.Button4.TabIndex = 236
        Me.Button4.Text = "순번"
        Me.Button4.UseVisualStyleBackColor = False
        '
        'Btn_Save
        '
        Me.Btn_Save.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Btn_Save.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Btn_Save.ForeColor = System.Drawing.Color.Black
        Me.Btn_Save.Location = New System.Drawing.Point(2, 142)
        Me.Btn_Save.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Btn_Save.Name = "Btn_Save"
        Me.Btn_Save.Size = New System.Drawing.Size(880, 62)
        Me.Btn_Save.TabIndex = 233
        Me.Btn_Save.Text = "저장"
        Me.Btn_Save.UseVisualStyleBackColor = False
        '
        'Bom_Insert
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(914, 438)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel2)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "Bom_Insert"
        Me.Text = "BOM 추가"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Insert_Com As Button
    Friend WithEvents Btn_Save As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Process_bigo As TextBox
    Friend WithEvents process_name As TextBox
    Friend WithEvents process_code As TextBox
    Friend WithEvents Unit As TextBox
    Friend WithEvents sun As TextBox
    Friend WithEvents Button7 As Button
    Friend WithEvents Button6 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Qty As TextBox
    Friend WithEvents Sub_Name As TextBox
    Friend WithEvents btnJobSearch As Button
    Friend WithEvents Sub_Code As TextBox
    Friend WithEvents Button9 As Button
End Class
