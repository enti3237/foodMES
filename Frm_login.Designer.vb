<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Frm_Login
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm_Login))
        Me.Com_Can = New System.Windows.Forms.Button()
        Me.Com_Ok = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Text_Man_Pass = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Text_Man_Code = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.PictureBox3 = New System.Windows.Forms.PictureBox()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Com_Can
        '
        Me.Com_Can.Location = New System.Drawing.Point(846, 36)
        Me.Com_Can.Name = "Com_Can"
        Me.Com_Can.Size = New System.Drawing.Size(81, 33)
        Me.Com_Can.TabIndex = 23
        Me.Com_Can.Text = "취소"
        Me.Com_Can.UseVisualStyleBackColor = True
        '
        'Com_Ok
        '
        Me.Com_Ok.BackColor = System.Drawing.Color.YellowGreen
        Me.Com_Ok.Font = New System.Drawing.Font("Californian FB", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Com_Ok.ForeColor = System.Drawing.SystemColors.InfoText
        Me.Com_Ok.Location = New System.Drawing.Point(519, 438)
        Me.Com_Ok.Name = "Com_Ok"
        Me.Com_Ok.Size = New System.Drawing.Size(102, 33)
        Me.Com_Ok.TabIndex = 22
        Me.Com_Ok.Text = "LOGIN"
        Me.Com_Ok.UseVisualStyleBackColor = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("문체부 제목 돋음체", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label2.Location = New System.Drawing.Point(417, 366)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(31, 19)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "PW"
        '
        'Text_Man_Pass
        '
        Me.Text_Man_Pass.Location = New System.Drawing.Point(421, 388)
        Me.Text_Man_Pass.Name = "Text_Man_Pass"
        Me.Text_Man_Pass.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.Text_Man_Pass.Size = New System.Drawing.Size(310, 21)
        Me.Text_Man_Pass.TabIndex = 20
        Me.Text_Man_Pass.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("문체부 제목 돋음체", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label1.Location = New System.Drawing.Point(417, 298)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(22, 19)
        Me.Label1.TabIndex = 19
        Me.Label1.Text = "ID"
        '
        'Text_Man_Code
        '
        Me.Text_Man_Code.BackColor = System.Drawing.SystemColors.Window
        Me.Text_Man_Code.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Text_Man_Code.Location = New System.Drawing.Point(421, 320)
        Me.Text_Man_Code.Name = "Text_Man_Code"
        Me.Text_Man_Code.Size = New System.Drawing.Size(310, 26)
        Me.Text_Man_Code.TabIndex = 18
        Me.Text_Man_Code.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.YellowGreen
        Me.Button1.Location = New System.Drawing.Point(838, 12)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(89, 18)
        Me.Button1.TabIndex = 24
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = False
        Me.Button1.Visible = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImage = CType(resources.GetObject("PictureBox1.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(-5, -66)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(1280, 358)
        Me.PictureBox1.TabIndex = 25
        Me.PictureBox1.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.BackgroundImage = CType(resources.GetObject("PictureBox2.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(1123, 425)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(152, 143)
        Me.PictureBox2.TabIndex = 26
        Me.PictureBox2.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.BackgroundImage = CType(resources.GetObject("PictureBox3.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(-5, 214)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(384, 354)
        Me.PictureBox3.TabIndex = 28
        Me.PictureBox3.TabStop = False
        '
        'Frm_Login
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.HighlightText
        Me.ClientSize = New System.Drawing.Size(1276, 563)
        Me.Controls.Add(Me.PictureBox3)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Com_Ok)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Com_Can)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Text_Man_Pass)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Text_Man_Code)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Frm_Login"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "로그인"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Com_Can As Button
    Friend WithEvents Com_Ok As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents Text_Man_Pass As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Text_Man_Code As TextBox
    Friend WithEvents Button1 As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents PictureBox2 As PictureBox
    Friend WithEvents PictureBox3 As PictureBox
End Class
