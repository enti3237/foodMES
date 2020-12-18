<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Excel_Panel
    Inherits System.Windows.Forms.UserControl

    'UserControl은 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
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
        Me.btnSave = New System.Windows.Forms.Button()
        Me.txtValue = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtY = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtX = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnRead = New System.Windows.Forms.Button()
        Me.btnWrite = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.TrackBar1 = New System.Windows.Forms.TrackBar()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        CType(Me.TrackBar1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(388, 393)
        Me.btnSave.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(77, 31)
        Me.btnSave.TabIndex = 22
        Me.btnSave.Text = "저장"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'txtValue
        '
        Me.txtValue.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtValue.Location = New System.Drawing.Point(158, 395)
        Me.txtValue.Margin = New System.Windows.Forms.Padding(2)
        Me.txtValue.Name = "txtValue"
        Me.txtValue.Size = New System.Drawing.Size(204, 25)
        Me.txtValue.TabIndex = 21
        Me.txtValue.Text = "Test"
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(137, 401)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(22, 15)
        Me.Label3.TabIndex = 20
        Me.Label3.Text = "값"
        '
        'txtY
        '
        Me.txtY.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtY.Location = New System.Drawing.Point(91, 395)
        Me.txtY.Margin = New System.Windows.Forms.Padding(2)
        Me.txtY.Name = "txtY"
        Me.txtY.Size = New System.Drawing.Size(39, 25)
        Me.txtY.TabIndex = 19
        Me.txtY.Text = "2"
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(72, 401)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(15, 15)
        Me.Label2.TabIndex = 18
        Me.Label2.Text = "Y"
        '
        'txtX
        '
        Me.txtX.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtX.Location = New System.Drawing.Point(27, 395)
        Me.txtX.Margin = New System.Windows.Forms.Padding(2)
        Me.txtX.Name = "txtX"
        Me.txtX.Size = New System.Drawing.Size(39, 25)
        Me.txtX.TabIndex = 17
        Me.txtX.Text = "22"
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 401)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(16, 15)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "X"
        '
        'btnRead
        '
        Me.btnRead.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnRead.Location = New System.Drawing.Point(526, 219)
        Me.btnRead.Margin = New System.Windows.Forms.Padding(2)
        Me.btnRead.Name = "btnRead"
        Me.btnRead.Size = New System.Drawing.Size(77, 31)
        Me.btnRead.TabIndex = 15
        Me.btnRead.Text = "읽기"
        Me.btnRead.UseVisualStyleBackColor = True
        '
        'btnWrite
        '
        Me.btnWrite.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnWrite.Location = New System.Drawing.Point(425, 219)
        Me.btnWrite.Margin = New System.Windows.Forms.Padding(2)
        Me.btnWrite.Name = "btnWrite"
        Me.btnWrite.Size = New System.Drawing.Size(77, 31)
        Me.btnWrite.TabIndex = 14
        Me.btnWrite.Text = "쓰기"
        Me.btnWrite.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.btnWrite)
        Me.Panel1.Controls.Add(Me.btnRead)
        Me.Panel1.Location = New System.Drawing.Point(2, 1)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(795, 378)
        Me.Panel1.TabIndex = 13
        '
        'TrackBar1
        '
        Me.TrackBar1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TrackBar1.LargeChange = 10
        Me.TrackBar1.Location = New System.Drawing.Point(645, 385)
        Me.TrackBar1.Margin = New System.Windows.Forms.Padding(2)
        Me.TrackBar1.Maximum = 100
        Me.TrackBar1.Minimum = 20
        Me.TrackBar1.Name = "TrackBar1"
        Me.TrackBar1.Size = New System.Drawing.Size(152, 56)
        Me.TrackBar1.SmallChange = 10
        Me.TrackBar1.TabIndex = 12
        Me.TrackBar1.TickFrequency = 10
        Me.TrackBar1.TickStyle = System.Windows.Forms.TickStyle.Both
        Me.TrackBar1.Value = 70
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(469, 393)
        Me.Button1.Margin = New System.Windows.Forms.Padding(2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(77, 31)
        Me.Button1.TabIndex = 23
        Me.Button1.Text = "출력"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button2.Location = New System.Drawing.Point(550, 393)
        Me.Button2.Margin = New System.Windows.Forms.Padding(2)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(77, 31)
        Me.Button2.TabIndex = 24
        Me.Button2.Text = "종료"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Excel_Panel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.txtValue)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtY)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtX)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.TrackBar1)
        Me.Name = "Excel_Panel"
        Me.Size = New System.Drawing.Size(799, 446)
        Me.Panel1.ResumeLayout(False)
        CType(Me.TrackBar1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnSave As Button
    Friend WithEvents txtValue As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents txtY As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents txtX As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents btnRead As Button
    Friend WithEvents btnWrite As Button
    Friend WithEvents Panel1 As Panel
    Friend WithEvents TrackBar1 As TrackBar
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
End Class
