<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Frm_PC_Report5
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
        Me.Di_Panel1 = New System.Windows.Forms.Panel()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Com_Excel_Print = New System.Windows.Forms.Button()
        Me.Search_Com = New System.Windows.Forms.Button()
        Me.Panel_Main = New System.Windows.Forms.Panel()
        Me.Grid_Main = New System.Windows.Forms.DataGridView()
        Me.Com_Excel = New System.Windows.Forms.Button()
        Me.Panel_Excel = New System.Windows.Forms.Panel()
        Me.Panel_Menu = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.E_Time = New System.Windows.Forms.TextBox()
        Me.E_Day = New System.Windows.Forms.TextBox()
        Me.S_Time = New System.Windows.Forms.TextBox()
        Me.S_Day = New System.Windows.Forms.TextBox()
        Me.Com_Contract = New System.Windows.Forms.Button()
        Me.Color_Label = New System.Windows.Forms.Label()
        Me.Di_Panel2 = New System.Windows.Forms.Panel()
        Me.Panel_Main.SuspendLayout()
        CType(Me.Grid_Main, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel_Menu.SuspendLayout()
        Me.Di_Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Di_Panel1
        '
        Me.Di_Panel1.BackColor = System.Drawing.Color.CornflowerBlue
        Me.Di_Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Di_Panel1.Location = New System.Drawing.Point(0, 80)
        Me.Di_Panel1.Margin = New System.Windows.Forms.Padding(0)
        Me.Di_Panel1.Name = "Di_Panel1"
        Me.Di_Panel1.Size = New System.Drawing.Size(1197, 2)
        Me.Di_Panel1.TabIndex = 84
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button2.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.Black
        Me.Button2.Location = New System.Drawing.Point(12, 11)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(340, 27)
        Me.Button2.TabIndex = 95
        Me.Button2.Text = "검색기간"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Com_Excel_Print
        '
        Me.Com_Excel_Print.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Com_Excel_Print.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Com_Excel_Print.Location = New System.Drawing.Point(472, 15)
        Me.Com_Excel_Print.Name = "Com_Excel_Print"
        Me.Com_Excel_Print.Size = New System.Drawing.Size(108, 54)
        Me.Com_Excel_Print.TabIndex = 13
        Me.Com_Excel_Print.Text = "Excel Printer"
        Me.Com_Excel_Print.UseVisualStyleBackColor = False
        Me.Com_Excel_Print.Visible = False
        '
        'Search_Com
        '
        Me.Search_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Search_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Search_Com.ForeColor = System.Drawing.Color.Black
        Me.Search_Com.Location = New System.Drawing.Point(358, 15)
        Me.Search_Com.Name = "Search_Com"
        Me.Search_Com.Size = New System.Drawing.Size(108, 54)
        Me.Search_Com.TabIndex = 11
        Me.Search_Com.Text = "검색"
        Me.Search_Com.UseVisualStyleBackColor = False
        '
        'Panel_Main
        '
        Me.Panel_Main.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel_Main.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel_Main.Controls.Add(Me.Grid_Main)
        Me.Panel_Main.Location = New System.Drawing.Point(12, 135)
        Me.Panel_Main.Name = "Panel_Main"
        Me.Panel_Main.Size = New System.Drawing.Size(1168, 534)
        Me.Panel_Main.TabIndex = 81
        '
        'Grid_Main
        '
        Me.Grid_Main.AllowUserToAddRows = False
        Me.Grid_Main.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Grid_Main.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Grid_Main.Location = New System.Drawing.Point(7, 11)
        Me.Grid_Main.Margin = New System.Windows.Forms.Padding(2)
        Me.Grid_Main.Name = "Grid_Main"
        Me.Grid_Main.RowTemplate.Height = 30
        Me.Grid_Main.Size = New System.Drawing.Size(1147, 505)
        Me.Grid_Main.TabIndex = 18
        '
        'Com_Excel
        '
        Me.Com_Excel.AutoSize = True
        Me.Com_Excel.FlatAppearance.BorderSize = 0
        Me.Com_Excel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange
        Me.Com_Excel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Com_Excel.Location = New System.Drawing.Point(110, 94)
        Me.Com_Excel.Margin = New System.Windows.Forms.Padding(0)
        Me.Com_Excel.Name = "Com_Excel"
        Me.Com_Excel.Size = New System.Drawing.Size(105, 25)
        Me.Com_Excel.TabIndex = 82
        Me.Com_Excel.Text = "공정 Excel"
        Me.Com_Excel.UseVisualStyleBackColor = True
        '
        'Panel_Excel
        '
        Me.Panel_Excel.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel_Excel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel_Excel.Cursor = System.Windows.Forms.Cursors.Default
        Me.Panel_Excel.Location = New System.Drawing.Point(3, 161)
        Me.Panel_Excel.Name = "Panel_Excel"
        Me.Panel_Excel.Size = New System.Drawing.Size(1164, 534)
        Me.Panel_Excel.TabIndex = 83
        '
        'Panel_Menu
        '
        Me.Panel_Menu.Controls.Add(Me.Label1)
        Me.Panel_Menu.Controls.Add(Me.E_Time)
        Me.Panel_Menu.Controls.Add(Me.E_Day)
        Me.Panel_Menu.Controls.Add(Me.S_Time)
        Me.Panel_Menu.Controls.Add(Me.S_Day)
        Me.Panel_Menu.Controls.Add(Me.Button2)
        Me.Panel_Menu.Controls.Add(Me.Com_Excel_Print)
        Me.Panel_Menu.Controls.Add(Me.Search_Com)
        Me.Panel_Menu.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel_Menu.Location = New System.Drawing.Point(0, 0)
        Me.Panel_Menu.Name = "Panel_Menu"
        Me.Panel_Menu.Size = New System.Drawing.Size(1197, 80)
        Me.Panel_Menu.TabIndex = 78
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(175, 50)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(14, 12)
        Me.Label1.TabIndex = 100
        Me.Label1.Text = "~"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'E_Time
        '
        Me.E_Time.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.E_Time.Location = New System.Drawing.Point(298, 44)
        Me.E_Time.Name = "E_Time"
        Me.E_Time.Size = New System.Drawing.Size(54, 25)
        Me.E_Time.TabIndex = 99
        Me.E_Time.Text = "00:12"
        Me.E_Time.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'E_Day
        '
        Me.E_Day.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.E_Day.Location = New System.Drawing.Point(194, 44)
        Me.E_Day.Name = "E_Day"
        Me.E_Day.Size = New System.Drawing.Size(98, 25)
        Me.E_Day.TabIndex = 98
        Me.E_Day.Text = "2019-03-01"
        Me.E_Day.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'S_Time
        '
        Me.S_Time.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.S_Time.Location = New System.Drawing.Point(116, 44)
        Me.S_Time.Name = "S_Time"
        Me.S_Time.Size = New System.Drawing.Size(54, 25)
        Me.S_Time.TabIndex = 97
        Me.S_Time.Text = "00:12"
        Me.S_Time.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'S_Day
        '
        Me.S_Day.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.S_Day.Location = New System.Drawing.Point(12, 44)
        Me.S_Day.Name = "S_Day"
        Me.S_Day.Size = New System.Drawing.Size(98, 25)
        Me.S_Day.TabIndex = 96
        Me.S_Day.Text = "2019-03-01"
        Me.S_Day.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Com_Contract
        '
        Me.Com_Contract.AutoSize = True
        Me.Com_Contract.FlatAppearance.BorderSize = 0
        Me.Com_Contract.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange
        Me.Com_Contract.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Com_Contract.Location = New System.Drawing.Point(12, 94)
        Me.Com_Contract.Margin = New System.Windows.Forms.Padding(0)
        Me.Com_Contract.Name = "Com_Contract"
        Me.Com_Contract.Size = New System.Drawing.Size(98, 25)
        Me.Com_Contract.TabIndex = 79
        Me.Com_Contract.Text = "공정별 내역"
        Me.Com_Contract.UseVisualStyleBackColor = True
        '
        'Color_Label
        '
        Me.Color_Label.Location = New System.Drawing.Point(180, 3)
        Me.Color_Label.Name = "Color_Label"
        Me.Color_Label.Size = New System.Drawing.Size(47, 8)
        Me.Color_Label.TabIndex = 0
        '
        'Di_Panel2
        '
        Me.Di_Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Di_Panel2.BackColor = System.Drawing.Color.Silver
        Me.Di_Panel2.Controls.Add(Me.Color_Label)
        Me.Di_Panel2.Location = New System.Drawing.Point(12, 121)
        Me.Di_Panel2.Margin = New System.Windows.Forms.Padding(2)
        Me.Di_Panel2.Name = "Di_Panel2"
        Me.Di_Panel2.Size = New System.Drawing.Size(1946, 5)
        Me.Di_Panel2.TabIndex = 80
        '
        'Frm_PC_Report5
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Di_Panel1)
        Me.Controls.Add(Me.Panel_Main)
        Me.Controls.Add(Me.Com_Excel)
        Me.Controls.Add(Me.Panel_Excel)
        Me.Controls.Add(Me.Panel_Menu)
        Me.Controls.Add(Me.Com_Contract)
        Me.Controls.Add(Me.Di_Panel2)
        Me.Name = "Frm_PC_Report5"
        Me.Size = New System.Drawing.Size(1197, 685)
        Me.Panel_Main.ResumeLayout(False)
        CType(Me.Grid_Main, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel_Menu.ResumeLayout(False)
        Me.Panel_Menu.PerformLayout()
        Me.Di_Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Di_Panel1 As Panel
    Friend WithEvents Button2 As Button
    Friend WithEvents Com_Excel_Print As Button
    Friend WithEvents Search_Com As Button
    Friend WithEvents Panel_Main As Panel
    Friend WithEvents Grid_Main As DataGridView
    Friend WithEvents Com_Excel As Button
    Friend WithEvents Panel_Excel As Panel
    Friend WithEvents Panel_Menu As Panel
    Friend WithEvents Label1 As Label
    Friend WithEvents E_Time As TextBox
    Friend WithEvents E_Day As TextBox
    Friend WithEvents S_Time As TextBox
    Friend WithEvents S_Day As TextBox
    Friend WithEvents Com_Contract As Button
    Friend WithEvents Color_Label As Label
    Friend WithEvents Di_Panel2 As Panel
End Class
