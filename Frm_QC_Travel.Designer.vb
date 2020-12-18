<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_QC_Travel
    Inherits System.Windows.Forms.UserControl

    'UserControl은 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
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
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Grid_Code = New System.Windows.Forms.DataGridView()
        Me.Panel_Menu = New System.Windows.Forms.Panel()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Print_Com = New System.Windows.Forms.Button()
        Me.Save_Com = New System.Windows.Forms.Button()
        Me.Insert_Com = New System.Windows.Forms.Button()
        Me.Search_Com = New System.Windows.Forms.Button()
        Me.Search_Text = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Di_Panel1 = New System.Windows.Forms.Panel()
        Me.Color_Label = New System.Windows.Forms.Label()
        Me.Di_Panel2 = New System.Windows.Forms.Panel()
        Me.Com_QC = New System.Windows.Forms.Button()
        Me.Com_Excel = New System.Windows.Forms.Button()
        Me.Excel_Panel = New System.Windows.Forms.Panel()
        Me.Panel2.SuspendLayout()
        CType(Me.Grid_Code, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel_Menu.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Di_Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.Grid_Code)
        Me.Panel2.Location = New System.Drawing.Point(3, 128)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1628, 551)
        Me.Panel2.TabIndex = 99
        '
        'Grid_Code
        '
        Me.Grid_Code.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.Grid_Code.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Grid_Code.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Grid_Code.Location = New System.Drawing.Point(0, 0)
        Me.Grid_Code.Name = "Grid_Code"
        Me.Grid_Code.RowTemplate.Height = 23
        Me.Grid_Code.Size = New System.Drawing.Size(1628, 551)
        Me.Grid_Code.TabIndex = 0
        '
        'Panel_Menu
        '
        Me.Panel_Menu.Controls.Add(Me.Button2)
        Me.Panel_Menu.Controls.Add(Me.Button1)
        Me.Panel_Menu.Controls.Add(Me.Print_Com)
        Me.Panel_Menu.Controls.Add(Me.Save_Com)
        Me.Panel_Menu.Controls.Add(Me.Insert_Com)
        Me.Panel_Menu.Controls.Add(Me.Search_Com)
        Me.Panel_Menu.Controls.Add(Me.Search_Text)
        Me.Panel_Menu.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel_Menu.Location = New System.Drawing.Point(0, 2)
        Me.Panel_Menu.Name = "Panel_Menu"
        Me.Panel_Menu.Size = New System.Drawing.Size(1438, 80)
        Me.Panel_Menu.TabIndex = 92
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button2.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button2.Location = New System.Drawing.Point(376, 14)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(108, 54)
        Me.Button2.TabIndex = 21
        Me.Button2.Text = "삭제"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button1.Location = New System.Drawing.Point(11, 14)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(131, 23)
        Me.Button1.TabIndex = 20
        Me.Button1.Text = "제품명"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Print_Com
        '
        Me.Print_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Print_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Print_Com.Location = New System.Drawing.Point(490, 14)
        Me.Print_Com.Name = "Print_Com"
        Me.Print_Com.Size = New System.Drawing.Size(108, 54)
        Me.Print_Com.TabIndex = 19
        Me.Print_Com.Text = "출력"
        Me.Print_Com.UseVisualStyleBackColor = False
        '
        'Save_Com
        '
        Me.Save_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Save_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Save_Com.ForeColor = System.Drawing.Color.Black
        Me.Save_Com.Location = New System.Drawing.Point(747, 14)
        Me.Save_Com.Name = "Save_Com"
        Me.Save_Com.Size = New System.Drawing.Size(108, 54)
        Me.Save_Com.TabIndex = 14
        Me.Save_Com.Text = "저장"
        Me.Save_Com.UseVisualStyleBackColor = False
        Me.Save_Com.Visible = False
        '
        'Insert_Com
        '
        Me.Insert_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Insert_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Insert_Com.Location = New System.Drawing.Point(262, 14)
        Me.Insert_Com.Name = "Insert_Com"
        Me.Insert_Com.Size = New System.Drawing.Size(108, 54)
        Me.Insert_Com.TabIndex = 12
        Me.Insert_Com.Text = "추가"
        Me.Insert_Com.UseVisualStyleBackColor = False
        '
        'Search_Com
        '
        Me.Search_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Search_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Search_Com.ForeColor = System.Drawing.Color.Black
        Me.Search_Com.Location = New System.Drawing.Point(148, 14)
        Me.Search_Com.Name = "Search_Com"
        Me.Search_Com.Size = New System.Drawing.Size(108, 54)
        Me.Search_Com.TabIndex = 11
        Me.Search_Com.Text = "검색"
        Me.Search_Com.UseVisualStyleBackColor = False
        '
        'Search_Text
        '
        Me.Search_Text.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Search_Text.Location = New System.Drawing.Point(12, 43)
        Me.Search_Text.Name = "Search_Text"
        Me.Search_Text.Size = New System.Drawing.Size(131, 25)
        Me.Search_Text.TabIndex = 10
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(53, 12)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "개정이력"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.DataGridView1)
        Me.Panel1.Location = New System.Drawing.Point(10, 685)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1624, 224)
        Me.Panel1.TabIndex = 98
        '
        'DataGridView1
        '
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(8, 29)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowTemplate.Height = 23
        Me.DataGridView1.Size = New System.Drawing.Size(1613, 195)
        Me.DataGridView1.TabIndex = 0
        '
        'Di_Panel1
        '
        Me.Di_Panel1.BackColor = System.Drawing.Color.CornflowerBlue
        Me.Di_Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Di_Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Di_Panel1.Margin = New System.Windows.Forms.Padding(0)
        Me.Di_Panel1.Name = "Di_Panel1"
        Me.Di_Panel1.Size = New System.Drawing.Size(1438, 2)
        Me.Di_Panel1.TabIndex = 97
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
        Me.Di_Panel2.Location = New System.Drawing.Point(-1, 117)
        Me.Di_Panel2.Margin = New System.Windows.Forms.Padding(2)
        Me.Di_Panel2.Name = "Di_Panel2"
        Me.Di_Panel2.Size = New System.Drawing.Size(1802, 5)
        Me.Di_Panel2.TabIndex = 94
        '
        'Com_QC
        '
        Me.Com_QC.AutoSize = True
        Me.Com_QC.FlatAppearance.BorderSize = 0
        Me.Com_QC.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange
        Me.Com_QC.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Com_QC.Location = New System.Drawing.Point(-1, 90)
        Me.Com_QC.Margin = New System.Windows.Forms.Padding(0)
        Me.Com_QC.Name = "Com_QC"
        Me.Com_QC.Size = New System.Drawing.Size(98, 25)
        Me.Com_QC.TabIndex = 93
        Me.Com_QC.Text = "공정검사"
        Me.Com_QC.UseVisualStyleBackColor = True
        '
        'Com_Excel
        '
        Me.Com_Excel.AutoSize = True
        Me.Com_Excel.FlatAppearance.BorderSize = 0
        Me.Com_Excel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange
        Me.Com_Excel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Com_Excel.Location = New System.Drawing.Point(97, 90)
        Me.Com_Excel.Margin = New System.Windows.Forms.Padding(0)
        Me.Com_Excel.Name = "Com_Excel"
        Me.Com_Excel.Size = New System.Drawing.Size(107, 25)
        Me.Com_Excel.TabIndex = 95
        Me.Com_Excel.Text = "공정검사성적서"
        Me.Com_Excel.UseVisualStyleBackColor = True
        '
        'Excel_Panel
        '
        Me.Excel_Panel.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Excel_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Excel_Panel.Location = New System.Drawing.Point(20, 131)
        Me.Excel_Panel.Name = "Excel_Panel"
        Me.Excel_Panel.Size = New System.Drawing.Size(1235, 471)
        Me.Excel_Panel.TabIndex = 96
        '
        'Frm_QC_Travel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel_Menu)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Di_Panel1)
        Me.Controls.Add(Me.Di_Panel2)
        Me.Controls.Add(Me.Com_QC)
        Me.Controls.Add(Me.Com_Excel)
        Me.Controls.Add(Me.Excel_Panel)
        Me.Name = "Frm_QC_Travel"
        Me.Size = New System.Drawing.Size(1438, 855)
        Me.Panel2.ResumeLayout(False)
        CType(Me.Grid_Code, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel_Menu.ResumeLayout(False)
        Me.Panel_Menu.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Di_Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Panel2 As Panel
    Friend WithEvents Grid_Code As DataGridView
    Friend WithEvents Panel_Menu As Panel
    Friend WithEvents Button2 As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents Print_Com As Button
    Friend WithEvents Save_Com As Button
    Friend WithEvents Insert_Com As Button
    Friend WithEvents Search_Com As Button
    Friend WithEvents Search_Text As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Di_Panel1 As Panel
    Friend WithEvents Color_Label As Label
    Friend WithEvents Di_Panel2 As Panel
    Friend WithEvents Com_QC As Button
    Friend WithEvents Com_Excel As Button
    Friend WithEvents Excel_Panel As Panel
End Class
