﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_PlanLG
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
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Color_Label = New System.Windows.Forms.Label()
        Me.Grid_LGPlan = New System.Windows.Forms.DataGridView()
        Me.Search_Text = New System.Windows.Forms.TextBox()
        Me.Search_Combo = New System.Windows.Forms.ComboBox()
        Me.Save_Com = New System.Windows.Forms.Button()
        Me.Load_Com = New System.Windows.Forms.Button()
        Me.Insert_Com = New System.Windows.Forms.Button()
        Me.LGPlan_Panel = New System.Windows.Forms.Panel()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Grid_Code = New System.Windows.Forms.DataGridView()
        Me.Excel_Com = New System.Windows.Forms.Button()
        Me.Di_Panel2 = New System.Windows.Forms.Panel()
        Me.Com_LGPlan = New System.Windows.Forms.Button()
        Me.Di_Panel1 = New System.Windows.Forms.Panel()
        Me.Panel_Menu = New System.Windows.Forms.Panel()
        Me.Search_Com = New System.Windows.Forms.Button()
        Me.Grid_Info_Combo = New System.Windows.Forms.ComboBox()
        Me.Grid_Info = New System.Windows.Forms.DataGridView()
        Me.Delete_Com = New System.Windows.Forms.Button()
        CType(Me.Grid_LGPlan, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.LGPlan_Panel.SuspendLayout()
        CType(Me.Grid_Code, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Di_Panel2.SuspendLayout()
        Me.Panel_Menu.SuspendLayout()
        CType(Me.Grid_Info, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Color_Label
        '
        Me.Color_Label.Location = New System.Drawing.Point(180, 3)
        Me.Color_Label.Name = "Color_Label"
        Me.Color_Label.Size = New System.Drawing.Size(47, 8)
        Me.Color_Label.TabIndex = 0
        '
        'Grid_LGPlan
        '
        Me.Grid_LGPlan.AllowUserToAddRows = False
        Me.Grid_LGPlan.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Grid_LGPlan.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Grid_LGPlan.Location = New System.Drawing.Point(7, 8)
        Me.Grid_LGPlan.Margin = New System.Windows.Forms.Padding(2)
        Me.Grid_LGPlan.Name = "Grid_LGPlan"
        Me.Grid_LGPlan.RowTemplate.Height = 30
        Me.Grid_LGPlan.Size = New System.Drawing.Size(837, 415)
        Me.Grid_LGPlan.TabIndex = 18
        '
        'Search_Text
        '
        Me.Search_Text.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Search_Text.Location = New System.Drawing.Point(12, 43)
        Me.Search_Text.Name = "Search_Text"
        Me.Search_Text.Size = New System.Drawing.Size(131, 25)
        Me.Search_Text.TabIndex = 10
        '
        'Search_Combo
        '
        Me.Search_Combo.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.Search_Combo.FormattingEnabled = True
        Me.Search_Combo.Location = New System.Drawing.Point(11, 14)
        Me.Search_Combo.Margin = New System.Windows.Forms.Padding(2)
        Me.Search_Combo.Name = "Search_Combo"
        Me.Search_Combo.Size = New System.Drawing.Size(132, 25)
        Me.Search_Combo.TabIndex = 8
        '
        'Save_Com
        '
        Me.Save_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Save_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Save_Com.ForeColor = System.Drawing.Color.Black
        Me.Save_Com.Location = New System.Drawing.Point(262, 14)
        Me.Save_Com.Name = "Save_Com"
        Me.Save_Com.Size = New System.Drawing.Size(108, 54)
        Me.Save_Com.TabIndex = 14
        Me.Save_Com.Text = "저장"
        Me.Save_Com.UseVisualStyleBackColor = False
        '
        'Load_Com
        '
        Me.Load_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Load_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Load_Com.Location = New System.Drawing.Point(594, 14)
        Me.Load_Com.Name = "Load_Com"
        Me.Load_Com.Size = New System.Drawing.Size(108, 54)
        Me.Load_Com.TabIndex = 13
        Me.Load_Com.Text = "Excel File"
        Me.Load_Com.UseVisualStyleBackColor = False
        '
        'Insert_Com
        '
        Me.Insert_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Insert_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Insert_Com.Location = New System.Drawing.Point(376, 14)
        Me.Insert_Com.Name = "Insert_Com"
        Me.Insert_Com.Size = New System.Drawing.Size(108, 54)
        Me.Insert_Com.TabIndex = 12
        Me.Insert_Com.Text = "추가"
        Me.Insert_Com.UseVisualStyleBackColor = False
        '
        'LGPlan_Panel
        '
        Me.LGPlan_Panel.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LGPlan_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LGPlan_Panel.Controls.Add(Me.Grid_LGPlan)
        Me.LGPlan_Panel.Location = New System.Drawing.Point(387, 135)
        Me.LGPlan_Panel.Name = "LGPlan_Panel"
        Me.LGPlan_Panel.Size = New System.Drawing.Size(858, 438)
        Me.LGPlan_Panel.TabIndex = 41
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button2.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.Black
        Me.Button2.Location = New System.Drawing.Point(9, 328)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(359, 27)
        Me.Button2.TabIndex = 37
        Me.Button2.Text = "상세 정보"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Grid_Code
        '
        Me.Grid_Code.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Grid_Code.Location = New System.Drawing.Point(9, 94)
        Me.Grid_Code.Name = "Grid_Code"
        Me.Grid_Code.RowTemplate.Height = 23
        Me.Grid_Code.Size = New System.Drawing.Size(359, 228)
        Me.Grid_Code.TabIndex = 35
        '
        'Excel_Com
        '
        Me.Excel_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Excel_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Excel_Com.Location = New System.Drawing.Point(708, 14)
        Me.Excel_Com.Name = "Excel_Com"
        Me.Excel_Com.Size = New System.Drawing.Size(143, 54)
        Me.Excel_Com.TabIndex = 15
        Me.Excel_Com.Text = "Excel파일로 저장"
        Me.Excel_Com.UseVisualStyleBackColor = False
        Me.Excel_Com.Visible = False
        '
        'Di_Panel2
        '
        Me.Di_Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Di_Panel2.BackColor = System.Drawing.Color.Silver
        Me.Di_Panel2.Controls.Add(Me.Color_Label)
        Me.Di_Panel2.Location = New System.Drawing.Point(387, 121)
        Me.Di_Panel2.Margin = New System.Windows.Forms.Padding(2)
        Me.Di_Panel2.Name = "Di_Panel2"
        Me.Di_Panel2.Size = New System.Drawing.Size(975, 5)
        Me.Di_Panel2.TabIndex = 40
        '
        'Com_LGPlan
        '
        Me.Com_LGPlan.AutoSize = True
        Me.Com_LGPlan.FlatAppearance.BorderSize = 0
        Me.Com_LGPlan.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(91, Byte), Integer), CType(CType(155, Byte), Integer), CType(CType(213, Byte), Integer))
        Me.Com_LGPlan.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Com_LGPlan.Location = New System.Drawing.Point(387, 94)
        Me.Com_LGPlan.Margin = New System.Windows.Forms.Padding(0)
        Me.Com_LGPlan.Name = "Com_LGPlan"
        Me.Com_LGPlan.Size = New System.Drawing.Size(99, 25)
        Me.Com_LGPlan.TabIndex = 39
        Me.Com_LGPlan.Text = "LG 계획"
        Me.Com_LGPlan.UseVisualStyleBackColor = True
        '
        'Di_Panel1
        '
        Me.Di_Panel1.BackColor = System.Drawing.Color.CornflowerBlue
        Me.Di_Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Di_Panel1.Location = New System.Drawing.Point(0, 80)
        Me.Di_Panel1.Margin = New System.Windows.Forms.Padding(0)
        Me.Di_Panel1.Name = "Di_Panel1"
        Me.Di_Panel1.Size = New System.Drawing.Size(1258, 2)
        Me.Di_Panel1.TabIndex = 34
        '
        'Panel_Menu
        '
        Me.Panel_Menu.Controls.Add(Me.Delete_Com)
        Me.Panel_Menu.Controls.Add(Me.Excel_Com)
        Me.Panel_Menu.Controls.Add(Me.Save_Com)
        Me.Panel_Menu.Controls.Add(Me.Load_Com)
        Me.Panel_Menu.Controls.Add(Me.Insert_Com)
        Me.Panel_Menu.Controls.Add(Me.Search_Com)
        Me.Panel_Menu.Controls.Add(Me.Search_Text)
        Me.Panel_Menu.Controls.Add(Me.Search_Combo)
        Me.Panel_Menu.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel_Menu.Location = New System.Drawing.Point(0, 0)
        Me.Panel_Menu.Name = "Panel_Menu"
        Me.Panel_Menu.Size = New System.Drawing.Size(1258, 80)
        Me.Panel_Menu.TabIndex = 33
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
        'Grid_Info_Combo
        '
        Me.Grid_Info_Combo.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.Grid_Info_Combo.FormattingEnabled = True
        Me.Grid_Info_Combo.Location = New System.Drawing.Point(77, 416)
        Me.Grid_Info_Combo.Margin = New System.Windows.Forms.Padding(2)
        Me.Grid_Info_Combo.Name = "Grid_Info_Combo"
        Me.Grid_Info_Combo.Size = New System.Drawing.Size(125, 25)
        Me.Grid_Info_Combo.TabIndex = 38
        Me.Grid_Info_Combo.Visible = False
        '
        'Grid_Info
        '
        Me.Grid_Info.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Grid_Info.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Grid_Info.Location = New System.Drawing.Point(11, 361)
        Me.Grid_Info.Name = "Grid_Info"
        Me.Grid_Info.RowTemplate.Height = 23
        Me.Grid_Info.Size = New System.Drawing.Size(359, 212)
        Me.Grid_Info.TabIndex = 36
        '
        'Delete_Com
        '
        Me.Delete_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Delete_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Delete_Com.Location = New System.Drawing.Point(490, 14)
        Me.Delete_Com.Name = "Delete_Com"
        Me.Delete_Com.Size = New System.Drawing.Size(98, 54)
        Me.Delete_Com.TabIndex = 16
        Me.Delete_Com.Text = "삭제"
        Me.Delete_Com.UseVisualStyleBackColor = False
        '
        'Frm_PlanLG
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.Controls.Add(Me.LGPlan_Panel)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Grid_Info)
        Me.Controls.Add(Me.Grid_Code)
        Me.Controls.Add(Me.Di_Panel2)
        Me.Controls.Add(Me.Com_LGPlan)
        Me.Controls.Add(Me.Di_Panel1)
        Me.Controls.Add(Me.Panel_Menu)
        Me.Controls.Add(Me.Grid_Info_Combo)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "Frm_PlanLG"
        Me.Size = New System.Drawing.Size(1258, 589)
        CType(Me.Grid_LGPlan, System.ComponentModel.ISupportInitialize).EndInit()
        Me.LGPlan_Panel.ResumeLayout(False)
        CType(Me.Grid_Code, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Di_Panel2.ResumeLayout(False)
        Me.Panel_Menu.ResumeLayout(False)
        Me.Panel_Menu.PerformLayout()
        CType(Me.Grid_Info, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents Color_Label As Label
    Friend WithEvents Grid_LGPlan As DataGridView
    Friend WithEvents Search_Text As TextBox
    Friend WithEvents Search_Combo As ComboBox
    Friend WithEvents Save_Com As Button
    Friend WithEvents Load_Com As Button
    Friend WithEvents Insert_Com As Button
    Friend WithEvents LGPlan_Panel As Panel
    Friend WithEvents Button2 As Button
    Friend WithEvents Grid_Code As DataGridView
    Friend WithEvents Excel_Com As Button
    Friend WithEvents Di_Panel2 As Panel
    Friend WithEvents Com_LGPlan As Button
    Friend WithEvents Di_Panel1 As Panel
    Friend WithEvents Panel_Menu As Panel
    Friend WithEvents Search_Com As Button
    Friend WithEvents Grid_Info_Combo As ComboBox
    Friend WithEvents Grid_Info As DataGridView
    Friend WithEvents Delete_Com As Button
End Class