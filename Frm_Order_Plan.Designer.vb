﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_Order_Plan
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
        Me.Grid_Main = New System.Windows.Forms.DataGridView()
        Me.Color_Label = New System.Windows.Forms.Label()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Calculation_Com = New System.Windows.Forms.Button()
        Me.Delete_Com = New System.Windows.Forms.Button()
        Me.Save_Com = New System.Windows.Forms.Button()
        Me.Com_Excel = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Com_Excel_Print = New System.Windows.Forms.Button()
        Me.Insert_Com = New System.Windows.Forms.Button()
        Me.Search_Text = New System.Windows.Forms.TextBox()
        Me.Di_Panel2 = New System.Windows.Forms.Panel()
        Me.Panel_Excel = New System.Windows.Forms.Panel()
        Me.Search_Com = New System.Windows.Forms.Button()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Search_Combo = New System.Windows.Forms.ComboBox()
        Me.Grid_Code = New System.Windows.Forms.DataGridView()
        Me.Grid_Info_Combo = New System.Windows.Forms.ComboBox()
        Me.Com_Main = New System.Windows.Forms.Button()
        Me.Grid_Info = New System.Windows.Forms.DataGridView()
        Me.Panel_Main = New System.Windows.Forms.Panel()
        Me.Di_Panel1 = New System.Windows.Forms.Panel()
        Me.Panel_Menu = New System.Windows.Forms.Panel()
        CType(Me.Grid_Main, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Di_Panel2.SuspendLayout()
        CType(Me.Grid_Code, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Grid_Info, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel_Main.SuspendLayout()
        Me.Panel_Menu.SuspendLayout()
        Me.SuspendLayout()
        '
        'Grid_Main
        '
        Me.Grid_Main.AllowUserToAddRows = False
        Me.Grid_Main.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Grid_Main.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Grid_Main.Location = New System.Drawing.Point(9, 9)
        Me.Grid_Main.Margin = New System.Windows.Forms.Padding(2)
        Me.Grid_Main.Name = "Grid_Main"
        Me.Grid_Main.RowTemplate.Height = 30
        Me.Grid_Main.Size = New System.Drawing.Size(867, 572)
        Me.Grid_Main.TabIndex = 18
        '
        'Color_Label
        '
        Me.Color_Label.Location = New System.Drawing.Point(180, 3)
        Me.Color_Label.Name = "Color_Label"
        Me.Color_Label.Size = New System.Drawing.Size(47, 8)
        Me.Color_Label.TabIndex = 0
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button3.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button3.Location = New System.Drawing.Point(718, 14)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(118, 54)
        Me.Button3.TabIndex = 20
        Me.Button3.Text = "발주서 생성"
        Me.Button3.UseVisualStyleBackColor = False
        '
        'Calculation_Com
        '
        Me.Calculation_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Calculation_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Calculation_Com.Location = New System.Drawing.Point(604, 14)
        Me.Calculation_Com.Name = "Calculation_Com"
        Me.Calculation_Com.Size = New System.Drawing.Size(108, 54)
        Me.Calculation_Com.TabIndex = 17
        Me.Calculation_Com.Text = "산출"
        Me.Calculation_Com.UseVisualStyleBackColor = False
        '
        'Delete_Com
        '
        Me.Delete_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Delete_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Delete_Com.Location = New System.Drawing.Point(490, 14)
        Me.Delete_Com.Name = "Delete_Com"
        Me.Delete_Com.Size = New System.Drawing.Size(108, 54)
        Me.Delete_Com.TabIndex = 16
        Me.Delete_Com.Text = "삭제"
        Me.Delete_Com.UseVisualStyleBackColor = False
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
        'Com_Excel
        '
        Me.Com_Excel.AutoSize = True
        Me.Com_Excel.FlatAppearance.BorderSize = 0
        Me.Com_Excel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange
        Me.Com_Excel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Com_Excel.Location = New System.Drawing.Point(487, 96)
        Me.Com_Excel.Margin = New System.Windows.Forms.Padding(0)
        Me.Com_Excel.Name = "Com_Excel"
        Me.Com_Excel.Size = New System.Drawing.Size(98, 25)
        Me.Com_Excel.TabIndex = 107
        Me.Com_Excel.Text = "산출 Excel"
        Me.Com_Excel.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button2.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.Black
        Me.Button2.Location = New System.Drawing.Point(11, 330)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(359, 27)
        Me.Button2.TabIndex = 102
        Me.Button2.Text = "상세 정보"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Com_Excel_Print
        '
        Me.Com_Excel_Print.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Com_Excel_Print.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Com_Excel_Print.Location = New System.Drawing.Point(842, 14)
        Me.Com_Excel_Print.Name = "Com_Excel_Print"
        Me.Com_Excel_Print.Size = New System.Drawing.Size(108, 54)
        Me.Com_Excel_Print.TabIndex = 13
        Me.Com_Excel_Print.Text = "Excel Printer"
        Me.Com_Excel_Print.UseVisualStyleBackColor = False
        Me.Com_Excel_Print.Visible = False
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
        'Search_Text
        '
        Me.Search_Text.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Search_Text.Location = New System.Drawing.Point(12, 43)
        Me.Search_Text.Name = "Search_Text"
        Me.Search_Text.Size = New System.Drawing.Size(131, 25)
        Me.Search_Text.TabIndex = 10
        '
        'Di_Panel2
        '
        Me.Di_Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Di_Panel2.BackColor = System.Drawing.Color.Silver
        Me.Di_Panel2.Controls.Add(Me.Color_Label)
        Me.Di_Panel2.Location = New System.Drawing.Point(389, 123)
        Me.Di_Panel2.Margin = New System.Windows.Forms.Padding(2)
        Me.Di_Panel2.Name = "Di_Panel2"
        Me.Di_Panel2.Size = New System.Drawing.Size(977, 5)
        Me.Di_Panel2.TabIndex = 105
        '
        'Panel_Excel
        '
        Me.Panel_Excel.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel_Excel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel_Excel.Cursor = System.Windows.Forms.Cursors.Default
        Me.Panel_Excel.Location = New System.Drawing.Point(393, 130)
        Me.Panel_Excel.Name = "Panel_Excel"
        Me.Panel_Excel.Size = New System.Drawing.Size(888, 596)
        Me.Panel_Excel.TabIndex = 108
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
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 12
        Me.ListBox1.Location = New System.Drawing.Point(712, 134)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(321, 592)
        Me.ListBox1.TabIndex = 110
        Me.ListBox1.Visible = False
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
        'Grid_Code
        '
        Me.Grid_Code.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Grid_Code.Location = New System.Drawing.Point(11, 96)
        Me.Grid_Code.Name = "Grid_Code"
        Me.Grid_Code.RowTemplate.Height = 23
        Me.Grid_Code.Size = New System.Drawing.Size(359, 228)
        Me.Grid_Code.TabIndex = 100
        '
        'Grid_Info_Combo
        '
        Me.Grid_Info_Combo.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.Grid_Info_Combo.FormattingEnabled = True
        Me.Grid_Info_Combo.Location = New System.Drawing.Point(79, 418)
        Me.Grid_Info_Combo.Margin = New System.Windows.Forms.Padding(2)
        Me.Grid_Info_Combo.Name = "Grid_Info_Combo"
        Me.Grid_Info_Combo.Size = New System.Drawing.Size(125, 25)
        Me.Grid_Info_Combo.TabIndex = 103
        Me.Grid_Info_Combo.Visible = False
        '
        'Com_Main
        '
        Me.Com_Main.AutoSize = True
        Me.Com_Main.FlatAppearance.BorderSize = 0
        Me.Com_Main.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange
        Me.Com_Main.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Com_Main.Location = New System.Drawing.Point(389, 96)
        Me.Com_Main.Margin = New System.Windows.Forms.Padding(0)
        Me.Com_Main.Name = "Com_Main"
        Me.Com_Main.Size = New System.Drawing.Size(98, 25)
        Me.Com_Main.TabIndex = 104
        Me.Com_Main.Text = "산출 내역"
        Me.Com_Main.UseVisualStyleBackColor = True
        '
        'Grid_Info
        '
        Me.Grid_Info.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Grid_Info.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Grid_Info.Location = New System.Drawing.Point(10, 363)
        Me.Grid_Info.Name = "Grid_Info"
        Me.Grid_Info.RowTemplate.Height = 23
        Me.Grid_Info.Size = New System.Drawing.Size(359, 373)
        Me.Grid_Info.TabIndex = 101
        '
        'Panel_Main
        '
        Me.Panel_Main.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel_Main.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel_Main.Controls.Add(Me.Grid_Main)
        Me.Panel_Main.Location = New System.Drawing.Point(389, 137)
        Me.Panel_Main.Name = "Panel_Main"
        Me.Panel_Main.Size = New System.Drawing.Size(888, 599)
        Me.Panel_Main.TabIndex = 106
        '
        'Di_Panel1
        '
        Me.Di_Panel1.BackColor = System.Drawing.Color.CornflowerBlue
        Me.Di_Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Di_Panel1.Location = New System.Drawing.Point(0, 80)
        Me.Di_Panel1.Margin = New System.Windows.Forms.Padding(0)
        Me.Di_Panel1.Name = "Di_Panel1"
        Me.Di_Panel1.Size = New System.Drawing.Size(1292, 2)
        Me.Di_Panel1.TabIndex = 99
        '
        'Panel_Menu
        '
        Me.Panel_Menu.Controls.Add(Me.Button3)
        Me.Panel_Menu.Controls.Add(Me.Calculation_Com)
        Me.Panel_Menu.Controls.Add(Me.Delete_Com)
        Me.Panel_Menu.Controls.Add(Me.Save_Com)
        Me.Panel_Menu.Controls.Add(Me.Com_Excel_Print)
        Me.Panel_Menu.Controls.Add(Me.Insert_Com)
        Me.Panel_Menu.Controls.Add(Me.Search_Com)
        Me.Panel_Menu.Controls.Add(Me.Search_Text)
        Me.Panel_Menu.Controls.Add(Me.Search_Combo)
        Me.Panel_Menu.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel_Menu.Location = New System.Drawing.Point(0, 0)
        Me.Panel_Menu.Name = "Panel_Menu"
        Me.Panel_Menu.Size = New System.Drawing.Size(1292, 80)
        Me.Panel_Menu.TabIndex = 98
        '
        'Frm_Order_Plan
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Com_Excel)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Di_Panel2)
        Me.Controls.Add(Me.Panel_Excel)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.Grid_Code)
        Me.Controls.Add(Me.Grid_Info_Combo)
        Me.Controls.Add(Me.Com_Main)
        Me.Controls.Add(Me.Grid_Info)
        Me.Controls.Add(Me.Panel_Main)
        Me.Controls.Add(Me.Di_Panel1)
        Me.Controls.Add(Me.Panel_Menu)
        Me.Name = "Frm_Order_Plan"
        Me.Size = New System.Drawing.Size(1292, 753)
        CType(Me.Grid_Main, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Di_Panel2.ResumeLayout(False)
        CType(Me.Grid_Code, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Grid_Info, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel_Main.ResumeLayout(False)
        Me.Panel_Menu.ResumeLayout(False)
        Me.Panel_Menu.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Grid_Main As DataGridView
    Friend WithEvents Color_Label As Label
    Friend WithEvents Button3 As Button
    Friend WithEvents Calculation_Com As Button
    Friend WithEvents Delete_Com As Button
    Friend WithEvents Save_Com As Button
    Friend WithEvents Com_Excel As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Com_Excel_Print As Button
    Friend WithEvents Insert_Com As Button
    Friend WithEvents Search_Text As TextBox
    Friend WithEvents Di_Panel2 As Panel
    Friend WithEvents Panel_Excel As Panel
    Friend WithEvents Search_Com As Button
    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents Search_Combo As ComboBox
    Friend WithEvents Grid_Code As DataGridView
    Friend WithEvents Grid_Info_Combo As ComboBox
    Friend WithEvents Com_Main As Button
    Friend WithEvents Grid_Info As DataGridView
    Friend WithEvents Panel_Main As Panel
    Friend WithEvents Di_Panel1 As Panel
    Friend WithEvents Panel_Menu As Panel
End Class