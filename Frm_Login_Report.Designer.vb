﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_Login_Report
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
        Me.Di_Panel1 = New System.Windows.Forms.Panel()
        Me.Panel_Main = New System.Windows.Forms.Panel()
        Me.Grid_Main = New System.Windows.Forms.DataGridView()
        Me.Color_Label = New System.Windows.Forms.Label()
        Me.Com_Excel = New System.Windows.Forms.Button()
        Me.Panel_Excel = New System.Windows.Forms.Panel()
        Me.Panel_Menu = New System.Windows.Forms.Panel()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Com_Excel_Print = New System.Windows.Forms.Button()
        Me.Search_Com = New System.Windows.Forms.Button()
        Me.Com_Contract = New System.Windows.Forms.Button()
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
        Me.Di_Panel1.Size = New System.Drawing.Size(1012, 2)
        Me.Di_Panel1.TabIndex = 126
        '
        'Panel_Main
        '
        Me.Panel_Main.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel_Main.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel_Main.Controls.Add(Me.Grid_Main)
        Me.Panel_Main.Location = New System.Drawing.Point(12, 123)
        Me.Panel_Main.Name = "Panel_Main"
        Me.Panel_Main.Size = New System.Drawing.Size(981, 471)
        Me.Panel_Main.TabIndex = 123
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
        Me.Grid_Main.Size = New System.Drawing.Size(960, 442)
        Me.Grid_Main.TabIndex = 18
        '
        'Color_Label
        '
        Me.Color_Label.Location = New System.Drawing.Point(180, 3)
        Me.Color_Label.Name = "Color_Label"
        Me.Color_Label.Size = New System.Drawing.Size(47, 8)
        Me.Color_Label.TabIndex = 0
        '
        'Com_Excel
        '
        Me.Com_Excel.AutoSize = True
        Me.Com_Excel.FlatAppearance.BorderSize = 0
        Me.Com_Excel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange
        Me.Com_Excel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Com_Excel.Location = New System.Drawing.Point(110, 82)
        Me.Com_Excel.Margin = New System.Windows.Forms.Padding(0)
        Me.Com_Excel.Name = "Com_Excel"
        Me.Com_Excel.Size = New System.Drawing.Size(105, 25)
        Me.Com_Excel.TabIndex = 124
        Me.Com_Excel.Text = "LOGIN Excel"
        Me.Com_Excel.UseVisualStyleBackColor = True
        '
        'Panel_Excel
        '
        Me.Panel_Excel.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel_Excel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel_Excel.Cursor = System.Windows.Forms.Cursors.Default
        Me.Panel_Excel.Location = New System.Drawing.Point(3, 149)
        Me.Panel_Excel.Name = "Panel_Excel"
        Me.Panel_Excel.Size = New System.Drawing.Size(977, 471)
        Me.Panel_Excel.TabIndex = 125
        '
        'Panel_Menu
        '
        Me.Panel_Menu.Controls.Add(Me.ComboBox2)
        Me.Panel_Menu.Controls.Add(Me.ComboBox1)
        Me.Panel_Menu.Controls.Add(Me.Button2)
        Me.Panel_Menu.Controls.Add(Me.Com_Excel_Print)
        Me.Panel_Menu.Controls.Add(Me.Search_Com)
        Me.Panel_Menu.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel_Menu.Location = New System.Drawing.Point(0, 0)
        Me.Panel_Menu.Name = "Panel_Menu"
        Me.Panel_Menu.Size = New System.Drawing.Size(1012, 80)
        Me.Panel_Menu.TabIndex = 120
        '
        'ComboBox2
        '
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(119, 49)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(233, 20)
        Me.ComboBox2.TabIndex = 97
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(13, 49)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(100, 20)
        Me.ComboBox1.TabIndex = 96
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
        Me.Button2.Text = "직원 선택"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Com_Excel_Print
        '
        Me.Com_Excel_Print.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Com_Excel_Print.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Com_Excel_Print.Location = New System.Drawing.Point(483, 15)
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
        Me.Search_Com.Location = New System.Drawing.Point(369, 15)
        Me.Search_Com.Name = "Search_Com"
        Me.Search_Com.Size = New System.Drawing.Size(108, 54)
        Me.Search_Com.TabIndex = 11
        Me.Search_Com.Text = "검색"
        Me.Search_Com.UseVisualStyleBackColor = False
        '
        'Com_Contract
        '
        Me.Com_Contract.AutoSize = True
        Me.Com_Contract.FlatAppearance.BorderSize = 0
        Me.Com_Contract.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange
        Me.Com_Contract.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Com_Contract.Location = New System.Drawing.Point(12, 82)
        Me.Com_Contract.Margin = New System.Windows.Forms.Padding(0)
        Me.Com_Contract.Name = "Com_Contract"
        Me.Com_Contract.Size = New System.Drawing.Size(98, 25)
        Me.Com_Contract.TabIndex = 121
        Me.Com_Contract.Text = "LOGIN 관리"
        Me.Com_Contract.UseVisualStyleBackColor = True
        '
        'Di_Panel2
        '
        Me.Di_Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Di_Panel2.BackColor = System.Drawing.Color.Silver
        Me.Di_Panel2.Controls.Add(Me.Color_Label)
        Me.Di_Panel2.Location = New System.Drawing.Point(12, 109)
        Me.Di_Panel2.Margin = New System.Windows.Forms.Padding(2)
        Me.Di_Panel2.Name = "Di_Panel2"
        Me.Di_Panel2.Size = New System.Drawing.Size(1491, 5)
        Me.Di_Panel2.TabIndex = 122
        '
        'Frm_Login_Report
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
        Me.Name = "Frm_Login_Report"
        Me.Size = New System.Drawing.Size(1012, 618)
        Me.Panel_Main.ResumeLayout(False)
        CType(Me.Grid_Main, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel_Menu.ResumeLayout(False)
        Me.Di_Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Di_Panel1 As Panel
    Friend WithEvents Panel_Main As Panel
    Friend WithEvents Grid_Main As DataGridView
    Friend WithEvents Color_Label As Label
    Friend WithEvents Com_Excel As Button
    Friend WithEvents Panel_Excel As Panel
    Friend WithEvents Panel_Menu As Panel
    Friend WithEvents ComboBox2 As ComboBox
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents Button2 As Button
    Friend WithEvents Com_Excel_Print As Button
    Friend WithEvents Search_Com As Button
    Friend WithEvents Com_Contract As Button
    Friend WithEvents Di_Panel2 As Panel
End Class
