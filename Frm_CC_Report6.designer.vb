﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Frm_CC_Report6
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
        Dim ChartArea1 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea()
        Dim Legend1 As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend()
        Dim Series1 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Me.Di_Panel2 = New System.Windows.Forms.Panel()
        Me.Color_Label = New System.Windows.Forms.Label()
        Me.Com_Contract = New System.Windows.Forms.Button()
        Me.Search_Combo_Name = New System.Windows.Forms.ComboBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.E_Time = New System.Windows.Forms.TextBox()
        Me.E_Day = New System.Windows.Forms.TextBox()
        Me.S_Time = New System.Windows.Forms.TextBox()
        Me.S_Day = New System.Windows.Forms.TextBox()
        Me.Grid_Main = New System.Windows.Forms.DataGridView()
        Me.Panel_Main = New System.Windows.Forms.Panel()
        Me.Panel_Excel = New System.Windows.Forms.Panel()
        Me.Chart1 = New System.Windows.Forms.DataVisualization.Charting.Chart()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Com_Excel_Print = New System.Windows.Forms.Button()
        Me.Search_Com = New System.Windows.Forms.Button()
        Me.Panel_Menu = New System.Windows.Forms.Panel()
        Me.Search_Combo = New System.Windows.Forms.ComboBox()
        Me.Di_Panel1 = New System.Windows.Forms.Panel()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Di_Panel2.SuspendLayout()
        CType(Me.Grid_Main, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel_Main.SuspendLayout()
        Me.Panel_Excel.SuspendLayout()
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel_Menu.SuspendLayout()
        Me.SuspendLayout()
        '
        'Di_Panel2
        '
        Me.Di_Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Di_Panel2.BackColor = System.Drawing.Color.Silver
        Me.Di_Panel2.Controls.Add(Me.Color_Label)
        Me.Di_Panel2.Location = New System.Drawing.Point(-275, 145)
        Me.Di_Panel2.Margin = New System.Windows.Forms.Padding(2)
        Me.Di_Panel2.Name = "Di_Panel2"
        Me.Di_Panel2.Size = New System.Drawing.Size(2224, 6)
        Me.Di_Panel2.TabIndex = 73
        '
        'Color_Label
        '
        Me.Color_Label.Location = New System.Drawing.Point(206, 4)
        Me.Color_Label.Name = "Color_Label"
        Me.Color_Label.Size = New System.Drawing.Size(54, 10)
        Me.Color_Label.TabIndex = 0
        '
        'Com_Contract
        '
        Me.Com_Contract.AutoSize = True
        Me.Com_Contract.FlatAppearance.BorderSize = 0
        Me.Com_Contract.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange
        Me.Com_Contract.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Com_Contract.Location = New System.Drawing.Point(13, 112)
        Me.Com_Contract.Margin = New System.Windows.Forms.Padding(0)
        Me.Com_Contract.Name = "Com_Contract"
        Me.Com_Contract.Size = New System.Drawing.Size(112, 31)
        Me.Com_Contract.TabIndex = 72
        Me.Com_Contract.Text = "납기일"
        Me.Com_Contract.UseVisualStyleBackColor = True
        '
        'Search_Combo_Name
        '
        Me.Search_Combo_Name.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.Search_Combo_Name.FormattingEnabled = True
        Me.Search_Combo_Name.Location = New System.Drawing.Point(565, 55)
        Me.Search_Combo_Name.Margin = New System.Windows.Forms.Padding(2)
        Me.Search_Combo_Name.Name = "Search_Combo_Name"
        Me.Search_Combo_Name.Size = New System.Drawing.Size(233, 31)
        Me.Search_Combo_Name.TabIndex = 102
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button1.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.Black
        Me.Button1.Location = New System.Drawing.Point(409, 15)
        Me.Button1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(389, 34)
        Me.Button1.TabIndex = 101
        Me.Button1.Text = "업체명"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(200, 62)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(18, 15)
        Me.Label1.TabIndex = 100
        Me.Label1.Text = "~"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'E_Time
        '
        Me.E_Time.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.E_Time.Location = New System.Drawing.Point(341, 55)
        Me.E_Time.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.E_Time.Name = "E_Time"
        Me.E_Time.Size = New System.Drawing.Size(61, 29)
        Me.E_Time.TabIndex = 99
        Me.E_Time.Text = "00:12"
        Me.E_Time.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'E_Day
        '
        Me.E_Day.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.E_Day.Location = New System.Drawing.Point(222, 55)
        Me.E_Day.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.E_Day.Name = "E_Day"
        Me.E_Day.Size = New System.Drawing.Size(111, 29)
        Me.E_Day.TabIndex = 98
        Me.E_Day.Text = "2019-03-01"
        Me.E_Day.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'S_Time
        '
        Me.S_Time.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.S_Time.Location = New System.Drawing.Point(133, 55)
        Me.S_Time.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.S_Time.Name = "S_Time"
        Me.S_Time.Size = New System.Drawing.Size(61, 29)
        Me.S_Time.TabIndex = 97
        Me.S_Time.Text = "00:12"
        Me.S_Time.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'S_Day
        '
        Me.S_Day.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.S_Day.Location = New System.Drawing.Point(14, 55)
        Me.S_Day.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.S_Day.Name = "S_Day"
        Me.S_Day.Size = New System.Drawing.Size(111, 29)
        Me.S_Day.TabIndex = 96
        Me.S_Day.Text = "2019-03-01"
        Me.S_Day.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Grid_Main
        '
        Me.Grid_Main.AllowUserToAddRows = False
        Me.Grid_Main.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Grid_Main.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Grid_Main.Location = New System.Drawing.Point(8, 14)
        Me.Grid_Main.Margin = New System.Windows.Forms.Padding(2)
        Me.Grid_Main.Name = "Grid_Main"
        Me.Grid_Main.RowTemplate.Height = 30
        Me.Grid_Main.Size = New System.Drawing.Size(1311, 631)
        Me.Grid_Main.TabIndex = 18
        '
        'Panel_Main
        '
        Me.Panel_Main.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel_Main.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel_Main.Controls.Add(Me.Grid_Main)
        Me.Panel_Main.Location = New System.Drawing.Point(14, 294)
        Me.Panel_Main.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Panel_Main.Name = "Panel_Main"
        Me.Panel_Main.Size = New System.Drawing.Size(1335, 667)
        Me.Panel_Main.TabIndex = 74
        '
        'Panel_Excel
        '
        Me.Panel_Excel.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel_Excel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel_Excel.Controls.Add(Me.Chart1)
        Me.Panel_Excel.Cursor = System.Windows.Forms.Cursors.Default
        Me.Panel_Excel.Location = New System.Drawing.Point(63, 174)
        Me.Panel_Excel.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Panel_Excel.Name = "Panel_Excel"
        Me.Panel_Excel.Size = New System.Drawing.Size(1330, 667)
        Me.Panel_Excel.TabIndex = 76
        '
        'Chart1
        '
        ChartArea1.Name = "ChartArea1"
        Me.Chart1.ChartAreas.Add(ChartArea1)
        Legend1.Name = "Legend1"
        Me.Chart1.Legends.Add(Legend1)
        Me.Chart1.Location = New System.Drawing.Point(35, 25)
        Me.Chart1.Name = "Chart1"
        Series1.ChartArea = "ChartArea1"
        Series1.Legend = "Legend1"
        Series1.Name = "납기일"
        Me.Chart1.Series.Add(Series1)
        Me.Chart1.Size = New System.Drawing.Size(1471, 691)
        Me.Chart1.TabIndex = 75
        Me.Chart1.Text = "Chart1"
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button2.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.Black
        Me.Button2.Location = New System.Drawing.Point(14, 14)
        Me.Button2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(389, 34)
        Me.Button2.TabIndex = 95
        Me.Button2.Text = "검색기간"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Com_Excel_Print
        '
        Me.Com_Excel_Print.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Com_Excel_Print.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Com_Excel_Print.Location = New System.Drawing.Point(933, 18)
        Me.Com_Excel_Print.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Com_Excel_Print.Name = "Com_Excel_Print"
        Me.Com_Excel_Print.Size = New System.Drawing.Size(123, 68)
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
        Me.Search_Com.Location = New System.Drawing.Point(804, 18)
        Me.Search_Com.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Search_Com.Name = "Search_Com"
        Me.Search_Com.Size = New System.Drawing.Size(123, 68)
        Me.Search_Com.TabIndex = 11
        Me.Search_Com.Text = "검색"
        Me.Search_Com.UseVisualStyleBackColor = False
        '
        'Panel_Menu
        '
        Me.Panel_Menu.Controls.Add(Me.Search_Combo_Name)
        Me.Panel_Menu.Controls.Add(Me.Button1)
        Me.Panel_Menu.Controls.Add(Me.Label1)
        Me.Panel_Menu.Controls.Add(Me.E_Time)
        Me.Panel_Menu.Controls.Add(Me.E_Day)
        Me.Panel_Menu.Controls.Add(Me.S_Time)
        Me.Panel_Menu.Controls.Add(Me.S_Day)
        Me.Panel_Menu.Controls.Add(Me.Button2)
        Me.Panel_Menu.Controls.Add(Me.Com_Excel_Print)
        Me.Panel_Menu.Controls.Add(Me.Search_Com)
        Me.Panel_Menu.Controls.Add(Me.Search_Combo)
        Me.Panel_Menu.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel_Menu.Location = New System.Drawing.Point(0, 2)
        Me.Panel_Menu.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Panel_Menu.Name = "Panel_Menu"
        Me.Panel_Menu.Size = New System.Drawing.Size(1368, 100)
        Me.Panel_Menu.TabIndex = 70
        '
        'Search_Combo
        '
        Me.Search_Combo.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.Search_Combo.FormattingEnabled = True
        Me.Search_Combo.Location = New System.Drawing.Point(409, 55)
        Me.Search_Combo.Margin = New System.Windows.Forms.Padding(2)
        Me.Search_Combo.Name = "Search_Combo"
        Me.Search_Combo.Size = New System.Drawing.Size(150, 31)
        Me.Search_Combo.TabIndex = 8
        '
        'Di_Panel1
        '
        Me.Di_Panel1.BackColor = System.Drawing.Color.CornflowerBlue
        Me.Di_Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Di_Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Di_Panel1.Margin = New System.Windows.Forms.Padding(0)
        Me.Di_Panel1.Name = "Di_Panel1"
        Me.Di_Panel1.Size = New System.Drawing.Size(1368, 2)
        Me.Di_Panel1.TabIndex = 71
        '
        'Button3
        '
        Me.Button3.AutoSize = True
        Me.Button3.FlatAppearance.BorderSize = 0
        Me.Button3.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button3.Location = New System.Drawing.Point(117, 112)
        Me.Button3.Margin = New System.Windows.Forms.Padding(0)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(112, 31)
        Me.Button3.TabIndex = 77
        Me.Button3.Text = "납기일 그래프"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Frm_CC_Report6
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Panel_Main)
        Me.Controls.Add(Me.Di_Panel2)
        Me.Controls.Add(Me.Com_Contract)
        Me.Controls.Add(Me.Panel_Excel)
        Me.Controls.Add(Me.Panel_Menu)
        Me.Controls.Add(Me.Di_Panel1)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "Frm_CC_Report6"
        Me.Size = New System.Drawing.Size(1368, 856)
        Me.Di_Panel2.ResumeLayout(False)
        CType(Me.Grid_Main, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel_Main.ResumeLayout(False)
        Me.Panel_Excel.ResumeLayout(False)
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel_Menu.ResumeLayout(False)
        Me.Panel_Menu.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Di_Panel2 As Panel
    Friend WithEvents Color_Label As Label
    Friend WithEvents Com_Contract As Button
    Friend WithEvents Search_Combo_Name As ComboBox
    Friend WithEvents Button1 As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents E_Time As TextBox
    Friend WithEvents E_Day As TextBox
    Friend WithEvents S_Time As TextBox
    Friend WithEvents S_Day As TextBox
    Friend WithEvents Grid_Main As DataGridView
    Friend WithEvents Panel_Main As Panel
    Friend WithEvents Panel_Excel As Panel
    Friend WithEvents Button2 As Button
    Friend WithEvents Com_Excel_Print As Button
    Friend WithEvents Search_Com As Button
    Friend WithEvents Panel_Menu As Panel
    Friend WithEvents Search_Combo As ComboBox
    Friend WithEvents Di_Panel1 As Panel
    Friend WithEvents Button3 As Button
    Friend WithEvents Chart1 As DataVisualization.Charting.Chart
End Class
