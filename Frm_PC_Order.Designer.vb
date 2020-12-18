<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_PC_Order
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
        Me.Insert_Com = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Com_Excel_Print = New System.Windows.Forms.Button()
        Me.Delete_Com = New System.Windows.Forms.Button()
        Me.Save_Com = New System.Windows.Forms.Button()
        Me.Panel_Menu = New System.Windows.Forms.Panel()
        Me.Search_Com = New System.Windows.Forms.Button()
        Me.Search_Text = New System.Windows.Forms.TextBox()
        Me.Search_Combo = New System.Windows.Forms.ComboBox()
        Me.Di_Panel1 = New System.Windows.Forms.Panel()
        Me.Main_Table = New System.Windows.Forms.TableLayoutPanel()
        Me.Log = New System.Windows.Forms.DataGridView()
        Me.Grid_Order = New System.Windows.Forms.DataGridView()
        Me.Title1 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Panel_Menu.SuspendLayout()
        Me.Main_Table.SuspendLayout()
        CType(Me.Log, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Grid_Order, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Insert_Com
        '
        Me.Insert_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Insert_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Insert_Com.Location = New System.Drawing.Point(320, 14)
        Me.Insert_Com.Name = "Insert_Com"
        Me.Insert_Com.Size = New System.Drawing.Size(150, 54)
        Me.Insert_Com.TabIndex = 12
        Me.Insert_Com.Text = "작업지시서 추가"
        Me.Insert_Com.UseVisualStyleBackColor = False
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Com_Excel_Print
        '
        Me.Com_Excel_Print.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Com_Excel_Print.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Com_Excel_Print.Location = New System.Drawing.Point(856, 14)
        Me.Com_Excel_Print.Name = "Com_Excel_Print"
        Me.Com_Excel_Print.Size = New System.Drawing.Size(108, 54)
        Me.Com_Excel_Print.TabIndex = 13
        Me.Com_Excel_Print.Text = "Excel Printer"
        Me.Com_Excel_Print.UseVisualStyleBackColor = False
        Me.Com_Excel_Print.Visible = False
        '
        'Delete_Com
        '
        Me.Delete_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Delete_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Delete_Com.Location = New System.Drawing.Point(490, 14)
        Me.Delete_Com.Name = "Delete_Com"
        Me.Delete_Com.Size = New System.Drawing.Size(150, 54)
        Me.Delete_Com.TabIndex = 16
        Me.Delete_Com.Text = "작업지시서 삭제"
        Me.Delete_Com.UseVisualStyleBackColor = False
        '
        'Save_Com
        '
        Me.Save_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Save_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Save_Com.ForeColor = System.Drawing.Color.Black
        Me.Save_Com.Location = New System.Drawing.Point(1009, 14)
        Me.Save_Com.Name = "Save_Com"
        Me.Save_Com.Size = New System.Drawing.Size(108, 54)
        Me.Save_Com.TabIndex = 14
        Me.Save_Com.Text = "저장"
        Me.Save_Com.UseVisualStyleBackColor = False
        '
        'Panel_Menu
        '
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
        Me.Panel_Menu.Size = New System.Drawing.Size(1273, 80)
        Me.Panel_Menu.TabIndex = 76
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
        'Di_Panel1
        '
        Me.Di_Panel1.BackColor = System.Drawing.Color.CornflowerBlue
        Me.Di_Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Di_Panel1.Location = New System.Drawing.Point(0, 80)
        Me.Di_Panel1.Margin = New System.Windows.Forms.Padding(0)
        Me.Di_Panel1.Name = "Di_Panel1"
        Me.Di_Panel1.Size = New System.Drawing.Size(1273, 2)
        Me.Di_Panel1.TabIndex = 77
        '
        'Main_Table
        '
        Me.Main_Table.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.Main_Table.ColumnCount = 2
        Me.Main_Table.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 90.0!))
        Me.Main_Table.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.Main_Table.Controls.Add(Me.Log, 0, 3)
        Me.Main_Table.Controls.Add(Me.Grid_Order, 0, 1)
        Me.Main_Table.Controls.Add(Me.Title1, 0, 0)
        Me.Main_Table.Controls.Add(Me.Button1, 0, 2)
        Me.Main_Table.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Main_Table.Location = New System.Drawing.Point(0, 82)
        Me.Main_Table.Name = "Main_Table"
        Me.Main_Table.RowCount = 4
        Me.Main_Table.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 3.0!))
        Me.Main_Table.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 64.0!))
        Me.Main_Table.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 3.0!))
        Me.Main_Table.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 30.0!))
        Me.Main_Table.Size = New System.Drawing.Size(1273, 760)
        Me.Main_Table.TabIndex = 104
        '
        'Log
        '
        Me.Log.AllowUserToAddRows = False
        Me.Log.BackgroundColor = System.Drawing.SystemColors.Control
        Me.Log.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Log.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Log.Location = New System.Drawing.Point(2, 532)
        Me.Log.Margin = New System.Windows.Forms.Padding(2)
        Me.Log.Name = "Log"
        Me.Log.RowTemplate.Height = 30
        Me.Log.Size = New System.Drawing.Size(1141, 226)
        Me.Log.TabIndex = 106
        '
        'Grid_Order
        '
        Me.Grid_Order.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.Grid_Order.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Grid_Order.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Grid_Order.Location = New System.Drawing.Point(3, 25)
        Me.Grid_Order.Name = "Grid_Order"
        Me.Grid_Order.ReadOnly = True
        Me.Grid_Order.RowTemplate.Height = 23
        Me.Grid_Order.Size = New System.Drawing.Size(1139, 480)
        Me.Grid_Order.TabIndex = 101
        '
        'Title1
        '
        Me.Title1.AutoSize = True
        Me.Title1.FlatAppearance.BorderSize = 0
        Me.Title1.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange
        Me.Title1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Title1.Font = New System.Drawing.Font("나눔고딕", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Title1.Location = New System.Drawing.Point(0, 0)
        Me.Title1.Margin = New System.Windows.Forms.Padding(0)
        Me.Title1.Name = "Title1"
        Me.Title1.Size = New System.Drawing.Size(105, 22)
        Me.Title1.TabIndex = 96
        Me.Title1.Text = "작업지시서 현황"
        Me.Title1.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.AutoSize = True
        Me.Button1.FlatAppearance.BorderSize = 0
        Me.Button1.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("나눔고딕", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button1.Location = New System.Drawing.Point(0, 508)
        Me.Button1.Margin = New System.Windows.Forms.Padding(0)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(133, 22)
        Me.Button1.TabIndex = 102
        Me.Button1.Text = "작업지시서 생성 내역"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Frm_PC_Order
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Main_Table)
        Me.Controls.Add(Me.Di_Panel1)
        Me.Controls.Add(Me.Panel_Menu)
        Me.Name = "Frm_PC_Order"
        Me.Size = New System.Drawing.Size(1273, 842)
        Me.Panel_Menu.ResumeLayout(False)
        Me.Panel_Menu.PerformLayout()
        Me.Main_Table.ResumeLayout(False)
        Me.Main_Table.PerformLayout()
        CType(Me.Log, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Grid_Order, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Insert_Com As Button
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents Com_Excel_Print As Button
    Friend WithEvents Delete_Com As Button
    Friend WithEvents Save_Com As Button
    Friend WithEvents Panel_Menu As Panel
    Friend WithEvents Search_Com As Button
    Friend WithEvents Search_Text As TextBox
    Friend WithEvents Search_Combo As ComboBox
    Friend WithEvents Di_Panel1 As Panel
    Friend WithEvents Main_Table As TableLayoutPanel
    Friend WithEvents Log As DataGridView
    Friend WithEvents Grid_Order As DataGridView
    Friend WithEvents Title1 As Button
    Friend WithEvents Button1 As Button
End Class
