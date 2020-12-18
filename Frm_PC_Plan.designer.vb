<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Frm_PC_Plan
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
        Me.Main_Table = New System.Windows.Forms.TableLayoutPanel()
        Me.Grid_Order = New System.Windows.Forms.DataGridView()
        Me.Grid_Main = New System.Windows.Forms.DataGridView()
        Me.Title1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Search_Combo = New System.Windows.Forms.ComboBox()
        Me.Search_Text = New System.Windows.Forms.TextBox()
        Me.Search_Com = New System.Windows.Forms.Button()
        Me.Insert_Com = New System.Windows.Forms.Button()
        Me.Com_Excel_Print = New System.Windows.Forms.Button()
        Me.Save_Com = New System.Windows.Forms.Button()
        Me.Delete_Com = New System.Windows.Forms.Button()
        Me.Calculation_Com = New System.Windows.Forms.Button()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.Btn_Produc = New System.Windows.Forms.Button()
        Me.Panel_Menu = New System.Windows.Forms.Panel()
        Me.Di_Panel1 = New System.Windows.Forms.Panel()
        Me.Main_Table.SuspendLayout()
        CType(Me.Grid_Order, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Grid_Main, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel_Menu.SuspendLayout()
        Me.SuspendLayout()
        '
        'Main_Table
        '
        Me.Main_Table.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.Main_Table.ColumnCount = 2
        Me.Main_Table.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.Main_Table.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.Main_Table.Controls.Add(Me.Grid_Order, 0, 1)
        Me.Main_Table.Controls.Add(Me.Grid_Main, 0, 3)
        Me.Main_Table.Controls.Add(Me.Title1, 0, 0)
        Me.Main_Table.Controls.Add(Me.Button2, 0, 2)
        Me.Main_Table.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Main_Table.Location = New System.Drawing.Point(0, 80)
        Me.Main_Table.Name = "Main_Table"
        Me.Main_Table.RowCount = 4
        Me.Main_Table.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 3.0!))
        Me.Main_Table.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 47.0!))
        Me.Main_Table.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 3.0!))
        Me.Main_Table.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 47.0!))
        Me.Main_Table.Size = New System.Drawing.Size(1442, 595)
        Me.Main_Table.TabIndex = 103
        '
        'Grid_Order
        '
        Me.Grid_Order.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.Grid_Order.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Grid_Order.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Grid_Order.Location = New System.Drawing.Point(3, 20)
        Me.Grid_Order.Name = "Grid_Order"
        Me.Grid_Order.ReadOnly = True
        Me.Grid_Order.RowTemplate.Height = 23
        Me.Grid_Order.Size = New System.Drawing.Size(1436, 273)
        Me.Grid_Order.TabIndex = 101
        '
        'Grid_Main
        '
        Me.Grid_Main.AllowUserToAddRows = False
        Me.Grid_Main.BackgroundColor = System.Drawing.SystemColors.Control
        Me.Grid_Main.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Grid_Main.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Grid_Main.Location = New System.Drawing.Point(2, 315)
        Me.Grid_Main.Margin = New System.Windows.Forms.Padding(2)
        Me.Grid_Main.Name = "Grid_Main"
        Me.Grid_Main.RowTemplate.Height = 30
        Me.Grid_Main.Size = New System.Drawing.Size(1438, 278)
        Me.Grid_Main.TabIndex = 99
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
        Me.Title1.Size = New System.Drawing.Size(65, 17)
        Me.Title1.TabIndex = 96
        Me.Title1.Text = "수주현황"
        Me.Title1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.AutoSize = True
        Me.Button2.FlatAppearance.BorderSize = 0
        Me.Button2.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Font = New System.Drawing.Font("나눔고딕", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button2.Location = New System.Drawing.Point(0, 296)
        Me.Button2.Margin = New System.Windows.Forms.Padding(0)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(109, 17)
        Me.Button2.TabIndex = 103
        Me.Button2.Text = "개별 소모량 산출"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Search_Combo
        '
        Me.Search_Combo.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.Search_Combo.FormattingEnabled = True
        Me.Search_Combo.Location = New System.Drawing.Point(15, 12)
        Me.Search_Combo.Margin = New System.Windows.Forms.Padding(2)
        Me.Search_Combo.Name = "Search_Combo"
        Me.Search_Combo.Size = New System.Drawing.Size(132, 25)
        Me.Search_Combo.TabIndex = 104
        '
        'Search_Text
        '
        Me.Search_Text.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Search_Text.Location = New System.Drawing.Point(16, 41)
        Me.Search_Text.Name = "Search_Text"
        Me.Search_Text.Size = New System.Drawing.Size(131, 25)
        Me.Search_Text.TabIndex = 105
        '
        'Search_Com
        '
        Me.Search_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Search_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Search_Com.ForeColor = System.Drawing.Color.Black
        Me.Search_Com.Location = New System.Drawing.Point(213, 12)
        Me.Search_Com.Name = "Search_Com"
        Me.Search_Com.Size = New System.Drawing.Size(108, 54)
        Me.Search_Com.TabIndex = 106
        Me.Search_Com.Text = "검색"
        Me.Search_Com.UseVisualStyleBackColor = False
        '
        'Insert_Com
        '
        Me.Insert_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Insert_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Insert_Com.Location = New System.Drawing.Point(1163, 12)
        Me.Insert_Com.Name = "Insert_Com"
        Me.Insert_Com.Size = New System.Drawing.Size(108, 54)
        Me.Insert_Com.TabIndex = 107
        Me.Insert_Com.Text = "추가"
        Me.Insert_Com.UseVisualStyleBackColor = False
        Me.Insert_Com.Visible = False
        '
        'Com_Excel_Print
        '
        Me.Com_Excel_Print.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Com_Excel_Print.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Com_Excel_Print.Location = New System.Drawing.Point(671, 12)
        Me.Com_Excel_Print.Name = "Com_Excel_Print"
        Me.Com_Excel_Print.Size = New System.Drawing.Size(108, 54)
        Me.Com_Excel_Print.TabIndex = 108
        Me.Com_Excel_Print.Text = "Excel Printer"
        Me.Com_Excel_Print.UseVisualStyleBackColor = False
        '
        'Save_Com
        '
        Me.Save_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Save_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Save_Com.ForeColor = System.Drawing.Color.Black
        Me.Save_Com.Location = New System.Drawing.Point(1049, 12)
        Me.Save_Com.Name = "Save_Com"
        Me.Save_Com.Size = New System.Drawing.Size(108, 54)
        Me.Save_Com.TabIndex = 109
        Me.Save_Com.Text = "저장"
        Me.Save_Com.UseVisualStyleBackColor = False
        Me.Save_Com.Visible = False
        '
        'Delete_Com
        '
        Me.Delete_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Delete_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Delete_Com.Location = New System.Drawing.Point(1277, 12)
        Me.Delete_Com.Name = "Delete_Com"
        Me.Delete_Com.Size = New System.Drawing.Size(108, 54)
        Me.Delete_Com.TabIndex = 110
        Me.Delete_Com.Text = "삭제"
        Me.Delete_Com.UseVisualStyleBackColor = False
        Me.Delete_Com.Visible = False
        '
        'Calculation_Com
        '
        Me.Calculation_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Calculation_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Calculation_Com.Location = New System.Drawing.Point(433, 12)
        Me.Calculation_Com.Name = "Calculation_Com"
        Me.Calculation_Com.Size = New System.Drawing.Size(108, 54)
        Me.Calculation_Com.TabIndex = 111
        Me.Calculation_Com.Text = "산출"
        Me.Calculation_Com.UseVisualStyleBackColor = False
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Location = New System.Drawing.Point(344, 16)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(82, 16)
        Me.RadioButton1.TabIndex = 112
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Order 산출"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(344, 45)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(83, 16)
        Me.RadioButton2.TabIndex = 113
        Me.RadioButton2.Text = "검증   산출"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'Btn_Produc
        '
        Me.Btn_Produc.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Btn_Produc.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Btn_Produc.Location = New System.Drawing.Point(547, 12)
        Me.Btn_Produc.Name = "Btn_Produc"
        Me.Btn_Produc.Size = New System.Drawing.Size(118, 54)
        Me.Btn_Produc.TabIndex = 114
        Me.Btn_Produc.Text = "작업지시서 생성"
        Me.Btn_Produc.UseVisualStyleBackColor = False
        '
        'Panel_Menu
        '
        Me.Panel_Menu.Controls.Add(Me.Di_Panel1)
        Me.Panel_Menu.Controls.Add(Me.Btn_Produc)
        Me.Panel_Menu.Controls.Add(Me.Search_Text)
        Me.Panel_Menu.Controls.Add(Me.RadioButton2)
        Me.Panel_Menu.Controls.Add(Me.Search_Combo)
        Me.Panel_Menu.Controls.Add(Me.RadioButton1)
        Me.Panel_Menu.Controls.Add(Me.Search_Com)
        Me.Panel_Menu.Controls.Add(Me.Calculation_Com)
        Me.Panel_Menu.Controls.Add(Me.Insert_Com)
        Me.Panel_Menu.Controls.Add(Me.Delete_Com)
        Me.Panel_Menu.Controls.Add(Me.Com_Excel_Print)
        Me.Panel_Menu.Controls.Add(Me.Save_Com)
        Me.Panel_Menu.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel_Menu.Location = New System.Drawing.Point(0, 0)
        Me.Panel_Menu.Name = "Panel_Menu"
        Me.Panel_Menu.Size = New System.Drawing.Size(1442, 80)
        Me.Panel_Menu.TabIndex = 115
        '
        'Di_Panel1
        '
        Me.Di_Panel1.BackColor = System.Drawing.Color.CornflowerBlue
        Me.Di_Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Di_Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Di_Panel1.Margin = New System.Windows.Forms.Padding(0)
        Me.Di_Panel1.Name = "Di_Panel1"
        Me.Di_Panel1.Size = New System.Drawing.Size(1442, 2)
        Me.Di_Panel1.TabIndex = 115
        '
        'Frm_PC_Plan
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Main_Table)
        Me.Controls.Add(Me.Panel_Menu)
        Me.Name = "Frm_PC_Plan"
        Me.Size = New System.Drawing.Size(1442, 675)
        Me.Main_Table.ResumeLayout(False)
        Me.Main_Table.PerformLayout()
        CType(Me.Grid_Order, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Grid_Main, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel_Menu.ResumeLayout(False)
        Me.Panel_Menu.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Main_Table As TableLayoutPanel
    Friend WithEvents Grid_Order As DataGridView
    Friend WithEvents Grid_Main As DataGridView
    Friend WithEvents Search_Combo As ComboBox
    Friend WithEvents Search_Text As TextBox
    Friend WithEvents Search_Com As Button
    Friend WithEvents Insert_Com As Button
    Friend WithEvents Com_Excel_Print As Button
    Friend WithEvents Save_Com As Button
    Friend WithEvents Delete_Com As Button
    Friend WithEvents Calculation_Com As Button
    Friend WithEvents RadioButton1 As RadioButton
    Friend WithEvents RadioButton2 As RadioButton
    Friend WithEvents Btn_Produc As Button
    Friend WithEvents Panel_Menu As Panel
    Friend WithEvents Title1 As Button
    Friend WithEvents Di_Panel1 As Panel
    Friend WithEvents Button2 As Button
End Class
