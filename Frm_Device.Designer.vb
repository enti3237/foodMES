<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_Device
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
        Me.Save_Com = New System.Windows.Forms.Button()
        Me.Delete_Com = New System.Windows.Forms.Button()
        Me.Insert_Com = New System.Windows.Forms.Button()
        Me.Search_Com = New System.Windows.Forms.Button()
        Me.Search_Text = New System.Windows.Forms.TextBox()
        Me.Di_Panel1 = New System.Windows.Forms.Panel()
        Me.Panel_Menu = New System.Windows.Forms.Panel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Process_Line_Grid = New System.Windows.Forms.DataGridView()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Delete__Bom = New System.Windows.Forms.Button()
        Me.Insert__Bom = New System.Windows.Forms.Button()
        Me.Grid_Bom_Combo = New System.Windows.Forms.ComboBox()
        Me.Process_Line_QC = New System.Windows.Forms.DataGridView()
        Me.Com_BOM = New System.Windows.Forms.Button()
        Me.Panel_Menu.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.Process_Line_Grid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        CType(Me.Process_Line_QC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Save_Com
        '
        Me.Save_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Save_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Save_Com.ForeColor = System.Drawing.Color.Black
        Me.Save_Com.Location = New System.Drawing.Point(560, 7)
        Me.Save_Com.Name = "Save_Com"
        Me.Save_Com.Size = New System.Drawing.Size(99, 54)
        Me.Save_Com.TabIndex = 14
        Me.Save_Com.Text = "저장"
        Me.Save_Com.UseVisualStyleBackColor = False
        Me.Save_Com.Visible = False
        '
        'Delete_Com
        '
        Me.Delete_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Delete_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Delete_Com.Location = New System.Drawing.Point(351, 3)
        Me.Delete_Com.Name = "Delete_Com"
        Me.Delete_Com.Size = New System.Drawing.Size(98, 54)
        Me.Delete_Com.TabIndex = 13
        Me.Delete_Com.Text = "삭제"
        Me.Delete_Com.UseVisualStyleBackColor = False
        '
        'Insert_Com
        '
        Me.Insert_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Insert_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Insert_Com.Location = New System.Drawing.Point(247, 5)
        Me.Insert_Com.Name = "Insert_Com"
        Me.Insert_Com.Size = New System.Drawing.Size(98, 54)
        Me.Insert_Com.TabIndex = 12
        Me.Insert_Com.Text = "추가"
        Me.Insert_Com.UseVisualStyleBackColor = False
        '
        'Search_Com
        '
        Me.Search_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Search_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Search_Com.ForeColor = System.Drawing.Color.Black
        Me.Search_Com.Location = New System.Drawing.Point(142, 5)
        Me.Search_Com.Name = "Search_Com"
        Me.Search_Com.Size = New System.Drawing.Size(99, 54)
        Me.Search_Com.TabIndex = 11
        Me.Search_Com.Text = "검색"
        Me.Search_Com.UseVisualStyleBackColor = False
        '
        'Search_Text
        '
        Me.Search_Text.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Search_Text.Location = New System.Drawing.Point(12, 34)
        Me.Search_Text.Name = "Search_Text"
        Me.Search_Text.Size = New System.Drawing.Size(124, 25)
        Me.Search_Text.TabIndex = 10
        '
        'Di_Panel1
        '
        Me.Di_Panel1.BackColor = System.Drawing.Color.CornflowerBlue
        Me.Di_Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Di_Panel1.Location = New System.Drawing.Point(0, 65)
        Me.Di_Panel1.Margin = New System.Windows.Forms.Padding(0)
        Me.Di_Panel1.Name = "Di_Panel1"
        Me.Di_Panel1.Size = New System.Drawing.Size(1509, 2)
        Me.Di_Panel1.TabIndex = 37
        '
        'Panel_Menu
        '
        Me.Panel_Menu.Controls.Add(Me.Button1)
        Me.Panel_Menu.Controls.Add(Me.Save_Com)
        Me.Panel_Menu.Controls.Add(Me.Delete_Com)
        Me.Panel_Menu.Controls.Add(Me.Insert_Com)
        Me.Panel_Menu.Controls.Add(Me.Search_Com)
        Me.Panel_Menu.Controls.Add(Me.Search_Text)
        Me.Panel_Menu.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel_Menu.Location = New System.Drawing.Point(0, 0)
        Me.Panel_Menu.Name = "Panel_Menu"
        Me.Panel_Menu.Size = New System.Drawing.Size(1509, 65)
        Me.Panel_Menu.TabIndex = 36
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button1.Enabled = False
        Me.Button1.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.Black
        Me.Button1.Location = New System.Drawing.Point(12, 7)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(123, 26)
        Me.Button1.TabIndex = 26
        Me.Button1.Text = "품명"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 70.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 30.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.Button2, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Process_Line_Grid, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel3, 1, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Grid_Bom_Combo, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Process_Line_QC, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Com_BOM, 0, 2)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 67)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 4
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 4.019257!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 45.68621!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 4.045454!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 46.24907!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(1509, 621)
        Me.TableLayoutPanel1.TabIndex = 64
        '
        'Button2
        '
        Me.Button2.AutoSize = True
        Me.Button2.FlatAppearance.BorderSize = 0
        Me.Button2.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Font = New System.Drawing.Font("굴림", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button2.Location = New System.Drawing.Point(0, 0)
        Me.Button2.Margin = New System.Windows.Forms.Padding(0)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(95, 23)
        Me.Button2.TabIndex = 60
        Me.Button2.Text = "설비"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Process_Line_Grid
        '
        Me.Process_Line_Grid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.Process_Line_Grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Process_Line_Grid.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Process_Line_Grid.Location = New System.Drawing.Point(3, 27)
        Me.Process_Line_Grid.Name = "Process_Line_Grid"
        Me.Process_Line_Grid.RowTemplate.Height = 23
        Me.Process_Line_Grid.Size = New System.Drawing.Size(1050, 277)
        Me.Process_Line_Grid.TabIndex = 52
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.FromArgb(CType(CType(91, Byte), Integer), CType(CType(155, Byte), Integer), CType(CType(213, Byte), Integer))
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.Delete__Bom)
        Me.Panel3.Controls.Add(Me.Insert__Bom)
        Me.Panel3.Location = New System.Drawing.Point(1059, 335)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(124, 132)
        Me.Panel3.TabIndex = 62
        '
        'Delete__Bom
        '
        Me.Delete__Bom.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Delete__Bom.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Delete__Bom.Location = New System.Drawing.Point(3, 63)
        Me.Delete__Bom.Name = "Delete__Bom"
        Me.Delete__Bom.Size = New System.Drawing.Size(114, 54)
        Me.Delete__Bom.TabIndex = 7
        Me.Delete__Bom.Text = "삭제"
        Me.Delete__Bom.UseVisualStyleBackColor = False
        '
        'Insert__Bom
        '
        Me.Insert__Bom.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Insert__Bom.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Insert__Bom.Location = New System.Drawing.Point(3, 3)
        Me.Insert__Bom.Name = "Insert__Bom"
        Me.Insert__Bom.Size = New System.Drawing.Size(114, 54)
        Me.Insert__Bom.TabIndex = 5
        Me.Insert__Bom.Text = "추가"
        Me.Insert__Bom.UseVisualStyleBackColor = False
        '
        'Grid_Bom_Combo
        '
        Me.Grid_Bom_Combo.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.Grid_Bom_Combo.FormattingEnabled = True
        Me.Grid_Bom_Combo.Location = New System.Drawing.Point(1058, 309)
        Me.Grid_Bom_Combo.Margin = New System.Windows.Forms.Padding(2)
        Me.Grid_Bom_Combo.Name = "Grid_Bom_Combo"
        Me.Grid_Bom_Combo.Size = New System.Drawing.Size(125, 25)
        Me.Grid_Bom_Combo.TabIndex = 21
        Me.Grid_Bom_Combo.Visible = False
        '
        'Process_Line_QC
        '
        Me.Process_Line_QC.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.Process_Line_QC.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Process_Line_QC.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Process_Line_QC.Location = New System.Drawing.Point(3, 335)
        Me.Process_Line_QC.Name = "Process_Line_QC"
        Me.Process_Line_QC.RowTemplate.Height = 23
        Me.Process_Line_QC.Size = New System.Drawing.Size(1050, 283)
        Me.Process_Line_QC.TabIndex = 34
        '
        'Com_BOM
        '
        Me.Com_BOM.AutoSize = True
        Me.Com_BOM.FlatAppearance.BorderSize = 0
        Me.Com_BOM.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange
        Me.Com_BOM.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Com_BOM.Font = New System.Drawing.Font("굴림", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Com_BOM.Location = New System.Drawing.Point(0, 307)
        Me.Com_BOM.Margin = New System.Windows.Forms.Padding(0)
        Me.Com_BOM.Name = "Com_BOM"
        Me.Com_BOM.Size = New System.Drawing.Size(95, 23)
        Me.Com_BOM.TabIndex = 46
        Me.Com_BOM.Text = "점검내역"
        Me.Com_BOM.UseVisualStyleBackColor = True
        '
        'Frm_Device
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.Di_Panel1)
        Me.Controls.Add(Me.Panel_Menu)
        Me.Name = "Frm_Device"
        Me.Size = New System.Drawing.Size(1509, 688)
        Me.Panel_Menu.ResumeLayout(False)
        Me.Panel_Menu.PerformLayout()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        CType(Me.Process_Line_Grid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        CType(Me.Process_Line_QC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Save_Com As Button
    Friend WithEvents Delete_Com As Button
    Friend WithEvents Insert_Com As Button
    Friend WithEvents Search_Com As Button
    Friend WithEvents Search_Text As TextBox
    Friend WithEvents Di_Panel1 As Panel
    Friend WithEvents Panel_Menu As Panel
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents Button2 As Button
    Friend WithEvents Process_Line_Grid As DataGridView
    Friend WithEvents Panel3 As Panel
    Friend WithEvents Delete__Bom As Button
    Friend WithEvents Insert__Bom As Button
    Friend WithEvents Grid_Bom_Combo As ComboBox
    Friend WithEvents Process_Line_QC As DataGridView
    Friend WithEvents Com_BOM As Button
    Friend WithEvents Button1 As Button
End Class
