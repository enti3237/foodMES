<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Frm_Contract2
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
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Panel_Excel = New System.Windows.Forms.Panel()
        Me.Panel_Menu = New System.Windows.Forms.Panel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Com_Excel = New System.Windows.Forms.Button()
        Me.Delete_Com = New System.Windows.Forms.Button()
        Me.Save_Com = New System.Windows.Forms.Button()
        Me.Com_Excel_Print = New System.Windows.Forms.Button()
        Me.Insert_Com = New System.Windows.Forms.Button()
        Me.Search_Com = New System.Windows.Forms.Button()
        Me.Search_Text = New System.Windows.Forms.TextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Log = New System.Windows.Forms.DataGridView()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Grid_Info = New System.Windows.Forms.DataGridView()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Delete__Bom = New System.Windows.Forms.Button()
        Me.Insert__Bom = New System.Windows.Forms.Button()
        Me.Grid_Bom_Combo = New System.Windows.Forms.ComboBox()
        Me.Grid_Main = New System.Windows.Forms.DataGridView()
        Me.Com_BOM = New System.Windows.Forms.Button()
        Me.Panel_Menu.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.Log, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Grid_Info, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        CType(Me.Grid_Main, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Di_Panel1
        '
        Me.Di_Panel1.BackColor = System.Drawing.Color.CornflowerBlue
        Me.Di_Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Di_Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Di_Panel1.Margin = New System.Windows.Forms.Padding(0)
        Me.Di_Panel1.Name = "Di_Panel1"
        Me.Di_Panel1.Size = New System.Drawing.Size(1332, 2)
        Me.Di_Panel1.TabIndex = 66
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Panel_Excel
        '
        Me.Panel_Excel.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel_Excel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel_Excel.Cursor = System.Windows.Forms.Cursors.Default
        Me.Panel_Excel.Location = New System.Drawing.Point(262, 103)
        Me.Panel_Excel.Name = "Panel_Excel"
        Me.Panel_Excel.Size = New System.Drawing.Size(939, 519)
        Me.Panel_Excel.TabIndex = 75
        '
        'Panel_Menu
        '
        Me.Panel_Menu.Controls.Add(Me.Button1)
        Me.Panel_Menu.Controls.Add(Me.Com_Excel)
        Me.Panel_Menu.Controls.Add(Me.Delete_Com)
        Me.Panel_Menu.Controls.Add(Me.Save_Com)
        Me.Panel_Menu.Controls.Add(Me.Com_Excel_Print)
        Me.Panel_Menu.Controls.Add(Me.Insert_Com)
        Me.Panel_Menu.Controls.Add(Me.Search_Com)
        Me.Panel_Menu.Controls.Add(Me.Search_Text)
        Me.Panel_Menu.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel_Menu.Location = New System.Drawing.Point(0, 2)
        Me.Panel_Menu.Name = "Panel_Menu"
        Me.Panel_Menu.Size = New System.Drawing.Size(1332, 80)
        Me.Panel_Menu.TabIndex = 65
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button1.Enabled = False
        Me.Button1.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.Black
        Me.Button1.Location = New System.Drawing.Point(12, 14)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(130, 26)
        Me.Button1.TabIndex = 75
        Me.Button1.Text = "거래처명"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Com_Excel
        '
        Me.Com_Excel.AutoSize = True
        Me.Com_Excel.FlatAppearance.BorderSize = 0
        Me.Com_Excel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange
        Me.Com_Excel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Com_Excel.Location = New System.Drawing.Point(620, 43)
        Me.Com_Excel.Margin = New System.Windows.Forms.Padding(0)
        Me.Com_Excel.Name = "Com_Excel"
        Me.Com_Excel.Size = New System.Drawing.Size(98, 25)
        Me.Com_Excel.TabIndex = 74
        Me.Com_Excel.Text = "Excel"
        Me.Com_Excel.UseVisualStyleBackColor = True
        '
        'Delete_Com
        '
        Me.Delete_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Delete_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Delete_Com.Location = New System.Drawing.Point(376, 14)
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
        Me.Save_Com.Location = New System.Drawing.Point(775, 14)
        Me.Save_Com.Name = "Save_Com"
        Me.Save_Com.Size = New System.Drawing.Size(108, 54)
        Me.Save_Com.TabIndex = 14
        Me.Save_Com.Text = "저장"
        Me.Save_Com.UseVisualStyleBackColor = False
        Me.Save_Com.Visible = False
        '
        'Com_Excel_Print
        '
        Me.Com_Excel_Print.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Com_Excel_Print.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Com_Excel_Print.Location = New System.Drawing.Point(489, 14)
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
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.CornflowerBlue
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 82)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1332, 2)
        Me.Panel1.TabIndex = 76
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 89.78979!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.21021!))
        Me.TableLayoutPanel1.Controls.Add(Me.Button3, 0, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.Log, 0, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.Button2, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Grid_Info, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel3, 1, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Grid_Bom_Combo, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Grid_Main, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Com_BOM, 0, 2)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 84)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 6
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 4.186598!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 32.5149!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 4.213884!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 27.80531!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 3.501776!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 27.77753!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(1332, 561)
        Me.TableLayoutPanel1.TabIndex = 77
        '
        'Button3
        '
        Me.Button3.AutoSize = True
        Me.Button3.FlatAppearance.BorderSize = 0
        Me.Button3.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button3.Font = New System.Drawing.Font("굴림", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button3.Location = New System.Drawing.Point(0, 383)
        Me.Button3.Margin = New System.Windows.Forms.Padding(0)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(95, 19)
        Me.Button3.TabIndex = 76
        Me.Button3.Text = "개정이력"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Log
        '
        Me.Log.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.Log.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Log.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Log.Location = New System.Drawing.Point(3, 405)
        Me.Log.Name = "Log"
        Me.Log.RowTemplate.Height = 23
        Me.Log.Size = New System.Drawing.Size(1189, 153)
        Me.Log.TabIndex = 66
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
        Me.Button2.Text = "수주정보"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Grid_Info
        '
        Me.Grid_Info.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.Grid_Info.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Grid_Info.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Grid_Info.Location = New System.Drawing.Point(3, 26)
        Me.Grid_Info.Name = "Grid_Info"
        Me.Grid_Info.RowTemplate.Height = 23
        Me.Grid_Info.Size = New System.Drawing.Size(1189, 176)
        Me.Grid_Info.TabIndex = 52
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.FromArgb(CType(CType(91, Byte), Integer), CType(CType(155, Byte), Integer), CType(CType(213, Byte), Integer))
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.Delete__Bom)
        Me.Panel3.Controls.Add(Me.Insert__Bom)
        Me.Panel3.Location = New System.Drawing.Point(1198, 231)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(124, 130)
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
        Me.Grid_Bom_Combo.Location = New System.Drawing.Point(1197, 207)
        Me.Grid_Bom_Combo.Margin = New System.Windows.Forms.Padding(2)
        Me.Grid_Bom_Combo.Name = "Grid_Bom_Combo"
        Me.Grid_Bom_Combo.Size = New System.Drawing.Size(125, 25)
        Me.Grid_Bom_Combo.TabIndex = 21
        Me.Grid_Bom_Combo.Visible = False
        '
        'Grid_Main
        '
        Me.Grid_Main.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.Grid_Main.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Grid_Main.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Grid_Main.Location = New System.Drawing.Point(3, 231)
        Me.Grid_Main.Name = "Grid_Main"
        Me.Grid_Main.RowTemplate.Height = 23
        Me.Grid_Main.Size = New System.Drawing.Size(1189, 149)
        Me.Grid_Main.TabIndex = 34
        '
        'Com_BOM
        '
        Me.Com_BOM.AutoSize = True
        Me.Com_BOM.FlatAppearance.BorderSize = 0
        Me.Com_BOM.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange
        Me.Com_BOM.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Com_BOM.Font = New System.Drawing.Font("굴림", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Com_BOM.Location = New System.Drawing.Point(0, 205)
        Me.Com_BOM.Margin = New System.Windows.Forms.Padding(0)
        Me.Com_BOM.Name = "Com_BOM"
        Me.Com_BOM.Size = New System.Drawing.Size(95, 23)
        Me.Com_BOM.TabIndex = 46
        Me.Com_BOM.Text = "세부정보"
        Me.Com_BOM.UseVisualStyleBackColor = True
        '
        'Frm_Contract2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel_Menu)
        Me.Controls.Add(Me.Panel_Excel)
        Me.Controls.Add(Me.Di_Panel1)
        Me.Name = "Frm_Contract2"
        Me.Size = New System.Drawing.Size(1332, 645)
        Me.Panel_Menu.ResumeLayout(False)
        Me.Panel_Menu.PerformLayout()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        CType(Me.Log, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Grid_Info, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        CType(Me.Grid_Main, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Di_Panel1 As Panel
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents Panel_Excel As Panel
    Friend WithEvents Panel_Menu As Panel
    Friend WithEvents Com_Excel As Button
    Friend WithEvents Delete_Com As Button
    Friend WithEvents Save_Com As Button
    Friend WithEvents Com_Excel_Print As Button
    Friend WithEvents Insert_Com As Button
    Friend WithEvents Search_Com As Button
    Friend WithEvents Search_Text As TextBox
    Friend WithEvents Panel1 As Panel
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents Button2 As Button
    Friend WithEvents Grid_Info As DataGridView
    Friend WithEvents Panel3 As Panel
    Friend WithEvents Delete__Bom As Button
    Friend WithEvents Insert__Bom As Button
    Friend WithEvents Grid_Bom_Combo As ComboBox
    Friend WithEvents Grid_Main As DataGridView
    Friend WithEvents Com_BOM As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Log As DataGridView
End Class
