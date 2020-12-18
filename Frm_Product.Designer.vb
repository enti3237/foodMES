<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Frm_Product
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm_Product))
        Me.Di_Panel1 = New System.Windows.Forms.Panel()
        Me.Panel_Menu = New System.Windows.Forms.Panel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Btn_Form = New System.Windows.Forms.Button()
        Me.Delete_Com = New System.Windows.Forms.Button()
        Me.Insert_Com = New System.Windows.Forms.Button()
        Me.Search_Com = New System.Windows.Forms.Button()
        Me.Search_Text = New System.Windows.Forms.TextBox()
        Me.Com_BOM = New System.Windows.Forms.Button()
        Me.Pic_File = New System.Windows.Forms.OpenFileDialog()
        Me.Grid_Bom_Combo = New System.Windows.Forms.ComboBox()
        Me.Grid_Pro_Combo = New System.Windows.Forms.ComboBox()
        Me.Grid_Pro = New System.Windows.Forms.DataGridView()
        Me.Tree_Image = New System.Windows.Forms.ImageList(Me.components)
        Me.Grid_Info = New System.Windows.Forms.DataGridView()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Grid_Bom = New System.Windows.Forms.DataGridView()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Delete__Bom = New System.Windows.Forms.Button()
        Me.Insert__Bom = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Save_delete = New System.Windows.Forms.Button()
        Me.Insert_Pro = New System.Windows.Forms.Button()
        Me.Panel_Menu.SuspendLayout()
        CType(Me.Grid_Pro, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Grid_Info, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Grid_Bom, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Di_Panel1
        '
        Me.Di_Panel1.BackColor = System.Drawing.Color.CornflowerBlue
        Me.Di_Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Di_Panel1.Location = New System.Drawing.Point(0, 86)
        Me.Di_Panel1.Margin = New System.Windows.Forms.Padding(0)
        Me.Di_Panel1.Name = "Di_Panel1"
        Me.Di_Panel1.Size = New System.Drawing.Size(2173, 2)
        Me.Di_Panel1.TabIndex = 32
        '
        'Panel_Menu
        '
        Me.Panel_Menu.Controls.Add(Me.Button1)
        Me.Panel_Menu.Controls.Add(Me.Btn_Form)
        Me.Panel_Menu.Controls.Add(Me.Delete_Com)
        Me.Panel_Menu.Controls.Add(Me.Insert_Com)
        Me.Panel_Menu.Controls.Add(Me.Search_Com)
        Me.Panel_Menu.Controls.Add(Me.Search_Text)
        Me.Panel_Menu.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel_Menu.Location = New System.Drawing.Point(0, 0)
        Me.Panel_Menu.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Panel_Menu.Name = "Panel_Menu"
        Me.Panel_Menu.Size = New System.Drawing.Size(2173, 86)
        Me.Panel_Menu.TabIndex = 31
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button1.Enabled = False
        Me.Button1.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.Black
        Me.Button1.Location = New System.Drawing.Point(14, 6)
        Me.Button1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(141, 32)
        Me.Button1.TabIndex = 25
        Me.Button1.Text = "품명"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Btn_Form
        '
        Me.Btn_Form.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Btn_Form.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Btn_Form.Location = New System.Drawing.Point(519, 6)
        Me.Btn_Form.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Btn_Form.Name = "Btn_Form"
        Me.Btn_Form.Size = New System.Drawing.Size(112, 68)
        Me.Btn_Form.TabIndex = 15
        Me.Btn_Form.Text = "Excel upload"
        Me.Btn_Form.UseVisualStyleBackColor = False
        '
        'Delete_Com
        '
        Me.Delete_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Delete_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Delete_Com.Location = New System.Drawing.Point(400, 6)
        Me.Delete_Com.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Delete_Com.Name = "Delete_Com"
        Me.Delete_Com.Size = New System.Drawing.Size(112, 68)
        Me.Delete_Com.TabIndex = 13
        Me.Delete_Com.Text = "삭제"
        Me.Delete_Com.UseVisualStyleBackColor = False
        '
        'Insert_Com
        '
        Me.Insert_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Insert_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Insert_Com.Location = New System.Drawing.Point(281, 6)
        Me.Insert_Com.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Insert_Com.Name = "Insert_Com"
        Me.Insert_Com.Size = New System.Drawing.Size(112, 68)
        Me.Insert_Com.TabIndex = 12
        Me.Insert_Com.Text = "추가"
        Me.Insert_Com.UseVisualStyleBackColor = False
        '
        'Search_Com
        '
        Me.Search_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Search_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Search_Com.ForeColor = System.Drawing.Color.Black
        Me.Search_Com.Location = New System.Drawing.Point(161, 6)
        Me.Search_Com.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Search_Com.Name = "Search_Com"
        Me.Search_Com.Size = New System.Drawing.Size(113, 68)
        Me.Search_Com.TabIndex = 11
        Me.Search_Com.Text = "검색"
        Me.Search_Com.UseVisualStyleBackColor = False
        '
        'Search_Text
        '
        Me.Search_Text.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Search_Text.Location = New System.Drawing.Point(13, 42)
        Me.Search_Text.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Search_Text.Name = "Search_Text"
        Me.Search_Text.Size = New System.Drawing.Size(141, 29)
        Me.Search_Text.TabIndex = 10
        '
        'Com_BOM
        '
        Me.Com_BOM.AutoSize = True
        Me.Com_BOM.Dock = System.Windows.Forms.DockStyle.Left
        Me.Com_BOM.FlatAppearance.BorderSize = 0
        Me.Com_BOM.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange
        Me.Com_BOM.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Com_BOM.Location = New System.Drawing.Point(0, 495)
        Me.Com_BOM.Margin = New System.Windows.Forms.Padding(0)
        Me.Com_BOM.Name = "Com_BOM"
        Me.Com_BOM.Size = New System.Drawing.Size(109, 67)
        Me.Com_BOM.TabIndex = 46
        Me.Com_BOM.Text = "레시피"
        Me.Com_BOM.UseVisualStyleBackColor = True
        '
        'Pic_File
        '
        Me.Pic_File.FileName = "OpenFileDialog1"
        '
        'Grid_Bom_Combo
        '
        Me.Grid_Bom_Combo.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.Grid_Bom_Combo.FormattingEnabled = True
        Me.Grid_Bom_Combo.Location = New System.Drawing.Point(1706, 497)
        Me.Grid_Bom_Combo.Margin = New System.Windows.Forms.Padding(2)
        Me.Grid_Bom_Combo.Name = "Grid_Bom_Combo"
        Me.Grid_Bom_Combo.Size = New System.Drawing.Size(142, 31)
        Me.Grid_Bom_Combo.TabIndex = 21
        Me.Grid_Bom_Combo.Visible = False
        '
        'Grid_Pro_Combo
        '
        Me.Grid_Pro_Combo.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.Grid_Pro_Combo.FormattingEnabled = True
        Me.Grid_Pro_Combo.Location = New System.Drawing.Point(1706, 992)
        Me.Grid_Pro_Combo.Margin = New System.Windows.Forms.Padding(2)
        Me.Grid_Pro_Combo.Name = "Grid_Pro_Combo"
        Me.Grid_Pro_Combo.Size = New System.Drawing.Size(142, 31)
        Me.Grid_Pro_Combo.TabIndex = 22
        Me.Grid_Pro_Combo.Visible = False
        '
        'Grid_Pro
        '
        Me.Grid_Pro.AllowUserToAddRows = False
        Me.Grid_Pro.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Grid_Pro.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Grid_Pro.Location = New System.Drawing.Point(2, 1035)
        Me.Grid_Pro.Margin = New System.Windows.Forms.Padding(2)
        Me.Grid_Pro.Name = "Grid_Pro"
        Me.Grid_Pro.RowHeadersWidth = 51
        Me.Grid_Pro.RowTemplate.Height = 30
        Me.Grid_Pro.Size = New System.Drawing.Size(1700, 363)
        Me.Grid_Pro.TabIndex = 19
        '
        'Tree_Image
        '
        Me.Tree_Image.ImageStream = CType(resources.GetObject("Tree_Image.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.Tree_Image.TransparentColor = System.Drawing.Color.Transparent
        Me.Tree_Image.Images.SetKeyName(0, "제품.png")
        Me.Tree_Image.Images.SetKeyName(1, "조립품.png")
        Me.Tree_Image.Images.SetKeyName(2, "단품.png")
        Me.Tree_Image.Images.SetKeyName(3, "공정.png")
        Me.Tree_Image.Images.SetKeyName(4, "dongil.ico")
        '
        'Grid_Info
        '
        Me.Grid_Info.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Grid_Info.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Grid_Info.Location = New System.Drawing.Point(3, 63)
        Me.Grid_Info.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Grid_Info.Name = "Grid_Info"
        Me.Grid_Info.RowHeadersWidth = 51
        Me.Grid_Info.RowTemplate.Height = 23
        Me.Grid_Info.Size = New System.Drawing.Size(1698, 428)
        Me.Grid_Info.TabIndex = 52
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Grid_Bom
        '
        Me.Grid_Bom.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Grid_Bom.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Grid_Bom.Location = New System.Drawing.Point(3, 566)
        Me.Grid_Bom.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Grid_Bom.Name = "Grid_Bom"
        Me.Grid_Bom.RowHeadersWidth = 51
        Me.Grid_Bom.RowTemplate.Height = 23
        Me.Grid_Bom.Size = New System.Drawing.Size(1698, 420)
        Me.Grid_Bom.TabIndex = 34
        '
        'Button2
        '
        Me.Button2.AutoSize = True
        Me.Button2.Dock = System.Windows.Forms.DockStyle.Left
        Me.Button2.FlatAppearance.BorderSize = 0
        Me.Button2.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Location = New System.Drawing.Point(0, 0)
        Me.Button2.Margin = New System.Windows.Forms.Padding(0)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(109, 59)
        Me.Button2.TabIndex = 60
        Me.Button2.Text = "제품정보"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 78.46295!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 21.53705!))
        Me.TableLayoutPanel1.Controls.Add(Me.Button2, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Com_BOM, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Grid_Pro, 0, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.Grid_Info, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Grid_Pro_Combo, 1, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel3, 1, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Button3, 0, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel1, 1, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.Grid_Bom_Combo, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Grid_Bom, 0, 3)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 88)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 6
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 4.249921!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 31.15587!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 4.792464!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 30.61715!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 3.142523!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 26.04207!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(2173, 1400)
        Me.TableLayoutPanel1.TabIndex = 62
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.FromArgb(CType(CType(91, Byte), Integer), CType(CType(155, Byte), Integer), CType(CType(213, Byte), Integer))
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.Delete__Bom)
        Me.Panel3.Controls.Add(Me.Insert__Bom)
        Me.Panel3.Location = New System.Drawing.Point(1707, 566)
        Me.Panel3.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(141, 158)
        Me.Panel3.TabIndex = 62
        '
        'Delete__Bom
        '
        Me.Delete__Bom.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Delete__Bom.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Delete__Bom.Location = New System.Drawing.Point(3, 79)
        Me.Delete__Bom.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Delete__Bom.Name = "Delete__Bom"
        Me.Delete__Bom.Size = New System.Drawing.Size(130, 68)
        Me.Delete__Bom.TabIndex = 7
        Me.Delete__Bom.Text = "삭제"
        Me.Delete__Bom.UseVisualStyleBackColor = False
        '
        'Insert__Bom
        '
        Me.Insert__Bom.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Insert__Bom.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Insert__Bom.Location = New System.Drawing.Point(3, 4)
        Me.Insert__Bom.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Insert__Bom.Name = "Insert__Bom"
        Me.Insert__Bom.Size = New System.Drawing.Size(130, 68)
        Me.Insert__Bom.TabIndex = 5
        Me.Insert__Bom.Text = "추가"
        Me.Insert__Bom.UseVisualStyleBackColor = False
        '
        'Button3
        '
        Me.Button3.AutoSize = True
        Me.Button3.Dock = System.Windows.Forms.DockStyle.Left
        Me.Button3.FlatAppearance.BorderSize = 0
        Me.Button3.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button3.Location = New System.Drawing.Point(0, 990)
        Me.Button3.Margin = New System.Windows.Forms.Padding(0)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(109, 43)
        Me.Button3.TabIndex = 61
        Me.Button3.Text = "공정"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(91, Byte), Integer), CType(CType(155, Byte), Integer), CType(CType(213, Byte), Integer))
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.Save_delete)
        Me.Panel1.Controls.Add(Me.Insert_Pro)
        Me.Panel1.Location = New System.Drawing.Point(1707, 1037)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(141, 159)
        Me.Panel1.TabIndex = 63
        '
        'Save_delete
        '
        Me.Save_delete.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Save_delete.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Save_delete.Location = New System.Drawing.Point(3, 79)
        Me.Save_delete.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Save_delete.Name = "Save_delete"
        Me.Save_delete.Size = New System.Drawing.Size(130, 68)
        Me.Save_delete.TabIndex = 7
        Me.Save_delete.Text = "삭제"
        Me.Save_delete.UseVisualStyleBackColor = False
        '
        'Insert_Pro
        '
        Me.Insert_Pro.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Insert_Pro.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Insert_Pro.Location = New System.Drawing.Point(3, 4)
        Me.Insert_Pro.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Insert_Pro.Name = "Insert_Pro"
        Me.Insert_Pro.Size = New System.Drawing.Size(130, 68)
        Me.Insert_Pro.TabIndex = 5
        Me.Insert_Pro.Text = "추가"
        Me.Insert_Pro.UseVisualStyleBackColor = False
        '
        'Frm_Product
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.Di_Panel1)
        Me.Controls.Add(Me.Panel_Menu)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "Frm_Product"
        Me.Size = New System.Drawing.Size(2173, 1488)
        Me.Panel_Menu.ResumeLayout(False)
        Me.Panel_Menu.PerformLayout()
        CType(Me.Grid_Pro, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Grid_Info, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Grid_Bom, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Di_Panel1 As Panel
    Friend WithEvents Panel_Menu As Panel
    Friend WithEvents Delete_Com As Button
    Friend WithEvents Insert_Com As Button
    Friend WithEvents Search_Com As Button
    Friend WithEvents Search_Text As TextBox
    Friend WithEvents Com_BOM As Button
    Friend WithEvents Pic_File As OpenFileDialog
    Friend WithEvents Grid_Pro_Combo As ComboBox
    Friend WithEvents Grid_Bom_Combo As ComboBox
    Friend WithEvents Grid_Pro As DataGridView
    Friend WithEvents Tree_Image As ImageList
    Friend WithEvents Grid_Info As DataGridView
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents Btn_Form As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents Grid_Bom As DataGridView
    Friend WithEvents Button2 As Button
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents Button3 As Button
    Friend WithEvents Panel3 As Panel
    Friend WithEvents Delete__Bom As Button
    Friend WithEvents Insert__Bom As Button
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Save_delete As Button
    Friend WithEvents Insert_Pro As Button
End Class
