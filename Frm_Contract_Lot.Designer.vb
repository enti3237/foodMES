﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Frm_Contract_Lot
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm_Contract_Lot))
        Me.SerialPort1 = New System.IO.Ports.SerialPort(Me.components)
        Me.Product_Tree = New System.Windows.Forms.TreeView()
        Me.Tree_Image = New System.Windows.Forms.ImageList(Me.components)
        Me.Search_Combo = New System.Windows.Forms.ComboBox()
        Me.Di_Panel2 = New System.Windows.Forms.Panel()
        Me.Color_Label = New System.Windows.Forms.Label()
        Me.Search_Com = New System.Windows.Forms.Button()
        Me.Search_Text = New System.Windows.Forms.TextBox()
        Me.Grid_Info_Combo = New System.Windows.Forms.ComboBox()
        Me.Text1 = New System.Windows.Forms.TextBox()
        Me.Com_Tree = New System.Windows.Forms.Button()
        Me.Grid_Info_Picture1 = New System.Windows.Forms.PictureBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Grid_Code = New System.Windows.Forms.DataGridView()
        Me.Panel_Menu = New System.Windows.Forms.Panel()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Text4 = New System.Windows.Forms.TextBox()
        Me.Text3 = New System.Windows.Forms.TextBox()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Tree_Panel = New System.Windows.Forms.Panel()
        Me.Pic_File = New System.Windows.Forms.OpenFileDialog()
        Me.Grid_Info = New System.Windows.Forms.DataGridView()
        Me.Di_Panel1 = New System.Windows.Forms.Panel()
        Me.Di_Panel2.SuspendLayout()
        CType(Me.Grid_Info_Picture1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Grid_Code, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel_Menu.SuspendLayout()
        Me.Tree_Panel.SuspendLayout()
        CType(Me.Grid_Info, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Product_Tree
        '
        Me.Product_Tree.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Product_Tree.Font = New System.Drawing.Font("맑은 고딕", 9.75!)
        Me.Product_Tree.Location = New System.Drawing.Point(8, 7)
        Me.Product_Tree.Margin = New System.Windows.Forms.Padding(2)
        Me.Product_Tree.Name = "Product_Tree"
        Me.Product_Tree.Size = New System.Drawing.Size(848, 660)
        Me.Product_Tree.TabIndex = 10
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
        'Search_Combo
        '
        Me.Search_Combo.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.Search_Combo.FormattingEnabled = True
        Me.Search_Combo.Location = New System.Drawing.Point(13, 5)
        Me.Search_Combo.Margin = New System.Windows.Forms.Padding(2)
        Me.Search_Combo.Name = "Search_Combo"
        Me.Search_Combo.Size = New System.Drawing.Size(125, 25)
        Me.Search_Combo.TabIndex = 8
        '
        'Di_Panel2
        '
        Me.Di_Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Di_Panel2.BackColor = System.Drawing.Color.Silver
        Me.Di_Panel2.Controls.Add(Me.Color_Label)
        Me.Di_Panel2.Location = New System.Drawing.Point(388, 115)
        Me.Di_Panel2.Margin = New System.Windows.Forms.Padding(2)
        Me.Di_Panel2.Name = "Di_Panel2"
        Me.Di_Panel2.Size = New System.Drawing.Size(1323, 5)
        Me.Di_Panel2.TabIndex = 82
        '
        'Color_Label
        '
        Me.Color_Label.Location = New System.Drawing.Point(180, 3)
        Me.Color_Label.Name = "Color_Label"
        Me.Color_Label.Size = New System.Drawing.Size(47, 5)
        Me.Color_Label.TabIndex = 0
        '
        'Search_Com
        '
        Me.Search_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Search_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Search_Com.ForeColor = System.Drawing.Color.Black
        Me.Search_Com.Location = New System.Drawing.Point(144, 6)
        Me.Search_Com.Name = "Search_Com"
        Me.Search_Com.Size = New System.Drawing.Size(99, 54)
        Me.Search_Com.TabIndex = 11
        Me.Search_Com.Text = "검색"
        Me.Search_Com.UseVisualStyleBackColor = False
        '
        'Search_Text
        '
        Me.Search_Text.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Search_Text.Location = New System.Drawing.Point(14, 35)
        Me.Search_Text.Name = "Search_Text"
        Me.Search_Text.Size = New System.Drawing.Size(124, 25)
        Me.Search_Text.TabIndex = 10
        '
        'Grid_Info_Combo
        '
        Me.Grid_Info_Combo.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.Grid_Info_Combo.FormattingEnabled = True
        Me.Grid_Info_Combo.Location = New System.Drawing.Point(81, 417)
        Me.Grid_Info_Combo.Margin = New System.Windows.Forms.Padding(2)
        Me.Grid_Info_Combo.Name = "Grid_Info_Combo"
        Me.Grid_Info_Combo.Size = New System.Drawing.Size(125, 25)
        Me.Grid_Info_Combo.TabIndex = 80
        Me.Grid_Info_Combo.Visible = False
        '
        'Text1
        '
        Me.Text1.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Text1.Location = New System.Drawing.Point(611, 39)
        Me.Text1.Name = "Text1"
        Me.Text1.Size = New System.Drawing.Size(98, 25)
        Me.Text1.TabIndex = 98
        Me.Text1.Text = "200"
        Me.Text1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Com_Tree
        '
        Me.Com_Tree.AutoSize = True
        Me.Com_Tree.FlatAppearance.BorderSize = 0
        Me.Com_Tree.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange
        Me.Com_Tree.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Com_Tree.Location = New System.Drawing.Point(388, 88)
        Me.Com_Tree.Margin = New System.Windows.Forms.Padding(0)
        Me.Com_Tree.Name = "Com_Tree"
        Me.Com_Tree.Size = New System.Drawing.Size(95, 25)
        Me.Com_Tree.TabIndex = 83
        Me.Com_Tree.Text = "구조"
        Me.Com_Tree.UseVisualStyleBackColor = True
        '
        'Grid_Info_Picture1
        '
        Me.Grid_Info_Picture1.Location = New System.Drawing.Point(71, 473)
        Me.Grid_Info_Picture1.Name = "Grid_Info_Picture1"
        Me.Grid_Info_Picture1.Size = New System.Drawing.Size(239, 127)
        Me.Grid_Info_Picture1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.Grid_Info_Picture1.TabIndex = 81
        Me.Grid_Info_Picture1.TabStop = False
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button2.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.Black
        Me.Button2.Location = New System.Drawing.Point(13, 329)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(359, 27)
        Me.Button2.TabIndex = 79
        Me.Button2.Text = "상세 정보"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button3.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button3.ForeColor = System.Drawing.Color.Black
        Me.Button3.Location = New System.Drawing.Point(923, 6)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(99, 58)
        Me.Button3.TabIndex = 106
        Me.Button3.Text = "인쇄"
        Me.Button3.UseVisualStyleBackColor = False
        '
        'Grid_Code
        '
        Me.Grid_Code.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Grid_Code.Location = New System.Drawing.Point(13, 88)
        Me.Grid_Code.Name = "Grid_Code"
        Me.Grid_Code.RowTemplate.Height = 23
        Me.Grid_Code.Size = New System.Drawing.Size(359, 228)
        Me.Grid_Code.TabIndex = 77
        '
        'Panel_Menu
        '
        Me.Panel_Menu.Controls.Add(Me.Button3)
        Me.Panel_Menu.Controls.Add(Me.Button6)
        Me.Panel_Menu.Controls.Add(Me.Text4)
        Me.Panel_Menu.Controls.Add(Me.Text3)
        Me.Panel_Menu.Controls.Add(Me.Button4)
        Me.Panel_Menu.Controls.Add(Me.Text1)
        Me.Panel_Menu.Controls.Add(Me.Button1)
        Me.Panel_Menu.Controls.Add(Me.Search_Com)
        Me.Panel_Menu.Controls.Add(Me.Search_Text)
        Me.Panel_Menu.Controls.Add(Me.Search_Combo)
        Me.Panel_Menu.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel_Menu.Location = New System.Drawing.Point(0, 2)
        Me.Panel_Menu.Name = "Panel_Menu"
        Me.Panel_Menu.Size = New System.Drawing.Size(1275, 69)
        Me.Panel_Menu.TabIndex = 75
        '
        'Button6
        '
        Me.Button6.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button6.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button6.ForeColor = System.Drawing.Color.Black
        Me.Button6.Location = New System.Drawing.Point(819, 6)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(98, 27)
        Me.Button6.TabIndex = 105
        Me.Button6.Text = "전체수량"
        Me.Button6.UseVisualStyleBackColor = False
        '
        'Text4
        '
        Me.Text4.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Text4.Location = New System.Drawing.Point(819, 39)
        Me.Text4.Name = "Text4"
        Me.Text4.Size = New System.Drawing.Size(98, 25)
        Me.Text4.TabIndex = 104
        Me.Text4.Text = "0"
        Me.Text4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Text3
        '
        Me.Text3.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Text3.Location = New System.Drawing.Point(715, 39)
        Me.Text3.Name = "Text3"
        Me.Text3.Size = New System.Drawing.Size(98, 25)
        Me.Text3.TabIndex = 102
        Me.Text3.Text = "2000"
        Me.Text3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button4.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button4.ForeColor = System.Drawing.Color.Black
        Me.Button4.Location = New System.Drawing.Point(715, 6)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(98, 27)
        Me.Button4.TabIndex = 101
        Me.Button4.Text = "제품 대 포장"
        Me.Button4.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button1.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.Black
        Me.Button1.Location = New System.Drawing.Point(611, 6)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(98, 27)
        Me.Button1.TabIndex = 97
        Me.Button1.Text = "제품 소포장 단위"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Tree_Panel
        '
        Me.Tree_Panel.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Tree_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Tree_Panel.Controls.Add(Me.Product_Tree)
        Me.Tree_Panel.Location = New System.Drawing.Point(391, 135)
        Me.Tree_Panel.Name = "Tree_Panel"
        Me.Tree_Panel.Size = New System.Drawing.Size(866, 677)
        Me.Tree_Panel.TabIndex = 84
        '
        'Pic_File
        '
        Me.Pic_File.FileName = "OpenFileDialog1"
        '
        'Grid_Info
        '
        Me.Grid_Info.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Grid_Info.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Grid_Info.Location = New System.Drawing.Point(13, 354)
        Me.Grid_Info.Name = "Grid_Info"
        Me.Grid_Info.RowTemplate.Height = 23
        Me.Grid_Info.Size = New System.Drawing.Size(359, 458)
        Me.Grid_Info.TabIndex = 78
        '
        'Di_Panel1
        '
        Me.Di_Panel1.BackColor = System.Drawing.Color.CornflowerBlue
        Me.Di_Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Di_Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Di_Panel1.Margin = New System.Windows.Forms.Padding(0)
        Me.Di_Panel1.Name = "Di_Panel1"
        Me.Di_Panel1.Size = New System.Drawing.Size(1275, 2)
        Me.Di_Panel1.TabIndex = 76
        '
        'Frm_Contract_Lot
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Di_Panel2)
        Me.Controls.Add(Me.Grid_Info_Combo)
        Me.Controls.Add(Me.Com_Tree)
        Me.Controls.Add(Me.Grid_Info_Picture1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Grid_Code)
        Me.Controls.Add(Me.Panel_Menu)
        Me.Controls.Add(Me.Tree_Panel)
        Me.Controls.Add(Me.Grid_Info)
        Me.Controls.Add(Me.Di_Panel1)
        Me.Name = "Frm_Contract_Lot"
        Me.Size = New System.Drawing.Size(1275, 832)
        Me.Di_Panel2.ResumeLayout(False)
        CType(Me.Grid_Info_Picture1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Grid_Code, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel_Menu.ResumeLayout(False)
        Me.Panel_Menu.PerformLayout()
        Me.Tree_Panel.ResumeLayout(False)
        CType(Me.Grid_Info, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents SerialPort1 As IO.Ports.SerialPort
    Friend WithEvents Product_Tree As TreeView
    Friend WithEvents Tree_Image As ImageList
    Friend WithEvents Search_Combo As ComboBox
    Friend WithEvents Di_Panel2 As Panel
    Friend WithEvents Color_Label As Label
    Friend WithEvents Search_Com As Button
    Friend WithEvents Search_Text As TextBox
    Friend WithEvents Grid_Info_Combo As ComboBox
    Friend WithEvents Text1 As TextBox
    Friend WithEvents Com_Tree As Button
    Friend WithEvents Grid_Info_Picture1 As PictureBox
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Grid_Code As DataGridView
    Friend WithEvents Panel_Menu As Panel
    Friend WithEvents Button6 As Button
    Friend WithEvents Text4 As TextBox
    Friend WithEvents Text3 As TextBox
    Friend WithEvents Button4 As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents Tree_Panel As Panel
    Friend WithEvents Pic_File As OpenFileDialog
    Friend WithEvents Grid_Info As DataGridView
    Friend WithEvents Di_Panel1 As Panel
End Class
