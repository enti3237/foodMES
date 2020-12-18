<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Frm_Customer
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
        Me.Customer_Grid = New System.Windows.Forms.DataGridView()
        Me.Di_Panel1 = New System.Windows.Forms.Panel()
        Me.Search_Text = New System.Windows.Forms.TextBox()
        Me.Search_Com = New System.Windows.Forms.Button()
        Me.Insert_Com = New System.Windows.Forms.Button()
        Me.Delete_Com = New System.Windows.Forms.Button()
        Me.Save_Com = New System.Windows.Forms.Button()
        Me.Panel_Menu = New System.Windows.Forms.Panel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        CType(Me.Customer_Grid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel_Menu.SuspendLayout()
        Me.SuspendLayout()
        '
        'Customer_Grid
        '
        Me.Customer_Grid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Customer_Grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Customer_Grid.Location = New System.Drawing.Point(21, 101)
        Me.Customer_Grid.Name = "Customer_Grid"
        Me.Customer_Grid.RowTemplate.Height = 23
        Me.Customer_Grid.Size = New System.Drawing.Size(1456, 412)
        Me.Customer_Grid.TabIndex = 24
        '
        'Di_Panel1
        '
        Me.Di_Panel1.BackColor = System.Drawing.Color.CornflowerBlue
        Me.Di_Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Di_Panel1.Location = New System.Drawing.Point(0, 71)
        Me.Di_Panel1.Margin = New System.Windows.Forms.Padding(0)
        Me.Di_Panel1.Name = "Di_Panel1"
        Me.Di_Panel1.Size = New System.Drawing.Size(1506, 2)
        Me.Di_Panel1.TabIndex = 25
        '
        'Search_Text
        '
        Me.Search_Text.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Search_Text.Location = New System.Drawing.Point(11, 37)
        Me.Search_Text.Name = "Search_Text"
        Me.Search_Text.Size = New System.Drawing.Size(124, 25)
        Me.Search_Text.TabIndex = 10
        '
        'Search_Com
        '
        Me.Search_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Search_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Search_Com.ForeColor = System.Drawing.Color.Black
        Me.Search_Com.Location = New System.Drawing.Point(141, 8)
        Me.Search_Com.Name = "Search_Com"
        Me.Search_Com.Size = New System.Drawing.Size(99, 54)
        Me.Search_Com.TabIndex = 11
        Me.Search_Com.Text = "검색"
        Me.Search_Com.UseVisualStyleBackColor = False
        '
        'Insert_Com
        '
        Me.Insert_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Insert_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Insert_Com.Location = New System.Drawing.Point(246, 8)
        Me.Insert_Com.Name = "Insert_Com"
        Me.Insert_Com.Size = New System.Drawing.Size(98, 54)
        Me.Insert_Com.TabIndex = 12
        Me.Insert_Com.Text = "추가"
        Me.Insert_Com.UseVisualStyleBackColor = False
        '
        'Delete_Com
        '
        Me.Delete_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Delete_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Delete_Com.Location = New System.Drawing.Point(350, 8)
        Me.Delete_Com.Name = "Delete_Com"
        Me.Delete_Com.Size = New System.Drawing.Size(98, 54)
        Me.Delete_Com.TabIndex = 13
        Me.Delete_Com.Text = "삭제"
        Me.Delete_Com.UseVisualStyleBackColor = False
        '
        'Save_Com
        '
        Me.Save_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Save_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Save_Com.ForeColor = System.Drawing.Color.Black
        Me.Save_Com.Location = New System.Drawing.Point(663, 8)
        Me.Save_Com.Name = "Save_Com"
        Me.Save_Com.Size = New System.Drawing.Size(99, 54)
        Me.Save_Com.TabIndex = 14
        Me.Save_Com.Text = "저장"
        Me.Save_Com.UseVisualStyleBackColor = False
        Me.Save_Com.Visible = False
        '
        'Panel_Menu
        '
        Me.Panel_Menu.Controls.Add(Me.Button2)
        Me.Panel_Menu.Controls.Add(Me.Button1)
        Me.Panel_Menu.Controls.Add(Me.Save_Com)
        Me.Panel_Menu.Controls.Add(Me.Delete_Com)
        Me.Panel_Menu.Controls.Add(Me.Insert_Com)
        Me.Panel_Menu.Controls.Add(Me.Search_Com)
        Me.Panel_Menu.Controls.Add(Me.Search_Text)
        Me.Panel_Menu.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel_Menu.Location = New System.Drawing.Point(0, 0)
        Me.Panel_Menu.Name = "Panel_Menu"
        Me.Panel_Menu.Size = New System.Drawing.Size(1506, 71)
        Me.Panel_Menu.TabIndex = 22
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button1.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button1.Location = New System.Drawing.Point(559, 8)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(98, 54)
        Me.Button1.TabIndex = 15
        Me.Button1.Text = "거래처등록"
        Me.Button1.UseVisualStyleBackColor = False
        Me.Button1.Visible = False
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button2.Enabled = False
        Me.Button2.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.Black
        Me.Button2.Location = New System.Drawing.Point(12, 8)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(123, 26)
        Me.Button2.TabIndex = 26
        Me.Button2.Text = "거래처명"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Frm_Customer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Di_Panel1)
        Me.Controls.Add(Me.Panel_Menu)
        Me.Controls.Add(Me.Customer_Grid)
        Me.Name = "Frm_Customer"
        Me.Size = New System.Drawing.Size(1506, 641)
        CType(Me.Customer_Grid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel_Menu.ResumeLayout(False)
        Me.Panel_Menu.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Customer_Grid As DataGridView
    Friend WithEvents Di_Panel1 As Panel
    Friend WithEvents Search_Text As TextBox
    Friend WithEvents Search_Com As Button
    Friend WithEvents Insert_Com As Button
    Friend WithEvents Delete_Com As Button
    Friend WithEvents Save_Com As Button
    Friend WithEvents Panel_Menu As Panel
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
End Class
