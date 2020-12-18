<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Frm_Temper
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
        Me.Process_Line_Grid = New System.Windows.Forms.DataGridView()
        Me.Save_Com = New System.Windows.Forms.Button()
        Me.Search_Com = New System.Windows.Forms.Button()
        Me.Search_Text = New System.Windows.Forms.TextBox()
        Me.Search_Combo = New System.Windows.Forms.ComboBox()
        Me.Di_Panel1 = New System.Windows.Forms.Panel()
        Me.Panel_Menu = New System.Windows.Forms.Panel()
        Me.Delete_Com = New System.Windows.Forms.Button()
        Me.Insert_Com = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.E_Time = New System.Windows.Forms.TextBox()
        Me.E_Day = New System.Windows.Forms.TextBox()
        Me.S_Time = New System.Windows.Forms.TextBox()
        Me.S_Day = New System.Windows.Forms.TextBox()
        Me.Button2 = New System.Windows.Forms.Button()
        CType(Me.Process_Line_Grid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel_Menu.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Process_Line_Grid
        '
        Me.Process_Line_Grid.AllowUserToAddRows = False
        Me.Process_Line_Grid.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Process_Line_Grid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.Process_Line_Grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Process_Line_Grid.Location = New System.Drawing.Point(12, 81)
        Me.Process_Line_Grid.Margin = New System.Windows.Forms.Padding(2)
        Me.Process_Line_Grid.Name = "Process_Line_Grid"
        Me.Process_Line_Grid.RowHeadersWidth = 51
        Me.Process_Line_Grid.RowTemplate.Height = 30
        Me.Process_Line_Grid.Size = New System.Drawing.Size(395, 536)
        Me.Process_Line_Grid.TabIndex = 44
        '
        'Save_Com
        '
        Me.Save_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Save_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Save_Com.ForeColor = System.Drawing.Color.Black
        Me.Save_Com.Location = New System.Drawing.Point(142, 3)
        Me.Save_Com.Name = "Save_Com"
        Me.Save_Com.Size = New System.Drawing.Size(99, 54)
        Me.Save_Com.TabIndex = 14
        Me.Save_Com.Text = "저장"
        Me.Save_Com.UseVisualStyleBackColor = False
        '
        'Search_Com
        '
        Me.Search_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Search_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Search_Com.ForeColor = System.Drawing.Color.Black
        Me.Search_Com.Location = New System.Drawing.Point(778, 5)
        Me.Search_Com.Name = "Search_Com"
        Me.Search_Com.Size = New System.Drawing.Size(99, 58)
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
        'Search_Combo
        '
        Me.Search_Combo.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.Search_Combo.FormattingEnabled = True
        Me.Search_Combo.Location = New System.Drawing.Point(11, 5)
        Me.Search_Combo.Margin = New System.Windows.Forms.Padding(2)
        Me.Search_Combo.Name = "Search_Combo"
        Me.Search_Combo.Size = New System.Drawing.Size(125, 25)
        Me.Search_Combo.TabIndex = 8
        '
        'Di_Panel1
        '
        Me.Di_Panel1.BackColor = System.Drawing.Color.CornflowerBlue
        Me.Di_Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Di_Panel1.Location = New System.Drawing.Point(0, 65)
        Me.Di_Panel1.Margin = New System.Windows.Forms.Padding(0)
        Me.Di_Panel1.Name = "Di_Panel1"
        Me.Di_Panel1.Size = New System.Drawing.Size(1258, 2)
        Me.Di_Panel1.TabIndex = 43
        '
        'Panel_Menu
        '
        Me.Panel_Menu.Controls.Add(Me.Label1)
        Me.Panel_Menu.Controls.Add(Me.E_Time)
        Me.Panel_Menu.Controls.Add(Me.E_Day)
        Me.Panel_Menu.Controls.Add(Me.S_Time)
        Me.Panel_Menu.Controls.Add(Me.S_Day)
        Me.Panel_Menu.Controls.Add(Me.Button2)
        Me.Panel_Menu.Controls.Add(Me.Delete_Com)
        Me.Panel_Menu.Controls.Add(Me.Insert_Com)
        Me.Panel_Menu.Controls.Add(Me.Save_Com)
        Me.Panel_Menu.Controls.Add(Me.Search_Com)
        Me.Panel_Menu.Controls.Add(Me.Search_Text)
        Me.Panel_Menu.Controls.Add(Me.Search_Combo)
        Me.Panel_Menu.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel_Menu.Location = New System.Drawing.Point(0, 0)
        Me.Panel_Menu.Name = "Panel_Menu"
        Me.Panel_Menu.Size = New System.Drawing.Size(1258, 65)
        Me.Panel_Menu.TabIndex = 42
        '
        'Delete_Com
        '
        Me.Delete_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Delete_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Delete_Com.Location = New System.Drawing.Point(351, 3)
        Me.Delete_Com.Name = "Delete_Com"
        Me.Delete_Com.Size = New System.Drawing.Size(98, 54)
        Me.Delete_Com.TabIndex = 16
        Me.Delete_Com.Text = "삭제"
        Me.Delete_Com.UseVisualStyleBackColor = False
        '
        'Insert_Com
        '
        Me.Insert_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Insert_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Insert_Com.Location = New System.Drawing.Point(247, 3)
        Me.Insert_Com.Name = "Insert_Com"
        Me.Insert_Com.Size = New System.Drawing.Size(98, 54)
        Me.Insert_Com.TabIndex = 15
        Me.Insert_Com.Text = "추가"
        Me.Insert_Com.UseVisualStyleBackColor = False
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(432, 82)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowTemplate.Height = 23
        Me.DataGridView1.Size = New System.Drawing.Size(806, 535)
        Me.DataGridView1.TabIndex = 45
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(595, 44)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(14, 12)
        Me.Label1.TabIndex = 106
        Me.Label1.Text = "~"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'E_Time
        '
        Me.E_Time.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.E_Time.Location = New System.Drawing.Point(718, 38)
        Me.E_Time.Name = "E_Time"
        Me.E_Time.Size = New System.Drawing.Size(54, 25)
        Me.E_Time.TabIndex = 105
        Me.E_Time.Text = "00:12"
        Me.E_Time.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'E_Day
        '
        Me.E_Day.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.E_Day.Location = New System.Drawing.Point(614, 38)
        Me.E_Day.Name = "E_Day"
        Me.E_Day.Size = New System.Drawing.Size(98, 25)
        Me.E_Day.TabIndex = 104
        Me.E_Day.Text = "2019-03-01"
        Me.E_Day.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'S_Time
        '
        Me.S_Time.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.S_Time.Location = New System.Drawing.Point(536, 38)
        Me.S_Time.Name = "S_Time"
        Me.S_Time.Size = New System.Drawing.Size(54, 25)
        Me.S_Time.TabIndex = 103
        Me.S_Time.Text = "00:12"
        Me.S_Time.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'S_Day
        '
        Me.S_Day.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.S_Day.Location = New System.Drawing.Point(432, 38)
        Me.S_Day.Name = "S_Day"
        Me.S_Day.Size = New System.Drawing.Size(98, 25)
        Me.S_Day.TabIndex = 102
        Me.S_Day.Text = "2019-03-01"
        Me.S_Day.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button2.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.Black
        Me.Button2.Location = New System.Drawing.Point(432, 5)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(340, 27)
        Me.Button2.TabIndex = 101
        Me.Button2.Text = "검색기간"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Frm_Temper
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.Process_Line_Grid)
        Me.Controls.Add(Me.Di_Panel1)
        Me.Controls.Add(Me.Panel_Menu)
        Me.Name = "Frm_Temper"
        Me.Size = New System.Drawing.Size(1258, 631)
        CType(Me.Process_Line_Grid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel_Menu.ResumeLayout(False)
        Me.Panel_Menu.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Process_Line_Grid As DataGridView
    Friend WithEvents Save_Com As Button
    Friend WithEvents Search_Com As Button
    Friend WithEvents Search_Text As TextBox
    Friend WithEvents Search_Combo As ComboBox
    Friend WithEvents Di_Panel1 As Panel
    Friend WithEvents Panel_Menu As Panel
    Friend WithEvents Delete_Com As Button
    Friend WithEvents Insert_Com As Button
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Label1 As Label
    Friend WithEvents E_Time As TextBox
    Friend WithEvents E_Day As TextBox
    Friend WithEvents S_Time As TextBox
    Friend WithEvents S_Day As TextBox
    Friend WithEvents Button2 As Button
End Class
