<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_Collection
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
        Dim ChartArea1 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea()
        Dim Legend1 As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend()
        Dim Series1 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Me.Panel_Menu = New System.Windows.Forms.Panel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Day2 = New System.Windows.Forms.Label()
        Me.Day1 = New System.Windows.Forms.Label()
        Me.RadioButton3 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.DateTimePicker2 = New System.Windows.Forms.DateTimePicker()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.Search_Com = New System.Windows.Forms.Button()
        Me.Di_Panel1 = New System.Windows.Forms.Panel()
        Me.Item_Grid = New System.Windows.Forms.DataGridView()
        Me.Customer_Grid = New System.Windows.Forms.DataGridView()
        Me.Chart1 = New System.Windows.Forms.DataVisualization.Charting.Chart()
        Me.Panel_Menu.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.Item_Grid, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Customer_Grid, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel_Menu
        '
        Me.Panel_Menu.Controls.Add(Me.Panel1)
        Me.Panel_Menu.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel_Menu.Location = New System.Drawing.Point(0, 2)
        Me.Panel_Menu.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Panel_Menu.Name = "Panel_Menu"
        Me.Panel_Menu.Size = New System.Drawing.Size(1522, 100)
        Me.Panel_Menu.TabIndex = 65
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Day2)
        Me.Panel1.Controls.Add(Me.Day1)
        Me.Panel1.Controls.Add(Me.RadioButton3)
        Me.Panel1.Controls.Add(Me.RadioButton2)
        Me.Panel1.Controls.Add(Me.RadioButton1)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.DateTimePicker2)
        Me.Panel1.Controls.Add(Me.DateTimePicker1)
        Me.Panel1.Controls.Add(Me.Search_Com)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1522, 78)
        Me.Panel1.TabIndex = 2
        '
        'Day2
        '
        Me.Day2.AutoSize = True
        Me.Day2.Location = New System.Drawing.Point(866, 39)
        Me.Day2.Name = "Day2"
        Me.Day2.Size = New System.Drawing.Size(50, 15)
        Me.Day2.TabIndex = 66
        Me.Day2.Text = "Label3"
        Me.Day2.Visible = False
        '
        'Day1
        '
        Me.Day1.AutoSize = True
        Me.Day1.Location = New System.Drawing.Point(709, 41)
        Me.Day1.Name = "Day1"
        Me.Day1.Size = New System.Drawing.Size(50, 15)
        Me.Day1.TabIndex = 65
        Me.Day1.Text = "Label2"
        Me.Day1.Visible = False
        '
        'RadioButton3
        '
        Me.RadioButton3.AutoSize = True
        Me.RadioButton3.Checked = True
        Me.RadioButton3.Location = New System.Drawing.Point(475, 38)
        Me.RadioButton3.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(58, 19)
        Me.RadioButton3.TabIndex = 62
        Me.RadioButton3.TabStop = True
        Me.RadioButton3.Text = "일간"
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(414, 38)
        Me.RadioButton2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(58, 19)
        Me.RadioButton2.TabIndex = 63
        Me.RadioButton2.Text = "월간"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Location = New System.Drawing.Point(354, 38)
        Me.RadioButton1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(58, 19)
        Me.RadioButton1.TabIndex = 64
        Me.RadioButton1.Text = "연간"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(174, 41)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(18, 15)
        Me.Label1.TabIndex = 61
        Me.Label1.Text = "~"
        '
        'DateTimePicker2
        '
        Me.DateTimePicker2.CustomFormat = "yyyy년 MM월 dd일"
        Me.DateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker2.Location = New System.Drawing.Point(194, 34)
        Me.DateTimePicker2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.DateTimePicker2.Name = "DateTimePicker2"
        Me.DateTimePicker2.Size = New System.Drawing.Size(153, 25)
        Me.DateTimePicker2.TabIndex = 59
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.CustomFormat = "yyyy년 MM월 dd일"
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker1.Location = New System.Drawing.Point(14, 34)
        Me.DateTimePicker1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(153, 25)
        Me.DateTimePicker1.TabIndex = 60
        '
        'Search_Com
        '
        Me.Search_Com.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Search_Com.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Search_Com.ForeColor = System.Drawing.Color.Black
        Me.Search_Com.Location = New System.Drawing.Point(536, 31)
        Me.Search_Com.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Search_Com.Name = "Search_Com"
        Me.Search_Com.Size = New System.Drawing.Size(123, 29)
        Me.Search_Com.TabIndex = 58
        Me.Search_Com.Text = "검색"
        Me.Search_Com.UseVisualStyleBackColor = False
        '
        'Di_Panel1
        '
        Me.Di_Panel1.BackColor = System.Drawing.Color.CornflowerBlue
        Me.Di_Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Di_Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Di_Panel1.Margin = New System.Windows.Forms.Padding(0)
        Me.Di_Panel1.Name = "Di_Panel1"
        Me.Di_Panel1.Size = New System.Drawing.Size(1522, 2)
        Me.Di_Panel1.TabIndex = 66
        '
        'Item_Grid
        '
        Me.Item_Grid.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Item_Grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Item_Grid.Location = New System.Drawing.Point(14, 128)
        Me.Item_Grid.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Item_Grid.Name = "Item_Grid"
        Me.Item_Grid.RowTemplate.Height = 23
        Me.Item_Grid.Size = New System.Drawing.Size(766, 642)
        Me.Item_Grid.TabIndex = 67
        '
        'Customer_Grid
        '
        Me.Customer_Grid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Customer_Grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Customer_Grid.Location = New System.Drawing.Point(14, 128)
        Me.Customer_Grid.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Customer_Grid.Name = "Customer_Grid"
        Me.Customer_Grid.RowTemplate.Height = 23
        Me.Customer_Grid.Size = New System.Drawing.Size(766, 414)
        Me.Customer_Grid.TabIndex = 68
        '
        'Chart1
        '
        Me.Chart1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        ChartArea1.Name = "ChartArea1"
        Me.Chart1.ChartAreas.Add(ChartArea1)
        Legend1.Name = "Legend1"
        Me.Chart1.Legends.Add(Legend1)
        Me.Chart1.Location = New System.Drawing.Point(798, 128)
        Me.Chart1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Chart1.Name = "Chart1"
        Series1.ChartArea = "ChartArea1"
        Series1.Legend = "Legend1"
        Series1.Name = "Series1"
        Me.Chart1.Series.Add(Series1)
        Me.Chart1.Size = New System.Drawing.Size(711, 621)
        Me.Chart1.TabIndex = 69
        Me.Chart1.Text = "  "
        '
        'Frm_Collection
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Chart1)
        Me.Controls.Add(Me.Customer_Grid)
        Me.Controls.Add(Me.Item_Grid)
        Me.Controls.Add(Me.Panel_Menu)
        Me.Controls.Add(Me.Di_Panel1)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "Frm_Collection"
        Me.Size = New System.Drawing.Size(1522, 806)
        Me.Panel_Menu.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.Item_Grid, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Customer_Grid, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel_Menu As Panel
    Friend WithEvents Di_Panel1 As Panel
    Friend WithEvents Item_Grid As DataGridView
    Friend WithEvents Customer_Grid As DataGridView
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Day2 As Label
    Friend WithEvents Day1 As Label
    Friend WithEvents RadioButton3 As RadioButton
    Friend WithEvents RadioButton2 As RadioButton
    Friend WithEvents RadioButton1 As RadioButton
    Friend WithEvents Label1 As Label
    Friend WithEvents DateTimePicker2 As DateTimePicker
    Friend WithEvents DateTimePicker1 As DateTimePicker
    Friend WithEvents Search_Com As Button
    Friend WithEvents Chart1 As DataVisualization.Charting.Chart
End Class
