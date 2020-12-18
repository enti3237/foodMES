<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Frm_graph
    Inherits System.Windows.Forms.Form

    'Form은 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
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
        Dim ChartArea1 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea()
        Dim Legend1 As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend()
        Dim Series1 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Dim Series2 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Dim Series3 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Dim Series4 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Dim Series5 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Dim Series6 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Me.Picture_Logo = New System.Windows.Forms.PictureBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Chart1 = New System.Windows.Forms.DataVisualization.Charting.Chart()
        CType(Me.Picture_Logo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Picture_Logo
        '
        Me.Picture_Logo.BackColor = System.Drawing.Color.White
        Me.Picture_Logo.Location = New System.Drawing.Point(22, 12)
        Me.Picture_Logo.Name = "Picture_Logo"
        Me.Picture_Logo.Size = New System.Drawing.Size(384, 111)
        Me.Picture_Logo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.Picture_Logo.TabIndex = 1
        Me.Picture_Logo.TabStop = False
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 600000
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("맑은 고딕", 72.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.DarkOliveGreen
        Me.Label1.Location = New System.Drawing.Point(500, 27)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(880, 128)
        Me.Label1.TabIndex = 191
        Me.Label1.Text = "CCP 공정 모니터링"
        '
        'Chart1
        '
        ChartArea1.Name = "ChartArea1"
        Me.Chart1.ChartAreas.Add(ChartArea1)
        Legend1.Font = New System.Drawing.Font("맑은 고딕", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Legend1.IsTextAutoFit = False
        Legend1.Name = "Legend1"
        Me.Chart1.Legends.Add(Legend1)
        Me.Chart1.Location = New System.Drawing.Point(12, 194)
        Me.Chart1.Name = "Chart1"
        Series1.BorderWidth = 20
        Series1.ChartArea = "ChartArea1"
        Series1.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline
        Series1.Font = New System.Drawing.Font("맑은 고딕", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Series1.Legend = "Legend1"
        Series1.Name = "Series1"
        Series2.BorderWidth = 20
        Series2.ChartArea = "ChartArea1"
        Series2.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline
        Series2.Font = New System.Drawing.Font("맑은 고딕", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Series2.Legend = "Legend1"
        Series2.Name = "Series2"
        Series3.BorderWidth = 20
        Series3.ChartArea = "ChartArea1"
        Series3.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline
        Series3.Font = New System.Drawing.Font("맑은 고딕", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Series3.Legend = "Legend1"
        Series3.Name = "Series3"
        Series4.BorderWidth = 20
        Series4.ChartArea = "ChartArea1"
        Series4.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline
        Series4.Font = New System.Drawing.Font("맑은 고딕", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Series4.Legend = "Legend1"
        Series4.Name = "Series4"
        Series5.BorderWidth = 20
        Series5.ChartArea = "ChartArea1"
        Series5.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline
        Series5.Font = New System.Drawing.Font("맑은 고딕", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Series5.Legend = "Legend1"
        Series5.Name = "Series5"
        Series6.BorderWidth = 20
        Series6.ChartArea = "ChartArea1"
        Series6.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline
        Series6.Font = New System.Drawing.Font("맑은 고딕", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Series6.Legend = "Legend1"
        Series6.Name = "Series6"
        Me.Chart1.Series.Add(Series1)
        Me.Chart1.Series.Add(Series2)
        Me.Chart1.Series.Add(Series3)
        Me.Chart1.Series.Add(Series4)
        Me.Chart1.Series.Add(Series5)
        Me.Chart1.Series.Add(Series6)
        Me.Chart1.Size = New System.Drawing.Size(1916, 894)
        Me.Chart1.TabIndex = 192
        Me.Chart1.Text = "Chart1"
        '
        'Frm_graph
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1940, 1100)
        Me.Controls.Add(Me.Chart1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Picture_Logo)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Frm_graph"
        Me.Text = "Frm_Monitoring"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.Picture_Logo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Picture_Logo As PictureBox
    Friend WithEvents Timer1 As Timer
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Panel1 As Panel
    Friend WithEvents PictureBox3 As PictureBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label97 As Label
    Friend WithEvents Label98 As Label
    Friend WithEvents Label99 As Label
    Friend WithEvents Label100 As Label
    Friend WithEvents Label101 As Label
    Friend WithEvents Label102 As Label
    Friend WithEvents Label103 As Label
    Friend WithEvents Label104 As Label
    Friend WithEvents Label105 As Label
    Friend WithEvents Label106 As Label
    Friend WithEvents Label107 As Label
    Friend WithEvents Label108 As Label
    Friend WithEvents Label109 As Label
    Friend WithEvents Label110 As Label
    Friend WithEvents Label111 As Label
    Friend WithEvents Label112 As Label
    Friend WithEvents Label113 As Label
    Friend WithEvents Label114 As Label
    Friend WithEvents Label115 As Label
    Friend WithEvents Label116 As Label
    Friend WithEvents Label117 As Label
    Friend WithEvents Label118 As Label
    Friend WithEvents Label119 As Label
    Friend WithEvents Label120 As Label
    Friend WithEvents Label121 As Label
    Friend WithEvents Label122 As Label
    Friend WithEvents Label123 As Label
    Friend WithEvents Label124 As Label
    Friend WithEvents Label125 As Label
    Friend WithEvents Label126 As Label
    Friend WithEvents Label127 As Label
    Friend WithEvents Label128 As Label
    Friend WithEvents Label129 As Label
    Friend WithEvents Label130 As Label
    Friend WithEvents Label131 As Label
    Friend WithEvents Label132 As Label
    Friend WithEvents Label133 As Label
    Friend WithEvents Label134 As Label
    Friend WithEvents Label135 As Label
    Friend WithEvents Label136 As Label
    Friend WithEvents Label137 As Label
    Friend WithEvents Label138 As Label
    Friend WithEvents Label139 As Label
    Friend WithEvents Label140 As Label
    Friend WithEvents Label141 As Label
    Friend WithEvents Label142 As Label
    Friend WithEvents Label143 As Label
    Friend WithEvents Label144 As Label
    Friend WithEvents Label145 As Label
    Friend WithEvents Label146 As Label
    Friend WithEvents Label147 As Label
    Friend WithEvents Label148 As Label
    Friend WithEvents Label149 As Label
    Friend WithEvents Label150 As Label
    Friend WithEvents Label151 As Label
    Friend WithEvents Label152 As Label
    Friend WithEvents Label153 As Label
    Friend WithEvents Label154 As Label
    Friend WithEvents Label155 As Label
    Friend WithEvents Label156 As Label
    Friend WithEvents Label157 As Label
    Friend WithEvents Label158 As Label
    Friend WithEvents Label159 As Label
    Friend WithEvents Label160 As Label
    Friend WithEvents Label161 As Label
    Friend WithEvents Label162 As Label
    Friend WithEvents Label163 As Label
    Friend WithEvents Label164 As Label
    Friend WithEvents Label165 As Label
    Friend WithEvents Label166 As Label
    Friend WithEvents Label167 As Label
    Friend WithEvents Label168 As Label
    Friend WithEvents Label169 As Label
    Friend WithEvents Label170 As Label
    Friend WithEvents Label171 As Label
    Friend WithEvents Label172 As Label
    Friend WithEvents Label173 As Label
    Friend WithEvents Label174 As Label
    Friend WithEvents Label175 As Label
    Friend WithEvents Label176 As Label
    Friend WithEvents Label177 As Label
    Friend WithEvents Label178 As Label
    Friend WithEvents Label179 As Label
    Friend WithEvents Label180 As Label
    Friend WithEvents Label181 As Label
    Friend WithEvents Label182 As Label
    Friend WithEvents Label183 As Label
    Friend WithEvents Label184 As Label
    Friend WithEvents Label185 As Label
    Friend WithEvents Label186 As Label
    Friend WithEvents Label187 As Label
    Friend WithEvents Label188 As Label
    Friend WithEvents Label189 As Label
    Friend WithEvents Label190 As Label
    Friend WithEvents Label191 As Label
    Friend WithEvents Label192 As Label
    Friend WithEvents Label193 As Label
    Friend WithEvents Label194 As Label
    Friend WithEvents Label195 As Label
    Friend WithEvents Label196 As Label
    Friend WithEvents Label197 As Label
    Friend WithEvents Label198 As Label
    Friend WithEvents Label199 As Label
    Friend WithEvents Label200 As Label
    Friend WithEvents Label201 As Label
    Friend WithEvents Label202 As Label
    Friend WithEvents Label203 As Label
    Friend WithEvents Label204 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents Chart1 As DataVisualization.Charting.Chart
End Class
