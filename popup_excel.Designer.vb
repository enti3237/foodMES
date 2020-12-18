<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class popup_excel
    Inherits System.Windows.Forms.Form

    'Form은 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
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
        Me.Btn_exSave = New System.Windows.Forms.Button()
        Me.Pro_Info = New System.Windows.Forms.DataGridView()
        CType(Me.Pro_Info, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Btn_exSave
        '
        Me.Btn_exSave.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Btn_exSave.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Btn_exSave.Location = New System.Drawing.Point(12, 390)
        Me.Btn_exSave.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Btn_exSave.Name = "Btn_exSave"
        Me.Btn_exSave.Size = New System.Drawing.Size(889, 68)
        Me.Btn_exSave.TabIndex = 18
        Me.Btn_exSave.Text = "저장"
        Me.Btn_exSave.UseVisualStyleBackColor = False
        '
        'Pro_Info
        '
        Me.Pro_Info.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Pro_Info.Location = New System.Drawing.Point(12, 11)
        Me.Pro_Info.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Pro_Info.MultiSelect = False
        Me.Pro_Info.Name = "Pro_Info"
        Me.Pro_Info.ReadOnly = True
        Me.Pro_Info.RowHeadersWidth = 51
        Me.Pro_Info.RowTemplate.Height = 27
        Me.Pro_Info.Size = New System.Drawing.Size(889, 355)
        Me.Pro_Info.TabIndex = 17
        '
        'popup_excel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(920, 471)
        Me.Controls.Add(Me.Btn_exSave)
        Me.Controls.Add(Me.Pro_Info)
        Me.Name = "popup_excel"
        Me.Text = "제품 업로드"
        CType(Me.Pro_Info, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Btn_exSave As Button
    Friend WithEvents Pro_Info As DataGridView
End Class
