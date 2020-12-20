<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class PCOrder_Insert
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
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Txt_crDay = New System.Windows.Forms.TextBox()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Txt_crDate = New System.Windows.Forms.TextBox()
        Me.Txt_standard = New System.Windows.Forms.TextBox()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Txt_qty = New System.Windows.Forms.TextBox()
        Me.Txt_plName = New System.Windows.Forms.TextBox()
        Me.Txt_cmName = New System.Windows.Forms.TextBox()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.btnJobSearch = New System.Windows.Forms.Button()
        Me.Txt_psEnd = New System.Windows.Forms.TextBox()
        Me.Button12 = New System.Windows.Forms.Button()
        Me.Txt_psBigo = New System.Windows.Forms.TextBox()
        Me.Button13 = New System.Windows.Forms.Button()
        Me.Txt_psStart = New System.Windows.Forms.TextBox()
        Me.Button11 = New System.Windows.Forms.Button()
        Me.Txt_psCode = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Txt_poBigo = New System.Windows.Forms.TextBox()
        Me.Txt_poCode = New System.Windows.Forms.TextBox()
        Me.Txt_poDay = New System.Windows.Forms.TextBox()
        Me.Button10 = New System.Windows.Forms.Button()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.Button9 = New System.Windows.Forms.Button()
        Me.PO_Code = New System.Windows.Forms.Button()
        Me.Btn_Save = New System.Windows.Forms.Button()
        Me.inv_PPCode = New System.Windows.Forms.TextBox()
        Me.Button_invisible1 = New System.Windows.Forms.Button()
        Me.inv_CRCode = New System.Windows.Forms.TextBox()
        Me.Button_invisible2 = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(91, Byte), Integer), CType(CType(155, Byte), Integer), CType(CType(213, Byte), Integer))
        Me.Panel1.Controls.Add(Me.inv_CRCode)
        Me.Panel1.Controls.Add(Me.Button_invisible2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Txt_crDay)
        Me.Panel1.Controls.Add(Me.Button6)
        Me.Panel1.Controls.Add(Me.Txt_crDate)
        Me.Panel1.Controls.Add(Me.Txt_standard)
        Me.Panel1.Controls.Add(Me.Button5)
        Me.Panel1.Controls.Add(Me.Button3)
        Me.Panel1.Controls.Add(Me.Txt_qty)
        Me.Panel1.Controls.Add(Me.Txt_plName)
        Me.Panel1.Controls.Add(Me.Txt_cmName)
        Me.Panel1.Controls.Add(Me.Button4)
        Me.Panel1.Controls.Add(Me.Button2)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(872, 93)
        Me.Panel1.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(822, 81)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(42, 12)
        Me.Label1.TabIndex = 242
        Me.Label1.Text = "Label1"
        Me.Label1.Visible = False
        '
        'Txt_crDay
        '
        Me.Txt_crDay.Enabled = False
        Me.Txt_crDay.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Txt_crDay.Location = New System.Drawing.Point(714, 40)
        Me.Txt_crDay.Name = "Txt_crDay"
        Me.Txt_crDay.ReadOnly = True
        Me.Txt_crDay.Size = New System.Drawing.Size(150, 25)
        Me.Txt_crDay.TabIndex = 246
        '
        'Button6
        '
        Me.Button6.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button6.Enabled = False
        Me.Button6.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button6.ForeColor = System.Drawing.Color.Black
        Me.Button6.Location = New System.Drawing.Point(582, 40)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(128, 27)
        Me.Button6.TabIndex = 245
        Me.Button6.Text = "납품예정일"
        Me.Button6.UseVisualStyleBackColor = False
        '
        'Txt_crDate
        '
        Me.Txt_crDate.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Txt_crDate.Location = New System.Drawing.Point(714, 7)
        Me.Txt_crDate.Name = "Txt_crDate"
        Me.Txt_crDate.Size = New System.Drawing.Size(150, 25)
        Me.Txt_crDate.TabIndex = 244
        '
        'Txt_standard
        '
        Me.Txt_standard.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Txt_standard.Location = New System.Drawing.Point(138, 40)
        Me.Txt_standard.Name = "Txt_standard"
        Me.Txt_standard.Size = New System.Drawing.Size(150, 25)
        Me.Txt_standard.TabIndex = 243
        '
        'Button5
        '
        Me.Button5.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button5.Enabled = False
        Me.Button5.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button5.ForeColor = System.Drawing.Color.Black
        Me.Button5.Location = New System.Drawing.Point(295, 40)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(128, 27)
        Me.Button5.TabIndex = 242
        Me.Button5.Text = "수주수량"
        Me.Button5.UseVisualStyleBackColor = False
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button3.Enabled = False
        Me.Button3.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button3.ForeColor = System.Drawing.Color.Black
        Me.Button3.Location = New System.Drawing.Point(582, 7)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(128, 27)
        Me.Button3.TabIndex = 241
        Me.Button3.Text = "수주일자"
        Me.Button3.UseVisualStyleBackColor = False
        '
        'Txt_qty
        '
        Me.Txt_qty.Enabled = False
        Me.Txt_qty.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Txt_qty.Location = New System.Drawing.Point(426, 40)
        Me.Txt_qty.Name = "Txt_qty"
        Me.Txt_qty.ReadOnly = True
        Me.Txt_qty.Size = New System.Drawing.Size(150, 25)
        Me.Txt_qty.TabIndex = 240
        '
        'Txt_plName
        '
        Me.Txt_plName.Enabled = False
        Me.Txt_plName.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Txt_plName.Location = New System.Drawing.Point(426, 7)
        Me.Txt_plName.Name = "Txt_plName"
        Me.Txt_plName.Size = New System.Drawing.Size(150, 25)
        Me.Txt_plName.TabIndex = 239
        '
        'Txt_cmName
        '
        Me.Txt_cmName.Enabled = False
        Me.Txt_cmName.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Txt_cmName.Location = New System.Drawing.Point(138, 7)
        Me.Txt_cmName.Name = "Txt_cmName"
        Me.Txt_cmName.ReadOnly = True
        Me.Txt_cmName.Size = New System.Drawing.Size(150, 25)
        Me.Txt_cmName.TabIndex = 238
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button4.Enabled = False
        Me.Button4.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button4.ForeColor = System.Drawing.Color.Black
        Me.Button4.Location = New System.Drawing.Point(8, 40)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(128, 27)
        Me.Button4.TabIndex = 237
        Me.Button4.Text = "제품규격"
        Me.Button4.UseVisualStyleBackColor = False
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button2.Enabled = False
        Me.Button2.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.Black
        Me.Button2.Location = New System.Drawing.Point(295, 7)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(128, 27)
        Me.Button2.TabIndex = 236
        Me.Button2.Text = "제품명"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button1.Enabled = False
        Me.Button1.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.Black
        Me.Button1.Location = New System.Drawing.Point(8, 7)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(128, 27)
        Me.Button1.TabIndex = 235
        Me.Button1.Text = "거래처명"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.White
        Me.Panel2.Controls.Add(Me.Button7)
        Me.Panel2.Controls.Add(Me.btnJobSearch)
        Me.Panel2.Controls.Add(Me.Txt_psEnd)
        Me.Panel2.Controls.Add(Me.Button12)
        Me.Panel2.Controls.Add(Me.Txt_psBigo)
        Me.Panel2.Controls.Add(Me.Button13)
        Me.Panel2.Controls.Add(Me.Txt_psStart)
        Me.Panel2.Controls.Add(Me.Button11)
        Me.Panel2.Controls.Add(Me.Txt_psCode)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.Txt_poBigo)
        Me.Panel2.Controls.Add(Me.Txt_poCode)
        Me.Panel2.Controls.Add(Me.Txt_poDay)
        Me.Panel2.Controls.Add(Me.Button10)
        Me.Panel2.Controls.Add(Me.Button8)
        Me.Panel2.Controls.Add(Me.Button9)
        Me.Panel2.Controls.Add(Me.PO_Code)
        Me.Panel2.Controls.Add(Me.Btn_Save)
        Me.Panel2.Location = New System.Drawing.Point(12, 111)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(872, 167)
        Me.Panel2.TabIndex = 4
        '
        'Button7
        '
        Me.Button7.BackColor = System.Drawing.Color.White
        Me.Button7.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button7.Font = New System.Drawing.Font("맑은 고딕", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button7.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.Button7.Location = New System.Drawing.Point(545, 49)
        Me.Button7.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(31, 29)
        Me.Button7.TabIndex = 246
        Me.Button7.Text = "⌕"
        Me.Button7.UseVisualStyleBackColor = False
        '
        'btnJobSearch
        '
        Me.btnJobSearch.BackColor = System.Drawing.Color.White
        Me.btnJobSearch.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnJobSearch.Font = New System.Drawing.Font("맑은 고딕", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.btnJobSearch.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.btnJobSearch.Location = New System.Drawing.Point(832, 50)
        Me.btnJobSearch.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnJobSearch.Name = "btnJobSearch"
        Me.btnJobSearch.Size = New System.Drawing.Size(32, 26)
        Me.btnJobSearch.TabIndex = 245
        Me.btnJobSearch.Text = "⌕"
        Me.btnJobSearch.UseVisualStyleBackColor = False
        '
        'Txt_psEnd
        '
        Me.Txt_psEnd.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Txt_psEnd.Location = New System.Drawing.Point(714, 49)
        Me.Txt_psEnd.Name = "Txt_psEnd"
        Me.Txt_psEnd.Size = New System.Drawing.Size(150, 25)
        Me.Txt_psEnd.TabIndex = 253
        '
        'Button12
        '
        Me.Button12.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button12.Enabled = False
        Me.Button12.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button12.ForeColor = System.Drawing.Color.Black
        Me.Button12.Location = New System.Drawing.Point(582, 49)
        Me.Button12.Name = "Button12"
        Me.Button12.Size = New System.Drawing.Size(128, 27)
        Me.Button12.TabIndex = 252
        Me.Button12.Text = "생산종료예정일"
        Me.Button12.UseVisualStyleBackColor = False
        '
        'Txt_psBigo
        '
        Me.Txt_psBigo.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Txt_psBigo.Location = New System.Drawing.Point(138, 84)
        Me.Txt_psBigo.Name = "Txt_psBigo"
        Me.Txt_psBigo.Size = New System.Drawing.Size(726, 25)
        Me.Txt_psBigo.TabIndex = 251
        '
        'Button13
        '
        Me.Button13.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button13.Enabled = False
        Me.Button13.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button13.ForeColor = System.Drawing.Color.Black
        Me.Button13.Location = New System.Drawing.Point(8, 84)
        Me.Button13.Name = "Button13"
        Me.Button13.Size = New System.Drawing.Size(128, 27)
        Me.Button13.TabIndex = 250
        Me.Button13.Text = "생산실적비고"
        Me.Button13.UseVisualStyleBackColor = False
        '
        'Txt_psStart
        '
        Me.Txt_psStart.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Txt_psStart.Location = New System.Drawing.Point(426, 49)
        Me.Txt_psStart.Name = "Txt_psStart"
        Me.Txt_psStart.Size = New System.Drawing.Size(150, 25)
        Me.Txt_psStart.TabIndex = 249
        '
        'Button11
        '
        Me.Button11.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button11.Enabled = False
        Me.Button11.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button11.ForeColor = System.Drawing.Color.Black
        Me.Button11.Location = New System.Drawing.Point(295, 49)
        Me.Button11.Name = "Button11"
        Me.Button11.Size = New System.Drawing.Size(128, 27)
        Me.Button11.TabIndex = 248
        Me.Button11.Text = "생산시작예정일"
        Me.Button11.UseVisualStyleBackColor = False
        '
        'Txt_psCode
        '
        Me.Txt_psCode.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Txt_psCode.Location = New System.Drawing.Point(138, 49)
        Me.Txt_psCode.Name = "Txt_psCode"
        Me.Txt_psCode.Size = New System.Drawing.Size(150, 25)
        Me.Txt_psCode.TabIndex = 247
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(813, 155)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(42, 12)
        Me.Label2.TabIndex = 241
        Me.Label2.Text = "Label2"
        Me.Label2.Visible = False
        '
        'Txt_poBigo
        '
        Me.Txt_poBigo.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Txt_poBigo.Location = New System.Drawing.Point(714, 14)
        Me.Txt_poBigo.Name = "Txt_poBigo"
        Me.Txt_poBigo.Size = New System.Drawing.Size(150, 25)
        Me.Txt_poBigo.TabIndex = 244
        '
        'Txt_poCode
        '
        Me.Txt_poCode.Enabled = False
        Me.Txt_poCode.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Txt_poCode.Location = New System.Drawing.Point(138, 14)
        Me.Txt_poCode.Name = "Txt_poCode"
        Me.Txt_poCode.ReadOnly = True
        Me.Txt_poCode.Size = New System.Drawing.Size(150, 25)
        Me.Txt_poCode.TabIndex = 243
        '
        'Txt_poDay
        '
        Me.Txt_poDay.Enabled = False
        Me.Txt_poDay.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Txt_poDay.Location = New System.Drawing.Point(426, 14)
        Me.Txt_poDay.Name = "Txt_poDay"
        Me.Txt_poDay.ReadOnly = True
        Me.Txt_poDay.Size = New System.Drawing.Size(150, 25)
        Me.Txt_poDay.TabIndex = 242
        '
        'Button10
        '
        Me.Button10.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button10.Enabled = False
        Me.Button10.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button10.ForeColor = System.Drawing.Color.Black
        Me.Button10.Location = New System.Drawing.Point(8, 49)
        Me.Button10.Name = "Button10"
        Me.Button10.Size = New System.Drawing.Size(128, 27)
        Me.Button10.TabIndex = 239
        Me.Button10.Text = "생산실적코드"
        Me.Button10.UseVisualStyleBackColor = False
        '
        'Button8
        '
        Me.Button8.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button8.Enabled = False
        Me.Button8.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button8.ForeColor = System.Drawing.Color.Black
        Me.Button8.Location = New System.Drawing.Point(295, 14)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(128, 27)
        Me.Button8.TabIndex = 238
        Me.Button8.Text = "생산지시일자"
        Me.Button8.UseVisualStyleBackColor = False
        '
        'Button9
        '
        Me.Button9.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button9.Enabled = False
        Me.Button9.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button9.ForeColor = System.Drawing.Color.Black
        Me.Button9.Location = New System.Drawing.Point(582, 14)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(128, 27)
        Me.Button9.TabIndex = 237
        Me.Button9.Text = "작업지시비고"
        Me.Button9.UseVisualStyleBackColor = False
        '
        'PO_Code
        '
        Me.PO_Code.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.PO_Code.Enabled = False
        Me.PO_Code.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.PO_Code.ForeColor = System.Drawing.Color.Black
        Me.PO_Code.Location = New System.Drawing.Point(8, 14)
        Me.PO_Code.Name = "PO_Code"
        Me.PO_Code.Size = New System.Drawing.Size(128, 27)
        Me.PO_Code.TabIndex = 236
        Me.PO_Code.Text = "작업지시서코드"
        Me.PO_Code.UseVisualStyleBackColor = False
        '
        'Btn_Save
        '
        Me.Btn_Save.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Btn_Save.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Btn_Save.ForeColor = System.Drawing.Color.Black
        Me.Btn_Save.Location = New System.Drawing.Point(8, 119)
        Me.Btn_Save.Name = "Btn_Save"
        Me.Btn_Save.Size = New System.Drawing.Size(856, 44)
        Me.Btn_Save.TabIndex = 233
        Me.Btn_Save.Text = "저장"
        Me.Btn_Save.UseVisualStyleBackColor = False
        '
        'inv_PPCode
        '
        Me.inv_PPCode.Enabled = False
        Me.inv_PPCode.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.inv_PPCode.Location = New System.Drawing.Point(674, 85)
        Me.inv_PPCode.Name = "inv_PPCode"
        Me.inv_PPCode.ReadOnly = True
        Me.inv_PPCode.Size = New System.Drawing.Size(150, 25)
        Me.inv_PPCode.TabIndex = 248
        Me.inv_PPCode.Visible = False
        '
        'Button_invisible1
        '
        Me.Button_invisible1.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button_invisible1.Enabled = False
        Me.Button_invisible1.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button_invisible1.ForeColor = System.Drawing.Color.Black
        Me.Button_invisible1.Location = New System.Drawing.Point(544, 85)
        Me.Button_invisible1.Name = "Button_invisible1"
        Me.Button_invisible1.Size = New System.Drawing.Size(128, 27)
        Me.Button_invisible1.TabIndex = 247
        Me.Button_invisible1.Text = "생산계획코드"
        Me.Button_invisible1.UseVisualStyleBackColor = False
        Me.Button_invisible1.Visible = False
        '
        'inv_CRCode
        '
        Me.inv_CRCode.Enabled = False
        Me.inv_CRCode.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.inv_CRCode.Location = New System.Drawing.Point(376, 73)
        Me.inv_CRCode.Name = "inv_CRCode"
        Me.inv_CRCode.ReadOnly = True
        Me.inv_CRCode.Size = New System.Drawing.Size(150, 25)
        Me.inv_CRCode.TabIndex = 250
        Me.inv_CRCode.Visible = False
        '
        'Button_invisible2
        '
        Me.Button_invisible2.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(222, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.Button_invisible2.Enabled = False
        Me.Button_invisible2.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button_invisible2.ForeColor = System.Drawing.Color.Black
        Me.Button_invisible2.Location = New System.Drawing.Point(246, 73)
        Me.Button_invisible2.Name = "Button_invisible2"
        Me.Button_invisible2.Size = New System.Drawing.Size(128, 27)
        Me.Button_invisible2.TabIndex = 249
        Me.Button_invisible2.Text = "수주코드"
        Me.Button_invisible2.UseVisualStyleBackColor = False
        Me.Button_invisible2.Visible = False
        '
        'PCOrder_Insert
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(894, 291)
        Me.Controls.Add(Me.inv_PPCode)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Button_invisible1)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "PCOrder_Insert"
        Me.Text = "작업지시서 추가"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Panel1 As Panel
    Friend WithEvents Txt_crDate As TextBox
    Friend WithEvents Txt_standard As TextBox
    Friend WithEvents Button5 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Txt_qty As TextBox
    Friend WithEvents Txt_plName As TextBox
    Friend WithEvents Txt_cmName As TextBox
    Friend WithEvents Button4 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents Txt_crDay As TextBox
    Friend WithEvents Button6 As Button
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Txt_psBigo As TextBox
    Friend WithEvents Button13 As Button
    Friend WithEvents Txt_psStart As TextBox
    Friend WithEvents Button11 As Button
    Friend WithEvents Txt_psCode As TextBox
    Friend WithEvents Button7 As Button
    Friend WithEvents btnJobSearch As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents Txt_poBigo As TextBox
    Friend WithEvents Txt_poCode As TextBox
    Friend WithEvents Txt_poDay As TextBox
    Friend WithEvents Button10 As Button
    Friend WithEvents Button8 As Button
    Friend WithEvents Button9 As Button
    Friend WithEvents PO_Code As Button
    Friend WithEvents Btn_Save As Button
    Friend WithEvents Txt_psEnd As TextBox
    Friend WithEvents Button12 As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents inv_CRCode As TextBox
    Friend WithEvents Button_invisible2 As Button
    Friend WithEvents inv_PPCode As TextBox
    Friend WithEvents Button_invisible1 As Button
End Class
