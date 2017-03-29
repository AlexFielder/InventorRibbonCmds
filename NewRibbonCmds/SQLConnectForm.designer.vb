<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SQLConnectForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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
    Private label1 As System.Windows.Forms.Label
    'Private cmdExecute As System.Windows.Forms.Button
    'Required by the Windows Form Designer
    Friend WithEvents cmdExecute As System.Windows.Forms.Button
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton3 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton4 As System.Windows.Forms.RadioButton
    'Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents SqlCommand1 As System.Data.SqlClient.SqlCommand
    Friend WithEvents DataSet1 As System.Data.DataSet
    Friend WithEvents srctbl As String
    'Private Sub InitializeComponent()
    '    components = New System.ComponentModel.Container
    '    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    '    Me.Text = "Form2"
    'End Sub

    Public Sub New()
        InitializeComponent()
    End Sub
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SQLConnectForm))
        Me.label1 = New System.Windows.Forms.Label()
        Me.cmdExecute = New System.Windows.Forms.Button()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.RadioButton3 = New System.Windows.Forms.RadioButton()
        Me.RadioButton4 = New System.Windows.Forms.RadioButton()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.BindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
        Me.Label25 = New System.Windows.Forms.Label()
        Me.tbUsed10 = New System.Windows.Forms.TextBox()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.tbUsed09 = New System.Windows.Forms.TextBox()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.tbUsed08 = New System.Windows.Forms.TextBox()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.tbUsed07 = New System.Windows.Forms.TextBox()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.tbUsed06 = New System.Windows.Forms.TextBox()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.tbUsed05 = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.tbUsed04 = New System.Windows.Forms.TextBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.tbUsed03 = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.tbUsed02 = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.tbUsed01 = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.tbProFin = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.tbTol3 = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.tbTol2 = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.tbTol1 = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.tbSurface = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.tbSubCon = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.tbScale = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.tbDimsIn = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.tbClassification = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.tbNumSheets = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.tbSheetNum = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.tbDate = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.tbIssue = New System.Windows.Forms.TextBox()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.tbTitle = New System.Windows.Forms.TextBox()
        Me.tbDwgNum = New System.Windows.Forms.TextBox()
        Me.tbChange = New System.Windows.Forms.TextBox()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'label1
        '
        Me.label1.Location = New System.Drawing.Point(0, 0)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(298, 31)
        Me.label1.TabIndex = 0
        Me.label1.Text = "Select a frame size to see the available default entries." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Then click Execute to " & _
            "insert a border:"
        '
        'cmdExecute
        '
        Me.cmdExecute.Location = New System.Drawing.Point(394, 31)
        Me.cmdExecute.Name = "cmdExecute"
        Me.cmdExecute.Size = New System.Drawing.Size(75, 23)
        Me.cmdExecute.TabIndex = 1
        Me.cmdExecute.Text = "Execute"
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Location = New System.Drawing.Point(11, 34)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(38, 17)
        Me.RadioButton1.TabIndex = 4
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "A3"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(107, 34)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(38, 17)
        Me.RadioButton2.TabIndex = 5
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "A2"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton3
        '
        Me.RadioButton3.AutoSize = True
        Me.RadioButton3.Location = New System.Drawing.Point(203, 34)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(38, 17)
        Me.RadioButton3.TabIndex = 6
        Me.RadioButton3.TabStop = True
        Me.RadioButton3.Text = "A1"
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'RadioButton4
        '
        Me.RadioButton4.AutoSize = True
        Me.RadioButton4.Location = New System.Drawing.Point(299, 34)
        Me.RadioButton4.Name = "RadioButton4"
        Me.RadioButton4.Size = New System.Drawing.Size(38, 17)
        Me.RadioButton4.TabIndex = 7
        Me.RadioButton4.TabStop = True
        Me.RadioButton4.Text = "A0"
        Me.RadioButton4.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.CheckBox1)
        Me.Panel1.Location = New System.Drawing.Point(10, 59)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(506, 40)
        Me.Panel1.TabIndex = 8
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(85, 4)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(214, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "<- Select whether this is an Assembly or not!"
        Me.Label2.Visible = False
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(3, 3)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(76, 17)
        Me.CheckBox1.TabIndex = 0
        Me.CheckBox1.Text = "Assembly?"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Timer1
        '
        Me.Timer1.Interval = 500
        '
        'Timer2
        '
        Me.Timer2.Interval = 5000
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(271, 425)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(67, 13)
        Me.Label25.TabIndex = 204
        Me.Label25.Text = "Used On 10:"
        '
        'tbUsed10
        '
        Me.tbUsed10.Location = New System.Drawing.Point(344, 422)
        Me.tbUsed10.Name = "tbUsed10"
        Me.tbUsed10.Size = New System.Drawing.Size(139, 20)
        Me.tbUsed10.TabIndex = 195
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(271, 399)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(67, 13)
        Me.Label26.TabIndex = 203
        Me.Label26.Text = "Used On 09:"
        '
        'tbUsed09
        '
        Me.tbUsed09.Location = New System.Drawing.Point(344, 396)
        Me.tbUsed09.Name = "tbUsed09"
        Me.tbUsed09.Size = New System.Drawing.Size(139, 20)
        Me.tbUsed09.TabIndex = 194
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(271, 373)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(67, 13)
        Me.Label23.TabIndex = 202
        Me.Label23.Text = "Used On 08:"
        '
        'tbUsed08
        '
        Me.tbUsed08.Location = New System.Drawing.Point(344, 370)
        Me.tbUsed08.Name = "tbUsed08"
        Me.tbUsed08.Size = New System.Drawing.Size(139, 20)
        Me.tbUsed08.TabIndex = 192
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(271, 347)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(67, 13)
        Me.Label24.TabIndex = 201
        Me.Label24.Text = "Used On 07:"
        '
        'tbUsed07
        '
        Me.tbUsed07.Location = New System.Drawing.Point(344, 344)
        Me.tbUsed07.Name = "tbUsed07"
        Me.tbUsed07.Size = New System.Drawing.Size(139, 20)
        Me.tbUsed07.TabIndex = 191
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(271, 321)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(67, 13)
        Me.Label21.TabIndex = 200
        Me.Label21.Text = "Used On 06:"
        '
        'tbUsed06
        '
        Me.tbUsed06.Location = New System.Drawing.Point(344, 318)
        Me.tbUsed06.Name = "tbUsed06"
        Me.tbUsed06.Size = New System.Drawing.Size(139, 20)
        Me.tbUsed06.TabIndex = 190
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(271, 295)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(67, 13)
        Me.Label22.TabIndex = 199
        Me.Label22.Text = "Used On 05:"
        '
        'tbUsed05
        '
        Me.tbUsed05.Location = New System.Drawing.Point(344, 292)
        Me.tbUsed05.Name = "tbUsed05"
        Me.tbUsed05.Size = New System.Drawing.Size(139, 20)
        Me.tbUsed05.TabIndex = 188
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(271, 269)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(67, 13)
        Me.Label19.TabIndex = 198
        Me.Label19.Text = "Used On 04:"
        '
        'tbUsed04
        '
        Me.tbUsed04.Location = New System.Drawing.Point(344, 266)
        Me.tbUsed04.Name = "tbUsed04"
        Me.tbUsed04.Size = New System.Drawing.Size(139, 20)
        Me.tbUsed04.TabIndex = 187
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(271, 243)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(67, 13)
        Me.Label20.TabIndex = 197
        Me.Label20.Text = "Used On 03:"
        '
        'tbUsed03
        '
        Me.tbUsed03.Location = New System.Drawing.Point(344, 240)
        Me.tbUsed03.Name = "tbUsed03"
        Me.tbUsed03.Size = New System.Drawing.Size(139, 20)
        Me.tbUsed03.TabIndex = 185
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(271, 217)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(67, 13)
        Me.Label18.TabIndex = 196
        Me.Label18.Text = "Used On 02:"
        '
        'tbUsed02
        '
        Me.tbUsed02.Location = New System.Drawing.Point(344, 214)
        Me.tbUsed02.Name = "tbUsed02"
        Me.tbUsed02.Size = New System.Drawing.Size(139, 20)
        Me.tbUsed02.TabIndex = 183
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(271, 191)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(67, 13)
        Me.Label17.TabIndex = 193
        Me.Label17.Text = "Used On 01:"
        '
        'tbUsed01
        '
        Me.tbUsed01.Location = New System.Drawing.Point(344, 188)
        Me.tbUsed01.Name = "tbUsed01"
        Me.tbUsed01.Size = New System.Drawing.Size(139, 20)
        Me.tbUsed01.TabIndex = 182
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(12, 425)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(88, 13)
        Me.Label16.TabIndex = 189
        Me.Label16.Text = "Protective Finish:"
        '
        'tbProFin
        '
        Me.tbProFin.Location = New System.Drawing.Point(107, 422)
        Me.tbProFin.Name = "tbProFin"
        Me.tbProFin.Size = New System.Drawing.Size(139, 20)
        Me.tbProFin.TabIndex = 176
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(270, 165)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(67, 13)
        Me.Label15.TabIndex = 186
        Me.Label15.Text = "Tolerance 3:"
        '
        'tbTol3
        '
        Me.tbTol3.Location = New System.Drawing.Point(344, 162)
        Me.tbTol3.Name = "tbTol3"
        Me.tbTol3.Size = New System.Drawing.Size(139, 20)
        Me.tbTol3.TabIndex = 181
        Me.tbTol3.Text = "UNLESS STATED."
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(270, 139)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(67, 13)
        Me.Label14.TabIndex = 184
        Me.Label14.Text = "Tolerance 2:"
        '
        'tbTol2
        '
        Me.tbTol2.Location = New System.Drawing.Point(344, 136)
        Me.tbTol2.Name = "tbTol2"
        Me.tbTol2.Size = New System.Drawing.Size(139, 20)
        Me.tbTol2.TabIndex = 179
        Me.tbTol2.Text = "ANGULAR: ±1°"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(270, 113)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(67, 13)
        Me.Label13.TabIndex = 180
        Me.Label13.Text = "Tolerance 1:"
        '
        'tbTol1
        '
        Me.tbTol1.Location = New System.Drawing.Point(344, 110)
        Me.tbTol1.Name = "tbTol1"
        Me.tbTol1.Size = New System.Drawing.Size(139, 20)
        Me.tbTol1.TabIndex = 178
        Me.tbTol1.Text = "LINEAR: ±0,5mm"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(23, 399)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(77, 13)
        Me.Label12.TabIndex = 177
        Me.Label12.Text = "Surface Finish:"
        '
        'tbSurface
        '
        Me.tbSurface.Location = New System.Drawing.Point(107, 396)
        Me.tbSurface.Name = "tbSurface"
        Me.tbSurface.Size = New System.Drawing.Size(139, 20)
        Me.tbSurface.TabIndex = 174
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(19, 373)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(81, 13)
        Me.Label11.TabIndex = 175
        Me.Label11.Text = "Sub Contractor:"
        '
        'tbSubCon
        '
        Me.tbSubCon.Location = New System.Drawing.Point(107, 370)
        Me.tbSubCon.Name = "tbSubCon"
        Me.tbSubCon.Size = New System.Drawing.Size(139, 20)
        Me.tbSubCon.TabIndex = 173
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(63, 347)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(37, 13)
        Me.Label10.TabIndex = 172
        Me.Label10.Text = "Scale:"
        '
        'tbScale
        '
        Me.tbScale.Location = New System.Drawing.Point(107, 344)
        Me.tbScale.Name = "tbScale"
        Me.tbScale.Size = New System.Drawing.Size(139, 20)
        Me.tbScale.TabIndex = 171
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(56, 321)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(44, 13)
        Me.Label9.TabIndex = 170
        Me.Label9.Text = "Dims in:"
        '
        'tbDimsIn
        '
        Me.tbDimsIn.Location = New System.Drawing.Point(107, 318)
        Me.tbDimsIn.Name = "tbDimsIn"
        Me.tbDimsIn.Size = New System.Drawing.Size(139, 20)
        Me.tbDimsIn.TabIndex = 169
        Me.tbDimsIn.Text = "mm"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(29, 295)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(71, 13)
        Me.Label8.TabIndex = 168
        Me.Label8.Text = "Classification:"
        '
        'tbClassification
        '
        Me.tbClassification.Location = New System.Drawing.Point(107, 292)
        Me.tbClassification.Name = "tbClassification"
        Me.tbClassification.Size = New System.Drawing.Size(139, 20)
        Me.tbClassification.TabIndex = 167
        Me.tbClassification.Text = "LEATHERY"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(43, 269)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(57, 13)
        Me.Label7.TabIndex = 166
        Me.Label7.Text = "Of Sheets:"
        '
        'tbNumSheets
        '
        Me.tbNumSheets.Location = New System.Drawing.Point(107, 266)
        Me.tbNumSheets.Name = "tbNumSheets"
        Me.tbNumSheets.Size = New System.Drawing.Size(139, 20)
        Me.tbNumSheets.TabIndex = 165
        Me.tbNumSheets.Text = "1"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(22, 243)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(78, 13)
        Me.Label6.TabIndex = 164
        Me.Label6.Text = "Sheet Number:"
        '
        'tbSheetNum
        '
        Me.tbSheetNum.Location = New System.Drawing.Point(107, 240)
        Me.tbSheetNum.Name = "tbSheetNum"
        Me.tbSheetNum.Size = New System.Drawing.Size(139, 20)
        Me.tbSheetNum.TabIndex = 163
        Me.tbSheetNum.Text = "1"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(67, 217)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(33, 13)
        Me.Label5.TabIndex = 162
        Me.Label5.Text = "Date:"
        '
        'tbDate
        '
        Me.tbDate.Location = New System.Drawing.Point(107, 214)
        Me.tbDate.Name = "tbDate"
        Me.tbDate.Size = New System.Drawing.Size(139, 20)
        Me.tbDate.TabIndex = 161
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(66, 191)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(35, 13)
        Me.Label4.TabIndex = 160
        Me.Label4.Text = "Issue:"
        '
        'tbIssue
        '
        Me.tbIssue.Location = New System.Drawing.Point(107, 188)
        Me.tbIssue.Name = "tbIssue"
        Me.tbIssue.Size = New System.Drawing.Size(139, 20)
        Me.tbIssue.TabIndex = 159
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(71, 113)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(30, 13)
        Me.Label27.TabIndex = 157
        Me.Label27.Text = "Title:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 139)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(89, 13)
        Me.Label3.TabIndex = 158
        Me.Label3.Text = "Drawing Number:"
        '
        'tbTitle
        '
        Me.tbTitle.Location = New System.Drawing.Point(107, 110)
        Me.tbTitle.Name = "tbTitle"
        Me.tbTitle.Size = New System.Drawing.Size(139, 20)
        Me.tbTitle.TabIndex = 155
        '
        'tbDwgNum
        '
        Me.tbDwgNum.Location = New System.Drawing.Point(107, 136)
        Me.tbDwgNum.Name = "tbDwgNum"
        Me.tbDwgNum.Size = New System.Drawing.Size(139, 20)
        Me.tbDwgNum.TabIndex = 156
        '
        'tbChange
        '
        Me.tbChange.Location = New System.Drawing.Point(107, 162)
        Me.tbChange.Name = "tbChange"
        Me.tbChange.Size = New System.Drawing.Size(139, 20)
        Me.tbChange.TabIndex = 157
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Location = New System.Drawing.Point(54, 165)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(47, 13)
        Me.Label28.TabIndex = 158
        Me.Label28.Text = "Change:"
        '
        'SQLConnectForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(541, 511)
        Me.Controls.Add(Me.Label25)
        Me.Controls.Add(Me.tbUsed10)
        Me.Controls.Add(Me.Label26)
        Me.Controls.Add(Me.tbUsed09)
        Me.Controls.Add(Me.Label23)
        Me.Controls.Add(Me.tbUsed08)
        Me.Controls.Add(Me.Label24)
        Me.Controls.Add(Me.tbUsed07)
        Me.Controls.Add(Me.Label21)
        Me.Controls.Add(Me.tbUsed06)
        Me.Controls.Add(Me.Label22)
        Me.Controls.Add(Me.tbUsed05)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.tbUsed04)
        Me.Controls.Add(Me.Label20)
        Me.Controls.Add(Me.tbUsed03)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.tbUsed02)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.tbUsed01)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.tbProFin)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.tbTol3)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.tbTol2)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.tbTol1)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.tbSurface)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.tbSubCon)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.tbScale)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.tbDimsIn)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.tbClassification)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.tbNumSheets)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.tbSheetNum)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.tbDate)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.tbIssue)
        Me.Controls.Add(Me.Label27)
        Me.Controls.Add(Me.Label28)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.tbTitle)
        Me.Controls.Add(Me.tbChange)
        Me.Controls.Add(Me.tbDwgNum)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.RadioButton4)
        Me.Controls.Add(Me.RadioButton3)
        Me.Controls.Add(Me.RadioButton2)
        Me.Controls.Add(Me.RadioButton1)
        Me.Controls.Add(Me.cmdExecute)
        Me.Controls.Add(Me.label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "SQLConnectForm"
        Me.Text = "Query Processor"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Private components As System.ComponentModel.IContainer
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Friend WithEvents BindingSource1 As System.Windows.Forms.BindingSource
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents tbUsed10 As System.Windows.Forms.TextBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents tbUsed09 As System.Windows.Forms.TextBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents tbUsed08 As System.Windows.Forms.TextBox
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents tbUsed07 As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents tbUsed06 As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents tbUsed05 As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents tbUsed04 As System.Windows.Forms.TextBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents tbUsed03 As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents tbUsed02 As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents tbUsed01 As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents tbProFin As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents tbTol3 As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents tbTol2 As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents tbTol1 As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents tbSurface As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents tbSubCon As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents tbScale As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents tbDimsIn As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents tbClassification As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents tbNumSheets As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents tbSheetNum As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents tbDate As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents tbIssue As System.Windows.Forms.TextBox
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents tbTitle As System.Windows.Forms.TextBox
    Friend WithEvents tbDwgNum As System.Windows.Forms.TextBox
    Friend WithEvents tbChange As System.Windows.Forms.TextBox
    Friend WithEvents Label28 As System.Windows.Forms.Label

End Class
