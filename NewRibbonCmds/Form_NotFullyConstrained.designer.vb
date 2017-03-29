<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_NotFullyConstrained
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.cmdQuit = New System.Windows.Forms.Button()
        Me.cmdOpenModel = New System.Windows.Forms.Button()
        Me.cmdScanActiveDoc = New System.Windows.Forms.Button()
        Me.chkIncludeUnknownConstraintStatus = New System.Windows.Forms.CheckBox()
        Me.chkEditSketch = New System.Windows.Forms.CheckBox()
        Me.List1 = New System.Windows.Forms.ListBox()
        Me.lblFound = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblFile = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cmdQuit
        '
        Me.cmdQuit.Location = New System.Drawing.Point(12, 12)
        Me.cmdQuit.Name = "cmdQuit"
        Me.cmdQuit.Size = New System.Drawing.Size(75, 23)
        Me.cmdQuit.TabIndex = 0
        Me.cmdQuit.Text = "Quit"
        Me.cmdQuit.UseVisualStyleBackColor = True
        '
        'cmdOpenModel
        '
        Me.cmdOpenModel.Location = New System.Drawing.Point(93, 12)
        Me.cmdOpenModel.Name = "cmdOpenModel"
        Me.cmdOpenModel.Size = New System.Drawing.Size(169, 23)
        Me.cmdOpenModel.TabIndex = 1
        Me.cmdOpenModel.Text = "Open Model and Scan Structure"
        Me.cmdOpenModel.UseVisualStyleBackColor = True
        '
        'cmdScanActiveDoc
        '
        Me.cmdScanActiveDoc.Location = New System.Drawing.Point(268, 12)
        Me.cmdScanActiveDoc.Name = "cmdScanActiveDoc"
        Me.cmdScanActiveDoc.Size = New System.Drawing.Size(142, 23)
        Me.cmdScanActiveDoc.TabIndex = 2
        Me.cmdScanActiveDoc.Text = "Scan Active Doc Structure"
        Me.cmdScanActiveDoc.UseVisualStyleBackColor = True
        '
        'chkIncludeUnknownConstraintStatus
        '
        Me.chkIncludeUnknownConstraintStatus.AutoSize = True
        Me.chkIncludeUnknownConstraintStatus.Location = New System.Drawing.Point(12, 41)
        Me.chkIncludeUnknownConstraintStatus.Name = "chkIncludeUnknownConstraintStatus"
        Me.chkIncludeUnknownConstraintStatus.Size = New System.Drawing.Size(496, 17)
        Me.chkIncludeUnknownConstraintStatus.TabIndex = 3
        Me.chkIncludeUnknownConstraintStatus.Text = "Include Unknown Constraint Status (construction geometry can produce unknown cons" & _
            "traint status)"
        Me.chkIncludeUnknownConstraintStatus.UseVisualStyleBackColor = True
        '
        'chkEditSketch
        '
        Me.chkEditSketch.AutoSize = True
        Me.chkEditSketch.Location = New System.Drawing.Point(12, 64)
        Me.chkEditSketch.Name = "chkEditSketch"
        Me.chkEditSketch.Size = New System.Drawing.Size(81, 17)
        Me.chkEditSketch.TabIndex = 4
        Me.chkEditSketch.Text = "Edit Sketch"
        Me.chkEditSketch.UseVisualStyleBackColor = True
        '
        'List1
        '
        Me.List1.FormattingEnabled = True
        Me.List1.Location = New System.Drawing.Point(12, 87)
        Me.List1.Name = "List1"
        Me.List1.Size = New System.Drawing.Size(560, 472)
        Me.List1.TabIndex = 5
        '
        'lblFound
        '
        Me.lblFound.AutoSize = True
        Me.lblFound.Location = New System.Drawing.Point(99, 65)
        Me.lblFound.Name = "lblFound"
        Me.lblFound.Size = New System.Drawing.Size(39, 13)
        Me.lblFound.TabIndex = 6
        Me.lblFound.Text = "Label1"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.Color.Red
        Me.Label1.Location = New System.Drawing.Point(174, 65)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(303, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Double click on sketch to open model and make sketch visible"
        '
        'lblFile
        '
        Me.lblFile.AutoSize = True
        Me.lblFile.Location = New System.Drawing.Point(12, 42)
        Me.lblFile.Name = "lblFile"
        Me.lblFile.Size = New System.Drawing.Size(33, 13)
        Me.lblFile.TabIndex = 8
        Me.lblFile.Text = "lblFile"
        '
        'Form_NotFullyConstrained
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(584, 562)
        Me.Controls.Add(Me.chkIncludeUnknownConstraintStatus)
        Me.Controls.Add(Me.lblFile)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblFound)
        Me.Controls.Add(Me.List1)
        Me.Controls.Add(Me.chkEditSketch)
        Me.Controls.Add(Me.cmdScanActiveDoc)
        Me.Controls.Add(Me.cmdOpenModel)
        Me.Controls.Add(Me.cmdQuit)
        Me.Name = "Form_NotFullyConstrained"
        Me.Text = "Form_NotFullyConstrained"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdQuit As System.Windows.Forms.Button
    Friend WithEvents cmdOpenModel As System.Windows.Forms.Button
    Friend WithEvents cmdScanActiveDoc As System.Windows.Forms.Button
    Friend WithEvents chkIncludeUnknownConstraintStatus As System.Windows.Forms.CheckBox
    Friend WithEvents chkEditSketch As System.Windows.Forms.CheckBox
    Friend WithEvents List1 As System.Windows.Forms.ListBox
    Friend WithEvents lblFound As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblFile As System.Windows.Forms.Label
End Class
