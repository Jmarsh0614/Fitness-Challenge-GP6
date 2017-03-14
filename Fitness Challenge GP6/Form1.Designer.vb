<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFitnessChallenge
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
        Me.mnuFitness = New System.Windows.Forms.MenuStrip()
        Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuClear = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.lblHeading = New System.Windows.Forms.Label()
        Me.btnWeightLoss = New System.Windows.Forms.Button()
        Me.lstWeightLoss = New System.Windows.Forms.ListBox()
        Me.lblAverageLoss = New System.Windows.Forms.Label()
        Me.mnuFitness.SuspendLayout()
        Me.SuspendLayout()
        '
        'mnuFitness
        '
        Me.mnuFitness.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile})
        Me.mnuFitness.Location = New System.Drawing.Point(0, 0)
        Me.mnuFitness.Name = "mnuFitness"
        Me.mnuFitness.Size = New System.Drawing.Size(659, 24)
        Me.mnuFitness.TabIndex = 0
        Me.mnuFitness.Text = "MenuStrip1"
        '
        'mnuFile
        '
        Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuClear, Me.mnuExit})
        Me.mnuFile.Name = "mnuFile"
        Me.mnuFile.Size = New System.Drawing.Size(37, 20)
        Me.mnuFile.Text = "&File"
        '
        'mnuClear
        '
        Me.mnuClear.Name = "mnuClear"
        Me.mnuClear.Size = New System.Drawing.Size(152, 22)
        Me.mnuClear.Text = "&Clear"
        '
        'mnuExit
        '
        Me.mnuExit.Name = "mnuExit"
        Me.mnuExit.Size = New System.Drawing.Size(152, 22)
        Me.mnuExit.Text = "E&xit"
        '
        'lblHeading
        '
        Me.lblHeading.AutoSize = True
        Me.lblHeading.Font = New System.Drawing.Font("Tahoma", 21.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHeading.ForeColor = System.Drawing.Color.Goldenrod
        Me.lblHeading.Location = New System.Drawing.Point(90, 20)
        Me.lblHeading.Name = "lblHeading"
        Me.lblHeading.Size = New System.Drawing.Size(283, 70)
        Me.lblHeading.TabIndex = 1
        Me.lblHeading.Text = "Fitness Challenge" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Team Weight Loss"
        Me.lblHeading.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnWeightLoss
        '
        Me.btnWeightLoss.AutoSize = True
        Me.btnWeightLoss.BackColor = System.Drawing.Color.Goldenrod
        Me.btnWeightLoss.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnWeightLoss.Location = New System.Drawing.Point(134, 93)
        Me.btnWeightLoss.Name = "btnWeightLoss"
        Me.btnWeightLoss.Size = New System.Drawing.Size(195, 33)
        Me.btnWeightLoss.TabIndex = 2
        Me.btnWeightLoss.Text = "Enter Weight Loss"
        Me.btnWeightLoss.UseVisualStyleBackColor = False
        '
        'lstWeightLoss
        '
        Me.lstWeightLoss.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstWeightLoss.FormattingEnabled = True
        Me.lstWeightLoss.ItemHeight = 23
        Me.lstWeightLoss.Location = New System.Drawing.Point(206, 132)
        Me.lstWeightLoss.Name = "lstWeightLoss"
        Me.lstWeightLoss.Size = New System.Drawing.Size(50, 188)
        Me.lstWeightLoss.TabIndex = 3
        '
        'lblAverageLoss
        '
        Me.lblAverageLoss.AutoSize = True
        Me.lblAverageLoss.BackColor = System.Drawing.Color.Goldenrod
        Me.lblAverageLoss.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAverageLoss.Location = New System.Drawing.Point(9, 353)
        Me.lblAverageLoss.Name = "lblAverageLoss"
        Me.lblAverageLoss.Size = New System.Drawing.Size(444, 23)
        Me.lblAverageLoss.TabIndex = 4
        Me.lblAverageLoss.Text = "Average weight loss of your team is XX.X lbs."
        Me.lblAverageLoss.Visible = False
        '
        'frmFitnessChallenge
        '
        Me.AcceptButton = Me.btnWeightLoss
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.Fitness_Challenge_GP6.My.Resources.Resources.Scale
        Me.ClientSize = New System.Drawing.Size(659, 436)
        Me.Controls.Add(Me.lblAverageLoss)
        Me.Controls.Add(Me.lstWeightLoss)
        Me.Controls.Add(Me.btnWeightLoss)
        Me.Controls.Add(Me.lblHeading)
        Me.Controls.Add(Me.mnuFitness)
        Me.MainMenuStrip = Me.mnuFitness
        Me.Name = "frmFitnessChallenge"
        Me.Text = "                                                                              Fit" &
    "ness Challenge"
        Me.mnuFitness.ResumeLayout(False)
        Me.mnuFitness.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents mnuFitness As MenuStrip
    Friend WithEvents mnuFile As ToolStripMenuItem
    Friend WithEvents mnuClear As ToolStripMenuItem
    Friend WithEvents mnuExit As ToolStripMenuItem
    Friend WithEvents lblHeading As Label
    Friend WithEvents btnWeightLoss As Button
    Friend WithEvents lstWeightLoss As ListBox
    Friend WithEvents lblAverageLoss As Label
End Class
