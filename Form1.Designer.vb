<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.txtInterval = New System.Windows.Forms.TextBox()
        Me.btnSelectDownloadPath = New System.Windows.Forms.Button()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.btnUpdateInterval = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtInterval
        '
        Me.txtInterval.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtInterval.Location = New System.Drawing.Point(57, 57)
        Me.txtInterval.Name = "txtInterval"
        Me.txtInterval.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtInterval.Size = New System.Drawing.Size(100, 37)
        Me.txtInterval.TabIndex = 0
        Me.txtInterval.Text = "10"
        '
        'btnSelectDownloadPath
        '
        Me.btnSelectDownloadPath.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSelectDownloadPath.Location = New System.Drawing.Point(57, 108)
        Me.btnSelectDownloadPath.Name = "btnSelectDownloadPath"
        Me.btnSelectDownloadPath.Size = New System.Drawing.Size(394, 48)
        Me.btnSelectDownloadPath.TabIndex = 1
        Me.btnSelectDownloadPath.Text = "Select Dowload Location"
        Me.btnSelectDownloadPath.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 201)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(498, 22)
        Me.StatusStrip1.TabIndex = 2
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(0, 15)
        '
        'btnUpdateInterval
        '
        Me.btnUpdateInterval.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUpdateInterval.Location = New System.Drawing.Point(264, 51)
        Me.btnUpdateInterval.Name = "btnUpdateInterval"
        Me.btnUpdateInterval.Size = New System.Drawing.Size(187, 51)
        Me.btnUpdateInterval.TabIndex = 3
        Me.btnUpdateInterval.Text = "Update Interval"
        Me.btnUpdateInterval.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(164, 62)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(94, 29)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Seconds"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(498, 223)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnUpdateInterval)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnSelectDownloadPath)
        Me.Controls.Add(Me.txtInterval)
        Me.MaximizeBox = False
        Me.Name = "Form1"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "Scotia - Outlook Email Downloader"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtInterval As TextBox
    Friend WithEvents btnSelectDownloadPath As Button
    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As ToolStripStatusLabel
    Friend WithEvents btnUpdateInterval As Button
    Friend WithEvents Label1 As Label
End Class
