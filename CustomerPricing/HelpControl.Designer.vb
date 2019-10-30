<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class HelpControl
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lnkDone = New System.Windows.Forms.LinkLabel()
        Me.lnkHelp = New System.Windows.Forms.LinkLabel()
        Me.wbHelp = New System.Windows.Forms.WebBrowser()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.lnkDone)
        Me.Panel1.Controls.Add(Me.lnkHelp)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(374, 19)
        Me.Panel1.TabIndex = 44
        '
        'lnkDone
        '
        Me.lnkDone.AutoSize = True
        Me.lnkDone.Dock = System.Windows.Forms.DockStyle.Left
        Me.lnkDone.Location = New System.Drawing.Point(29, 0)
        Me.lnkDone.Name = "lnkDone"
        Me.lnkDone.Size = New System.Drawing.Size(33, 13)
        Me.lnkDone.TabIndex = 41
        Me.lnkDone.TabStop = True
        Me.lnkDone.Text = "Done"
        '
        'lnkHelp
        '
        Me.lnkHelp.AutoSize = True
        Me.lnkHelp.Dock = System.Windows.Forms.DockStyle.Left
        Me.lnkHelp.Location = New System.Drawing.Point(0, 0)
        Me.lnkHelp.Name = "lnkHelp"
        Me.lnkHelp.Size = New System.Drawing.Size(29, 13)
        Me.lnkHelp.TabIndex = 40
        Me.lnkHelp.TabStop = True
        Me.lnkHelp.Text = "Help"
        '
        'wbHelp
        '
        Me.wbHelp.Dock = System.Windows.Forms.DockStyle.Fill
        Me.wbHelp.Location = New System.Drawing.Point(0, 19)
        Me.wbHelp.MinimumSize = New System.Drawing.Size(20, 20)
        Me.wbHelp.Name = "wbHelp"
        Me.wbHelp.Size = New System.Drawing.Size(374, 167)
        Me.wbHelp.TabIndex = 43
        Me.wbHelp.Url = New System.Uri("", System.UriKind.Relative)
        '
        'HelpControl
        '
        'Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        'Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        'Me.AutoSize = True
        Me.Controls.Add(Me.wbHelp)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "HelpControl"
        Me.Size = New System.Drawing.Size(374, 186)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents lnkDone As System.Windows.Forms.LinkLabel
    Friend WithEvents lnkHelp As System.Windows.Forms.LinkLabel
    Friend WithEvents wbHelp As System.Windows.Forms.WebBrowser

End Class
