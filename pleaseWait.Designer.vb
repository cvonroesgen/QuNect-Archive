<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class pleaseWait
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
        Me.pb = New System.Windows.Forms.ProgressBar()
        Me.SuspendLayout()
        '
        'pb
        '
        Me.pb.Location = New System.Drawing.Point(44, 30)
        Me.pb.Name = "pb"
        Me.pb.Size = New System.Drawing.Size(283, 23)
        Me.pb.TabIndex = 0
        '
        'pleaseWait
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(367, 88)
        Me.Controls.Add(Me.pb)
        Me.Name = "pleaseWait"
        Me.Text = "Please wait"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pb As System.Windows.Forms.ProgressBar
End Class
