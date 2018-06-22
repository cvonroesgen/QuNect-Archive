<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRazor
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmRazor))
        Me.ckbSSO = New System.Windows.Forms.CheckBox()
        Me.pb = New System.Windows.Forms.ProgressBar()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ckbDetectProxy = New System.Windows.Forms.CheckBox()
        Me.btnListTables = New System.Windows.Forms.Button()
        Me.RetrieveTheTableReportsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.tvAppsTables = New System.Windows.Forms.TreeView()
        Me.lblAppToken = New System.Windows.Forms.Label()
        Me.txtAppToken = New System.Windows.Forms.TextBox()
        Me.lblServer = New System.Windows.Forms.Label()
        Me.txtServer = New System.Windows.Forms.TextBox()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.lblUsername = New System.Windows.Forms.Label()
        Me.txtUsername = New System.Windows.Forms.TextBox()
        Me.btnAnalyze = New System.Windows.Forms.Button()
        Me.cmbTests = New System.Windows.Forms.ComboBox()
        Me.lblResult = New System.Windows.Forms.Label()
        Me.tvFields = New System.Windows.Forms.TreeView()
        Me.cmbPassword = New System.Windows.Forms.ComboBox()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ckbSSO
        '
        Me.ckbSSO.AutoSize = True
        Me.ckbSSO.Location = New System.Drawing.Point(541, 31)
        Me.ckbSSO.Name = "ckbSSO"
        Me.ckbSSO.Size = New System.Drawing.Size(70, 17)
        Me.ckbSSO.TabIndex = 49
        Me.ckbSSO.Text = "Use SSO"
        Me.ckbSSO.UseVisualStyleBackColor = True
        '
        'pb
        '
        Me.pb.Location = New System.Drawing.Point(284, 81)
        Me.pb.Maximum = 1000
        Me.pb.Name = "pb"
        Me.pb.Size = New System.Drawing.Size(511, 23)
        Me.pb.TabIndex = 47
        Me.pb.Visible = False
        '
        'ckbDetectProxy
        '
        Me.ckbDetectProxy.AutoSize = True
        Me.ckbDetectProxy.Location = New System.Drawing.Point(633, 31)
        Me.ckbDetectProxy.Name = "ckbDetectProxy"
        Me.ckbDetectProxy.Size = New System.Drawing.Size(188, 17)
        Me.ckbDetectProxy.TabIndex = 48
        Me.ckbDetectProxy.Text = "Automatically detect proxy settings"
        Me.ckbDetectProxy.UseVisualStyleBackColor = True
        '
        'btnListTables
        '
        Me.btnListTables.Location = New System.Drawing.Point(84, 121)
        Me.btnListTables.Name = "btnListTables"
        Me.btnListTables.Size = New System.Drawing.Size(211, 23)
        Me.btnListTables.TabIndex = 34
        Me.btnListTables.Text = "Show Tables You Have Access To"
        Me.btnListTables.UseVisualStyleBackColor = True
        '
        'RetrieveTheTableReportsToolStripMenuItem
        '
        Me.RetrieveTheTableReportsToolStripMenuItem.Name = "RetrieveTheTableReportsToolStripMenuItem"
        Me.RetrieveTheTableReportsToolStripMenuItem.Size = New System.Drawing.Size(354, 22)
        Me.RetrieveTheTableReportsToolStripMenuItem.Text = "Retrieve the table reports for the selected application."
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.RetrieveTheTableReportsToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(355, 26)
        '
        'tvAppsTables
        '
        Me.tvAppsTables.ContextMenuStrip = Me.ContextMenuStrip1
        Me.tvAppsTables.Location = New System.Drawing.Point(10, 150)
        Me.tvAppsTables.Name = "tvAppsTables"
        Me.tvAppsTables.Size = New System.Drawing.Size(370, 642)
        Me.tvAppsTables.TabIndex = 33
        Me.tvAppsTables.Visible = False
        '
        'lblAppToken
        '
        Me.lblAppToken.AutoSize = True
        Me.lblAppToken.Location = New System.Drawing.Point(13, 65)
        Me.lblAppToken.Name = "lblAppToken"
        Me.lblAppToken.Size = New System.Drawing.Size(148, 13)
        Me.lblAppToken.TabIndex = 32
        Me.lblAppToken.Text = "QuickBase Application Token"
        '
        'txtAppToken
        '
        Me.txtAppToken.Location = New System.Drawing.Point(10, 84)
        Me.txtAppToken.Name = "txtAppToken"
        Me.txtAppToken.Size = New System.Drawing.Size(258, 20)
        Me.txtAppToken.TabIndex = 31
        '
        'lblServer
        '
        Me.lblServer.AutoSize = True
        Me.lblServer.Location = New System.Drawing.Point(295, 12)
        Me.lblServer.Name = "lblServer"
        Me.lblServer.Size = New System.Drawing.Size(93, 13)
        Me.lblServer.TabIndex = 30
        Me.lblServer.Text = "QuickBase Server"
        '
        'txtServer
        '
        Me.txtServer.Location = New System.Drawing.Point(295, 31)
        Me.txtServer.Name = "txtServer"
        Me.txtServer.Size = New System.Drawing.Size(226, 20)
        Me.txtServer.TabIndex = 29
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(148, 31)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(141, 20)
        Me.txtPassword.TabIndex = 27
        '
        'lblUsername
        '
        Me.lblUsername.AutoSize = True
        Me.lblUsername.Location = New System.Drawing.Point(13, 12)
        Me.lblUsername.Name = "lblUsername"
        Me.lblUsername.Size = New System.Drawing.Size(110, 13)
        Me.lblUsername.TabIndex = 26
        Me.lblUsername.Text = "QuickBase Username"
        '
        'txtUsername
        '
        Me.txtUsername.Location = New System.Drawing.Point(10, 31)
        Me.txtUsername.Name = "txtUsername"
        Me.txtUsername.Size = New System.Drawing.Size(120, 20)
        Me.txtUsername.TabIndex = 25
        '
        'btnAnalyze
        '
        Me.btnAnalyze.Location = New System.Drawing.Point(714, 122)
        Me.btnAnalyze.Name = "btnAnalyze"
        Me.btnAnalyze.Size = New System.Drawing.Size(107, 22)
        Me.btnAnalyze.TabIndex = 50
        Me.btnAnalyze.Text = "Analyze"
        Me.btnAnalyze.UseVisualStyleBackColor = True
        Me.btnAnalyze.Visible = False
        '
        'cmbTests
        '
        Me.cmbTests.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTests.FormattingEnabled = True
        Me.cmbTests.Location = New System.Drawing.Point(392, 122)
        Me.cmbTests.Name = "cmbTests"
        Me.cmbTests.Size = New System.Drawing.Size(316, 21)
        Me.cmbTests.TabIndex = 51
        Me.cmbTests.Visible = False
        '
        'lblResult
        '
        Me.lblResult.AutoSize = True
        Me.lblResult.Location = New System.Drawing.Point(397, 154)
        Me.lblResult.Name = "lblResult"
        Me.lblResult.Size = New System.Drawing.Size(0, 13)
        Me.lblResult.TabIndex = 52
        '
        'tvFields
        '
        Me.tvFields.Location = New System.Drawing.Point(392, 150)
        Me.tvFields.Name = "tvFields"
        Me.tvFields.Size = New System.Drawing.Size(429, 642)
        Me.tvFields.TabIndex = 53
        '
        'cmbPassword
        '
        Me.cmbPassword.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPassword.FormattingEnabled = True
        Me.cmbPassword.Items.AddRange(New Object() {"Please choose...", "QuickBase Password", "QuickBase User Token"})
        Me.cmbPassword.Location = New System.Drawing.Point(148, 4)
        Me.cmbPassword.Name = "cmbPassword"
        Me.cmbPassword.Size = New System.Drawing.Size(141, 21)
        Me.cmbPassword.TabIndex = 77
        '
        'frmRazor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(864, 804)
        Me.Controls.Add(Me.cmbPassword)
        Me.Controls.Add(Me.tvFields)
        Me.Controls.Add(Me.lblResult)
        Me.Controls.Add(Me.cmbTests)
        Me.Controls.Add(Me.btnAnalyze)
        Me.Controls.Add(Me.ckbSSO)
        Me.Controls.Add(Me.pb)
        Me.Controls.Add(Me.ckbDetectProxy)
        Me.Controls.Add(Me.btnListTables)
        Me.Controls.Add(Me.tvAppsTables)
        Me.Controls.Add(Me.lblAppToken)
        Me.Controls.Add(Me.txtAppToken)
        Me.Controls.Add(Me.lblServer)
        Me.Controls.Add(Me.txtServer)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.lblUsername)
        Me.Controls.Add(Me.txtUsername)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmRazor"
        Me.Text = "QuNect Razor"
        Me.ContextMenuStrip1.ResumeLayout(false)
        Me.ResumeLayout(false)
        Me.PerformLayout

End Sub
    Friend WithEvents ckbSSO As System.Windows.Forms.CheckBox
    Friend WithEvents pb As System.Windows.Forms.ProgressBar
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents ckbDetectProxy As System.Windows.Forms.CheckBox
    Friend WithEvents btnListTables As System.Windows.Forms.Button
    Friend WithEvents RetrieveTheTableReportsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents tvAppsTables As System.Windows.Forms.TreeView
    Friend WithEvents lblAppToken As System.Windows.Forms.Label
    Friend WithEvents txtAppToken As System.Windows.Forms.TextBox
    Friend WithEvents lblServer As System.Windows.Forms.Label
    Friend WithEvents txtServer As System.Windows.Forms.TextBox
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents lblUsername As System.Windows.Forms.Label
    Friend WithEvents txtUsername As System.Windows.Forms.TextBox
    Friend WithEvents btnAnalyze As System.Windows.Forms.Button
    Friend WithEvents cmbTests As System.Windows.Forms.ComboBox
    Friend WithEvents lblResult As System.Windows.Forms.Label
    Friend WithEvents tvFields As TreeView
    Friend WithEvents cmbPassword As ComboBox
End Class
