<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class archive
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(archive))
        Me.txtUsername = New System.Windows.Forms.TextBox()
        Me.lblUsername = New System.Windows.Forms.Label()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.lblServer = New System.Windows.Forms.Label()
        Me.txtServer = New System.Windows.Forms.TextBox()
        Me.lblAppToken = New System.Windows.Forms.Label()
        Me.txtAppToken = New System.Windows.Forms.TextBox()
        Me.tvAppsTables = New System.Windows.Forms.TreeView()
        Me.btnListTables = New System.Windows.Forms.Button()
        Me.btnFolder = New System.Windows.Forms.Button()
        Me.lblBackupFolder = New System.Windows.Forms.Label()
        Me.lstArchiveFields = New System.Windows.Forms.ListBox()
        Me.btnAddToArchiveList = New System.Windows.Forms.Button()
        Me.btnArchive = New System.Windows.Forms.Button()
        Me.lblAttachments = New System.Windows.Forms.Label()
        Me.txtBackupFolder = New System.Windows.Forms.TextBox()
        Me.btnRemoveFromArchiveList = New System.Windows.Forms.Button()
        Me.lblTables = New System.Windows.Forms.Label()
        Me.lstFieldsToKeep = New System.Windows.Forms.ListBox()
        Me.lblFieldsToKeep = New System.Windows.Forms.Label()
        Me.lblFieldsToArchive = New System.Windows.Forms.Label()
        Me.lstReports = New System.Windows.Forms.ListBox()
        Me.lblReports = New System.Windows.Forms.Label()
        Me.btnFields = New System.Windows.Forms.Button()
        Me.btnBytes = New System.Windows.Forms.Button()
        Me.ckbDetectProxy = New System.Windows.Forms.CheckBox()
        Me.btnAddAllToArchiveList = New System.Windows.Forms.Button()
        Me.btnRemoveAllFromArchiveList = New System.Windows.Forms.Button()
        Me.cmbPassword = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'txtUsername
        '
        Me.txtUsername.Location = New System.Drawing.Point(19, 24)
        Me.txtUsername.Name = "txtUsername"
        Me.txtUsername.Size = New System.Drawing.Size(120, 20)
        Me.txtUsername.TabIndex = 1
        '
        'lblUsername
        '
        Me.lblUsername.AutoSize = True
        Me.lblUsername.Location = New System.Drawing.Point(22, 5)
        Me.lblUsername.Name = "lblUsername"
        Me.lblUsername.Size = New System.Drawing.Size(110, 13)
        Me.lblUsername.TabIndex = 0
        Me.lblUsername.Text = "QuickBase Username"
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(157, 24)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(139, 20)
        Me.txtPassword.TabIndex = 2
        '
        'lblServer
        '
        Me.lblServer.AutoSize = True
        Me.lblServer.Location = New System.Drawing.Point(302, 5)
        Me.lblServer.Name = "lblServer"
        Me.lblServer.Size = New System.Drawing.Size(93, 13)
        Me.lblServer.TabIndex = 5
        Me.lblServer.Text = "QuickBase Server"
        '
        'txtServer
        '
        Me.txtServer.Location = New System.Drawing.Point(305, 24)
        Me.txtServer.Name = "txtServer"
        Me.txtServer.Size = New System.Drawing.Size(225, 20)
        Me.txtServer.TabIndex = 3
        '
        'lblAppToken
        '
        Me.lblAppToken.AutoSize = True
        Me.lblAppToken.Location = New System.Drawing.Point(22, 53)
        Me.lblAppToken.Name = "lblAppToken"
        Me.lblAppToken.Size = New System.Drawing.Size(148, 13)
        Me.lblAppToken.TabIndex = 7
        Me.lblAppToken.Text = "QuickBase Application Token"
        '
        'txtAppToken
        '
        Me.txtAppToken.Location = New System.Drawing.Point(19, 72)
        Me.txtAppToken.Name = "txtAppToken"
        Me.txtAppToken.Size = New System.Drawing.Size(277, 20)
        Me.txtAppToken.TabIndex = 5
        '
        'tvAppsTables
        '
        Me.tvAppsTables.HideSelection = False
        Me.tvAppsTables.Location = New System.Drawing.Point(12, 147)
        Me.tvAppsTables.Name = "tvAppsTables"
        Me.tvAppsTables.Size = New System.Drawing.Size(369, 284)
        Me.tvAppsTables.TabIndex = 8
        '
        'btnListTables
        '
        Me.btnListTables.Location = New System.Drawing.Point(305, 117)
        Me.btnListTables.Name = "btnListTables"
        Me.btnListTables.Size = New System.Drawing.Size(76, 23)
        Me.btnListTables.TabIndex = 9
        Me.btnListTables.Text = "List Tables"
        Me.btnListTables.UseVisualStyleBackColor = True
        '
        'btnFolder
        '
        Me.btnFolder.Location = New System.Drawing.Point(802, 22)
        Me.btnFolder.Name = "btnFolder"
        Me.btnFolder.Size = New System.Drawing.Size(28, 23)
        Me.btnFolder.TabIndex = 10
        Me.btnFolder.Text = "..."
        Me.btnFolder.UseVisualStyleBackColor = True
        '
        'lblBackupFolder
        '
        Me.lblBackupFolder.AutoSize = True
        Me.lblBackupFolder.Location = New System.Drawing.Point(565, 5)
        Me.lblBackupFolder.Name = "lblBackupFolder"
        Me.lblBackupFolder.Size = New System.Drawing.Size(248, 13)
        Me.lblBackupFolder.TabIndex = 0
        Me.lblBackupFolder.Text = "Folder to Backup To (this in addition to the archive)"
        '
        'lstArchiveFields
        '
        Me.lstArchiveFields.FormattingEnabled = True
        Me.lstArchiveFields.Location = New System.Drawing.Point(424, 467)
        Me.lstArchiveFields.Name = "lstArchiveFields"
        Me.lstArchiveFields.Size = New System.Drawing.Size(397, 251)
        Me.lstArchiveFields.Sorted = True
        Me.lstArchiveFields.TabIndex = 12
        '
        'btnAddToArchiveList
        '
        Me.btnAddToArchiveList.Location = New System.Drawing.Point(387, 562)
        Me.btnAddToArchiveList.Name = "btnAddToArchiveList"
        Me.btnAddToArchiveList.Size = New System.Drawing.Size(31, 24)
        Me.btnAddToArchiveList.TabIndex = 13
        Me.btnAddToArchiveList.Text = "->"
        Me.btnAddToArchiveList.UseVisualStyleBackColor = True
        Me.btnAddToArchiveList.Visible = False
        '
        'btnArchive
        '
        Me.btnArchive.Location = New System.Drawing.Point(424, 80)
        Me.btnArchive.Name = "btnArchive"
        Me.btnArchive.Size = New System.Drawing.Size(397, 23)
        Me.btnArchive.TabIndex = 14
        Me.btnArchive.Text = "Archive"
        Me.btnArchive.UseVisualStyleBackColor = True
        Me.btnArchive.Visible = False
        '
        'lblAttachments
        '
        Me.lblAttachments.AutoSize = True
        Me.lblAttachments.Location = New System.Drawing.Point(296, 58)
        Me.lblAttachments.Name = "lblAttachments"
        Me.lblAttachments.Size = New System.Drawing.Size(0, 13)
        Me.lblAttachments.TabIndex = 15
        '
        'txtBackupFolder
        '
        Me.txtBackupFolder.Enabled = False
        Me.txtBackupFolder.Location = New System.Drawing.Point(559, 24)
        Me.txtBackupFolder.Name = "txtBackupFolder"
        Me.txtBackupFolder.Size = New System.Drawing.Size(237, 20)
        Me.txtBackupFolder.TabIndex = 4
        '
        'btnRemoveFromArchiveList
        '
        Me.btnRemoveFromArchiveList.Location = New System.Drawing.Point(387, 610)
        Me.btnRemoveFromArchiveList.Name = "btnRemoveFromArchiveList"
        Me.btnRemoveFromArchiveList.Size = New System.Drawing.Size(31, 24)
        Me.btnRemoveFromArchiveList.TabIndex = 18
        Me.btnRemoveFromArchiveList.Text = "<-"
        Me.btnRemoveFromArchiveList.UseVisualStyleBackColor = True
        Me.btnRemoveFromArchiveList.Visible = False
        '
        'lblTables
        '
        Me.lblTables.AutoSize = True
        Me.lblTables.Location = New System.Drawing.Point(16, 131)
        Me.lblTables.Name = "lblTables"
        Me.lblTables.Size = New System.Drawing.Size(135, 13)
        Me.lblTables.TabIndex = 20
        Me.lblTables.Text = "Tables you have access to"
        '
        'lstFieldsToKeep
        '
        Me.lstFieldsToKeep.FormattingEnabled = True
        Me.lstFieldsToKeep.Location = New System.Drawing.Point(12, 467)
        Me.lstFieldsToKeep.Name = "lstFieldsToKeep"
        Me.lstFieldsToKeep.Size = New System.Drawing.Size(369, 251)
        Me.lstFieldsToKeep.Sorted = True
        Me.lstFieldsToKeep.TabIndex = 22
        '
        'lblFieldsToKeep
        '
        Me.lblFieldsToKeep.AutoSize = True
        Me.lblFieldsToKeep.Location = New System.Drawing.Point(16, 451)
        Me.lblFieldsToKeep.Name = "lblFieldsToKeep"
        Me.lblFieldsToKeep.Size = New System.Drawing.Size(74, 13)
        Me.lblFieldsToKeep.TabIndex = 23
        Me.lblFieldsToKeep.Text = "Fields to Keep"
        '
        'lblFieldsToArchive
        '
        Me.lblFieldsToArchive.AutoSize = True
        Me.lblFieldsToArchive.Location = New System.Drawing.Point(425, 451)
        Me.lblFieldsToArchive.Name = "lblFieldsToArchive"
        Me.lblFieldsToArchive.Size = New System.Drawing.Size(85, 13)
        Me.lblFieldsToArchive.TabIndex = 24
        Me.lblFieldsToArchive.Text = "Fields to Archive"
        '
        'lstReports
        '
        Me.lstReports.FormattingEnabled = True
        Me.lstReports.Location = New System.Drawing.Point(424, 154)
        Me.lstReports.Name = "lstReports"
        Me.lstReports.Size = New System.Drawing.Size(397, 277)
        Me.lstReports.Sorted = True
        Me.lstReports.TabIndex = 25
        '
        'lblReports
        '
        Me.lblReports.AutoSize = True
        Me.lblReports.Location = New System.Drawing.Point(626, 138)
        Me.lblReports.Name = "lblReports"
        Me.lblReports.Size = New System.Drawing.Size(195, 13)
        Me.lblReports.TabIndex = 26
        Me.lblReports.Text = "Report to Choose Records for Archiving"
        '
        'btnFields
        '
        Me.btnFields.Location = New System.Drawing.Point(428, 116)
        Me.btnFields.Name = "btnFields"
        Me.btnFields.Size = New System.Drawing.Size(136, 23)
        Me.btnFields.TabIndex = 27
        Me.btnFields.Text = "List Reports and Fields"
        Me.btnFields.UseVisualStyleBackColor = True
        Me.btnFields.Visible = False
        '
        'btnBytes
        '
        Me.btnBytes.Location = New System.Drawing.Point(426, 53)
        Me.btnBytes.Name = "btnBytes"
        Me.btnBytes.Size = New System.Drawing.Size(394, 21)
        Me.btnBytes.TabIndex = 28
        Me.btnBytes.Text = "Test"
        Me.btnBytes.UseVisualStyleBackColor = True
        Me.btnBytes.Visible = False
        '
        'ckbDetectProxy
        '
        Me.ckbDetectProxy.AutoSize = True
        Me.ckbDetectProxy.Location = New System.Drawing.Point(19, 98)
        Me.ckbDetectProxy.Name = "ckbDetectProxy"
        Me.ckbDetectProxy.Size = New System.Drawing.Size(188, 17)
        Me.ckbDetectProxy.TabIndex = 30
        Me.ckbDetectProxy.Text = "Automatically detect proxy settings"
        Me.ckbDetectProxy.UseVisualStyleBackColor = True
        '
        'btnAddAllToArchiveList
        '
        Me.btnAddAllToArchiveList.Location = New System.Drawing.Point(387, 532)
        Me.btnAddAllToArchiveList.Name = "btnAddAllToArchiveList"
        Me.btnAddAllToArchiveList.Size = New System.Drawing.Size(31, 24)
        Me.btnAddAllToArchiveList.TabIndex = 31
        Me.btnAddAllToArchiveList.Text = "->>"
        Me.btnAddAllToArchiveList.UseVisualStyleBackColor = True
        Me.btnAddAllToArchiveList.Visible = False
        '
        'btnRemoveAllFromArchiveList
        '
        Me.btnRemoveAllFromArchiveList.Location = New System.Drawing.Point(387, 640)
        Me.btnRemoveAllFromArchiveList.Name = "btnRemoveAllFromArchiveList"
        Me.btnRemoveAllFromArchiveList.Size = New System.Drawing.Size(31, 24)
        Me.btnRemoveAllFromArchiveList.TabIndex = 32
        Me.btnRemoveAllFromArchiveList.Text = "<<-"
        Me.btnRemoveAllFromArchiveList.UseVisualStyleBackColor = True
        Me.btnRemoveAllFromArchiveList.Visible = False
        '
        'cmbPassword
        '
        Me.cmbPassword.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPassword.FormattingEnabled = True
        Me.cmbPassword.Items.AddRange(New Object() {"Please choose...", "QuickBase Password", "QuickBase User Token"})
        Me.cmbPassword.Location = New System.Drawing.Point(157, 2)
        Me.cmbPassword.Name = "cmbPassword"
        Me.cmbPassword.Size = New System.Drawing.Size(141, 21)
        Me.cmbPassword.TabIndex = 77
        '
        'archive
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(842, 742)
        Me.Controls.Add(Me.cmbPassword)
        Me.Controls.Add(Me.btnRemoveAllFromArchiveList)
        Me.Controls.Add(Me.btnAddAllToArchiveList)
        Me.Controls.Add(Me.ckbDetectProxy)
        Me.Controls.Add(Me.btnBytes)
        Me.Controls.Add(Me.btnFields)
        Me.Controls.Add(Me.lblReports)
        Me.Controls.Add(Me.lstReports)
        Me.Controls.Add(Me.lblFieldsToArchive)
        Me.Controls.Add(Me.lblFieldsToKeep)
        Me.Controls.Add(Me.lstFieldsToKeep)
        Me.Controls.Add(Me.lblTables)
        Me.Controls.Add(Me.btnRemoveFromArchiveList)
        Me.Controls.Add(Me.txtBackupFolder)
        Me.Controls.Add(Me.lblAttachments)
        Me.Controls.Add(Me.btnArchive)
        Me.Controls.Add(Me.btnAddToArchiveList)
        Me.Controls.Add(Me.lstArchiveFields)
        Me.Controls.Add(Me.lblBackupFolder)
        Me.Controls.Add(Me.btnFolder)
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
        Me.Name = "archive"
        Me.Text = "QuNect Archive"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtUsername As System.Windows.Forms.TextBox
    Friend WithEvents lblUsername As System.Windows.Forms.Label
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents lblServer As System.Windows.Forms.Label
    Friend WithEvents txtServer As System.Windows.Forms.TextBox
    Friend WithEvents lblAppToken As System.Windows.Forms.Label
    Friend WithEvents txtAppToken As System.Windows.Forms.TextBox
    Friend WithEvents tvAppsTables As System.Windows.Forms.TreeView
    Friend WithEvents btnListTables As System.Windows.Forms.Button
    Friend WithEvents btnFolder As System.Windows.Forms.Button
    Friend WithEvents lblBackupFolder As System.Windows.Forms.Label
    Friend WithEvents lstArchiveFields As System.Windows.Forms.ListBox
    Friend WithEvents btnAddToArchiveList As System.Windows.Forms.Button
    Friend WithEvents btnArchive As System.Windows.Forms.Button
    Friend WithEvents lblAttachments As System.Windows.Forms.Label
    Friend WithEvents txtBackupFolder As System.Windows.Forms.TextBox
    Friend WithEvents btnRemoveFromArchiveList As System.Windows.Forms.Button
    Friend WithEvents lblTables As System.Windows.Forms.Label
    Friend WithEvents lstFieldsToKeep As System.Windows.Forms.ListBox
    Friend WithEvents lblFieldsToKeep As System.Windows.Forms.Label
    Friend WithEvents lblFieldsToArchive As System.Windows.Forms.Label
    Friend WithEvents lstReports As System.Windows.Forms.ListBox
    Friend WithEvents lblReports As System.Windows.Forms.Label
    Friend WithEvents btnFields As System.Windows.Forms.Button
    Friend WithEvents btnBytes As System.Windows.Forms.Button
    Friend WithEvents ckbDetectProxy As System.Windows.Forms.CheckBox
    Friend WithEvents btnAddAllToArchiveList As System.Windows.Forms.Button
    Friend WithEvents btnRemoveAllFromArchiveList As System.Windows.Forms.Button
    Friend WithEvents cmbPassword As ComboBox
End Class
