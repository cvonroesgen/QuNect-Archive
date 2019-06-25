
'to deploy run the publish wizard which is configured to put the files on disk in the C:\QuNect\QuNectArchive\QuNectArchive\publish\ directory
'rename the Application Files folder in that directory to ApplicationFiles
'then edit the file C:\QuNect\QuNectArchive\QuNectArchive\publish\QuNectArchive.application and change Application Files to ApplicationFiles
'then run the MageUI tool from All Programs->Microsoft Visual Studio 2010->Microsoft SDK Tools->Manifest Generation and Editing Tool
'aka "C:\Program Files (x86)\Microsoft SDKs\Windows\v7.0A\Bin\NETFX 4.0 Tools\mageui.exe"
'open C:\QuNect\QuNectArchive\QuNectArchive\publish\QuNectArchive.application and then save it!
'ten ftp the files from the  C:\QuNect\QuNectArchive\QuNectArchive\publish\ directory to ftp.qunect.com/trial/QuNectArchive


'Windows app that lets you pick a table to archive.
'It will create a table in the same application with a file attachment field and a numeric field that holds the Record ID# of the earliest record in the CSV payload of the corresponding file attachment.
'The application will also blank out all data entry fields except for a select few chosen by the user.
'The key field will be spared of course. The key field and the Record ID# field will also always appear in the CSV payload.
'The fields left behind will not appear in the CSV payload. Also it will copy all fields to a CSV payload on local disc for safekeeping just like QuNectBackup does.
'The licensing will be done through QuNect ODBC for QuickBase. QuNect Archive's Windows application will also create a formula URL field that will allow restoration of records on a per record basis.
'This button will run a JavaScript file located in QuNect ODBC for QuickBase user defined page.
'How do I prevent someone from restoring a record over new data? Need to also create a Numeric field called Archived.
'This will contain the Record ID# of the record in the archive table that contains the archived record in CSV form, if a record has been gutted. If it's been restored then it will be blank.
'Users will be instructed to make their edit forms for that table display only the restore button if the Archive checkbox is checked.
Imports system.xml
Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Data.Odbc
Imports System.Text.RegularExpressions

Public Class archive
    Private qdb As QuickBaseClient
    Private Const AppName = "QuNectArchive"
    Private Const QuNectODBCParentDBID = "bcks8a7y3"
    Private cmdLineArgs() As String
    Private automode As Boolean = False
    Private config As String 'dbid of the table to archive, fid of the field QuNect Archived, fids to keep, dbid of the archive table
    Private schema As XmlDocument
    Private configHash As New Hashtable()
    Private reportNameToQid As New Hashtable()
    Private fieldLabelsToFIDs As New Hashtable
    Private Const recordsPerArchive = 100

    Private Sub archive_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed

    End Sub


    Private Sub archive_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ServicePointManager.SecurityProtocol = ServicePointManager.SecurityProtocol Or System.Net.SecurityProtocolType.Tls12
        txtUsername.Text = GetSetting(AppName, "Credentials", "username")
        cmbPassword.SelectedIndex = CInt(GetSetting(AppName, "Credentials", "passwordOrToken", "0"))
        txtPassword.Text = GetSetting(AppName, "Credentials", "password")
        txtServer.Text = GetSetting(AppName, "Credentials", "server", "www.quickbase.com")
        txtAppToken.Text = GetSetting(AppName, "Credentials", "apptoken", "b2fr52jcykx3tnbwj8s74b8ed55b")
        Dim detectProxySetting As String = GetSetting(AppName, "Credentials", "detectproxysettings", "0")
        If detectProxySetting = "1" Then
            ckbDetectProxy.Checked = True
        Else
            ckbDetectProxy.Checked = False
        End If
        qdb = New QuickBaseClient(txtUsername.Text, txtPassword.Text)
        txtBackupFolder.Text = GetSetting(AppName, "location", "path")
        cmdLineArgs = System.Environment.GetCommandLineArgs()
        If cmdLineArgs.Length > 1 Then
            If cmdLineArgs(1) = "auto" Then
                automode = True
                Try
                    listTables()
                Catch ex As Exception

                End Try

                archive(False)
                Me.Close()
            End If
        End If
        Dim myBuildInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(Application.ExecutablePath)
        Me.Text = "QuNect Archive " & myBuildInfo.ProductVersion
    End Sub

    Private Sub txtUsername_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUsername.TextChanged
        SaveSetting(AppName, "Credentials", "username", txtUsername.Text)
    End Sub

    Private Sub txtPassword_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPassword.TextChanged
        SaveSetting(AppName, "Credentials", "password", txtPassword.Text)
    End Sub

    Private Sub btnListTables_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnListTables.Click
        Try
            listTables()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, AppName)
        End Try

    End Sub
    Private Sub listTables()
        If txtPassword.Text.Contains(";") Then
            Throw New System.Exception("Although Quick Base allows semicolons in passwords the ODBC standard does not permit semicolons." & vbCrLf & "Please change your Quick Base password to eliminate semicolons or use a Quick Base user token instead of a password.")
        End If

        Me.Cursor = Cursors.WaitCursor
        Dim connectionString As String
        connectionString = "uid=" & txtUsername.Text
        connectionString &= ";pwd=" & txtPassword.Text
        connectionString &= ";driver={QuNect ODBC for QuickBase};"
        connectionString &= ";quickbaseserver=" & txtServer.Text
        connectionString &= ";APPTOKEN=" & txtAppToken.Text
        If cmbPassword.SelectedIndex = 0 Then
            cmbPassword.Focus()
            Throw New System.Exception("Please indicate whether you are using a password or a user token.")
        ElseIf cmbPassword.SelectedIndex = 1 Then
            connectionString &= ";PWDISPASSWORD=1"
        Else
            connectionString &= ";PWDISPASSWORD=0"
        End If
        Dim quNectConn As OdbcConnection = New OdbcConnection(connectionString)
        Try
            quNectConn.Open()
            Dim ver As String = quNectConn.ServerVersion
            If CInt(ver.Substring(2, 2)) < 12 Then
                MsgBox("Please install the 2012 version or later of QuNect ODBC for QuickBase", MsgBoxStyle.OkOnly, AppName)
                quNectConn.Dispose()
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
        Catch excpt As Exception
            If Not automode Then
                MsgBox("A license for QuNect ODBC for QuickBase is required to run QuNectArchive. " & excpt.Message(), MsgBoxStyle.OkOnly, AppName)
            End If
            Me.Cursor = Cursors.Default
            Exit Sub
        Finally

            quNectConn.Dispose()
        End Try
        Dim tableXML As XmlDocument
        Dim tableNodes As XmlNodeList = Nothing

        qdb.setServer(txtServer.Text, True)
        qdb.Authenticate(txtUsername.Text, txtPassword.Text)
        qdb.setAppToken(txtAppToken.Text)
        Try
            tableXML = qdb.GetGrantedDBs(True, True, False)
            tableNodes = tableXML.SelectNodes("/*/databases/dbinfo")
        Catch ex As Exception
            If Not automode Then
                MsgBox(ex.Message, MsgBoxStyle.OkOnly, AppName)
            End If
        Finally
            Me.Cursor = Cursors.Default
        End Try
        If tableNodes Is Nothing Then
            Exit Sub
        End If
        tvAppsTables.BeginUpdate()
        tvAppsTables.Nodes.Clear()
        Dim dbName As String
        Dim applicationName As String = ""
        Dim prevAppName As String = ""
        Dim dbid As String
        Dim i As Integer
        For i = 0 To tableNodes.Count - 1
            dbName = tableNodes(i).SelectSingleNode("dbname").InnerText
            applicationName = dbName.Split(":")(0)
            dbid = tableNodes(i).SelectSingleNode("dbid").InnerText

            If applicationName <> prevAppName Then
                tvAppsTables.Nodes.Add(applicationName)
                prevAppName = applicationName
            End If
            Dim tableName As String = dbName
            If dbName.Length > applicationName.Length Then
                tableName = dbName.Substring(applicationName.Length + 1)
            End If
            tvAppsTables.Nodes(tvAppsTables.Nodes.Count - 1).Nodes.Add(tableName & " " & dbid)
        Next
        tvAppsTables.EndUpdate()
        Dim dbidToArchive As String = GetSetting(AppName, "archive", "dbid")
        If dbidToArchive <> "" Then
            Dim tvAppNode As TreeNode
            For Each tvAppNode In tvAppsTables.Nodes
                Dim tvTableNode As TreeNode
                For Each tvTableNode In tvAppNode.Nodes
                    If tvTableNode.Text.EndsWith(" " & dbidToArchive) Then
                        tvAppsTables.SelectedNode = tvTableNode
                    End If
                Next
            Next
            If tvAppsTables.SelectedNode IsNot Nothing Then
                displayFields(tvAppsTables.SelectedNode.Text)
            End If
        End If
            lstArchiveFields.Visible = True
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub txtServer_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtServer.TextChanged
        SaveSetting(AppName, "Credentials", "server", txtServer.Text)
    End Sub
    Private Sub btnFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFolder.Click
        Dim MyFolderBrowser As New System.Windows.Forms.FolderBrowserDialog
        ' Description that displays above the dialog box control. 

        MyFolderBrowser.Description = "Select the Folder"
        ' Sets the root folder where the browsing starts from 
        MyFolderBrowser.RootFolder = Environment.SpecialFolder.MyComputer
        Dim dlgResult As DialogResult = MyFolderBrowser.ShowDialog()

        If dlgResult = Windows.Forms.DialogResult.OK Then
            txtBackupFolder.Text = MyFolderBrowser.SelectedPath
            SaveSetting(AppName, "location", "path", txtBackupFolder.Text)
        End If
    End Sub
    Private Sub btnAddToArchiveList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddToArchiveList.Click
        If lstFieldsToKeep.SelectedIndex = -1 Then
            Exit Sub
        End If
        lstArchiveFields.Items.Add(lstFieldsToKeep.Items(lstFieldsToKeep.SelectedIndex).ToString)
        lstFieldsToKeep.Items.RemoveAt(lstFieldsToKeep.SelectedIndex)
    End Sub
    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemoveFromArchiveList.Click
        If lstArchiveFields.SelectedIndex = -1 Then
            Exit Sub
        End If
        lstFieldsToKeep.Items.Add(lstArchiveFields.Items(lstArchiveFields.SelectedIndex).ToString)
        lstArchiveFields.Items.RemoveAt(lstArchiveFields.SelectedIndex)
    End Sub
    Private Sub displayFields(ByVal appTable As String)
        Me.Cursor = Cursors.WaitCursor
        qdb.setServer(txtServer.Text, True)
        qdb.Authenticate(txtUsername.Text, txtPassword.Text)
        qdb.setAppToken(txtAppToken.Text)
        Try
            Dim i As Integer
            Dim dbid As String = appTable.Substring(appTable.LastIndexOf(" ") + 1)
            schema = qdb.GetSchema(dbid)
            Dim configNode As XmlNode = schema.SelectSingleNode("/*/table/variables/var[@name='QuNectArchive " + dbid + "']")
            If configNode Is Nothing Then
                configNode = schema.SelectSingleNode("/*/table/variables/var[@name='QuNectArchive']")
            End If
            configHash.Clear()
            If Not configNode Is Nothing Then
                config = configNode.InnerText
                Dim configs As String() = config.Split(vbCrLf)
                For i = 0 To configs.Length - 1
                    configHash.Add(configs(i), configs(i))
                Next
            End If
            Dim qid As String = GetSetting(AppName, "archive", "qid")
            Dim reports As XmlNodeList = schema.SelectNodes("/*/table/queries/query[qytype='table']")
            lstReports.Items.Clear()
            reportNameToQid.Clear()
            For i = 0 To reports.Count - 1
                Dim reportNameNode As XmlNode = reports(i).SelectSingleNode("qyname")
                If reportNameNode Is Nothing Then
                    Continue For
                End If

                Dim reportName As String = reportNameNode.InnerText()
                Try
                    Dim lastListBoxItem As Integer = lstReports.Items.Add(reportName)
                    Dim thisQid As String = reports(i).SelectSingleNode("@id").InnerText()
                    reportNameToQid.Add(reportName, thisQid)
                    If thisQid = qid Then
                        lstReports.SelectedIndex = lastListBoxItem
                    End If
                Catch excpt As Exception
                    Continue For
                End Try
            Next
            Dim fields As XmlNodeList = schema.SelectNodes("/*/table/fields/field[(append_only!=1 or @field_type='email') and not(mastag) and unique!=1 and required!=1 and  not(@role) and @base_type='text' and not(@field_type='userid' or @field_type='file')and not(@mode)]")

            fieldLabelsToFIDs.Clear()
            Try
                For i = 0 To fields.Count - 1
                    Dim label As String = getFieldLabelFromNode(fields(i))
                    fieldLabelsToFIDs.Add(label, fields(i).SelectSingleNode("@id").InnerText)
                Next
            Catch labelDupe As Exception
                Throw New ArgumentException("Two fields with the same name: '" & getFieldLabelFromNode(fields(i)) & "'")
            End Try


            lstArchiveFields.Items.Clear()
            Dim fidsToArchive As String = GetSetting(AppName, "archive", "fids")
            Dim fids As String() = fidsToArchive.Split(".")
            If fids.GetLength(0) > 0 Then
                Dim labelsToArchive As New Hashtable()
                For i = 0 To fids.Length - 1
                    Dim fid As String = fids(i)
                    Dim fieldNode As XmlNode = schema.SelectSingleNode("/*/table/fields/field[@id='" & fid & "']")
                    If fieldNode Is Nothing Then
                        Continue For
                    End If
                    Dim label As String = getFieldLabelFromNode(fieldNode)
                    lstArchiveFields.Items.Add(label)
                    labelsToArchive.Add(label, label)
                Next
                lstFieldsToKeep.Items.Clear()
                For i = 0 To fields.Count - 1
                    Dim label As String = getFieldLabelFromNode(fields(i))
                    If labelsToArchive.ContainsKey(label) Then
                        Continue For
                    End If
                    lstFieldsToKeep.Items.Add(label)
                Next
            Else
                For i = 0 To fields.Count - 1
                    Dim label As String = getFieldLabelFromNode(fields(i))
                    If configHash.ContainsKey(label) Then
                        Continue For
                    End If
                    lstArchiveFields.Items.Add(label)
                Next
                lstFieldsToKeep.Items.Clear()
                For Each label As DictionaryEntry In configHash
                    If fieldLabelsToFIDs.ContainsKey(label.Key) Then
                        lstFieldsToKeep.Items.Add(label.Key)
                    End If
                Next
            End If
            btnAddToArchiveList.Visible = True
            btnRemoveFromArchiveList.Visible = True
            btnAddAllToArchiveList.Visible = True
            btnRemoveAllFromArchiveList.Visible = True
            btnArchive.Text = "Archive " & appTable
            btnBytes.Text = "How many bytes would I archive if I clicked on the button below?"
            btnArchive.Visible = True
            btnBytes.Visible = True
        Catch excpt As Exception
            MsgBox("Could not get schema of " & appTable & " " & excpt.Message(), MsgBoxStyle.OkOnly, AppName)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub
    Function getFieldLabelFromNode(fieldNode As XmlNode) As String
        getFieldLabelFromNode = fieldNode.SelectSingleNode("label").InnerText
        Dim parentFieldIDNode As XmlNode = fieldNode.SelectSingleNode("parentFieldID")
        If parentFieldIDNode IsNot Nothing Then
            getFieldLabelFromNode = schema.SelectSingleNode("/*/table/fields/field[@id='" & parentFieldIDNode.InnerText & "']/label").InnerText & ": " & getFieldLabelFromNode
        End If
    End Function
    Private Sub tvAppsTables_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvAppsTables.AfterSelect
        If Not tvAppsTables.SelectedNode Is Nothing Then
            If tvAppsTables.SelectedNode.Parent Is Nothing Then
                btnFields.Visible = False
            Else
                btnFields.Visible = True
            End If
        End If
        lstFieldsToKeep.Items.Clear()
        lstArchiveFields.Items.Clear()
        lstReports.Items.Clear()
        btnAddToArchiveList.Visible = False
        btnRemoveFromArchiveList.Visible = False
        btnAddAllToArchiveList.Visible = False
        btnRemoveAllFromArchiveList.Visible = False
        btnArchive.Visible = False
        btnBytes.Visible = False
    End Sub

    Private Sub tvAppsTables_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tvAppsTables.DoubleClick
        If Not tvAppsTables.SelectedNode.Parent Is Nothing Then
            displayFields(tvAppsTables.SelectedNode.FullPath())        
        End If       
    End Sub

    Private Sub backup_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        lstArchiveFields.Width = Me.Width - 30 - lstArchiveFields.Left
        lstArchiveFields.Height = Me.Height - 60 - lstArchiveFields.Top
        lstFieldsToKeep.Height = Me.Height - 60 - lstFieldsToKeep.Top
        lstReports.Width = Me.Width - 30 - lstReports.Left
        btnArchive.Width = Me.Width - 30 - btnArchive.Left
        btnBytes.Width = Me.Width - 30 - btnBytes.Left
    End Sub

    Private Sub btnArchive_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnArchive.Click
        archive(False)
    End Sub
    Private Sub btnBytes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBytes.Click
        archive(True)
    End Sub
    Private Sub archive(ByVal countBytesOnly As Boolean)
        Me.Cursor = Cursors.WaitCursor
        pleaseWait.Show()
        pleaseWait.Top = Me.Top + btnFields.Top
        pleaseWait.Left = Me.Left + btnFields.Left
        If tvAppsTables.SelectedNode Is Nothing Then
            pleaseWait.Close()
            MsgBox("Please choose a table for archiving.", MsgBoxStyle.OkOnly, AppName)
            tvAppsTables.Focus()
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        If lstArchiveFields.Items.Count = 0 Then
            pleaseWait.Close()
            MsgBox("Please choose at least one field for archiving.", MsgBoxStyle.OkOnly, AppName)
            lstArchiveFields.Focus()
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Dim folderPath As String = txtBackupFolder.Text
        If folderPath = "" Then
            pleaseWait.Close()
            MsgBox("Please choose a folder.", MsgBoxStyle.OkOnly, AppName)
            txtBackupFolder.Focus()
            Me.Cursor = Cursors.Default
            Exit Sub
        End If

        If lstReports.SelectedIndex = -1 Then
            pleaseWait.Close()
            MsgBox("Please choose a report.", MsgBoxStyle.OkOnly, AppName)
            lstReports.Focus()
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Dim refFID As String = ""
        Dim dbid As String = ""
        Dim dbName As String = tvAppsTables.SelectedNode.FullPath().ToString()
        Dim i As Integer        
        dbid = dbName.Substring(dbName.LastIndexOf(" ") + 1)
        Dim archivedbid As String = ""
        Dim fileFID As String = ""
        Dim keyfid As String = "3"
        schema = qdb.GetSchema(dbid)
        Dim keyField As XmlNode = schema.SelectSingleNode("/*/table/original/key_fid")
        If Not keyField Is Nothing Then
            keyfid = keyField.InnerText
        End If
        'now we have to run the report to get the csv records and put them into a file attachement field while blanking out the text fields and setting the reference field
        'first we'll get the rids from the report then we'll get records in 100 chunks and create records
        'let's get the clist first
        Dim clistArray(lstArchiveFields.Items.Count) As String

        For i = 0 To lstArchiveFields.Items.Count - 1
            clistArray(i) = fieldLabelsToFIDs(lstArchiveFields.Items(i).ToString)
        Next
        'now add the key fid to the end

        clistArray(clistArray.GetUpperBound(0)) = keyfid
        'now we need to get all the rids from the report
        Dim qid As String = reportNameToQid(lstReports.Items(lstReports.SelectedIndex))

        Dim query As String = ""
        Dim formulaDBID As String = "dbid()"
        Dim parentDBID As String = schema.SelectSingleNode("/*/table/original/app_id").InnerText

        'here we need to check to see if the backup table exists
        'we need to create an archive table
        Dim archiveTableName As String = "QuNect Archive for " & dbid
        Dim archiveTableAlias As String = archiveTableName
        archiveTableAlias = archiveTableAlias.ToUpper()
        archiveTableAlias = Regex.Replace(archiveTableAlias, "[^A-Z0-9]", "_")
        archiveTableAlias = "_DBID_" & archiveTableAlias
        'first need to check to see if the archive table already exists

        Dim parentSchema As XmlDocument = qdb.GetSchema(parentDBID)
        Dim aliasNode As XmlNode = parentSchema.SelectSingleNode("/*/table/chdbids/chdbid[@name='" & archiveTableAlias.ToLower() & "']")

        If aliasNode Is Nothing Then
            If Not countBytesOnly Then
                archivedbid = qdb.CreateTable(parentDBID, archiveTableName, "CSV archives")
            End If
        Else
            archivedbid = aliasNode.InnerText
        End If

        Dim sourceNode As XmlNode = parentSchema.SelectSingleNode("/*/table/chdbids/chdbid[. = '" & dbid & "']/@name")
        If Not sourceNode Is Nothing Then
            formulaDBID = "[" & sourceNode.InnerText & "]"
        End If
        'then we need a reference field in the table to be archived that points back to this new table
        Dim refNode As XmlNode = schema.SelectSingleNode("/*/table/fields/field[mastag='" & archiveTableAlias & "']")
        If refNode Is Nothing Then
            If Not countBytesOnly Then
                refFID = qdb.AddField(dbid, "QuNect Archive Reference", "float", False)
                qdb.SetFieldProperties(dbid, refFID, "mastag", archiveTableAlias, "foreignkey", "1")
            End If
        Else
            refFID = refNode.SelectSingleNode("@id").InnerText
        End If
        'here I should probably add a criteria to prevent archiving of records that were already archived
        If refFID <> "" Then
            query = "{'" & refFID & "'.EX.''}"
        End If

        Dim criteriaNode As XmlNode = schema.SelectSingleNode("/*/table/queries/query[@id=" & qid & "]/qycrit")
        If Not criteriaNode Is Nothing Then
            If query <> "" Then
                query &= "AND"
            End If
            query &= criteriaNode.InnerText
        End If
        Dim clist As String = "3." & keyfid
        If keyfid = "3" Then
            clist = "3"
        End If
        Dim xmlRids As XmlDocument = qdb.DoQuery(dbid, query, clist, "3", "")
        Dim ridNodeList As XmlNodeList = xmlRids.SelectNodes("/*/table/records/record")

        If Not countBytesOnly Then
            Me.Cursor = Cursors.WaitCursor
            Dim connectionString As String
            connectionString = "FIELDNAMECHARACTERS=all;uid=" & txtUsername.Text
            connectionString &= ";pwd=" & txtPassword.Text
            connectionString &= ";driver={QuNect ODBC for QuickBase};"
            connectionString &= ";quickbaseserver=" & txtServer.Text
            connectionString &= ";APPTOKEN=" & txtAppToken.Text
            If ckbDetectProxy.Checked Then
                connectionString &= ";DETECTPROXY=1"
            End If
            If cmbPassword.SelectedIndex = 0 Then
                cmbPassword.Focus()
                Throw New System.Exception("Please indicate whether you are using a password or a user token.")
            ElseIf cmbPassword.SelectedIndex = 1 Then
                connectionString &= ";PWDISPASSWORD=1"
            Else
                connectionString &= ";PWDISPASSWORD=0"
            End If
            Dim quNectConnFIDs As OdbcConnection = New OdbcConnection(connectionString & ";usefids=1")
            Try
                quNectConnFIDs.Open()
            Catch excpt As Exception
                If Not automode Then
                    MsgBox(excpt.Message(), MsgBoxStyle.OkOnly, AppName)
                End If
                quNectConnFIDs.Dispose()
                Me.Cursor = Cursors.Default
                pleaseWait.Close()
                Exit Sub
            End Try
            Dim quNectConn As OdbcConnection = New OdbcConnection(connectionString)
            Try
                quNectConn.Open()
            Catch excpt As Exception
                If Not automode Then
                    MsgBox(excpt.Message(), MsgBoxStyle.OkOnly, AppName)
                End If
                quNectConnFIDs.Close()
                quNectConnFIDs.Dispose()
                quNectConn.Dispose()
                Me.Cursor = Cursors.Default
                pleaseWait.Close()
                Exit Sub
            End Try


            If backupTable(dbName, dbid, quNectConn, quNectConnFIDs, ridNodeList) = DialogResult.Cancel Then
                MsgBox("Could not backup table, canceling archive operation.", MsgBoxStyle.OkOnly, AppName)
                quNectConn.Close()
                quNectConn.Dispose()
                quNectConnFIDs.Close()
                quNectConnFIDs.Dispose()
                Me.Cursor = Cursors.Default
                pleaseWait.Close()
                Exit Sub
            End If
            quNectConn.Close()
            quNectConn.Dispose()
            quNectConnFIDs.Close()
            quNectConnFIDs.Dispose()

            'now we need to check to see if this has a file attachment field
            Dim archiveSchema As XmlDocument = qdb.GetSchema(archivedbid)
            Dim fileNode As XmlNode = archiveSchema.SelectSingleNode("/*/table/fields/field[@field_type='file']")

            If fileNode Is Nothing Then
                fileFID = qdb.AddField(archivedbid, "CSV Archive", "file", False)
            Else
                fileFID = fileNode.SelectSingleNode("@id").InnerText
            End If

            Dim buttonNode As XmlNode = schema.SelectSingleNode("/*/table/fields/field[label='Retrieve from Archive']")
            Dim buttonFID As String
            If buttonNode Is Nothing Then
                buttonFID = qdb.AddField(dbid, "Retrieve from Archive", "url", True)
            Else
                buttonFID = buttonNode.SelectSingleNode("@id").InnerText
            End If
            Dim formula As String = ""


            keyField = schema.SelectSingleNode("/*/table/fields/field[@id=" & keyfid & "]")
            Dim keyFieldLabel As String = getFieldLabelFromNode(keyField)
            formula &= "var Text pagename = ""QuNectArchive.js"";" & vbCrLf
            formula &= "var Text cfg = ""key="" & urlencode([" & keyFieldLabel & "]) & ""&keyfid=" & keyfid & "&filefid=" & fileFID & "&reffid=" & refFID & "&apptoken=" & txtAppToken.Text & "&dbid="" & " & formulaDBID & " & ""&archivedbid=" & archivedbid & "&filerid="" & urlencode([QuNect Archive Reference]);" & vbCrLf
            formula &= "if([QuNect Archive Reference] = 0, """", " & vbCrLf
            formula &= """javascript:var cfg = '"" & URLEncode($cfg) & ""';if(typeof(qnctdg) != 'undefined'){void(qnctdg.display(cfg))}else{void($.getScript('/db/"" & Dbid() & ""?a=dbpage&pagename="" & $pagename & ""',function(){qnctdg = new QuNectArchive(cfg)}))}"""
            formula &= ")" & vbCrLf

            qdb.SetFieldProperties(dbid, buttonFID, "formula", formula, "appears_as", "Retrieve from Archive")
            Dim QuNectArchiveJS As String = qdb.GetDBPage(QuNectODBCParentDBID, "QuNectArchive.js")
            qdb.AddReplaceDBPage(dbid, "QuNectArchive.js", "1", QuNectArchiveJS)
            Dim config As String = ""
            Dim hardReturn As String = ""
            For i = 0 To lstFieldsToKeep.Items.Count - 1
                config &= hardReturn & lstFieldsToKeep.Items(i).ToString()
                hardReturn = vbCrLf
            Next
            qdb.SetDBvar(parentDBID, "QuNectArchive " + dbid, config)

        End If

        Dim bytesArchived = 0
        Dim numArchived As Integer = 0
        pleaseWait.pb.Value = 0
        Application.DoEvents()
        pleaseWait.pb.Maximum = ridNodeList.Count
        Dim xpathIndex As String = "2"
        If keyfid = "3" Then
            xpathIndex = "1"
        End If
        For i = 0 To ridNodeList.Count - 1 Step recordsPerArchive
            If i + recordsPerArchive > pleaseWait.pb.Maximum Then
                pleaseWait.pb.Value = pleaseWait.pb.Maximum
            Else
                pleaseWait.pb.Value = i + recordsPerArchive
            End If
            Application.DoEvents()
            Dim lowRid As Integer = i
            Dim highRid As Integer = i + recordsPerArchive - 1
            If ridNodeList.Count - 1 < i + recordsPerArchive - 1 Then
                highRid = ridNodeList.Count - 1
            End If
            Dim n As Integer
            Dim oRRids As String = ""
            Dim strOr As String = ""
            For n = i To highRid
                oRRids &= strOr & ridNodeList(n).SelectSingleNode("f").InnerText
                strOr = " OR "
            Next


            Dim strCSV As String = String.Join(",", clistArray) & vbCrLf
            Dim genResultsTable As String = qdb.GenResultsTable(dbid, "{'3'.EX.'" & oRRids & "'}", String.Join(".", clistArray), "3", "sortorder-A.csv")
            Dim bytes As Integer = genResultsTable.Length
            bytes -= (highRid - lowRid + 1) * (2 + (lstArchiveFields.Items.Count))
            bytesArchived += bytes
            strCSV &= genResultsTable
            If Not countBytesOnly Then
                'need to create a new record in the archive table
                Dim sTempFileName As String = System.IO.Path.GetTempFileName()
                Dim swTemp As New System.IO.StreamWriter(sTempFileName)
                swTemp.Write(strCSV)
                swTemp.Close()
                Dim fsTemp As New System.IO.FileStream(sTempFileName, IO.FileMode.Open, FileAccess.Read)
                Dim archiveRID As String = qdb.AddRecord(archivedbid, "", fileFID, fsTemp)
                fsTemp.Close()
                System.IO.File.Delete(sTempFileName)
                'now we need to hollow out the records and update the reference field value
                Dim j As Integer
                strCSV = ""
                Dim k As Integer
                For j = lowRid To highRid
                    For k = 0 To clistArray.Length - 2
                        strCSV &= """"","
                    Next
                    strCSV &= """" & ridNodeList(j).SelectSingleNode("f[" & xpathIndex & "]").InnerText & """, " & archiveRID & vbCrLf
                Next
                Dim recordids(0) As Integer
                numArchived += qdb.ImportFromCSV(dbid, strCSV, String.Join(".", clistArray) & "." & refFID, recordids, False)
            End If
        Next
        Me.Cursor = Cursors.Default
        pleaseWait.Close()
        If countBytesOnly Then
            MsgBox("About " & bytesArchived & " bytes would be archived from " & ridNodeList.Count & " records.", MsgBoxStyle.OkOnly, AppName)
        Else
            MsgBox("About " & bytesArchived & " bytes archived from " & ridNodeList.Count & " records.", MsgBoxStyle.OkOnly, AppName)
        End If

    End Sub
    Private Function backupTable(ByVal dbName As String, ByVal dbid As String, ByVal quNectConn As OdbcConnection, ByVal quNectConnFIDs As OdbcConnection, ByRef ridNodeList As XmlNodeList) As DialogResult
        backupTable = DialogResult.OK
        Dim quickBaseSQL As String = "select * from """ & dbid & """"

        Dim quNectCmd As OdbcCommand = New OdbcCommand(quickBaseSQL, quNectConnFIDs)
        Dim dr As OdbcDataReader
        Try
            dr = quNectCmd.ExecuteReader()
        Catch excpt As Exception
            If Not automode Then
                backupTable = MsgBox("Could not get field identifiers for table " & dbid & " because " & excpt.Message() & vbCrLf & "Would you like to continue?", MsgBoxStyle.OkCancel, AppName)
            End If
            quNectCmd.Dispose()
            Exit Function
        End Try
        If Not dr.HasRows Then
            Exit Function
        End If
        Dim i
        Dim clist As String = ""
        Dim period As String = ""
        For i = 0 To dr.FieldCount - 1
            clist &= period & dr.GetName(i).Replace("fid", "")
            period = "."
        Next
        dr.Close()
        quNectCmd.Dispose()
        Dim folderPath As String = txtBackupFolder.Text
        folderPath &= "\" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss")
        Directory.CreateDirectory(folderPath)
        Dim filenamePrefix As String = dbName.Replace("/", "").Replace("\", "").Replace(":", "").Replace(":", "_").Replace("?", "").Replace("""", "").Replace("<", "").Replace(">", "").Replace("|", "")
        If filenamePrefix.Length > 229 Then
            filenamePrefix = filenamePrefix.Substring(filenamePrefix.Length - 229)
        End If
        Dim filepath As String = folderPath & "\" & filenamePrefix & ".fids"
        Dim objWriter As System.IO.StreamWriter
        Try
            objWriter = New System.IO.StreamWriter(filepath)
        Catch excpt As Exception
            If Not automode Then
                backupTable = MsgBox("Could not open file " & filepath & " because " & excpt.Message() & vbCrLf & "Would you like to continue?", MsgBoxStyle.OkCancel, AppName)
            End If
            Exit Function
        End Try
        objWriter.Write(clist)
        objWriter.Close()
        'here we need to open a file
        'filename prefix can only be 229 characters in length

        filepath = folderPath & "\" & filenamePrefix & ".csv"
        Try
            objWriter = New System.IO.StreamWriter(filepath)
        Catch excpt As Exception
            If Not automode Then
                backupTable = MsgBox("Could not open file " & filepath & " because " & excpt.Message() & vbCrLf & "Would you like to continue?", MsgBoxStyle.OkCancel, AppName)
            End If
            Exit Function
        End Try
        quickBaseSQL &= " WHERE fid3 IN ("
        Dim j As Integer
        For j = 0 To ridNodeList.Count - 1 Step 100
            Application.DoEvents()
            Dim lowRid As Integer = j
            Dim highRid As Integer = j + 100 - 1
            If ridNodeList.Count - 1 < j + 100 - 1 Then
                highRid = ridNodeList.Count - 1
            End If
            Dim n As Integer
            Dim commaRids As String = ""
            Dim strComma As String = ""
            For n = j To highRid
                commaRids &= strComma & ridNodeList(n).SelectSingleNode("f").InnerText
                strComma = ","
            Next

            quNectCmd = New OdbcCommand(quickBaseSQL & commaRids & ")", quNectConn)
            Try
                dr = quNectCmd.ExecuteReader()
            Catch excpt As Exception
                If Not automode Then
                    backupTable = MsgBox("Could not backup table " & filenamePrefix & " because " & excpt.Message() & vbCrLf & "Would you like to continue?", MsgBoxStyle.OkCancel, AppName)
                End If
                quNectCmd.Dispose()
                Exit Function
            End Try
            If Not dr.HasRows Then
                Exit Function
            End If

            If j = 0 Then
                For i = 0 To dr.FieldCount - 1
                    objWriter.Write("""")
                    objWriter.Write(Replace(CStr(dr.GetName(i)), """", """"""))
                    objWriter.Write(""",")
                Next
            End If
            objWriter.Write(vbCrLf)
            Dim k As Integer = 0
            pleaseWait.pb.Maximum = 100
            While (dr.Read())
                pleaseWait.pb.Value = k Mod 100
                Application.DoEvents()
                k += 1
                For i = 0 To dr.FieldCount - 1
                    If dr.GetValue(i) Is Nothing Then
                        objWriter.Write(",")
                    Else
                        objWriter.Write("""")
                        objWriter.Write(Replace(dr.GetValue(i).ToString(), """", """"""))
                        objWriter.Write(""",")
                    End If
                Next
                objWriter.Write(vbCrLf)
            End While
            dr.Close()
        Next j
        objWriter.Close()
        quNectCmd.Dispose()
    End Function

    Private Sub txtAppToken_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAppToken.TextChanged
        SaveSetting(AppName, "Credentials", "apptoken", txtAppToken.Text)
    End Sub


    Private Sub btnFields_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFields.Click
        If Not tvAppsTables.SelectedNode Is Nothing And Not tvAppsTables.SelectedNode.Parent Is Nothing Then
            displayFields(tvAppsTables.SelectedNode.FullPath())
        Else
            MsgBox("Please select a table first.", MsgBoxStyle.OkOnly, AppName)
        End If

    End Sub

    Private Sub ckbDetectProxy_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ckbDetectProxy.CheckStateChanged
        If ckbDetectProxy.Checked Then
            SaveSetting(AppName, "Credentials", "detectproxysettings", "1")
        Else
            SaveSetting(AppName, "Credentials", "detectproxysettings", "0")
        End If
    End Sub

    Private Sub btnAddAllToArchiveList_Click(sender As Object, e As EventArgs) Handles btnAddAllToArchiveList.Click
        While lstFieldsToKeep.Items.Count
            lstArchiveFields.Items.Add(lstFieldsToKeep.Items(0).ToString)
            lstFieldsToKeep.Items.RemoveAt(0)
        End While
    End Sub

    Private Sub btnRemoveAllFromArchiveList_Click(sender As Object, e As EventArgs) Handles btnRemoveAllFromArchiveList.Click
        While lstArchiveFields.Items.Count
            lstFieldsToKeep.Items.Add(lstArchiveFields.Items(0).ToString)
            lstArchiveFields.Items.RemoveAt(0)
        End While
    End Sub
    Private Sub cmbPassword_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbPassword.SelectedIndexChanged
        SaveSetting(AppName, "Credentials", "passwordOrToken", cmbPassword.SelectedIndex)
        If cmbPassword.SelectedIndex = 0 Then
            txtPassword.Enabled = False
        Else
            txtPassword.Enabled = True
        End If
    End Sub
End Class

Public Class QuickBaseClient

    Private Password As String
    Private UserName As String
    Private strProxyPassword As String
    Private strProxyUsername As String
    Private ticket As String
    Private apptoken As String
    Private QDBHost As String = "www.quickbase.com"
    Private useHTTPS As Boolean = True
    Public GMTOffset As Single


    Public errorcode As Integer
    Public errortext As String
    Public errordetail As String
    Public httpContentLengthProgress As Integer
    Public httpContentLength As Integer

    Private Const OB32CHARACTERS As String = "abcdefghijkmnpqrstuvwxyz23456789"
    Private Const Map64 As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    Private Const MILLISECONDS_IN_A_DAY As Double = 86400000.0#
    Private Const DAYS_BETWEEN_JAVASCRIPT_AND_MICROSOFT_DATE_REFERENCES As Double = 25569.0#
    Function makeCSVCells(ByRef cells As ArrayList) As String
        Dim i As Integer
        Dim cell As String
        makeCSVCells = ""
        For i = 0 To cells.Count - 1
            If cells(i) Is Nothing Then
                cell = ""
            Else
                cell = cells(i).ToString()
            End If
            makeCSVCells = makeCSVCells & """" & cell.Replace("""", """""") & ""","
        Next
    End Function

    Function encode32(ByVal strDecimal As String) As String

        Dim ob32 As String = ""
        Dim intDecimal As Integer
        intDecimal = CInt(strDecimal)
        Dim remainder As Integer

        Do While (intDecimal > 0)
            remainder = intDecimal Mod 32
            ob32 = Mid(OB32CHARACTERS, CInt(remainder) + 1, 1) & ob32
            intDecimal = intDecimal \ 32
        Loop
        encode32 = ob32

    End Function
    Public Function getTextByFID(ByRef recordNode As XmlNode, ByRef fid As String) As String
        Dim cell As XmlNode = recordNode.SelectSingleNode("f[@id=" & fid & "]")
        If cell Is Nothing Then
            Err.Raise(vbObjectError + 5, "QuickBase.QuickBaseClient", "Could not find fid " & fid)
        End If
        getTextByFID = cell.InnerText
    End Function
    Public Function makeClist(ByVal fids As Hashtable) As String
        Dim period As String = ""
        makeClist = ""
        For Each fid As DictionaryEntry In fids
            makeClist = makeClist & period & fid.Value
            period = "."
        Next
    End Function
    Public Function makeClist(ByRef fids As ArrayList) As String
        Dim period As String = ""
        makeClist = ""
        Dim i As Integer
        For i = 0 To fids.Count - 1
            makeClist = makeClist & period & fids(i)
            period = "."
        Next
    End Function
    Public Function makeClist(ByRef fids() As String) As String
        Dim period As String = ""
        makeClist = ""
        Dim fid As String
        For Each fid In fids
            makeClist = makeClist & period & fid
            period = "."
        Next
    End Function
    Public Function FieldAddChoices(ByVal dbid As String, ByVal fid As String, ByVal ParamArray NameValues() As Object) As Integer
        Dim xmlQDBRequest As XmlDocument
        Dim firstfield As Integer
        Dim lastfield As Integer
        Dim i As Integer

        xmlQDBRequest = InitXMLRequest()
        addParameter(xmlQDBRequest, "fid", fid)
        lastfield = UBound(NameValues)
        firstfield = LBound(NameValues)
        For i = firstfield To lastfield
            addParameter(xmlQDBRequest, "choice", CStr(NameValues(i)))
        Next i

        xmlQDBRequest = APIXMLPost(dbid, "API_FieldAddChoices", xmlQDBRequest, useHTTPS)
        FieldAddChoices = CInt(xmlQDBRequest.DocumentElement.SelectSingleNode("/*/numadded").InnerText)
    End Function

    Public Function CreateDatabase(ByVal dbname As String, ByVal dbdesc As String) As String
        Dim xmlQDBRequest As XmlDocument

        xmlQDBRequest = InitXMLRequest()
        addParameter(xmlQDBRequest, "dbname", dbname)
        addParameter(xmlQDBRequest, "dbdesc", dbdesc)
        CreateDatabase = ""
        CreateDatabase = APIXMLPost("main", "API_CreateDatabase", xmlQDBRequest, useHTTPS).DocumentElement.SelectSingleNode("/*/dbid").InnerText
    End Function
    Public Sub DeleteDatabase(ByVal dbid As String)
        Dim xmlQDBRequest As XmlDocument

        xmlQDBRequest = InitXMLRequest()
        addParameter(xmlQDBRequest, "dbid", dbid)
        Call APIXMLPost(dbid, "API_DeleteDatabase", xmlQDBRequest, useHTTPS)
    End Sub
    Public Function AddField(ByVal dbid As String, ByVal label As String, ByVal fieldtype As String, ByVal Formula As Boolean) As String
        Dim xmlQDBRequest As XmlDocument

        xmlQDBRequest = InitXMLRequest()
        addParameter(xmlQDBRequest, "label", label)
        addParameter(xmlQDBRequest, "type", fieldtype)
        If Formula Then
            addParameter(xmlQDBRequest, "mode", "virtual")
        End If
        AddField = ""
        AddField = APIXMLPost(dbid, "API_AddField", xmlQDBRequest, useHTTPS).DocumentElement.SelectSingleNode("/*/fid").InnerText
    End Function
    Public Sub SetFieldProperties(ByVal dbid As String, ByVal fid As String, ByVal ParamArray NameValues() As Object)
        Dim xmlQDBRequest As XmlDocument
        Dim lastfield As Integer
        Dim firstfield As Integer
        Dim i As Integer

        xmlQDBRequest = InitXMLRequest()
        addParameter(xmlQDBRequest, "fid", fid)
        lastfield = UBound(NameValues)
        firstfield = LBound(NameValues)
        i = 0
        For i = firstfield To lastfield Step 2
            addParameter(xmlQDBRequest, CStr(NameValues(i)), CStr(NameValues(i + 1)))
        Next i
        Call APIXMLPost(dbid, "API_SetFieldProperties", xmlQDBRequest, useHTTPS)
    End Sub

    Public Function Authenticate(ByVal strUsername As String, ByVal strPassword As String) As Integer
        UserName = strUsername
        Password = strPassword
        ticket = ""
        Authenticate = 0
    End Function
    Public Function proxyAuthenticate(ByVal strUsername As String, ByVal strPassword As String) As Integer
        strProxyUsername = strUsername
        strProxyPassword = strPassword
        proxyAuthenticate = 0
    End Function
    Function downloadAttachedFile(ByVal dbid As String, ByVal rid As String, ByVal fid As String, ByVal DownloadDirectory As String, ByVal Filename As String) As String
        Filename = makeValidFilename(Filename)
        downloadAttachedFile = HTTPPost("", True, "/up/" & dbid & "/a/r" & rid & "/e" & fid & "/?ticket=" & getTicket() & "&apptoken=" & apptoken, "text/html", "", DownloadDirectory & "\" & Filename)
    End Function
    Public Function FindDBByName(ByVal dbname As String) As String
        Dim xmlQDBRequest As XmlDocument

        xmlQDBRequest = InitXMLRequest()
        addParameter(xmlQDBRequest, "dbname", dbname)
        FindDBByName = ""
        FindDBByName = APIXMLPost("main", "API_FindDBByName", xmlQDBRequest, useHTTPS).DocumentElement.SelectSingleNode("/*/dbid").InnerText
    End Function
    Public Function CloneDatabase(ByVal sourcedbid As String, ByVal Name As String, ByVal Description As String) As String
        Dim xmlQDBRequest As XmlDocument
        Dim xmlNewDBID As XmlNode
        xmlQDBRequest = InitXMLRequest()
        addParameter(xmlQDBRequest, "newdbname", Name)
        addParameter(xmlQDBRequest, "newdbdesc", Description)
        CloneDatabase = ""
        xmlQDBRequest = APIXMLPost(sourcedbid, "API_CloneDatabase", xmlQDBRequest, useHTTPS)
        If Not xmlQDBRequest.HasChildNodes Then
            Err.Raise(vbObjectError + 5, "QuickBase.QuickBaseClient", "Please login with an user account that has permission to create applications.")
        End If
        xmlNewDBID = xmlQDBRequest.DocumentElement.SelectSingleNode("/*/newdbid")
        If xmlNewDBID Is Nothing Then
            Err.Raise(vbObjectError + 5, "QuickBase.QuickBaseClient", "Please login with an user account that has permission to create applications in only one billing account.")
        Else
            CloneDatabase = xmlNewDBID.InnerText
        End If
    End Function
    Public Function ImportFromCSV(ByVal dbid As String, ByVal CSV As String, ByVal clist As String, ByRef rids() As Integer, ByVal skipfirst As Boolean) As Integer
        Dim xmlQDBRequest As XmlDocument
        Dim RidNodeList As XmlNodeList 'XmlNodeList

        xmlQDBRequest = InitXMLRequest()
        addParameter(xmlQDBRequest, "clist", clist)
        addParameter(xmlQDBRequest, "msInUTC", "1")
        If skipfirst Then
            addParameter(xmlQDBRequest, "skipfirst", "1")
        End If
        addCDATAParameter(xmlQDBRequest, "records_csv", CSV)
        xmlQDBRequest = APIXMLPost(dbid, "API_ImportFromCSV", xmlQDBRequest, useHTTPS)
        RidNodeList = xmlQDBRequest.SelectNodes("/*/rids/rid")
        Dim ridListLength As Integer
        Dim i As Integer
        ridListLength = RidNodeList.Count
        If ridListLength > 0 Then
            ReDim rids(ridListLength - 1)
            For i = 0 To ridListLength - 1
                rids(i) = CInt(RidNodeList(i).InnerText)
            Next i
        End If
        On Error Resume Next
        ImportFromCSV = CInt(xmlQDBRequest.DocumentElement.SelectSingleNode("/*/num_recs_added").InnerText)
        ImportFromCSV = CInt(xmlQDBRequest.DocumentElement.SelectSingleNode("/*/num_recs_updated").InnerText)
        xmlQDBRequest = Nothing
    End Function
    Public Function AddRecordByArray(ByVal dbid As String, ByRef update_id As String, ByRef NameValues(,) As Object) As String
        Dim xmlQDBRequest As XmlDocument
        Dim firstfield As Integer
        Dim lastfield As Integer
        Dim i As Integer
        AddRecordByArray = ""

        xmlQDBRequest = InitXMLRequest()
        lastfield = UBound(NameValues, 2)
        firstfield = LBound(NameValues, 2)

        For i = firstfield To lastfield
            If IsDBNull(NameValues(0, i)) Then
                Err.Raise(vbObjectError + 2, "QuickBase.QuickBaseClient", "AddRecordByArray: Please do use null for field names or fids")
                Exit Function
            End If
            If IsDBNull(NameValues(1, i)) Then
                NameValues(1, i) = CObj("")
            End If

            If (IsNumeric(NameValues(0, i)) And Not IsDate(NameValues(0, i))) Then
                addFieldParameter(xmlQDBRequest, "fid", CStr(NameValues(0, i)), NameValues(1, i))
            Else
                addFieldParameter(xmlQDBRequest, "name", makeAlphaNumLowerCase(CStr(NameValues(0, i))), NameValues(1, i))
            End If
        Next i
        xmlQDBRequest = APIXMLPost(dbid, "API_AddRecord", xmlQDBRequest, useHTTPS)
        update_id = xmlQDBRequest.DocumentElement.SelectSingleNode("/*/update_id").InnerText
        AddRecordByArray = xmlQDBRequest.DocumentElement.SelectSingleNode("/*/rid").InnerText
    End Function

    Public Function EditRecordByArray(ByVal dbid As String, ByVal rid As String, ByRef update_id As String, ByRef NameValues(,) As Object) As String
        Dim xmlQDBRequest As XmlDocument
        Dim firstfield As Integer
        Dim lastfield As Integer
        Dim i As Integer
        EditRecordByArray = ""

        xmlQDBRequest = InitXMLRequest()
        lastfield = UBound(NameValues, 2)
        firstfield = LBound(NameValues, 2)



        For i = firstfield To lastfield
            If (IsNumeric(NameValues(0, i)) And Not IsDate(NameValues(0, i))) Then
                addFieldParameter(xmlQDBRequest, "fid", CStr(NameValues(0, i)), NameValues(1, i))
            Else
                addFieldParameter(xmlQDBRequest, "name", makeAlphaNumLowerCase(CStr(NameValues(0, i))), NameValues(1, i))
            End If
        Next i
        addParameter(xmlQDBRequest, "rid", rid)
        If update_id <> "" Then
            addParameter(xmlQDBRequest, "update_id", update_id)
        End If
        EditRecordByArray = APIXMLPost(dbid, "API_EditRecord", xmlQDBRequest, useHTTPS).DocumentElement.SelectSingleNode("/*/update_id").InnerText
    End Function
    Public Function AddRecord(ByVal dbid As String, ByRef update_id As String, ByVal ParamArray NameValues() As Object) As String
        Dim xmlQDBRequest As XmlDocument
        Dim firstfield As Integer
        Dim lastfield As Integer
        Dim i As Integer
        AddRecord = ""

        xmlQDBRequest = InitXMLRequest()
        lastfield = UBound(NameValues)
        firstfield = LBound(NameValues)
        If ((lastfield - firstfield + 1) Mod 2) <> 0 Then
            Err.Raise(vbObjectError + 3, "QuickBase.QuickBaseClient", "AddRecord: Please use an even number of arguements after the DBID")
            Exit Function
        End If


        For i = firstfield To lastfield Step 2
            If (IsNumeric(NameValues(i)) And Not IsDate(NameValues(i))) Then
                addFieldParameter(xmlQDBRequest, "fid", CStr(NameValues(i)), NameValues(i + 1))
            Else
                addFieldParameter(xmlQDBRequest, "name", makeAlphaNumLowerCase(CStr(NameValues(i))), NameValues(i + 1))
            End If
        Next i
        xmlQDBRequest = APIXMLPost(dbid, "API_AddRecord", xmlQDBRequest, useHTTPS)
        update_id = xmlQDBRequest.DocumentElement.SelectSingleNode("/*/update_id").InnerText
        AddRecord = xmlQDBRequest.DocumentElement.SelectSingleNode("/*/rid").InnerText
    End Function
    Public Function EditRecord(ByVal dbid As String, ByVal rid As String, ByRef update_id As String, ByVal ParamArray NameValues() As Object) As String
        Dim xmlQDBRequest As XmlDocument
        Dim firstfield As Integer
        Dim lastfield As Integer
        Dim i As Integer
        EditRecord = ""

        xmlQDBRequest = InitXMLRequest()
        lastfield = UBound(NameValues)
        firstfield = LBound(NameValues)
        If ((lastfield - firstfield + 1) Mod 2) <> 0 Then
            Err.Raise(vbObjectError + 4, "QuickBase.QuickBaseClient", "EditRecord: Please use an even number of arguements.")
            Exit Function
        End If


        For i = firstfield To lastfield Step 2
            If (IsNumeric(NameValues(i)) And Not IsDate(NameValues(i))) Then
                addFieldParameter(xmlQDBRequest, "fid", CStr(NameValues(i)), NameValues(i + 1))
            Else
                addFieldParameter(xmlQDBRequest, "name", makeAlphaNumLowerCase(CStr(NameValues(i))), NameValues(i + 1))
            End If
        Next i
        addParameter(xmlQDBRequest, "rid", rid)
        If update_id <> "" Then
            addParameter(xmlQDBRequest, "update_id", update_id)
        End If
        EditRecord = APIXMLPost(dbid, "API_EditRecord", xmlQDBRequest, useHTTPS).DocumentElement.SelectSingleNode("/*/update_id").InnerText

    End Function
    Public Function DeleteRecord(ByVal dbid As String, ByVal rid As Object) As String
        Dim xmlQDBRequest As XmlDocument
        DeleteRecord = ""

        xmlQDBRequest = InitXMLRequest()
        addParameter(xmlQDBRequest, "rid", CStr(rid))

        DeleteRecord = APIXMLPost(dbid, "API_DeleteRecord", xmlQDBRequest, useHTTPS).DocumentElement.SelectSingleNode("/*/rid").InnerText

    End Function
    Public Function GetSchema(ByVal dbid As String) As XmlDocument
        Dim xmlQDBRequest As XmlDocument

        xmlQDBRequest = InitXMLRequest()
        GetSchema = APIXMLPost(dbid, "API_GetSchema", xmlQDBRequest, useHTTPS)
    End Function

    Public Function GetGrantedDBs(ByVal withEmbeddedTables As Boolean, ByVal excludeParents As Boolean, ByVal adminOnly As Boolean) As XmlDocument
        Dim xmlQDBRequest As XmlDocument

        xmlQDBRequest = InitXMLRequest()
        If withEmbeddedTables Then
            addParameter(xmlQDBRequest, "withEmbeddedTables", "1")
        Else
            addParameter(xmlQDBRequest, "withEmbeddedTables", "0")
        End If
        If excludeParents Then
            addParameter(xmlQDBRequest, "excludeParents", "1")
        Else
            addParameter(xmlQDBRequest, "excludeParents", "0")
        End If
        If adminOnly Then
            addParameter(xmlQDBRequest, "adminOnly", "1")
        End If
        addParameter(xmlQDBRequest, "realmAppsOnly", "true")
        GetGrantedDBs = APIXMLPost("main", "API_GrantedDBs", xmlQDBRequest, useHTTPS)
    End Function
    Public Function GetDBInfo(ByVal dbid As String) As XmlDocument
        Dim xmlQDBRequest As XmlDocument

        xmlQDBRequest = InitXMLRequest()
        GetDBInfo = APIXMLPost(dbid, "API_GetDBInfo", xmlQDBRequest, useHTTPS)
    End Function

    Public Function ChangeRecordOwner(ByVal dbid As String, ByVal rid As Object, ByVal Owner As String) As Boolean
        Dim xmlQDBRequest As XmlDocument

        xmlQDBRequest = InitXMLRequest()
        addParameter(xmlQDBRequest, "rid", CStr(rid))
        addParameter(xmlQDBRequest, "newowner", Owner)
        On Error GoTo noChange
        Call APIXMLPost(dbid, "API_ChangeRecordOwner", xmlQDBRequest, useHTTPS)
        ChangeRecordOwner = True
        Exit Function

noChange:
        ChangeRecordOwner = False
        Exit Function
    End Function
    Public Function DoQuery(ByVal dbid As String, ByVal query As String, ByVal clist As String, ByVal slist As String, ByVal options As String) As XmlDocument
        Dim xmlQDBRequest As XmlDocument

        xmlQDBRequest = InitXMLRequest()
        If CStr(Val(query)) = query Then
            addParameter(xmlQDBRequest, "qid", query)
        Else
            addParameter(xmlQDBRequest, "query", query)
        End If
        addParameter(xmlQDBRequest, "clist", clist)
        addParameter(xmlQDBRequest, "slist", slist)
        addParameter(xmlQDBRequest, "options", options)
        addParameter(xmlQDBRequest, "fmt", "structured")
        DoQuery = APIXMLPost(dbid, "API_DoQuery", xmlQDBRequest, useHTTPS)
    End Function
    Public Function GenResultsTable(ByVal dbid As String, ByVal query As String, ByVal clist As String, ByVal slist As String, ByVal options As String) As String
        Dim xmlQDBRequest As XmlDocument

        xmlQDBRequest = InitXMLRequest()

        If Left(query, 1) = "{" And Right(query, 1) = "}" Then
            addParameter(xmlQDBRequest, "query", query)
        ElseIf CStr(Val(query)) = query Then
            addParameter(xmlQDBRequest, "qid", query)
        Else
            addParameter(xmlQDBRequest, "qname", query)
        End If
        addParameter(xmlQDBRequest, "clist", clist)
        addParameter(xmlQDBRequest, "slist", slist)
        addParameter(xmlQDBRequest, "options", options)
        GenResultsTable = APIHTMLPost(dbid, "API_GenResultsTable", xmlQDBRequest, useHTTPS)
    End Function

    Public Function DoQueryAsArray(ByVal dbid As String, ByVal query As String, ByVal clist As String, ByVal slist As String, ByVal options As String) As Object(,)
        Dim xmlQDBResponse As New XmlDocument
        xmlQDBResponse = DoQuery(dbid, query, clist, slist, options)
        Dim QDBRecord(,) As Object

        Dim i As Integer
        Dim j As Integer
        Dim FieldNodeList As XmlNodeList
        Dim RecordNodeList As XmlNodeList
        Dim intFields As Integer
        Dim strFieldValue As String

        FieldNodeList = xmlQDBResponse.DocumentElement.SelectNodes("/*/table/fields/field")
        intFields = FieldNodeList.Count

        RecordNodeList = xmlQDBResponse.DocumentElement.SelectNodes("/*/table/records/record")

        ReDim QDBRecord(RecordNodeList.Count + 1, intFields)

        For i = 0 To intFields - 1
            QDBRecord(0, i) = FieldNodeList(i).SelectSingleNode("label").InnerText
        Next i

        For i = 1 To RecordNodeList.Count
            For j = 0 To intFields - 1
                On Error Resume Next
                strFieldValue = RecordNodeList(i - 1).SelectSingleNode("f[" & CStr(j) & "]").InnerText
                Select Case FieldNodeList(j).SelectSingleNode("@base_type").InnerText
                    Case "float"
                        If strFieldValue <> "" Then
                            QDBRecord(i, j) = makeDouble(strFieldValue)
                        End If
                    Case "text"
                        QDBRecord(i, j) = Replace(strFieldValue, Chr(10), vbCrLf)
                    Case "bool"
                        QDBRecord(i, j) = CBool(strFieldValue)
                    Case "int64"
                        If strFieldValue <> "" Then
                            If FieldNodeList(j).SelectSingleNode("@field_type").InnerText <> "date" Then
                                QDBRecord(i, j) = CDbl(strFieldValue) / MILLISECONDS_IN_A_DAY
                            Else
                                QDBRecord(i, j) = int64ToDate(strFieldValue)
                            End If
                        End If
                    Case "int32"
                        If FieldNodeList(j).SelectSingleNode("@field_type").InnerText = "userid" Then
                            On Error Resume Next
                            Dim tempLong As Integer
                            tempLong = CInt(strFieldValue)
                            If Err.Number = 0 Then
                                QDBRecord(i, j) = xmlQDBResponse.SelectSingleNode("/*/table/lusers/luser[@id='" & strFieldValue & "']").InnerText
                            Else
                                QDBRecord(i, j) = strFieldValue
                            End If
                            On Error GoTo 0
                        Else
                            QDBRecord(i, j) = CLng(strFieldValue)
                        End If
                End Select
            Next j
        Next i
        DoQueryAsArray = QDBRecord
    End Function

    Public Function APIXMLPost(ByVal dbid As String, ByVal action As String, ByRef xmlQDBRequest As XmlDocument, ByVal useHTTPS As Boolean) As XmlDocument

        Dim script As String
        Dim content As String
        Dim req As HttpWebRequest
        Dim resp As HttpWebResponse
        Dim xmlStream As Stream
        Dim xmlTxtReader As XmlTextReader
        Dim xmlDoc As XmlDocument

        script = QDBHost & "/db/" & dbid & "?act=" & action
        If useHTTPS Then
            script = "https://" & script
        Else
            script = "http://" & script
        End If
        content = "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" & xmlQDBRequest.OuterXml
        req = CType(WebRequest.Create(script), HttpWebRequest)
        req.ContentType = "text/xml"
        req.Method = "POST"
        Dim byteRequestArray As Byte() = Encoding.UTF8.GetBytes(content)
        req.ContentLength = byteRequestArray.Length
        Dim reqStream As Stream = req.GetRequestStream()
        reqStream.Write(byteRequestArray, 0, byteRequestArray.Length)
        resp = CType(req.GetResponse(), HttpWebResponse)
        reqStream.Close()
        'create a new stream that can be placed into an XmlTextReader
        xmlStream = resp.GetResponseStream()
        xmlTxtReader = New XmlTextReader(xmlStream)
        xmlTxtReader.XmlResolver = Nothing
        'create a new Xml document
        xmlDoc = New XmlDocument
        xmlDoc.Load(xmlTxtReader)
        xmlStream.Close()
        On Error Resume Next
        errorcode = CInt(resp.Headers("QUICKBASE-ERRCODE"))
        ticket = xmlDoc.DocumentElement.SelectSingleNode("/*/ticket").InnerText
        errortext = xmlDoc.DocumentElement.SelectSingleNode("/*/errtext").InnerText
        If xmlDoc.DocumentElement.SelectSingleNode("/*/errdetail") Is Nothing Then
            errordetail = xmlDoc.DocumentElement.SelectSingleNode("/*/errtext").InnerText
        Else
            errordetail = xmlDoc.DocumentElement.SelectSingleNode("/*/errdetail").InnerText
        End If
        On Error GoTo 0
        If errorcode <> 0 Then
            Err.Raise(vbObjectError + CInt(errorcode), "QuickBase.QuickBaseClient", script & ": " & errordetail)
        End If


        APIXMLPost = xmlDoc
    End Function
    Public Function APIHTMLPost(ByVal dbid As String, ByVal action As String, ByRef xmlQDBRequest As XmlDocument, ByVal useHTTPS As Boolean) As String

        Dim script As String


        script = "/db/" & dbid & "?act=" & action
        APIHTMLPost = HTTPPost(QDBHost, useHTTPS, script, "text/xml", xmlQDBRequest.OuterXml, "")

    End Function
    Private Function HTTPPost(ByVal QDBHost As String, ByVal useHTTPS As Boolean, ByVal script As String, ByVal contentType As String, ByVal content As String, ByVal fileName As String) As String
        Dim url As String
        Dim Client As WebClient

        Client = New WebClient
        Client.Headers.Add("Content-Type", contentType)
        url = QDBHost & script
        If useHTTPS Then
            url = "https://" & url
        Else
            url = "http://" & url
        End If
        Dim byteRequestArray As Byte() = Encoding.UTF8.GetBytes(content)

        Dim byteResponseArray As Byte() = Client.UploadData(url, "POST", byteRequestArray)
        If fileName = "" Then
            HTTPPost = Encoding.UTF8.GetString(byteResponseArray)
        Else
            'check if write file exists 
            If File.Exists(Path:=fileName) Then
                'delete file
                File.Delete(Path:=fileName)
            End If

            'create a fileStream instance to pass to BinaryWriter object
            Dim fsWrite As FileStream
            fsWrite = New FileStream(Path:=fileName, _
                mode:=FileMode.CreateNew, access:=FileAccess.Write)

            'create binary writer instance
            Dim bWrite As BinaryWriter
            bWrite = New BinaryWriter(output:=fsWrite)
            'write bytes out 
            bWrite.Write(byteResponseArray, 0, byteResponseArray.Length)


            'close the writer 
            bWrite.Close()

            fsWrite.Close()


            HTTPPost = fileName
        End If

    End Function
    Public Function setAppToken(ByVal aapptoken As String) As String
        apptoken = aapptoken
        setAppToken = apptoken
    End Function
    Public Function getServer() As String
        getServer = QDBHost
    End Function

    Public Function getTicket() As String
        If ticket = "" Then
            Dim xmlQDBRequest As XmlDocument

            xmlQDBRequest = InitXMLRequest()
            Call APIXMLPost("main", "API_Authenticate", xmlQDBRequest, useHTTPS)
        End If
        getTicket = ticket
    End Function

    Public Function InitXMLRequest() As XmlDocument
        Dim xmlQDBRequest As New XmlDocument
        Dim Root As XmlElement

        Root = xmlQDBRequest.CreateElement("qdbapi")
        xmlQDBRequest.AppendChild(Root)
        If Len(ticket) <> 0 Then
            addParameter(xmlQDBRequest, "ticket", ticket)
        ElseIf UserName <> "" Then
            addParameter(xmlQDBRequest, "username", UserName)
            addParameter(xmlQDBRequest, "password", Password)
        End If
        If Len(apptoken) <> 0 Then
            addParameter(xmlQDBRequest, "apptoken", apptoken)
        End If
        InitXMLRequest = xmlQDBRequest
    End Function

    Public Sub addParameter(ByRef xmlQDBRequest As XmlDocument, ByVal Name As String, ByVal Value As String)
        Dim Root As XmlElement
        Dim ElementNode As XmlNode
        Dim TextNode As XmlNode

        Root = xmlQDBRequest.DocumentElement
        ElementNode = xmlQDBRequest.CreateNode(XmlNodeType.Element, Name, "")
        TextNode = xmlQDBRequest.CreateNode(XmlNodeType.Text, "", "")
        TextNode.InnerText = Value
        ElementNode.AppendChild(TextNode)
        Root.AppendChild(ElementNode)
        Root = Nothing
        ElementNode = Nothing
        TextNode = Nothing
    End Sub

    Public Sub addParameterWithAttribute(ByRef xmlQDBRequest As XmlDocument, ByVal Name As String, ByVal AttributeName As String, ByVal AttributeValue As String, ByVal Value As String)
        Dim Root As XmlElement
        Dim ElementNode As XmlNode
        Dim TextNode As XmlNode
        Dim Attribute As XmlAttribute

        Root = xmlQDBRequest.DocumentElement
        ElementNode = xmlQDBRequest.CreateNode(XmlNodeType.Element, Name, "")
        TextNode = xmlQDBRequest.CreateNode(XmlNodeType.Text, "", "")
        TextNode.InnerText = Value
        Attribute = xmlQDBRequest.CreateAttribute(AttributeName)
        Attribute.Value = AttributeValue
        ElementNode.Attributes.Append(Attribute)

        ElementNode.AppendChild(TextNode)
        Root.AppendChild(ElementNode)
        Root = Nothing
        ElementNode = Nothing
        TextNode = Nothing
    End Sub


    Public Sub addCDATAParameter(ByRef xmlQDBRequest As XmlDocument, ByVal Name As String, ByVal Value As String)
        Dim Root As XmlElement
        Dim ElementNode As XmlNode
        Dim CDATANode As XmlNode

        Root = xmlQDBRequest.DocumentElement
        ElementNode = xmlQDBRequest.CreateNode(XmlNodeType.Element, Name, "")
        CDATANode = xmlQDBRequest.CreateNode(XmlNodeType.CDATA, "", "")
        CDATANode.InnerText = Value
        ElementNode.AppendChild(CDATANode)
        Root.AppendChild(ElementNode)
        Root = Nothing
        ElementNode = Nothing
        CDATANode = Nothing
    End Sub

    Public Sub addFieldParameter(ByRef xmlQDBRequest As XmlDocument, ByVal attrName As String, ByVal Name As String, ByVal Value As Object)
        Dim Root As XmlElement
        Dim ElementNode As XmlNode
        Dim TextNode As XmlNode
        Dim attrField As XmlAttribute
        Dim attrFileName As XmlAttribute


        Root = xmlQDBRequest.DocumentElement
        ElementNode = xmlQDBRequest.CreateNode(XmlNodeType.Element, "field", "")
        attrField = xmlQDBRequest.CreateAttribute(attrName)
        attrField.Value = Name
        Call ElementNode.Attributes.SetNamedItem(attrField)


        If TypeName(Value) = "FileStream" Then
            attrFileName = xmlQDBRequest.CreateAttribute("filename")
            attrFileName.Value = DirectCast(Value, FileStream).Name
            Call ElementNode.Attributes.SetNamedItem(attrFileName)
        End If

        TextNode = xmlQDBRequest.CreateNode(XmlNodeType.Text, "", "")
        If TypeName(Value) = "FileStream" Then
            TextNode.InnerText = fileEncode64(DirectCast(Value, FileStream))
        Else
            TextNode.InnerText = CStr(Value)
        End If
        ElementNode.AppendChild(TextNode)

        Root.AppendChild(ElementNode)
        Root = Nothing
        ElementNode = Nothing
        attrField = Nothing
        TextNode = Nothing
    End Sub
    Function int64ToDate(ByVal int64 As String) As Date
        int64ToDate = Date.FromOADate(DAYS_BETWEEN_JAVASCRIPT_AND_MICROSOFT_DATE_REFERENCES + int64toDateCommon(int64))
    End Function
    Private Function int64toDateCommon(ByVal int64 As String) As Double
        If int64 = "" Then
            Exit Function
        End If
        Dim dblTemp As Double
        dblTemp = makeDouble(int64)
        If dblTemp <= -59011459200001.0# Then
            dblTemp = -59011459200000.0#
        ElseIf dblTemp > 255611376000000.0# Then
            dblTemp = 255611376000000.0#
        Else
            int64toDateCommon = (dblTemp / MILLISECONDS_IN_A_DAY)
        End If
    End Function

    Function int64ToDuration(ByVal int64 As String) As Date
        int64ToDuration = Date.FromOADate(int64toDateCommon(int64))
    End Function

    Function makeAlphaNumLowerCase(ByVal strString As String) As String
        Dim i As Integer
        Dim chrString As String

        makeAlphaNumLowerCase = ""
        For i = 1 To Len(strString)
            chrString = Mid(strString, i, 1)
            If System.Char.IsLetterOrDigit(chrString, 0) Then
                makeAlphaNumLowerCase = makeAlphaNumLowerCase & chrString
            Else
                makeAlphaNumLowerCase = makeAlphaNumLowerCase & "_"
            End If
        Next i
        makeAlphaNumLowerCase = LCase(makeAlphaNumLowerCase)
    End Function
    Public Sub setGMTOffset(ByVal offsetHours As Single)
        GMTOffset = offsetHours
    End Sub
    Public Sub setServer(ByVal strHost As String, ByVal HTTPS As Boolean)
        If strHost <> "" Then
            QDBHost = strHost
            useHTTPS = HTTPS
        Else
            QDBHost = "www.quickbase.com"
            useHTTPS = True
        End If
    End Sub

    Public Function getDBLastModified(ByVal dbid As String) As Date
        Dim qdbResponse As New XmlDocument
        Dim strInt64Time As String

        qdbResponse = GetDBInfo(dbid)
        strInt64Time = qdbResponse.DocumentElement.SelectSingleNode("/*/lastRecModTime").InnerText
        If Left(strInt64Time, 1) = "-" Then
            getDBLastModified = #1/1/1970#
        Else
            getDBLastModified = int64ToDate(qdbResponse.DocumentElement.SelectSingleNode("/*/lastRecModTime").InnerText)
        End If
    End Function

    Function getCompleteCSVSnapshot(ByVal dbid As String) As String
        Dim FieldNodeList As XmlNodeList
        Dim xmlNode As XmlNode
        Dim qdbResponse As New XmlDocument
        Dim clist As String = ""

        qdbResponse = GetSchema(dbid)
        FieldNodeList = qdbResponse.DocumentElement.SelectNodes("/*/table/fields/field/@id")
        For Each xmlNode In FieldNodeList
            clist = clist + xmlNode.InnerText & "."
        Next xmlNode
        getCompleteCSVSnapshot = GenResultsTable(dbid, "{'0'.CT.''}", clist, "", "csv")
    End Function
    Function getRecordAsArray(ByVal dbid As String, ByVal clist As String, ByVal ridFID As String, ByVal rid As String, ByRef QDBRecord(,) As Object) As String
        Dim xmlQDBResponse As New XmlDocument
        xmlQDBResponse = DoQuery(dbid, "{'" & ridFID & "'.EX.'" & rid & "'", clist, "", "")
        Dim strFieldValue As String
        Dim i As Integer
        Dim FieldNodeList As XmlNodeList
        Dim FieldDefNodeList As XmlNodeList

        FieldDefNodeList = xmlQDBResponse.DocumentElement.SelectNodes("/*/table/fields/field")
        FieldNodeList = xmlQDBResponse.DocumentElement.SelectNodes("/*/table/records/record/f")

        ReDim QDBRecord(1, FieldNodeList.Count - 1)
        For i = 0 To FieldNodeList.Count - 1
            QDBRecord(0, i) = FieldNodeList(i).SelectSingleNode("@id").InnerText
            strFieldValue = FieldNodeList(i).SelectSingleNode(".").InnerText

            Select Case FieldDefNodeList(i).SelectSingleNode("@base_type").InnerText
                Case "float"
                    If strFieldValue <> "" Then
                        QDBRecord(1, i) = makeDouble(strFieldValue)
                    End If
                Case "text"
                    QDBRecord(1, i) = Replace(strFieldValue, Chr(10), vbCrLf)
                Case "bool"
                    QDBRecord(1, i) = CBool(strFieldValue)
                Case "int64"
                    If strFieldValue <> "" Then
                        If FieldDefNodeList(i).SelectSingleNode("@field_type").InnerText <> "date" Then
                            QDBRecord(1, i) = CDbl(strFieldValue) / MILLISECONDS_IN_A_DAY
                        Else
                            QDBRecord(1, i) = int64ToDate(strFieldValue)
                        End If
                    End If
                Case "int32"
                    If FieldDefNodeList(i).SelectSingleNode("@field_type").InnerText = "userid" Then
                        On Error Resume Next
                        Dim tempLong As Integer
                        tempLong = CInt(strFieldValue)
                        If Err.Number = 0 Then
                            QDBRecord(1, i) = xmlQDBResponse.SelectSingleNode("/*/table/lusers/luser[@id='" & strFieldValue & "']").InnerText
                        Else
                            QDBRecord(1, i) = strFieldValue
                        End If
                        On Error GoTo 0
                    Else
                        QDBRecord(1, i) = CLng(strFieldValue)
                    End If
            End Select
        Next i
        getRecordAsArray = xmlQDBResponse.DocumentElement.SelectSingleNode("/*/table/records/record/f[@id=/*/table/fields/field[@role='modified']/@id]").InnerText
    End Function
    Public Function makeDouble(ByVal strString As String) As Double
        Dim i As Integer
        Dim chrString As String
        Dim strChar As String
        Dim resultString As String

        On Error Resume Next
        makeDouble = CDbl(strString)
        If Err.Number = 0 Then
            Exit Function
        End If
        On Error GoTo 0
        resultString = ""
        For i = 1 To Len(strString)
            strChar = Mid(strString, i, 1)
            If (((Not System.Char.IsLetter(strChar, 0)) And System.Char.IsLetterOrDigit(strChar, 0)) Or strChar = "." Or strChar = "-") Then
                resultString = resultString & strChar
            End If
        Next i
        On Error Resume Next
        makeDouble = CDbl(resultString)
        Exit Function
    End Function


    Function fileEncode64(ByVal fileToUpload As FileStream) As String
        Dim triplicate As Integer
        Dim i As Integer
        Dim outputText As String
        Dim fileLength As Integer
        Dim fileTriads As Integer
        Dim firstByte(0) As Byte
        Dim secondByte(0) As Byte
        Dim thirdByte(0) As Byte
        Dim fileRemainder As Integer

        fileLength = CInt(fileToUpload.Length)
        fileRemainder = CInt(fileLength Mod 3)
        fileTriads = fileLength \ 3
        If fileRemainder > 0 Then
            outputText = Space((fileTriads + 1) * 4)
        Else
            outputText = Space(fileTriads * 4)
        End If


        For i = 0 To fileTriads - 1             ' loop through octets
            'build 24 bit triplicate
            fileToUpload.Read(firstByte, 0, 1)
            fileToUpload.Read(secondByte, 0, 1)
            fileToUpload.Read(thirdByte, 0, 1)

            triplicate = (CInt(firstByte(0)) * 65536) + (CInt(secondByte(0)) * CInt(256)) + CInt(thirdByte(0))
            'extract four 6 bit quartets from triplicate
            Mid(outputText, (i * 4) + 1) = Mid(Map64, (triplicate \ 262144) + 1, 1) & Mid(Map64, ((triplicate And 258048) \ 4096) + 1, 1) & Mid(Map64, ((triplicate And 4032) \ 64) + 1, 1) & Mid(Map64, (triplicate And 63) + 1, 1)
        Next                                                    ' next octet
        Select Case fileRemainder
            Case 1
                fileToUpload.Read(firstByte, 0, 1)
                triplicate = (firstByte(0) * 65536)
                Mid(outputText, (i * 4) + 1) = Mid(Map64, (triplicate \ 262144) + 1, 1) & Mid(Map64, ((triplicate And 258048) \ 4096) + 1, 1) & "="
            Case 2
                fileToUpload.Read(firstByte, 0, 1)
                fileToUpload.Read(secondByte, 0, 1)
                triplicate = (firstByte(0) * 65536) + (secondByte(0) * 256)
                Mid(outputText, (i * 4) + 1) = Mid(Map64, (triplicate \ 262144) + 1, 1) & Mid(Map64, ((triplicate And 258048) \ 4096) + 1, 1) & Mid(Map64, ((triplicate And 4032) \ 64) + 1, 1) & "="
        End Select
        fileEncode64 = outputText
    End Function

    Function makeValidFilename(ByVal strString As String) As String
        Dim i As Integer
        Dim byteChar As String
        makeValidFilename = ""
        For i = 1 To Len(strString)
            byteChar = Mid(strString, i, 1)
            If byteChar = "\" Or byteChar = "/" Or _
               byteChar = ":" Or byteChar = "*" Or _
               Asc(byteChar) = 63 Or byteChar = """" Or _
               byteChar = "<" Or byteChar = ">" Or _
               byteChar = "|" Or byteChar = "'" _
            Then
                makeValidFilename = makeValidFilename & "_"
            Else
                makeValidFilename = makeValidFilename + byteChar
            End If
        Next i
    End Function
    Public Function getServerStatus() As XmlDocument
        Dim xmlQDBRequest As XmlDocument
        xmlQDBRequest = InitXMLRequest()
        getServerStatus = APIXMLPost("main", "API_OBStatus", xmlQDBRequest, useHTTPS)
    End Function
    Public Function GetNumRecords(ByVal dbid As String) As Integer
        Dim xmlQDBRequest As XmlDocument

        xmlQDBRequest = InitXMLRequest()
        GetNumRecords = CInt(APIXMLPost(dbid, "API_GetNumRecords", xmlQDBRequest, useHTTPS).SelectSingleNode("/*/num_records").InnerText)
    End Function

    Public Function GetDBPage(ByVal dbid As String, ByVal page As String) As String
        Dim xmlQDBRequest As XmlDocument

        xmlQDBRequest = InitXMLRequest()
        If CStr(Val(page)) = page Then
            addParameter(xmlQDBRequest, "pageid", page)
        Else
            addParameter(xmlQDBRequest, "pagename", page)
        End If
        GetDBPage = APIXMLPost(dbid, "API_GetDBPage", xmlQDBRequest, useHTTPS).SelectSingleNode("/*/pagebody").InnerText
    End Function

    Public Function AddReplaceDBPage(ByVal dbid As String, ByVal page As String, ByVal pagetype As String, ByVal pagebody As String) As XmlDocument
        Dim xmlQDBRequest As XmlDocument

        xmlQDBRequest = InitXMLRequest()
        If CStr(Val(page)) = page Then
            addParameter(xmlQDBRequest, "pageid", page)
        Else
            addParameter(xmlQDBRequest, "pagename", page)
        End If
        addParameter(xmlQDBRequest, "pagetype", pagetype)
        addParameter(xmlQDBRequest, "pagebody", pagebody)
        AddReplaceDBPage = APIXMLPost(dbid, "API_AddReplaceDBPage", xmlQDBRequest, useHTTPS)
    End Function

    Public Function ListDBPages(ByVal dbid As String) As XmlDocument
        Dim xmlQDBRequest As XmlDocument

        xmlQDBRequest = InitXMLRequest()
        ListDBPages = APIXMLPost(dbid, "API_ListDBPages", xmlQDBRequest, useHTTPS)
    End Function
    Public Function PurgeRecords(ByVal dbid As String, ByVal query As String) As XmlDocument
        Dim xmlQDBRequest As XmlDocument
        xmlQDBRequest = InitXMLRequest()
        If Left(query, 1) = "{" And Right(query, 1) = "}" Then
            addParameter(xmlQDBRequest, "query", query)
        ElseIf CStr(Val(query)) = query Then
            addParameter(xmlQDBRequest, "qid", query)
        Else
            addParameter(xmlQDBRequest, "qname", query)
        End If


        PurgeRecords = APIXMLPost(dbid, "API_PurgeRecords", xmlQDBRequest, useHTTPS)

    End Function
    Public Sub New(ByVal uid As String, ByVal pwd As String)
        UserName = uid
        Password = pwd
        GMTOffset = -7
    End Sub

    Public Function RenameApp(ByVal dbid As String, ByVal newappname As String) As Boolean
        Dim xmlQDBRequest As XmlDocument
        xmlQDBRequest = InitXMLRequest()
        addParameter(xmlQDBRequest, "newappname", newappname)
        On Error GoTo exception
        Call APIXMLPost(dbid, "API_RenameApp", xmlQDBRequest, useHTTPS)
        RenameApp = True
        Exit Function
exception:
        RenameApp = False
    End Function

    Public Function GetDBvar(ByVal dbid As String, ByVal varname As String) As String
        Dim xmlQDBRequest As XmlDocument
        xmlQDBRequest = InitXMLRequest()
        addParameter(xmlQDBRequest, "varname", varname)
        xmlQDBRequest = APIXMLPost(dbid, "API_GetDBvar", xmlQDBRequest, useHTTPS)
        GetDBvar = xmlQDBRequest.DocumentElement.SelectSingleNode("/*/value").InnerText
    End Function

    Public Function CreateTable(ByVal application_dbid As String, ByVal tname As String, ByVal pnoun As String) As String
        Dim xmlQDBRequest As XmlDocument
        xmlQDBRequest = InitXMLRequest()
        addParameter(xmlQDBRequest, "pnoun", pnoun)
        addParameter(xmlQDBRequest, "tname", tname)
        xmlQDBRequest = APIXMLPost(application_dbid, "API_CreateTable", xmlQDBRequest, useHTTPS)
        CreateTable = xmlQDBRequest.DocumentElement.SelectSingleNode("/*/newdbid").InnerText
    End Function

    Public Function AddUserToRole(ByVal dbid As String, ByVal userid As String, ByVal roleid As String) As Boolean
        Dim xmlQDBRequest As XmlDocument
        xmlQDBRequest = InitXMLRequest()
        addParameter(xmlQDBRequest, "userid", userid)
        addParameter(xmlQDBRequest, "roleid", roleid)
        On Error GoTo exception
        Call APIXMLPost(dbid, "API_AddUserToRole", xmlQDBRequest, useHTTPS)
        AddUserToRole = True
        Exit Function
exception:
        AddUserToRole = False
    End Function

    Public Function GetOneTimeTicket() As String
        Dim xmlQDBRequest As XmlDocument
        xmlQDBRequest = InitXMLRequest()
        xmlQDBRequest = APIXMLPost("main", "API_GetOneTimeTicket", xmlQDBRequest, useHTTPS)
        GetOneTimeTicket = xmlQDBRequest.DocumentElement.SelectSingleNode("/*/ticket").InnerText
    End Function

    Public Function FieldRemoveChoices(ByVal dbid As String, ByVal fid As String, ByVal ParamArray NameValues() As Object) As Integer
        Dim xmlQDBRequest As XmlDocument
        Dim firstfield As Integer
        Dim lastfield As Integer
        Dim i As Integer

        xmlQDBRequest = InitXMLRequest()
        addParameter(xmlQDBRequest, "fid", fid)

        lastfield = UBound(NameValues)
        firstfield = LBound(NameValues)
        For i = firstfield To lastfield
            addParameter(xmlQDBRequest, "choice", CStr(NameValues(i)))
        Next i

        xmlQDBRequest = APIXMLPost(dbid, "API_FieldRemoveChoices", xmlQDBRequest, useHTTPS)
        FieldRemoveChoices = CInt(xmlQDBRequest.DocumentElement.SelectSingleNode("/*/numremoved").InnerText)
    End Function

    Public Function DeleteField(ByVal dbid As String, ByVal fid As String) As Boolean
        Dim xmlQDBRequest As XmlDocument
        xmlQDBRequest = InitXMLRequest()
        addParameter(xmlQDBRequest, "fid", fid)
        On Error GoTo exception
        Call APIXMLPost(dbid, "API_DeleteField", xmlQDBRequest, useHTTPS)
        DeleteField = True
        Exit Function
exception:
        DeleteField = False
    End Function

    Public Function GenAddRecordForm(ByVal dbid As String, ByRef fieldValues As Hashtable) As String
        Dim xmlQDBRequest As XmlDocument
        xmlQDBRequest = InitXMLRequest()
        Dim fieldValueEnumerator As IDictionaryEnumerator = fieldValues.GetEnumerator()
        While fieldValueEnumerator.MoveNext()
            addParameterWithAttribute(xmlQDBRequest, "field", "name", fieldValueEnumerator.Key.ToString, fieldValueEnumerator.Value.ToString())
        End While
        GenAddRecordForm = APIHTMLPost(dbid, "API_GenAddRecordForm", xmlQDBRequest, useHTTPS)
    End Function

    Public Function ChangeUserRole(ByVal dbid As String, ByVal userid As String, ByVal roleid As String, ByVal newroleid As String) As Boolean
        Dim xmlQDBRequest As XmlDocument
        xmlQDBRequest = InitXMLRequest()
        addParameter(xmlQDBRequest, "userid", userid)
        addParameter(xmlQDBRequest, "roleid", roleid)
        addParameter(xmlQDBRequest, "newroleid", newroleid)
        On Error GoTo exception
        Call APIXMLPost(dbid, "API_ChangeUserRole", xmlQDBRequest, useHTTPS)
        ChangeUserRole = True
        Exit Function
exception:
        ChangeUserRole = False
    End Function

    Public Function SetDBvar(ByVal dbid As String, ByVal varname As String, ByVal value As String) As Boolean
        Dim xmlQDBRequest As XmlDocument
        xmlQDBRequest = InitXMLRequest()
        addParameter(xmlQDBRequest, "varname", varname)
        addParameter(xmlQDBRequest, "value", value)
        On Error GoTo exception
        Call APIXMLPost(dbid, "API_SetDBvar", xmlQDBRequest, useHTTPS)
        SetDBvar = True
        Exit Function
exception:
        SetDBvar = False
    End Function

    Public Function ProvisionUser(ByVal dbid As String, ByVal roleid As String, ByVal email As String, ByVal fname As String, ByVal lname As String) As String
        Dim xmlQDBRequest As XmlDocument
        xmlQDBRequest = InitXMLRequest()
        addParameter(xmlQDBRequest, "roleid", roleid)
        addParameter(xmlQDBRequest, "email", email)
        addParameter(xmlQDBRequest, "fname", fname)
        addParameter(xmlQDBRequest, "lname", lname)
        xmlQDBRequest = APIXMLPost(dbid, "API_ProvisionUser", xmlQDBRequest, useHTTPS)
        ProvisionUser = xmlQDBRequest.DocumentElement.SelectSingleNode("/*/userid").InnerText
    End Function

    Public Function GetRecordInfo(ByVal dbid As String, ByVal rid As String) As XmlDocument
        Dim xmlQDBRequest As XmlDocument
        xmlQDBRequest = InitXMLRequest()
        addParameter(xmlQDBRequest, "rid", rid)
        GetRecordInfo = APIXMLPost(dbid, "API_GetRecordInfo", xmlQDBRequest, useHTTPS)
    End Function

    Public Function UserRoles(ByVal dbid As String) As XmlDocument
        Dim xmlQDBRequest As XmlDocument
        xmlQDBRequest = InitXMLRequest()
        UserRoles = APIXMLPost(dbid, "API_UserRoles", xmlQDBRequest, useHTTPS)
    End Function

    Public Function GetUserRole(ByVal dbid As String, ByVal userid As String) As XmlDocument
        Dim xmlQDBRequest As XmlDocument
        xmlQDBRequest = InitXMLRequest()
        addParameter(xmlQDBRequest, "userid", userid)
        GetUserRole = APIXMLPost(dbid, "API_GetUserRole", xmlQDBRequest, useHTTPS)
    End Function

    Public Function RemoveUserFromRole(ByVal dbid As String, ByVal userid As String, ByVal roleid As String) As Boolean
        Dim xmlQDBRequest As XmlDocument
        xmlQDBRequest = InitXMLRequest()
        addParameter(xmlQDBRequest, "userid", userid)
        addParameter(xmlQDBRequest, "roleid", roleid)
        On Error GoTo exception
        Call APIXMLPost(dbid, "API_RemoveUserFromRole", xmlQDBRequest, useHTTPS)
        RemoveUserFromRole = True
        Exit Function
exception:
        RemoveUserFromRole = False
    End Function

    Public Function GetRecordAsHTML(ByVal dbid As String, ByVal rid As String) As String
        Dim xmlQDBRequest As XmlDocument
        xmlQDBRequest = InitXMLRequest()
        addParameter(xmlQDBRequest, "rid", rid)
        GetRecordAsHTML = APIHTMLPost(dbid, "API_GetRecordAsHTML", xmlQDBRequest, useHTTPS)
    End Function

    Public Function SendInvitation(ByVal dbid As String, ByVal userid As String) As Boolean
        Dim xmlQDBRequest As XmlDocument
        xmlQDBRequest = InitXMLRequest()
        addParameter(xmlQDBRequest, "userid", userid)
        On Error GoTo exception
        Call APIXMLPost(dbid, "API_SendInvitation", xmlQDBRequest, useHTTPS)
        SendInvitation = True
        Exit Function
exception:
        SendInvitation = True
    End Function

    Public Function GetUserInfo(ByVal email As String) As XmlDocument
        Dim xmlQDBRequest As XmlDocument
        xmlQDBRequest = InitXMLRequest()
        addParameter(xmlQDBRequest, "email", email)
        GetUserInfo = APIXMLPost("main", "API_GetUserInfo", xmlQDBRequest, useHTTPS)
    End Function

    Public Function GetRoleInfo(ByVal dbid As String) As XmlDocument
        Dim xmlQDBRequest As XmlDocument
        xmlQDBRequest = InitXMLRequest()
        GetRoleInfo = APIXMLPost(dbid, "API_GetRoleInfo", xmlQDBRequest, useHTTPS)
    End Function

End Class