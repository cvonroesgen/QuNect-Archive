
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
    Private Class qdbVersion
        Public year As Integer
        Public major As Integer
        Public minor As Integer
    End Class
    Private qdbVer As qdbVersion = New qdbVersion
    Private Const AppName = "QuNectArchive"
    Private Const QuNectODBCParentDBID = "bcks8a7y3"
    Private cmdLineArgs() As String
    Private automode As Boolean = False
    Private config As String 'dbid of the table to archive, fid of the field QuNect Archived, fids to keep, dbid of the archive table
    Private schema As XmlDocument
    Private configHash As New Hashtable
    Private reportNameToQid As New Hashtable
    Private fieldLabelsToFIDs As New Hashtable
    Private fidsToFieldLabels As New Hashtable
    Private recordsPerArchive = 90
    Private dbidToAppName As New Dictionary(Of String, String)

    Sub showHideControls()
        cmbPassword.Visible = txtUsername.Text.Length > 0
        txtPassword.Visible = cmbPassword.Visible And cmbPassword.SelectedIndex <> 0
        txtServer.Visible = txtPassword.Visible And txtPassword.Text.Length > 0
        lblServer.Visible = txtServer.Visible
        lblAppToken.Visible = cmbPassword.Visible And cmbPassword.SelectedIndex = 1
        txtAppToken.Visible = lblAppToken.Visible
        btnAppToken.Visible = lblAppToken.Visible
        btnUserToken.Visible = cmbPassword.Visible And cmbPassword.SelectedIndex = 2
        ckbDetectProxy.Visible = txtServer.Text.Length > 0 And txtServer.Visible
        cmbPassword.Visible = txtUsername.Text.Length > 0
        txtPassword.Visible = cmbPassword.SelectedIndex > 0
        txtServer.Visible = txtUsername.Text.Length > 0 And txtPassword.Text.Length > 0 And cmbPassword.SelectedIndex > 0
        lblServer.Visible = txtServer.Visible
        txtAppToken.Visible = txtUsername.Text.Length > 0 And cmbPassword.SelectedIndex = 1 And txtPassword.Text.Length > 0 And txtServer.Text.Length > 0
        lblAppToken.Visible = txtAppToken.Visible
        btnAppToken.Visible = txtAppToken.Visible
        btnUserToken.Visible = txtUsername.Text.Length > 0 And cmbPassword.SelectedIndex = 2
        btnListTables.Visible = txtUsername.Text.Length > 0 And cmbPassword.SelectedIndex > 0 And txtPassword.Text.Length > 0 And txtServer.Text.Length > 0

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
        recordsPerArchive = GetSetting(AppName, "archive", "bytesPerRecord", 90)
        txtBackupFolder.Text = GetSetting(AppName, "location", "path")
        cmdLineArgs = System.Environment.GetCommandLineArgs()
        If cmdLineArgs.Length > 1 Then
            If cmdLineArgs(1) = "auto" Then
                automode = True
                Try
                    listTables()
                    archive(False)
                    Me.Close()
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.OkOnly, AppName)
                End Try


            End If
        End If
        Dim myBuildInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(Application.ExecutablePath)
        Me.Text = "QuNect Archive " & myBuildInfo.ProductVersion
        showHideControls()
    End Sub

    Private Sub txtUsername_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUsername.TextChanged
        SaveSetting(AppName, "Credentials", "username", txtUsername.Text)
        showHideControls()
    End Sub

    Private Sub txtPassword_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPassword.TextChanged
        SaveSetting(AppName, "Credentials", "password", txtPassword.Text)
        showHideControls()
    End Sub

    Private Sub btnListTables_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnListTables.Click
        Try
            listTables()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, AppName)
        End Try

    End Sub
    Private Function buildConnectionString(appDBID As String) As String
        If txtPassword.Text.Contains(";") Then
            Throw New System.Exception("Although Quick Base allows semicolons in passwords the ODBC standard does not permit semicolons." & vbCrLf & "Please change your Quick Base password to eliminate semicolons or use a Quick Base user token instead of a password.")
            Return ""
        End If
        buildConnectionString = "FIELDNAMECHARACTERS=all;uid=" & txtUsername.Text
        buildConnectionString &= ";pwd=" & txtPassword.Text
        buildConnectionString &= ";driver={QuNect ODBC for QuickBase};"
        buildConnectionString &= ";quickbaseserver=" & txtServer.Text
        If appDBID <> "" Then
            buildConnectionString &= ";appid=" & appDBID
            buildConnectionString &= ";appname=" & tvAppsTables.SelectedNode.Parent.Text
        End If

        If ckbDetectProxy.Checked Then
            buildConnectionString &= ";DETECTPROXY=1"
        End If



        If cmbPassword.SelectedIndex = 0 Then
            cmbPassword.Focus()
            Throw New System.Exception("Please indicate whether you are using a password or a user token.")
            Return ""
        ElseIf cmbPassword.SelectedIndex = 1 Then
            buildConnectionString &= ";PWDISPASSWORD=1"
            buildConnectionString &= ";APPTOKEN=" & txtAppToken.Text
        Else
            buildConnectionString &= ";PWDISPASSWORD=0"
        End If
#If DEBUG Then
        'buildConnectionString &= ";LOGSQL=1"
#End If

    End Function
    Sub listTablesFromGetSchema(tables As DataTable, appToDBID As Dictionary(Of String, String))

        tvAppsTables.BeginUpdate()
        tvAppsTables.Nodes.Clear()
        tvAppsTables.ShowNodeToolTips = True
        Dim dbName As String
        Dim applicationName As String = ""
        Dim prevAppName As String = ""
        Dim dbid As String
        pleaseWait.pb.Value = 0
        pleaseWait.pb.Visible = True
        pleaseWait.pb.Maximum = tables.Rows.Count
        Dim getDBIDfromdbName As New Regex("([a-z0-9~]+)$")
        Dim dbidCollection As New Collection

        Dim i As Integer
        dbidToAppName.Clear()
        For i = 0 To tables.Rows.Count - 1
            pleaseWait.pb.Value = i
            Application.DoEvents()
            dbName = tables.Rows(i)(2)
            applicationName = tables.Rows(i)(0)
            Dim dbidMatch As Match = getDBIDfromdbName.Match(dbName)
            dbid = dbidMatch.Value
            If Not dbidToAppName.ContainsKey(dbid) Then
                dbidToAppName.Add(dbid, applicationName)
            End If
            If applicationName <> prevAppName Then

                Dim appNode As TreeNode = tvAppsTables.Nodes.Add(applicationName)
                appNode.Tag = appToDBID(applicationName)
                prevAppName = applicationName
            End If
            Dim tableName As String = dbName

            Dim tableNode As TreeNode = tvAppsTables.Nodes(tvAppsTables.Nodes.Count - 1).Nodes.Add(tableName)
            tableNode.Tag = dbid

        Next
        dbid = GetSetting(AppName, "archive", "dbid", "")
        For i = 0 To tvAppsTables.Nodes.Count - 1
            Dim j As Integer
            For j = 0 To tvAppsTables.Nodes(i).Nodes.Count - 1
                If tvAppsTables.Nodes(i).Nodes(j).Tag = dbid Then
                    tvAppsTables.SelectedNode = tvAppsTables.Nodes(i).Nodes(j)
                    Exit For
                End If
            Next
        Next
        pleaseWait.pb.Visible = False
        tvAppsTables.EndUpdate()
        pleaseWait.pb.Value = 0
        If tvAppsTables.SelectedNode IsNot Nothing Then
            displayFields(tvAppsTables.SelectedNode.FullPath())
        End If
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub listTables()
        Me.Cursor = Cursors.WaitCursor
        tvAppsTables.Visible = True
        Try
            Dim connectionString As String = buildConnectionString("")
            Dim quNectConn As OdbcConnection = New OdbcConnection(connectionString)
            quNectConn.Open()
            Dim ver As String = quNectConn.ServerVersion
            Dim m As Match = Regex.Match(ver, "\d+\.(\d+)\.(\d+)\.(\d+)")
            qdbVer.year = CInt(m.Groups(1).Value)
            qdbVer.major = CInt(m.Groups(2).Value)
            qdbVer.minor = CInt(m.Groups(3).Value)

            If qdbVer.year < 20 Then
                MsgBox("You are running the 20" & qdbVer.year & " version of QuNect ODBC for QuickBase. Please install the latest version from https://qunect.com/download/QuNect.exe", MsgBoxStyle.OkOnly, AppName)
                quNectConn.Dispose()
                Me.Cursor = Cursors.Default
                Exit Sub
            ElseIf qdbVer.major = 8 And qdbVer.minor < 78 Then
                MsgBox("Please install the latest version of QuNect ODBC for QuickBase from https://qunect.com/download/QuNect.exe", MsgBoxStyle.OkOnly, AppName)
                quNectConn.Dispose()
                Me.Cursor = Cursors.Default
                Exit Sub
            End If

            Dim tableOfTables As DataTable = quNectConn.GetSchema("Tables")
            Dim appToDBID As New Dictionary(Of String, String)
            Using command As OdbcCommand = New OdbcCommand("SELECT * FROM apps", quNectConn)
                Dim dr As OdbcDataReader = command.ExecuteReader()
                While dr.Read
                    appToDBID.Add(dr.GetString(0), dr.GetString(2))
                End While
            End Using
            listTablesFromGetSchema(tableOfTables, appToDBID)
            Me.Cursor = Cursors.Default
            quNectConn.Close()
            quNectConn.Dispose()
        Catch excpt As Exception
            Me.Cursor = Cursors.Default
            If excpt.Message.Contains("Data source name not found") Then
                MsgBox("Please install QuNect ODBC for QuickBase from http://qunect.com/download/QuNect.exe and try again.", MsgBoxStyle.OkOnly, AppName)
            Else
                MsgBox(excpt.Message, MsgBoxStyle.OkOnly, AppName)
            End If
            Exit Sub
        End Try
    End Sub



    Private Sub txtServer_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtServer.TextChanged
        SaveSetting(AppName, "Credentials", "server", txtServer.Text)
        showHideControls()
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
    Private Function getQidFromReportName(reportName As String)
        Dim m As Match = Regex.Match(reportName, "^[^~]+~(\d+)$")
        If m.Groups.Count > 1 Then
            Return m.Groups(1).Value
        End If
        Return ""
    End Function
    Private Sub displayFields(ByVal appTable As String)
        Me.Cursor = Cursors.WaitCursor
        Try
            Dim appDBID As String = tvAppsTables.SelectedNode.Parent.Tag
            Dim tableToArchive As String = tvAppsTables.SelectedNode.Text
            Dim m As Match = Regex.Match(tableToArchive, "^[^:]+: (.*) [a-kmnp-z2-9]+$")
            tableToArchive = m.Groups(1).Value
            Dim connectionString As String = buildConnectionString(appDBID)
            Dim quNectConn As OdbcConnection = New OdbcConnection(connectionString)
            quNectConn.Open()
            Dim i As Integer
            Dim dbid As String = appTable.Substring(appTable.LastIndexOf(" ") + 1)
            SaveSetting(AppName, "archive", "dbid", dbid)
            Dim quNectCmd As OdbcCommand = Nothing
            Dim dr As OdbcDataReader
            Try
                quNectCmd = New OdbcCommand("SELECT * FROM " & appDBID & "~vars WHERE Name = 'QuNectArchive%'", quNectConn)
                dr = quNectCmd.ExecuteReader()
            Catch excpt As Exception
                If quNectCmd IsNot Nothing Then
                    quNectCmd.Dispose()
                End If
                Exit Sub
            End Try
            configHash.Clear()
            If dr.HasRows Then
                dr.Read()
                config = dr.GetString(1)
                Dim configs As String() = config.Split(vbCrLf)
                For i = 0 To configs.Length - 1
                    configHash.Add(configs(i), configs(i))
                Next
            End If
            'need to focus in on one app
            Dim tableOfTables As DataTable = quNectConn.GetSchema("Views")
            lstReports.Items.Clear()
            reportNameToQid.Clear()
            Dim qid As String = GetSetting(AppName, "archive", "qid")
            For i = 0 To tableOfTables.Rows.Count - 1
                Application.DoEvents()
                Dim reportName As String = tableOfTables.Rows(i)(2)
                m = Regex.Match(reportName, "^([^:]+)")
                Dim tableName As String = m.Groups(1).Value

                If tableToArchive <> tableName Then
                    Continue For
                End If
                Try
                    Dim lastListBoxItem As Integer = lstReports.Items.Add(reportName)
                    Dim thisQid As String = getQidFromReportName(reportName)
                    If thisQid = qid Then
                        lstReports.SelectedIndex = lastListBoxItem
                    End If
                    reportNameToQid.Add(reportName, thisQid)
                Catch excpt As Exception
                    Continue For
                End Try
            Next

            Try
                fieldLabelsToFIDs.Clear()
                fidsToFieldLabels.Clear()
                If quNectCmd IsNot Nothing Then
                    quNectCmd.Dispose()
                End If
                quNectConn.Close()
                quNectConn = New OdbcConnection(connectionString)
                quNectConn.Open()
                Dim sql As String = "SELECT COLUMN_NAME, fid FROM " & dbid & "~fields WHERE (append_only = 0 OR field_type='email') and isunique = 0 and required = 0 and base_type = 'text' and field_type NOT IN ('userid', 'file') And mastag = '' And role = '' And mode = ''"

                quNectCmd = quNectConn.CreateCommand()
                quNectCmd.CommandText = sql
                dr = quNectCmd.ExecuteReader()


                While dr.Read()
                    Dim label As String = dr.GetString(0)
                    Dim fid As String = dr.GetString(1)
                    fieldLabelsToFIDs.Add(label, fid)
                    fidsToFieldLabels.Add(fid, label)
                End While
            Catch labelDupe As Exception
                If quNectCmd IsNot Nothing Then
                    quNectCmd.Dispose()
                End If
                Throw New ArgumentException("Two fields with the same name: '" & dr.GetString(0) & "'")
            End Try


            lstArchiveFields.Items.Clear()
            lstFieldsToKeep.Items.Clear()

            Dim fidsToArchive As New Hashtable()
            Dim fidsToArchiveSetting As String = GetSetting(AppName, "archive", "fids")
            Dim fids As String() = fidsToArchiveSetting.Split(".")
            For Each fid As String In fids
                fidsToArchive.Add(fid, fid)
            Next

            If fidsToArchive.Count > 0 Then

                For Each field As DictionaryEntry In fieldLabelsToFIDs
                    Dim label As String = field.Key
                    Dim fid As String = field.Value
                    If fidsToArchive.ContainsKey(fid) Then
                        lstArchiveFields.Items.Add(label)
                    Else
                        lstFieldsToKeep.Items.Add(label)
                    End If
                Next
            Else
                For Each field As DictionaryEntry In fieldLabelsToFIDs
                    Dim label As String = field.Key
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
        showHideControls()
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
        Try
            archive(False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, AppName)
        End Try
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
        Else
            Dim fieldLabel As Object
            Dim fids As String = ""
            Dim period As String = ""
            For Each fieldLabel In lstArchiveFields.Items
                fids &= period & fieldLabelsToFIDs(fieldLabel.ToString())
                period = "."
            Next
            SaveSetting(AppName, "archive", "fids", fids)
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
        Else
            SaveSetting(AppName, "archive", "qid", getQidFromReportName(lstReports.Items(lstReports.SelectedIndex)))
        End If
        Dim refFID As String = ""
        Dim dbid As String = ""
        Dim dbName As String = tvAppsTables.SelectedNode.FullPath().ToString()
        Dim i As Integer
        dbid = dbName.Substring(dbName.LastIndexOf(" ") + 1)
        Dim archivedbid As String = ""
        Dim fileFID As String = ""
        Dim keyfid As String = "3"
        Dim keyFieldLabel As String = ""

        Dim connectionString As String = buildConnectionString("")
        Dim quNectConn As OdbcConnection = New OdbcConnection(connectionString)
        quNectConn.Open()
        Dim quNectCmd As OdbcCommand = Nothing

        Dim dr As OdbcDataReader
        Try
            quNectCmd = New OdbcCommand("SELECT COLUMN_NAME, fid, iskey FROM " & dbid & "~fields", quNectConn)
            dr = quNectCmd.ExecuteReader()
        Catch excpt As Exception
            If quNectCmd IsNot Nothing Then
                quNectCmd.Dispose()
            End If
            Exit Sub
        End Try

        fieldLabelsToFIDs.Clear()
        fidsToFieldLabels.Clear()
        Try
            While dr.Read()
                Dim label As String = dr("COLUMN_NAME")
                Dim fid As String = dr("fid")
                Dim iskey As Boolean = dr("iskey")
                fieldLabelsToFIDs.Add(label, fid)
                fidsToFieldLabels.Add(fid, label)
                If iskey Then
                    keyfid = fid
                    keyFieldLabel = label
                End If
            End While
        Catch labelDupe As Exception
            Throw New ArgumentException("Two fields with the same name: '" & dr.GetString(0) & "'")
        End Try




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

        Dim formulaDBID As String = "dbid()"
        Dim parentDBID As String = tvAppsTables.SelectedNode.Parent.Tag

        'here we need to check to see if the backup table exists
        'we need to create an archive table
        Dim archiveTableName As String = "QuNect Archive for " & dbid
        Dim archiveTableAlias As String = archiveTableName
        archiveTableAlias = archiveTableAlias.ToUpper()
        archiveTableAlias = Regex.Replace(archiveTableAlias, "[^A-Z0-9]", "_")
        archiveTableAlias = "_DBID_" & archiveTableAlias
        'first need to check to see if the archive table already exists

        Try
            If quNectCmd IsNot Nothing Then
                quNectCmd.Dispose()
            End If
            quNectCmd = New OdbcCommand("SELECT * FROM " & parentDBID & "~aliases where Name = '" & archiveTableAlias.ToLower() & "'", quNectConn)
            dr = quNectCmd.ExecuteReader()
        Catch excpt As Exception
            If quNectCmd IsNot Nothing Then
                quNectCmd.Dispose()
            End If
            Exit Sub
        End Try

        If dr.HasRows Then
            dr.Read()
            archivedbid = dr("Value")
        Else
            dr.Close()
            If Not countBytesOnly Then
                Try
                    If quNectCmd IsNot Nothing Then
                        quNectCmd.Dispose()
                    End If
                    Dim quNectCreateTableConn As OdbcConnection = New OdbcConnection(connectionString)
                    quNectCreateTableConn.Open()
                    quNectCmd = New OdbcCommand("CREATE TABLE """ & archiveTableName & " " & parentDBID & """ (""CSV Archive"" longVarBinary)", quNectCreateTableConn)
                    quNectCmd.ExecuteNonQuery()
                    quNectCmd = New OdbcCommand("SELECT @@newdbid", quNectCreateTableConn)
                    dr = quNectCmd.ExecuteReader()
                    dr.Read()
                    archivedbid = dr.GetString(0)
                    quNectCreateTableConn.Close()

                Catch excpt As Exception
                    If quNectCmd IsNot Nothing Then
                        quNectCmd.Dispose()
                    End If
                    Exit Sub
                End Try
            End If
        End If

        refFID = getRefFID(dbid, archiveTableAlias)
        If refFID = "" Then
            If Not countBytesOnly Then
                If quNectCmd IsNot Nothing Then
                    quNectCmd.Dispose()
                End If
                Dim quNectAlterTableConn As OdbcConnection = New OdbcConnection(connectionString)
                quNectAlterTableConn.Open()
                quNectCmd = New OdbcCommand("ALTER TABLE " & dbid & " ADD ""QuNect Archive Reference"" float", quNectAlterTableConn)
                quNectCmd.ExecuteNonQuery()
                quNectCmd = New OdbcCommand("ALTER TABLE " & dbid & " ADD CONSTRAINT doesNotMatter FOREIGN KEY (""QuNect Archive Reference"") REFERENCES """ & archiveTableName & """ (fid" & keyfid & ") ", quNectAlterTableConn)
                quNectCmd.ExecuteNonQuery()
                quNectAlterTableConn.Close()
                'Now we need to get the brand new reffid
                refFID = getRefFID(dbid, archiveTableAlias)
            End If
        End If

        Try
            If quNectCmd IsNot Nothing Then
                quNectCmd.Dispose()
            End If
            'here I should probably add a criteria to prevent archiving of records that were already archived
            Dim sql As String = "SELECT fid3 FROM ""All Columns Please " & lstReports.Items(lstReports.SelectedIndex) & """"
            If refFID <> "" Then
                sql &= " WHERE ""QuNect Archive Reference"" Is NULL"
            End If
            quNectCmd = New OdbcCommand(sql, quNectConn)
            dr = quNectCmd.ExecuteReader()
        Catch excpt As Exception
            If quNectCmd IsNot Nothing Then
                quNectCmd.Dispose()
            End If
            Exit Sub
        End Try

        Dim ridList As New List(Of Integer)()

        While dr.Read()
            ridList.Add(dr(0))
        End While




        If Not countBytesOnly Then
            Me.Cursor = Cursors.WaitCursor
            connectionString = buildConnectionString("")

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
            quNectConn = New OdbcConnection(connectionString)
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


            If backupTable(dbName, dbid, quNectConn, quNectConnFIDs, ridList) = DialogResult.Cancel Then
                MsgBox("Could Not backup table, canceling archive operation.", MsgBoxStyle.OkOnly, AppName)
                quNectConn.Close()
                quNectConn.Dispose()
                quNectConnFIDs.Close()
                quNectConnFIDs.Dispose()
                Me.Cursor = Cursors.Default
                pleaseWait.Close()
                Exit Sub
            End If
            quNectConnFIDs.Close()
            quNectConnFIDs.Dispose()


            'now we need to check to see if the archive table has a file attachment field
            Try
                If quNectCmd IsNot Nothing Then
                    quNectCmd.Dispose()
                End If
                quNectConn.Close()
                quNectConn = New OdbcConnection(connectionString)
                quNectConn.Open()
                Dim sql As String = "SELECT fid FROM " & archivedbid & "~fields WHERE field_type='file'"

                quNectCmd = quNectConn.CreateCommand()
                quNectCmd.CommandText = sql
                dr = quNectCmd.ExecuteReader()

                While dr.Read()
                    fileFID = dr.GetString(0)
                End While
            Catch excpt As Exception
                If quNectCmd IsNot Nothing Then
                    quNectCmd.Dispose()
                End If
                Throw New ArgumentException("could not access archive " & archivedbid & " for file attachment field information.")
            End Try

            If fileFID = "" Then
                Try
                    If quNectCmd IsNot Nothing Then
                        quNectCmd.Dispose()
                    End If
                    quNectConn.Close()
                    quNectConn = New OdbcConnection(connectionString)
                    quNectConn.Open()
                    Dim sql As String = "ALTER TABLE " & archivedbid & "~fields ADD ""CSV Archive"" longVarBinary"

                    quNectCmd = quNectConn.CreateCommand()
                    quNectCmd.CommandText = sql
                    quNectCmd.ExecuteNonQuery()
                Catch excpt As Exception
                    If quNectCmd IsNot Nothing Then
                        quNectCmd.Dispose()
                    End If
                    Throw New ArgumentException("Could not add 'CSV Archive' file attachment field to archive " & archivedbid & ".")
                End Try
            End If

            'now we need to check to see if the table to be archived has a restore button field
            Dim buttonFID As String = ""
            Try
                If quNectCmd IsNot Nothing Then
                    quNectCmd.Dispose()
                End If
                quNectConn.Close()
                quNectConn = New OdbcConnection(connectionString)
                quNectConn.Open()
                Dim sql As String = "SELECT fid FROM " & dbid & "~fields WHERE label='Retrieve from Archive'"

                quNectCmd = quNectConn.CreateCommand()
                quNectCmd.CommandText = sql
                dr = quNectCmd.ExecuteReader()

                While dr.Read()
                    buttonFID = dr.GetString(0)
                End While
            Catch excpt As Exception
                If quNectCmd IsNot Nothing Then
                    quNectCmd.Dispose()
                End If
                Throw New ArgumentException("could not access table to be archived " & dbid & " for button field information.")
            End Try

            If buttonFID = "" Then

                Try
                    If quNectCmd IsNot Nothing Then
                        quNectCmd.Dispose()
                    End If
                    quNectConn.Close()
                    quNectConn = New OdbcConnection(connectionString)
                    quNectConn.Open()
                    Dim sql As String = "ALTER TABLE " & dbid & " ADD ""Retrieve from Archive"" formula_url (255)"

                    quNectCmd = quNectConn.CreateCommand()
                    quNectCmd.CommandText = sql
                    quNectCmd.ExecuteNonQuery()
                Catch excpt As Exception
                    If quNectCmd IsNot Nothing Then
                        quNectCmd.Dispose()
                    End If
                    Throw New ArgumentException("Could not add 'Retrieve from Archive' button field to table to be archived " & dbid & ".")
                End Try
                'buttonFID = qdb.AddField(dbid, "Retrieve from Archive", "url", True)
            End If
            Dim formula As String = ""

            formula &= "var Text pagename = ""QuNectArchive.js"";" & vbCrLf
            formula &= "var Text cfg = ""key="" & urlencode([" & keyFieldLabel & "]) & ""&keyfid=" & keyfid & "&filefid=" & fileFID & "&reffid=" & refFID & "&apptoken=" & txtAppToken.Text & "&dbid="" & " & formulaDBID & " & ""&archivedbid=" & archivedbid & "&filerid="" & urlencode([QuNect Archive Reference]);" & vbCrLf
            formula &= "if([QuNect Archive Reference] = 0, """", " & vbCrLf
            formula &= """javascript:var cfg = '"" & URLEncode($cfg) & ""';if(typeof(qnctdg) != 'undefined'){void(qnctdg.display(cfg))}else{void($.getScript('/db/"" & Dbid() & ""?a=dbpage&pagename="" & $pagename & ""',function(){qnctdg = new QuNectArchive(cfg)}))}"""
            formula &= ")" & vbCrLf
            Try
                If quNectCmd IsNot Nothing Then
                    quNectCmd.Dispose()
                End If
                quNectConn.Close()
                quNectConn = New OdbcConnection(connectionString)
                quNectConn.Open()
                Dim sql As String = "UPDATE " & dbid & "~fields SET formula = '" & formula.Replace("'", "''") & "', appears_as = 'Retrieve from Archive' WHERE label = 'Retrieve from Archive'"

                quNectCmd = quNectConn.CreateCommand()
                quNectCmd.CommandText = sql
                quNectCmd.ExecuteNonQuery()
            Catch excpt As Exception
                If quNectCmd IsNot Nothing Then
                    quNectCmd.Dispose()
                End If
                'Throw New ArgumentException("Could not modify formula of  'Retrieve from Archive' button field to archive " & dbid & ".")
            End Try

            Try
                If quNectCmd IsNot Nothing Then
                    quNectCmd.Dispose()
                End If
                quNectConn.Close()
                quNectConn = New OdbcConnection(connectionString)
                quNectConn.Open()
                Dim sql As String = "INSERT INTO " & parentDBID & "~pages (Name, Value) Values ('QuNectArchive.js', '" & My.Resources.QuNectArchive.Replace("'", "''") & "')"

                quNectCmd = quNectConn.CreateCommand()
                quNectCmd.CommandText = sql
                quNectCmd.ExecuteNonQuery()
            Catch excpt As Exception
                If quNectCmd IsNot Nothing Then
                    quNectCmd.Dispose()
                End If
                Throw New ArgumentException("Could not add JavaScript page 'QuNectArchive.js' to " & parentDBID & ".")
            End Try

        End If

        Dim bytesArchived = 0
        Dim numArchived As Integer = 0
        pleaseWait.pb.Value = 0
        Application.DoEvents()
        pleaseWait.pb.Maximum = ridList.Count
        Dim xpathIndex As String = "2"
        If keyfid = "3" Then
            xpathIndex = "1"
        End If
        For i = 0 To ridList.Count - 1 Step recordsPerArchive
            If i + recordsPerArchive > pleaseWait.pb.Maximum Then
                pleaseWait.pb.Value = pleaseWait.pb.Maximum
            Else
                pleaseWait.pb.Value = i + recordsPerArchive
            End If
            Application.DoEvents()
            Dim lowRid As Integer = i
            Dim highRid As Integer = i + recordsPerArchive - 1
            If ridList.Count - 1 < i + recordsPerArchive - 1 Then
                highRid = ridList.Count - 1
            End If
            Dim n As Integer
            Dim inRids As String = " WHERE fid3 IN ("
            Dim comma As String = ""
            Dim orRids As String = ""
            Dim orOp As String = ""
            For n = i To highRid
                inRids &= comma & ridList(n)
                orRids &= orOp & ridList(n)
                comma = ","
                orOp = " OR "
                numArchived += 1
            Next
            inRids &= ")"

            Dim strCSV As String = String.Join(",", clistArray) & vbCrLf
            Dim genResultsTable = getCSVFromTable(dbid, clistArray, inRids, quNectConn)
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

                'need to create a record in the archive
                Dim importSQL As String = "INSERT INTO " & archivedbid & " (fid" & fileFID & ") VALUES ('" & sTempFileName & "')"
                quNectCmd = New OdbcCommand(importSQL, quNectConn)
                quNectCmd.ExecuteNonQuery()
                System.IO.File.Delete(sTempFileName)
                quNectCmd = New OdbcCommand("SELECT @@IDENTITY", quNectConn)
                dr = quNectCmd.ExecuteReader()
                dr.Read()
                Dim newArchiveRID As String = dr.GetString(0)
                'now we need to hollow out the records and update the reference field value
                Dim updateSQL As String = "UPDATE " & dbid & " SET "
                Dim archiveFieldCounter As Integer
                comma = ""
                For archiveFieldCounter = 0 To lstArchiveFields.Items.Count - 1
                    updateSQL &= comma & """" & lstArchiveFields.Items(archiveFieldCounter) & """ = ''"
                    comma = ","
                Next
                updateSQL &= ", fid" & refFID & " = " & newArchiveRID & inRids

                quNectCmd = New OdbcCommand(updateSQL, quNectConn)
                quNectCmd.ExecuteNonQuery()
            End If
        Next
        quNectConn.Close()
        quNectConn.Dispose()

        Me.Cursor = Cursors.Default
        pleaseWait.Close()
        If countBytesOnly Then
            If ridList.Count > 0 Then
                recordsPerArchive = 10000000 \ CInt(bytesArchived / ridList.Count)
                If recordsPerArchive > 90 Then
                    recordsPerArchive = 90
                ElseIf recordsPerArchive < 1 Then
                    recordsPerArchive = 1
                End If
                SaveSetting(AppName, "archive", "bytesPerRecord", CStr(recordsPerArchive))
            End If
            MsgBox("About " & bytesArchived & " bytes would be archived from " & ridList.Count & " records.", MsgBoxStyle.OkOnly, AppName)
        Else
            MsgBox("About " & bytesArchived & " bytes archived from " & ridList.Count & " records.", MsgBoxStyle.OkOnly, AppName)
        End If

    End Sub
    Private Function getRefFID(dbid As String, dbidAlias As String) As String
        Dim quNectConn As OdbcConnection = New OdbcConnection(buildConnectionString(""))
        quNectConn.Open()
        Dim quNectCmd As OdbcCommand = Nothing
        Dim dr As OdbcDataReader
        Try
            quNectCmd = New OdbcCommand("Select fid FROM " & dbid & "~fields WHERE mastag = '" & dbidAlias & "'", quNectConn)
            dr = quNectCmd.ExecuteReader()
        Catch excpt As Exception
            If quNectCmd IsNot Nothing Then
                quNectCmd.Dispose()
            End If
            Return ""
        End Try

        If dr.HasRows Then
            dr.Read()
            getRefFID = dr("fid")
        Else
            getRefFID = ""
        End If
        quNectConn.Close()
    End Function
    Private Function backupTable(ByVal dbName As String, ByVal dbid As String, ByVal quNectConn As OdbcConnection, ByVal quNectConnFIDs As OdbcConnection, ByRef ridList As List(Of Integer)) As DialogResult
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
        For j = 0 To ridList.Count - 1 Step 100
            Application.DoEvents()
            Dim lowRid As Integer = j
            Dim highRid As Integer = j + 100 - 1
            If ridList.Count - 1 < j + 100 - 1 Then
                highRid = ridList.Count - 1
            End If
            Dim n As Integer
            Dim commaRids As String = ""
            Dim strComma As String = ""
            For n = j To highRid
                commaRids &= strComma & ridList(n)
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
    Private Function getCSVFromTable(ByVal dbid As String, ByRef clistArray() As String, whereClause As String, ByRef quNectConn As OdbcConnection) As String

        Dim quickBaseSQL As String = "SELECT fid" & String.Join(", fid", clistArray) & " FROM """ & dbid & """" & whereClause

        Dim quNectCmd As OdbcCommand = New OdbcCommand(quickBaseSQL, quNectConn)
        Dim dr As OdbcDataReader
        Try
            dr = quNectCmd.ExecuteReader()
        Catch excpt As Exception
            Throw New ArgumentException("Could not get CSV data from " & dbid & ". " & excpt.Message)
        End Try
        If Not dr.HasRows Then
            Return ""
        End If
        getCSVFromTable = "fid" & String.Join(", fid", clistArray) & vbCrLf
        Dim k As Integer = 0
        Dim i As Integer
        pleaseWait.pb.Maximum = 100
        While (dr.Read())
            pleaseWait.pb.Value = k Mod 100
            Application.DoEvents()
            k += 1
            Dim comma As String = ""
            For i = 0 To dr.FieldCount - 1
                getCSVFromTable &= comma
                If Not dr.GetValue(i) Is Nothing Then
                    getCSVFromTable &= """"
                    getCSVFromTable &= Replace(dr.GetValue(i).ToString(), """", """""")
                    getCSVFromTable &= """"
                End If
                comma = ","
            Next
            getCSVFromTable &= vbCrLf
        End While
        dr.Close()
        quNectCmd.Dispose()
    End Function

    Private Sub txtAppToken_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAppToken.TextChanged
        SaveSetting(AppName, "Credentials", "apptoken", txtAppToken.Text)
        showHideControls()
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
        showHideControls()
    End Sub

    Private Sub lstReports_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstReports.SelectedIndexChanged
        Dim thisQid As String = getQidFromReportName(lstReports.Items(lstReports.SelectedIndex))
        SaveSetting(AppName, "archive", "qid", thisQid)
        showHideControls()
    End Sub
    Private Sub btnAppToken_Click(sender As Object, e As EventArgs) Handles btnAppToken.Click
        Process.Start("https://qunect.com/flash/AppToken.html")
    End Sub

    Private Sub btnUserToken_Click(sender As Object, e As EventArgs) Handles btnUserToken.Click
        Process.Start("https://qunect.com/flash/UserToken.html")
    End Sub

    Private Sub txtBackupFolder_TextChanged(sender As Object, e As EventArgs) Handles txtBackupFolder.TextChanged
        showHideControls()
    End Sub
End Class

