Imports System.Xml
Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Data.Odbc
Imports System.Text.RegularExpressions
Imports System.Threading
Public Class frmRazor
    Private Const AppName = "QuNectRazor"
    Private Const qunectRazorVersion = "1.0.0.11"
    Private cmdLineArgs() As String
    Private automode As Boolean = False
    Private connectionString As String = ""
    Private appdbid As String = ""
    Private qdbAppName As String = ""
    Private Class qdbVersion
        Public year As Integer
        Public major As Integer
        Public minor As Integer
    End Class
    Private qdbVer As qdbVersion = New qdbVersion
    Private fieldCharacterSettings = New ArrayList
    Private sqlKeyWordDict As Dictionary(Of String, Boolean)
    Private sqlKeyWords() As String = {"ADD", "ALL", "ALTER", "AND", "ANY", "AS", "ASC", "AVG", "BETWEEN", "BY", "CONSTRAINT", "COUNT", "CREATE", "CREATEOQ", "DELETE", "DELETEOQ", "DESC", "DISTINCT", "DROP", "EXISTS", "FOREIGN", "FROM", "GROUP", "HAVING", "IN", "INDEX", "INNER", "INSERT", "INSERTOQ", "INTO", "IS", "JOIN", "LEFT", "LIKE", "KEY", "MAX", "MIN", "NOT", "NULL", "ON", "OR", "ORDER", "OUTER", "PRIMARY", "REFERENCES", "SELECT", "SET", "SQL", "SUM", "TABLE", "TOP", "UNIQUE", "UPDATE", "UPDATEOQ", "USER", "VALUES", "WHERE"}
    Private Class tests
        Public Const dupes = "Check for Duplicate Column Names"
        Public Const sqlKeyWords = "Check for SQL Keywords in Column Names"
        Public Const empty = "Find Empty Columns"
        Public Const fields = "List non Report Link/Formula URL fields"
    End Class
    Private Sub razor_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        tvFields.Visible = False
        lblResult.Visible = False
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
        Dim samlSetting As String = GetSetting(AppName, "Credentials", "samlsetting", "0")
        If samlSetting = "1" Then
            ckbSSO.Checked = True
        Else
            ckbSSO.Checked = False
        End If


        Dim myBuildInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(Application.ExecutablePath)
        Me.Text = "QuNect Razor " & qunectRazorVersion

        sqlKeyWordDict = sqlKeyWords.ToDictionary(Function(value As String)
                                                      Return value
                                                  End Function,
                   Function(value As String)
                           Return True
                       End Function)

        fieldCharacterSettings.Add("all")
        fieldCharacterSettings.Add("all characters")
        fieldCharacterSettings.Add("msaccess")
        fieldCharacterSettings.Add("ms access allowed characters")
        fieldCharacterSettings.Add("lnuhs")
        fieldCharacterSettings.Add("letters numbers underscores hyphens spaces")
        fieldCharacterSettings.Add("lnus")
        fieldCharacterSettings.Add("letters numbers underscores spaces")
        fieldCharacterSettings.Add("lnu")
        fieldCharacterSettings.Add("letters numbers underscores")
        fieldCharacterSettings.Add("lnunc")
        fieldCharacterSettings.Add("letters numbers underscores no colons")
        cmbTests.Items.Add("Choose an Analysis")
        cmbTests.Items.Add(tests.dupes)
        cmbTests.Items.Add(tests.sqlKeyWords)
        cmbTests.Items.Add(tests.empty)
        cmbTests.Items.Add(tests.fields)
        cmbTests.SelectedIndex = 0
    End Sub

    Private Sub txtUsername_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUsername.TextChanged
        SaveSetting(AppName, "Credentials", "username", txtUsername.Text)
    End Sub

    Private Sub txtPassword_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPassword.TextChanged
        SaveSetting(AppName, "Credentials", "password", txtPassword.Text)
    End Sub

    Private Sub btnListTables_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnListTables.Click
        qdbAppName = ""
        appdbid = ""
        listTables()
    End Sub
    Private Sub listTables()
        Me.Cursor = Cursors.WaitCursor


        Dim quNectConn As OdbcConnection = getquNectConn("all")
        If Not quNectConn Is Nothing Then
            tvAppsTables.Visible = True
            Dim tables As DataTable = quNectConn.GetSchema("Tables")
            listTablesFromGetSchema(tables)
            quNectConn.Dispose()
        End If
        Me.Cursor = Cursors.Default
    End Sub
    Private Function getquNectConn(fieldCharacterSetting As String) As OdbcConnection
        Dim connectionString As String = ""
        Try
            connectionString = buildConnectionString(fieldCharacterSetting)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, AppName)
            Return Nothing
        End Try
        Dim quNectConn As OdbcConnection = New OdbcConnection(connectionString)
        Try
            quNectConn.Open()
        Catch excpt As Exception
            Me.Cursor = Cursors.Default
            If excpt.Message.StartsWith("ERROR [IM003]") Or excpt.Message.Contains("Data source name not found") Then
                MsgBox("Please install QuNect ODBC for QuickBase from http://qunect.com/download/QuNect.exe and try again.")
            Else
                MsgBox(excpt.Message.Substring(13))
            End If
            Return Nothing
            Exit Function
        End Try

        Dim ver As String = quNectConn.ServerVersion
        Dim m As Match = Regex.Match(ver, "\d+\.(\d+)\.(\d+)\.(\d+)")
        qdbVer.year = CInt(m.Groups(1).Value)
        qdbVer.major = CInt(m.Groups(2).Value)
        qdbVer.minor = CInt(m.Groups(3).Value)
        If (qdbVer.year < 15) Or ((qdbVer.year = 15) And ((qdbVer.major <= 5) And (qdbVer.minor < 78))) Then
            MsgBox("You are running the " & ver & " version of QuNect ODBC for QuickBase. Please install the latest version from http://qunect.com/download/QuNect.exe")
            quNectConn.Dispose()
            Me.Cursor = Cursors.Default
            Return Nothing
            Exit Function
        End If
        Return quNectConn
    End Function
    Sub timeoutCallback(ByVal result As System.IAsyncResult)
        If Not automode Then
            Me.Cursor = Cursors.Default
            MsgBox("Operation timed out. Please try again.")
        End If
    End Sub
    Sub listTablesFromGetSchema(tables As DataTable)

        Dim dbids As String = GetSetting(AppName, "razor", "dbids")
        Dim dbidArray As New ArrayList
        dbidArray.AddRange(dbids.Split(";"c))
        Dim i As Integer
        Dim dbidCollection As New Collection
        For i = 0 To dbidArray.Count - 1
            Try
                dbidCollection.Add(dbidArray(i), dbidArray(i))
            Catch excpt As Exception
                'ignore dupes
            End Try
        Next

        tvAppsTables.BeginUpdate()
        tvAppsTables.Nodes.Clear()
        tvAppsTables.ShowNodeToolTips = True
        Dim dbName As String
        Dim applicationName As String = ""
        Dim prevAppName As String = ""
        Dim dbid As String
        pb.Value = 0
        pb.Visible = True
        pb.Maximum = tables.Rows.Count
        Dim getDBIDfromdbName As New Regex("([a-z0-9~]+)$")

        For i = 0 To tables.Rows.Count - 1
            pb.Value = i
            Application.DoEvents()
            dbName = tables.Rows(i)(2)
            applicationName = tables.Rows(i)(0)
            Dim dbidMatch As Match = getDBIDfromdbName.Match(dbName)
            dbid = dbidMatch.Value
            If applicationName <> prevAppName Then
                Dim appNode As TreeNode = tvAppsTables.Nodes.Add(applicationName)
                appNode.ToolTipText = "Right click me to get all my table reports as well as tables!"
                appNode.Tag = dbid
                prevAppName = applicationName
            End If
            Dim tableName As String = dbName
            If appdbid.Length = 0 And dbName.Length > applicationName.Length Then
                tableName = dbName.Substring(applicationName.Length + 2)
            End If
            Dim tableNode As TreeNode = tvAppsTables.Nodes(tvAppsTables.Nodes.Count - 1).Nodes.Add(tableName)
            tableNode.Tag = dbid
        Next
        pb.Visible = False
        tvAppsTables.EndUpdate()
        pb.Value = 0
        Me.Cursor = Cursors.Default
    End Sub


    Private Sub txtServer_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtServer.TextChanged
        SaveSetting(AppName, "Credentials", "server", txtServer.Text)
    End Sub


    Private Sub btnRazor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        razor()
    End Sub
    Private Function buildConnectionString(fieldNameCharacters As String) As String
        buildConnectionString = "FIELDNAMECHARACTERS=" & fieldNameCharacters & ";uid=" & txtUsername.Text
        buildConnectionString &= ";pwd=" & txtPassword.Text
        buildConnectionString &= ";driver={QuNect ODBC for QuickBase};"
        buildConnectionString &= ";quickbaseserver=" & txtServer.Text
        buildConnectionString &= ";APPTOKEN=" & txtAppToken.Text
        If ckbDetectProxy.Checked Then
            buildConnectionString &= ";DETECTPROXY=1"
        End If
        If ckbSSO.Checked Then
            buildConnectionString &= ";SAML=1"
        End If
        If appdbid.Length Then
            buildConnectionString &= ";APPID=" & appdbid & ";APPNAME=" & qdbAppName
        End If
        If cmbPassword.SelectedIndex = 0 Then
            cmbPassword.Focus()
            Throw New System.Exception("Please indicate whether you are using a password or a user token.")
            Return ""
        ElseIf cmbPassword.SelectedIndex = 1 Then
            buildConnectionString &= ";PWDISPASSWORD=1"
        Else
            buildConnectionString &= ";PWDISPASSWORD=0"
        End If
    End Function
    Private Sub razor()
        'here we need to go through the list and razor
        Dim connectionString As String = ""
        Try
            connectionString = buildConnectionString("all")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, AppName)
            Exit Sub
        End Try
        Me.Cursor = Cursors.WaitCursor
        Dim quNectConnFIDs As OdbcConnection = New OdbcConnection(connectionString & ";usefids=1")

        Try
            quNectConnFIDs.Open()
        Catch excpt As Exception
            If Not automode Then
                MsgBox(excpt.Message())
            End If
            quNectConnFIDs.Dispose()
            Me.Cursor = Cursors.Default
            Exit Sub
        End Try

        Dim quNectConn As OdbcConnection = New OdbcConnection(connectionString)
        Try
            quNectConn.Open()
        Catch excpt As Exception
            If Not automode Then
                MsgBox(excpt.Message())
            End If
            quNectConn.Dispose()
            Me.Cursor = Cursors.Default
            Exit Sub
        End Try


        quNectConn.Close()
        quNectConn.Dispose()
        quNectConnFIDs.Close()
        quNectConnFIDs.Dispose()
        Me.Cursor = Cursors.Default
    End Sub
    Private Function razorTable(ByVal dbName As String, ByVal dbid As String, ByVal quNectConn As OdbcConnection, ByVal quNectConnFIDs As OdbcConnection) As DialogResult
        'we need to get the schema of the table
        Dim restrictions(2) As String
        restrictions(2) = dbid
        Dim columns As DataTable = quNectConn.GetSchema("Columns", restrictions)
        'now we can look for formula fileURL fields
        razorTable = DialogResult.OK
        Dim quickBaseSQL As String = "select * from """ & dbid & """"

        Dim quNectCmd As OdbcCommand = New OdbcCommand(quickBaseSQL, quNectConnFIDs)
        Dim dr As OdbcDataReader
        Try
            dr = quNectCmd.ExecuteReader()
        Catch excpt As Exception
            If Not automode Then
                razorTable = MsgBox("Could not get field identifiers for table " & dbid & " because " & excpt.Message() & vbCrLf & "Would you like to continue?", MsgBoxStyle.OkCancel, AppName)
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
        quNectCmd = New OdbcCommand(quickBaseSQL, quNectConn)
        Try
            dr = quNectCmd.ExecuteReader()
        Catch excpt As Exception
            quNectCmd.Dispose()
            Exit Function
        End Try
        If Not dr.HasRows Then
            Exit Function
        End If

        dr.Close()
        quNectCmd.Dispose()
    End Function
    Private Function ChangeCharacter(s As String, replaceWith As Char, idx As Integer) As String
        Dim sb As New StringBuilder(s)
        sb(idx) = replaceWith
        Return sb.ToString()
    End Function

    Private Function UrlDecode(text As String) As String
        text = text.Replace("+", " ")
        Return System.Uri.UnescapeDataString(text)
    End Function
    Private Sub txtAppToken_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAppToken.TextChanged
        SaveSetting(AppName, "Credentials", "apptoken", txtAppToken.Text)
    End Sub
    Private Sub ckbDetectProxy_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ckbDetectProxy.CheckStateChanged
        If ckbDetectProxy.Checked Then
            SaveSetting(AppName, "Credentials", "detectproxysettings", "1")
        Else
            SaveSetting(AppName, "Credentials", "detectproxysettings", "0")
        End If
    End Sub
    Private Sub ckbSSO_CheckStateChanged(sender As Object, e As EventArgs) Handles ckbSSO.CheckStateChanged
        If ckbSSO.Checked Then
            SaveSetting(AppName, "Credentials", "samlsetting", "1")
        Else
            SaveSetting(AppName, "Credentials", "samlsetting", "0")
        End If
    End Sub
    Private Sub ContextMenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles ContextMenuStrip1.ItemClicked
        If (qdbVer.year < 16) Or ((qdbVer.year = 16) And ((qdbVer.major <= 6) And (qdbVer.minor < 20))) Then
            MsgBox("To access this feature please install the latest version from http://qunect.com/download/QuNect.exe")
            Exit Sub
        End If
        'here we need to reconnect with the appid in the connection string
        If tvAppsTables.SelectedNode Is Nothing Then
            Exit Sub
        End If
        appdbid = tvAppsTables.SelectedNode.Tag
        qdbAppName = tvAppsTables.SelectedNode.Text
        listTables()
    End Sub

    Private Sub tvAppsTables_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles tvAppsTables.NodeMouseClick
        If e.Node.Level <> 0 Then
            btnAnalyze.Visible = True
            cmbTests.Visible = True
            tvAppsTables.SelectedNode = e.Node
            Exit Sub
        End If


    End Sub

    Private Sub btnAnalyze_Click(sender As Object, e As EventArgs) Handles btnAnalyze.Click
        lblResult.Visible = False
        tvFields.Visible = False
        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()
        Try
            If cmbTests.Text = tests.dupes Then
                lblResult.Visible = True
                findDupeColumnNames()
            ElseIf cmbTests.Text = tests.sqlKeyWords Then
                lblResult.Visible = True
                findSQLKeyWords()
            ElseIf cmbTests.Text = tests.empty Then
                lblResult.Visible = True
                findEmptyFields()
            ElseIf cmbTests.Text = tests.fields Then
                tvFields.Visible = True
                listFields()
            Else
                MsgBox("Please Choose a Test")
            End If
        Catch ex As Exception
            MsgBox("Sorry we could not perform your analysis because: " & ex.Message)
        End Try
        Me.Cursor = Cursors.Default
        tvAppsTables.Focus()
    End Sub
    Private Sub findEmptyFields()
        Dim colArray As New ArrayList
        Dim quNectConn As OdbcConnection = getquNectConn("all;IGNOREDUPEFIELDNAMES=1;TEXTNULL=1")
        Dim rowcount = 0
        If Not quNectConn Is Nothing Then
            Dim restrictions(3) As String
            restrictions(2) = tvAppsTables.SelectedNode.Tag 'catalog, owner, table, column
            Dim columns = quNectConn.GetSchema("Columns", restrictions)
            Dim findBooleans As Regex = New Regex("Checkbox", RegexOptions.IgnoreCase)
            Dim quickBaseSQL As String = "SELECT count(1) from [" & tvAppsTables.SelectedNode.Tag & "]"
            Dim quNectCmd = New OdbcCommand(quickBaseSQL, quNectConn)
            Try
                Dim dr = quNectCmd.ExecuteReader()
                dr.Read()
                rowcount = dr.GetValue(0)
                If rowcount = 0 Then
                    lblResult.Text = "There are no rows in " & tvAppsTables.SelectedNode.Text
                    Exit Sub
                End If
            Catch excpt As Exception
                quNectCmd.Dispose()
                Exit Sub
            End Try
            For i = 0 To columns.Rows.Count - 1
                Application.DoEvents()
                Dim ColumnName As String = columns.Rows(i)("COLUMN_NAME") 'index 3 for column name, index 11 for remarks has field type and then the fid
                Dim remarks As String = columns.Rows(i)("REMARKS")
                Dim whereClause = " IS NOT NULL"
                Dim isCheckBox = findBooleans.Match(remarks).Success
                If isCheckBox Then
                    whereClause = " = 1"
                End If
                quNectCmd = New OdbcCommand(quickBaseSQL & " WHERE [" & ColumnName & "]" & whereClause, quNectConn)
                Try
                    Dim dr = quNectCmd.ExecuteReader()
                    dr.Read()
                    If dr.GetValue(0) = 0 Then
                        colArray.Add(ColumnName & " is empty. " & columns.Rows(i)("REMARKS"))
                    End If
                Catch excpt As Exception
                    quNectCmd.Dispose()
                    Exit Sub
                End Try
            Next
            quNectConn.Dispose()
        End If
        Dim message As String = ""
        For Each col In colArray
            message &= vbCrLf & col
        Next
        lblResult.Text = "Out of the " & rowcount & " rows you have access to, the following columns have no data." & vbCrLf & message
    End Sub
    Private Sub listFields()
        Me.Cursor = Cursors.WaitCursor
        Dim columnNames As New ArrayList
        getAllTypesOfColumnNames(columnNames)
        Dim quNectConn As OdbcConnection = getquNectConn("all;IGNOREDUPEFIELDNAMES=1;TEXTNULL=1")
        Dim rowcount = 0
        tvFields.BeginUpdate()
        tvFields.Nodes.Clear()
        If Not quNectConn Is Nothing Then
            Dim restrictions(3) As String
            restrictions(2) = tvAppsTables.SelectedNode.Tag 'catalog, owner, table, column
            Dim columns = quNectConn.GetSchema("Columns", restrictions)
            For i = 0 To columns.Rows.Count - 1
                Application.DoEvents()
                Dim ColumnName As String = columns.Rows(i)("COLUMN_NAME") 'index 3 for column name, index 11 for remarks has field type and then the fid
                Dim fieldNode As TreeNode = tvFields.Nodes.Add(ColumnName)
                fieldNode.Nodes.Add(columns.Rows(i)("REMARKS"))
                For j As Integer = 0 To fieldCharacterSettings.Count - 1 Step 2
                    fieldNode.Nodes.Add(columnNames.Item(j / 2).Rows(i)("COLUMN_NAME") & " (" & fieldCharacterSettings(j + 1) & ")")
                Next
            Next
            quNectConn.Dispose()
        End If
        tvFields.EndUpdate()
        Me.Cursor = Cursors.Default
    End Sub
    Private Function findSQLKeyWordsForFieldSetting(fieldCharaterSetting As String) As ArrayList
        Dim colArray As New ArrayList
        Dim quNectConn As OdbcConnection = getquNectConn(fieldCharaterSetting & ";IGNOREDUPEFIELDNAMES=1")
        If Not quNectConn Is Nothing Then
            Dim restrictions(3) As String
            restrictions(2) = tvAppsTables.SelectedNode.Tag 'catalog, owner, table, column
            Dim columns = quNectConn.GetSchema("Columns", restrictions)

            For i = 0 To columns.Rows.Count - 1
                Application.DoEvents()
                Dim ColumnName As String = columns.Rows(i)("COLUMN_NAME") 'index 3 for column name, index 11 for remarks has field type and then the fid
                Dim wordsInColumnName() As String = ColumnName.Split(" ")
                For Each word In wordsInColumnName
                    If sqlKeyWordDict.ContainsKey(word.ToUpper) Then
                        'now we have a problem
                        colArray.Add(ColumnName & " contains " & word.ToUpper & ". " & columns.Rows(i)("REMARKS"))
                    End If
                Next
            Next
            quNectConn.Dispose()
        End If
        Return colArray
    End Function
    Private Sub findSQLKeyWords()
        Dim message As String = "Field names containing SQL keywords when spaces are not allowed." & vbCrLf
        Dim colArray As ArrayList
        colArray = findSQLKeyWordsForFieldSetting("lnu")
        For Each col In colArray
            message &= vbCrLf & col
        Next
        If colArray.Count = 0 Then
            message = "No field names contain SQL keywords when spaces are not allowed." & vbCrLf
        End If
        message &= vbCrLf & "Field names containing SQL keywords when spaces are allowed."
        colArray = findSQLKeyWordsForFieldSetting("all")
        For Each col In colArray
            message &= vbCrLf & col
        Next
        lblResult.Text = message
    End Sub
    Private Sub findDupeColumnNames()
        Dim message As String = ""
        Dim comma As String = vbCrLf & vbTab
        Dim columns As DataTable = Nothing
        Dim allCharacterColumns As DataTable = Nothing
        For i As Integer = 0 To fieldCharacterSettings.Count - 1 Step 2
            Dim dupeColumns As ArrayList = findDupeColumnNamesForFieldSetting(fieldCharacterSettings(i), fieldCharacterSettings(i + 1), columns, allCharacterColumns)
            Dim dupeColumnString As String = ""
            For Each dupeColumn In dupeColumns
                dupeColumnString &= comma & dupeColumn
            Next
            If dupeColumnString <> "" Then
                message &= fieldCharacterSettings(i + 1) & ": " & dupeColumnString & vbCrLf & vbCrLf
            Else
                message &= fieldCharacterSettings(i + 1) & ":" & vbCrLf & vbTab & "No Duplicates" & vbCrLf & vbCrLf
            End If
        Next
        lblResult.Text = "Duplicate field names are listed below for each field character replacement setting." & vbCrLf & vbCrLf & message
    End Sub
    Private Function findDupeColumnNamesForFieldSetting(fieldCharacterSetting As String, fieldCharacterSettingDescription As String, ByRef columns As DataTable, ByRef allCharacterColumns As DataTable) As ArrayList
        Dim quNectConn As OdbcConnection = getquNectConn(fieldCharacterSetting & ";IGNOREDUPEFIELDNAMES=1")
        Dim colDictionary = New Dictionary(Of String, Boolean)
        Dim colArray As New ArrayList
        Dim checkingUnalteredNames = False
        If Not quNectConn Is Nothing Then
            Dim restrictions(3) As String
            restrictions(2) = tvAppsTables.SelectedNode.Tag 'catalog, owner, table, column
            columns = quNectConn.GetSchema("Columns", restrictions)
            If allCharacterColumns Is Nothing Then
                allCharacterColumns = columns
                checkingUnalteredNames = True
            End If
            Dim colUnique As New Dictionary(Of String, Boolean)
            For i = 0 To columns.Rows.Count - 1
                Application.DoEvents()
                Dim ColumnName As String = columns.Rows(i)("COLUMN_NAME") 'index 3 for column name, index 11 for remarks has field type and then the fid
                If checkingUnalteredNames Then
                    Dim wordsInColumnName() As String = ColumnName.Split(" ")
                    For Each word In wordsInColumnName
                        If sqlKeyWordDict.ContainsKey(word.ToUpper) Then
                            'now we have a problem
                        End If
                    Next
                End If
                If colUnique.ContainsKey(ColumnName) Then
                    If Not colDictionary.ContainsKey(ColumnName) Then
                        colDictionary.Add(ColumnName, True)
                    End If
                Else
                    colUnique.Add(ColumnName, True)
                End If
            Next
            If colDictionary.Count Then
                For i = 0 To columns.Rows.Count - 1
                    Application.DoEvents()
                    Dim ColumnName As String = columns.Rows(i)("COLUMN_NAME") 'index 3 for column name, index 11 for remarks has field type and then the fid
                    If colDictionary.ContainsKey(ColumnName) Then
                        colArray.Add(allCharacterColumns.Rows(i)("COLUMN_NAME") & " becomes " & ColumnName & " " & columns.Rows(i)("REMARKS"))
                    End If
                Next
            End If
            quNectConn.Dispose()
        End If
        Return colArray
    End Function
    Private Sub getAllTypesOfColumnNames(ColumnTables As ArrayList)

        For i As Integer = 0 To fieldCharacterSettings.Count - 1 Step 2
            ColumnTables.Add(getColumnNamesForFieldSetting(fieldCharacterSettings(i)))
        Next

    End Sub
    Private Function getColumnNamesForFieldSetting(fieldCharacterSetting As String) As DataTable
        Dim quNectConn As OdbcConnection = getquNectConn(fieldCharacterSetting & ";IGNOREDUPEFIELDNAMES=1")

        Dim restrictions(3) As String
        restrictions(2) = tvAppsTables.SelectedNode.Tag 'catalog, owner, table, column
        getColumnNamesForFieldSetting = quNectConn.GetSchema("Columns", restrictions)
        quNectConn.Dispose()

    End Function
    Private Sub cmbPassword_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbPassword.SelectedIndexChanged
        SaveSetting(AppName, "Credentials", "passwordOrToken", cmbPassword.SelectedIndex)
        If cmbPassword.SelectedIndex = 0 Then
            txtPassword.Enabled = False
        Else
            txtPassword.Enabled = True
        End If
    End Sub
End Class

