Public Class frmEpsgDatabase
    'Setect or view an EPSG database.

#Region " Variable Declarations - All the variables used in this form and this application." '=================================================================================================

    'Variables used to connect to a database and open a table:
    Dim connString As String
    Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
    Dim ds As DataSet = New DataSet
    Dim da As OleDb.OleDbDataAdapter
    Dim tables As DataTableCollection = ds.Tables

#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Properties - All the properties used in this form and this application" '============================================================================================================

    Private _formNo As Integer = -1 'Multiple instances of this form can be displayed. FormNo is the index number of the form in ConversionList.
    'If the form is included in Main.ConversionList() then FormNo will be > -1 --> when exiting set Main.ClosedFormNo and call Main.ConversionFormClosed(). 
    Public Property FormNo As Integer
        Get
            Return _formNo
        End Get
        Set(ByVal value As Integer)
            _formNo = value
        End Set
    End Property

    Private _epsgDatabasePath = "" 'The path of the EPSG Database. This database contains a comprehensive set of coordinate reference system parameters.
    Property EpsgDatabasePath
        Get
            Return _epsgDatabasePath
        End Get
        Set(value)
            _epsgDatabasePath = value
        End Set
    End Property

    'The TableName property stores the name of the table selected for viewing.
    Private _tableName As String
    Public Property TableName As String
        Get
            Return _tableName
        End Get
        Set(value As String)
            _tableName = value
        End Set
    End Property

    'The Query property stores the text of the query used to display table values in the GridDataView on this form.
    Private _query As String
    Public Property Query() As String
        Get
            Return _query
        End Get
        Set(ByVal value As String)
            _query = value
        End Set
    End Property

#End Region 'Properties -----------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Process XML files - Read and write XML files." '=====================================================================================================================================

    Private Sub SaveFormSettings()
        'Save the form settings in an XML document.
        Dim settingsData = <?xml version="1.0" encoding="utf-8"?>
                           <!---->
                           <FormSettings>
                               <Left><%= Me.Left %></Left>
                               <Top><%= Me.Top %></Top>
                               <Width><%= Me.Width %></Width>
                               <Height><%= Me.Height %></Height>
                               <!---->
                               <TableName><%= TableName %></TableName>
                               <Query><%= Query %></Query>
                               <ViewAllRecords><%= rbViewAll.Checked %></ViewAllRecords>
                               <NRecords><%= txtViewNRecords.Text %></NRecords>
                           </FormSettings>

        'Add code to include other settings to save after the comment line <!---->

        'Dim SettingsFileName As String = "FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & ".xml"
        Dim SettingsFileName As String = "FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & "_" & FormNo & ".xml"
        Main.Project.SaveXmlSettings(SettingsFileName, settingsData)
    End Sub

    Private Sub RestoreFormSettings()
        'Read the form settings from an XML document.

        'Dim SettingsFileName As String = "FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & ".xml"
        Dim SettingsFileName As String = "FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & "_" & FormNo & ".xml"

        If Main.Project.SettingsFileExists(SettingsFileName) Then
            Dim Settings As System.Xml.Linq.XDocument
            Main.Project.ReadXmlSettings(SettingsFileName, Settings)

            If IsNothing(Settings) Then 'There is no Settings XML data.
                Exit Sub
            End If

            'Restore form position and size:
            If Settings.<FormSettings>.<Left>.Value <> Nothing Then Me.Left = Settings.<FormSettings>.<Left>.Value
            If Settings.<FormSettings>.<Top>.Value <> Nothing Then Me.Top = Settings.<FormSettings>.<Top>.Value
            If Settings.<FormSettings>.<Height>.Value <> Nothing Then Me.Height = Settings.<FormSettings>.<Height>.Value
            If Settings.<FormSettings>.<Width>.Value <> Nothing Then Me.Width = Settings.<FormSettings>.<Width>.Value

            'Add code to read other saved setting here:
            If Settings.<FormSettings>.<TableName>.Value <> Nothing Then TableName = Settings.<FormSettings>.<TableName>.Value
            If Settings.<FormSettings>.<Query>.Value <> Nothing Then Query = Settings.<FormSettings>.<Query>.Value

            If Settings.<FormSettings>.<NRecords>.Value <> Nothing Then txtViewNRecords.Text = Settings.<FormSettings>.<NRecords>.Value
            If Settings.<FormSettings>.<ViewAllRecords>.Value <> Nothing Then
                If Settings.<FormSettings>.<ViewAllRecords>.Value = "true" Then
                    rbViewAll.Checked = True
                Else
                    rbViewFirst.Checked = True
                End If
            End If


            CheckFormPos()
        End If
    End Sub

    Private Sub CheckFormPos()
        'Check that the form can be seen on a screen.

        Dim MinWidthVisible As Integer = 192 'Minimum number of X pixels visible. The form will be moved if this many form pixels are not visible.
        Dim MinHeightVisible As Integer = 64 'Minimum number of Y pixels visible. The form will be moved if this many form pixels are not visible.

        Dim FormRect As New Rectangle(Me.Left, Me.Top, Me.Width, Me.Height)
        Dim WARect As Rectangle = Screen.GetWorkingArea(FormRect) 'The Working Area rectangle - the usable area of the screen containing the form.

        'Check if the top of the form is above the top of the Working Area:
        If Me.Top < WARect.Top Then
            Me.Top = WARect.Top
        End If

        'Check if the top of the form is too close to the bottom of the Working Area:
        If (Me.Top + MinHeightVisible) > (WARect.Top + WARect.Height) Then
            Me.Top = WARect.Top + WARect.Height - MinHeightVisible
        End If

        'Check if the left edge of the form is too close to the right edge of the Working Area:
        If (Me.Left + MinWidthVisible) > (WARect.Left + WARect.Width) Then
            Me.Left = WARect.Left + WARect.Width - MinWidthVisible
        End If

        'Check if the right edge of the form is too close to the left edge of the Working Area:
        If (Me.Left + Me.Width - MinWidthVisible) < WARect.Left Then
            Me.Left = WARect.Left - Me.Width + MinWidthVisible
        End If

    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message) 'Save the form settings before the form is minimised:
        If m.Msg = &H112 Then 'SysCommand
            If m.WParam.ToInt32 = &HF020 Then 'Form is being minimised
                SaveFormSettings()
            End If
        End If
        MyBase.WndProc(m)
    End Sub

#End Region 'Process XML Files ----------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Display Methods - Code used to display this form." '============================================================================================================================

    Private Sub Form_Load(sender As Object, e As EventArgs) Handles Me.Load

        'rbViewAll.Checked = True
        txtViewNRecords.Text = "500"

        RestoreFormSettings()   'Restore the form settings

        'txtDatabase.Text = Main.EpsgDatabasePath
        txtDatabase.Text = EpsgDatabasePath
        FillCmbSelectTable()

        If TableName <> "" Then
            'Select the table iame in the combobox
            cmbSelectTable.SelectedIndex = cmbSelectTable.FindStringExact(TableName)
        End If

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Form
        If FormNo > -1 Then
            Main.ClosedFormNo = FormNo 'The Main form property ClosedFormNo is set to this form number. This is used in the ChartFormClosed method to select the correct form to set to nothing.
        End If

        'SaveFormSettings()
        Me.Close() 'Close the form
    End Sub

    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If WindowState = FormWindowState.Normal Then
            SaveFormSettings()
        Else
            'Dont save settings if the form is minimised.
        End If
    End Sub

    Private Sub EpsgDatabaseForm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        If FormNo > -1 Then
            Main.EpsgDatabaseFormClosed()
        End If
    End Sub




#End Region 'Form Display Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Open and Close Forms - Code used to open and close other forms." '===================================================================================================================

#End Region 'Open and Close Forms -------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Methods - The main actions performed by this form." '===========================================================================================================================

    Private Sub btnFindDatabase_Click(sender As Object, e As EventArgs) Handles btnFindDatabase.Click
        'Find a database file:

        If txtDatabase.Text <> "" Then
            Dim fInfo As New System.IO.FileInfo(txtDatabase.Text)
            OpenFileDialog1.InitialDirectory = fInfo.DirectoryName
            OpenFileDialog1.Filter = "Database |*.accdb; *.mdb"
            OpenFileDialog1.FileName = fInfo.Name
        Else
            If Main.Project.Type = ADVL_Utilities_Library_1.Project.Types.Archive Then
                OpenFileDialog1.InitialDirectory = System.Environment.SpecialFolder.MyDocuments
                OpenFileDialog1.Filter = "Database |*.accdb; *.mdb"
                OpenFileDialog1.FileName = ""
            Else
                OpenFileDialog1.InitialDirectory = Main.Project.Path
                OpenFileDialog1.Filter = "Database |*.accdb; *.mdb"
                OpenFileDialog1.FileName = ""
            End If

        End If

        If OpenFileDialog1.ShowDialog() = vbOK Then
            'Main.EpsgDatabasePath = OpenFileDialog1.FileName
            EpsgDatabasePath = OpenFileDialog1.FileName
            Main.EpsgDatabasePath = EpsgDatabasePath

            'txtDatabase.Text = Main.EpsgDatabasePath
            txtDatabase.Text = EpsgDatabasePath
            FillCmbSelectTable()
        End If
    End Sub

    Private Sub FillCmbSelectTable()
        'Fill the cmbSelectTable listbox with the availalble tables in the selected database.

        'If Main.EpsgDatabasePath = "" Then
        If EpsgDatabasePath = "" Then
            Main.Message.AddWarning("No EPSG database has been selected." & vbCrLf)
            Exit Sub
        End If

        'If Not System.IO.File.Exists(Main.EpsgDatabasePath) Then
        If Not System.IO.File.Exists(EpsgDatabasePath) Then
            Main.Message.AddWarning("Selected EPSG database can not be found." & vbCrLf)
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        cmbSelectTable.Items.Clear()

        'Specify the connection string:
        'Access 2003
        'connectionString = "provider=Microsoft.Jet.OLEDB.4.0;" + _
        '"data source = " + txtDatabase.Text

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + EpsgDatabasePath
        '"data source = " + Main.EpsgDatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        'This error occurs on the above line (conn.Open()):
        'Additional information: The 'Microsoft.ACE.OLEDB.12.0' provider is not registered on the local machine.
        'Fix attempt: 
        'http://www.microsoft.com/en-us/download/confirmation.aspx?id=23734
        'Download AccessDatabaseEngine.exe
        'Run the file to install the 2007 Office System Driver: Data Connectivity Components.


        Dim restrictions As String() = New String() {Nothing, Nothing, Nothing, "TABLE"} 'This restriction removes system tables
        dt = conn.GetSchema("Tables", restrictions)

        'Fill lstSelectTable
        Dim dr As DataRow
        Dim I As Integer 'Loop index
        Dim MaxI As Integer

        MaxI = dt.Rows.Count
        For I = 0 To MaxI - 1
            dr = dt.Rows(0)
            cmbSelectTable.Items.Add(dt.Rows(I).Item(2).ToString)
        Next I

        conn.Close()

    End Sub

    Private Sub cmbSelectTable_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSelectTable.SelectedIndexChanged
        'Update DataGridView1:

        If IsNothing(cmbSelectTable.SelectedItem) Then
            Exit Sub
        End If

        TableName = cmbSelectTable.SelectedItem.ToString

        If rbViewAll.Checked Then
            Query = "Select * From [" & TableName & "]"
        Else
            If txtViewNRecords.Text.Trim = "" Then txtViewNRecords.Text = "100"
            Query = "Select Top " & txtViewNRecords.Text & " * From [" & TableName & "]"
        End If

        txtQuery.Text = Query

        'connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & Main.EpsgDatabasePath
        connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & EpsgDatabasePath
        myConnection.ConnectionString = connString
        myConnection.Open()

        da = New OleDb.OleDbDataAdapter(Query, myConnection)

        da.MissingSchemaAction = MissingSchemaAction.AddWithKey 'This statement is required to obtain the correct result from the statement: ds.Tables(0).Columns(0).MaxLength (This fixes a Microsoft bug: http://support.microsoft.com/kb/317175 )

        ds.Clear()
        ds.Reset()

        da.FillSchema(ds, SchemaType.Source, TableName)

        da.Fill(ds, TableName)

        DataGridView1.AutoGenerateColumns = True

        DataGridView1.EditMode = DataGridViewEditMode.EditOnKeystroke

        DataGridView1.DataSource = ds.Tables(0)
        DataGridView1.AutoResizeColumns()

        DataGridView1.Update()
        DataGridView1.Refresh()
        myConnection.Close()

        txtNRecords.Text = ds.Tables(0).Rows.Count
    End Sub

    Private Sub btnApplyQuery_Click(sender As Object, e As EventArgs) Handles btnApplyQuery.Click
        'Apply the SQL Query in txtQuary:
        ApplyQuery()

        'Try
        '    Query = txtQuery.Text

        '    connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & Main.EpsgDatabasePath
        '    myConnection.ConnectionString = connString
        '    myConnection.Open()

        '    da = New OleDb.OleDbDataAdapter(Query, myConnection)

        '    da.MissingSchemaAction = MissingSchemaAction.AddWithKey 'This statement is required to obtain the correct result from the statement: ds.Tables(0).Columns(0).MaxLength (This fixes a Microsoft bug: http://support.microsoft.com/kb/317175 )

        '    ds.Clear()
        '    ds.Reset()

        '    da.FillSchema(ds, SchemaType.Source, TableName)

        '    'Try
        '    da.Fill(ds, TableName)

        '    DataGridView1.AutoGenerateColumns = True

        '    DataGridView1.EditMode = DataGridViewEditMode.EditOnKeystroke

        '    DataGridView1.DataSource = ds.Tables(0)
        '    DataGridView1.AutoResizeColumns()

        '    DataGridView1.Update()
        '    DataGridView1.Refresh()
        '    'Catch ex As Exception
        '    '    Main.Message.AddWarning("Error: " & ex.Message & vbCrLf)
        '    'End Try


        '    txtNRecords.Text = ds.Tables(0).Rows.Count
        '    myConnection.Close()
        'Catch ex As Exception
        '    Main.Message.AddWarning(ex.Message & vbCrLf)
        '    myConnection.Close()
        'End Try

    End Sub

    Private Sub ApplyQuery()

        Try
            Query = txtQuery.Text

            'connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & Main.EpsgDatabasePath
            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & EpsgDatabasePath
            myConnection.ConnectionString = connString
            myConnection.Open()

            da = New OleDb.OleDbDataAdapter(Query, myConnection)

            da.MissingSchemaAction = MissingSchemaAction.AddWithKey 'This statement is required to obtain the correct result from the statement: ds.Tables(0).Columns(0).MaxLength (This fixes a Microsoft bug: http://support.microsoft.com/kb/317175 )

            ds.Clear()
            ds.Reset()

            da.FillSchema(ds, SchemaType.Source, TableName)

            'Try
            da.Fill(ds, TableName)

            DataGridView1.AutoGenerateColumns = True

            DataGridView1.EditMode = DataGridViewEditMode.EditOnKeystroke

            DataGridView1.DataSource = ds.Tables(0)
            DataGridView1.AutoResizeColumns()

            DataGridView1.Update()
            DataGridView1.Refresh()
            'Catch ex As Exception
            '    Main.Message.AddWarning("Error: " & ex.Message & vbCrLf)
            'End Try


            txtNRecords.Text = ds.Tables(0).Rows.Count
            myConnection.Close()
        Catch ex As Exception
            Main.Message.AddWarning(ex.Message & vbCrLf)
            myConnection.Close()
        End Try
    End Sub

    Private Sub rbViewFirst_CheckedChanged(sender As Object, e As EventArgs) Handles rbViewFirst.CheckedChanged

    End Sub

    Private Sub rbViewAll_CheckedChanged(sender As Object, e As EventArgs) Handles rbViewAll.CheckedChanged

        Dim Query As String
        If rbViewAll.Checked Then
            Query = "Select * From [" & TableName & "]"
        Else
            Query = "Select Top " & txtViewNRecords.Text & " * From [" & TableName & "]"
        End If

        If txtQuery.Text = Query Then

        Else
            txtQuery.Text = Query
            ApplyQuery()
        End If

    End Sub

#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Events - Events that can be triggered by this form." '==========================================================================================================================

#End Region 'Form Events ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class