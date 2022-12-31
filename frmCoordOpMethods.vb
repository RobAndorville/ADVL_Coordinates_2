Public Class frmCoordOpMethods
    'This form template shows the code required for a new form.

#Region " Variable Declarations - All the variables used in this form and this application." '=================================================================================================

    Dim MethodSearch As DataSet = New DataSet

#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Properties - All the properties used in this form and this application" '============================================================================================================

    Private _epsgDatabasePath = "" 'The path of the EPSG Database. This database contains a comprehensive set of coordinate reference system parameters.
    Property EpsgDatabasePath
        Get
            Return _epsgDatabasePath
        End Get
        Set(value)
            _epsgDatabasePath = value
        End Set
    End Property

    Private _selMethodCode As Integer = -1 'The code number of the selected Coordinate Operation Method.
    Property SelMethodCode As Integer
        Get
            Return _selMethodCode
        End Get
        Set(value As Integer)
            _selMethodCode = value
            ShowMethodInfo
        End Set
    End Property

    Private _selRowNo As Integer = -1 'The selected row number in the list of coordinate operation methods.
    Property SelRowNo As Integer
        Get
            Return _selRowNo
        End Get
        Set(value As Integer)
            _selRowNo = value
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
                               <InputMethodQuery><%= txtQuery.Text %></InputMethodQuery>
                               <SelectedTabIndex><%= TabControl1.SelectedIndex %></SelectedTabIndex>
                               <InputSplit1Dist><%= SplitContainer1.SplitterDistance %></InputSplit1Dist>
                               <SelMethodCode><%= SelMethodCode %></SelMethodCode>
                               <SelRowNo><%= SelRowNo %></SelRowNo>
                           </FormSettings>

        'Add code to include other settings to save after the comment line <!---->

        Dim SettingsFileName As String = "FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & ".xml"
        Main.Project.SaveXmlSettings(SettingsFileName, settingsData)
    End Sub

    Private Sub RestoreFormSettings()
        'Read the form settings from an XML document.

        Dim SettingsFileName As String = "FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & ".xml"

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
            If Settings.<FormSettings>.<SelectedTabIndex>.Value <> Nothing Then TabControl1.SelectedIndex = Settings.<FormSettings>.<SelectedTabIndex>.Value
            If Settings.<FormSettings>.<InputSplit1Dist>.Value <> Nothing Then SplitContainer1.SplitterDistance = Settings.<FormSettings>.<InputSplit1Dist>.Value
            If Settings.<FormSettings>.<InputMethodQuery>.Value <> Nothing Then
                txtQuery.Text = Settings.<FormSettings>.<InputMethodQuery>.Value
                ApplyInputQuery()
            End If
            If Settings.<FormSettings>.<SelMethodCode>.Value <> Nothing Then
                SelMethodCode = Settings.<FormSettings>.<SelMethodCode>.Value
            End If
            If Settings.<FormSettings>.<SelRowNo>.Value <> Nothing Then SelRowNo = Settings.<FormSettings>.<SelRowNo>.Value

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

        udRowNo.Minimum = -1
        udRowNo.Increment = 1
        txtQuery.Text = "Select COORD_OP_METHOD_CODE, COORD_OP_METHOD_NAME, REVERSE_OP, DEPRECATED, REMARKS From [Coordinate_Operation Method]"
        RestoreFormSettings()   'Restore the form settings
        ApplyInputQuery()

        'udRowNo.Minimum = -1
        'udRowNo.Increment = 1
        If MethodSearch.Tables.Contains("List") Then
            udRowNo.Maximum = MethodSearch.Tables("List").Rows.Count - 1
            udRowNo.Value = SelRowNo
        Else
            udRowNo.Maximum = -1
            udRowNo.Value = -1
        End If

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Form
        Me.Close() 'Close the form
    End Sub

    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If WindowState = FormWindowState.Normal Then
            SaveFormSettings()
        Else
            'Dont save settings if the form is minimised.
        End If
    End Sub



#End Region 'Form Display Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Open and Close Forms - Code used to open and close other forms." '===================================================================================================================

#End Region 'Open and Close Forms -------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Methods - The main actions performed by this form." '===========================================================================================================================

    Private Sub btnFind_Click(sender As Object, e As EventArgs) Handles btnFind.Click
        If txtFind.Text.Trim = "" Then
            txtQuery.Text = "Select COORD_OP_METHOD_CODE, COORD_OP_METHOD_NAME, REVERSE_OP, DEPRECATED, REMARKS From [Coordinate_Operation Method]"
            ApplyInputQuery()
        Else
            txtQuery.Text = "Select COORD_OP_METHOD_CODE, COORD_OP_METHOD_NAME, REVERSE_OP, DEPRECATED, REMARKS From [Coordinate_Operation Method] Where COORD_OP_METHOD_NAME Like '%" & txtFind.Text.Trim & "%'"
            ApplyInputQuery()
        End If
    End Sub

    Private Sub ApplyInputQuery()
        'Apply the CRS search query.

        If EpsgDatabasePath = "" Then
            Main.Message.AddWarning("No EPSG database has been selected." & vbCrLf)
            Exit Sub
        End If

        If Not System.IO.File.Exists(EpsgDatabasePath) Then
            Main.Message.AddWarning("Selected EPSG database can not be found." & vbCrLf)
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + EpsgDatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(txtQuery.Text, conn)
        If MethodSearch.Tables.Contains("List") Then MethodSearch.Tables("List").Clear() 'Clear any previous search results.
        da.Fill(MethodSearch, "List")
        conn.Close()

        dgvMethodList.DataSource = MethodSearch.Tables("List")

        dgvMethodList.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgvMethodList.AutoResizeColumns()

        If MethodSearch.Tables.Contains("List") Then
            udRowNo.Maximum = MethodSearch.Tables("List").Rows.Count - 1
            udRowNo.Value = -1
            txtNInputRecords.Text = MethodSearch.Tables("List").Rows.Count
        Else
            txtNInputRecords.Text = "0"
        End If



    End Sub

    Private Sub btnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click
        'Select the Coordinate Operation Method
        Dim SelRowCount As Integer = dgvMethodList.SelectedRows.Count

        If SelRowCount = 0 Then

        ElseIf SelRowCount = 1 Then
            'Dim SelRow As Integer = dgvMethodList.SelectedRows(0).Index
            SelRowNo = dgvMethodList.SelectedRows(0).Index
            udRowNo.Value = SelRowNo
            'Dim SelMethodCode As Integer = dgvMethodList.Rows(SelRow).Cells(0).Value
            'SelMethodCode = dgvMethodList.Rows(SelRow).Cells(0).Value
            SelMethodCode = dgvMethodList.Rows(SelRowNo).Cells(0).Value

        Else

        End If
    End Sub

    Public Sub ShowMethodInfo()

        If EpsgDatabasePath = "" Then
            Main.Message.AddWarning("No EPSG database has been selected." & vbCrLf)
            Exit Sub
        End If

        If Not System.IO.File.Exists(EpsgDatabasePath) Then
            Main.Message.AddWarning("Selected EPSG database can not be found." & vbCrLf)
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + EpsgDatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter("Select * From [Coordinate_Operation Method] Where COORD_OP_METHOD_CODE = " & SelMethodCode, conn)
        If MethodSearch.Tables.Contains("SelMethod") Then MethodSearch.Tables("SelMethod").Clear() 'Clear any previous search results.
        da.Fill(MethodSearch, "SelMethod")
        conn.Close()

        'txtInputCrsProjMethodName.Text = Conversion.InputCRS.ProjectionCoordOpMethod.Name
        'txtInputCrsProjMethodReversable.Text = Conversion.InputCRS.ProjectionCoordOpMethod.ReverseOp.ToString
        'txtInputCrsProjMethodFormula.Text = Conversion.InputCRS.ProjectionCoordOpMethod.Formula
        'txtInputCrsProjMethodExample.Text = Conversion.InputCRS.ProjectionCoordOpMethod.Example
        'txtInputCrsProjMethodCode.Text = Conversion.InputCRS.ProjectionCoordOpMethod.Code
        'txtInputCrsProjMethodRemarks.Text = Conversion.InputCRS.ProjectionCoordOpMethod.Remarks

        If MethodSearch.Tables("SelMethod").Rows.Count = 0 Then
            txtMethodName.Text = ""
            txtMethodReversable.Text = ""
            txtMethodFormula.Text = ""
            txtMethodExample.Text = ""
            txtMethodCode.Text = ""
            txtMethodRemarks.Text = ""
        ElseIf MethodSearch.Tables("SelMethod").Rows.Count = 1 Then
            txtMethodName.Text = MethodSearch.Tables("SelMethod").Rows(0).Item("COORD_OP_METHOD_NAME")
            txtMethodReversable.Text = MethodSearch.Tables("SelMethod").Rows(0).Item("REVERSE_OP")
            'txtMethodFormula.Text = MethodSearch.Tables("SelMethod").Rows(0).Item("FORMULA")
            If IsDBNull(MethodSearch.Tables("SelMethod").Rows(0).Item("FORMULA")) Then txtMethodFormula.Text = "" Else txtMethodFormula.Text = MethodSearch.Tables("SelMethod").Rows(0).Item("FORMULA")
            'txtMethodExample.Text = MethodSearch.Tables("SelMethod").Rows(0).Item("EXAMPLE")
            If IsDBNull(MethodSearch.Tables("SelMethod").Rows(0).Item("EXAMPLE")) Then txtMethodRemarks.Text = "" Else txtMethodExample.Text = MethodSearch.Tables("SelMethod").Rows(0).Item("EXAMPLE")
            txtMethodCode.Text = MethodSearch.Tables("SelMethod").Rows(0).Item("COORD_OP_METHOD_CODE")
            'txtMethodRemarks.Text = MethodSearch.Tables("SelMethod").Rows(0).Item("REMARKS")
            If IsDBNull(MethodSearch.Tables("SelMethod").Rows(0).Item("REMARKS")) Then txtMethodRemarks.Text = "" Else txtMethodRemarks.Text = MethodSearch.Tables("SelMethod").Rows(0).Item("REMARKS")
        Else
                txtMethodName.Text = MethodSearch.Tables("SelMethod").Rows(0).Item("COORD_OP_METHOD_NAME")
            txtMethodReversable.Text = MethodSearch.Tables("SelMethod").Rows(0).Item("REVERSE_OP")
            'txtMethodFormula.Text = MethodSearch.Tables("SelMethod").Rows(0).Item("FORMULA")
            If IsDBNull(MethodSearch.Tables("SelMethod").Rows(0).Item("FORMULA")) Then txtMethodFormula.Text = "" Else txtMethodFormula.Text = MethodSearch.Tables("SelMethod").Rows(0).Item("FORMULA")
            'txtMethodExample.Text = MethodSearch.Tables("SelMethod").Rows(0).Item("EXAMPLE")
            If IsDBNull(MethodSearch.Tables("SelMethod").Rows(0).Item("EXAMPLE")) Then txtMethodRemarks.Text = "" Else txtMethodExample.Text = MethodSearch.Tables("SelMethod").Rows(0).Item("EXAMPLE")
            txtMethodCode.Text = MethodSearch.Tables("SelMethod").Rows(0).Item("COORD_OP_METHOD_CODE")
            'txtMethodRemarks.Text = MethodSearch.Tables("SelMethod").Rows(0).Item("REMARKS")
            If IsDBNull(MethodSearch.Tables("SelMethod").Rows(0).Item("REMARKS")) Then txtMethodRemarks.Text = "" Else txtMethodRemarks.Text = MethodSearch.Tables("SelMethod").Rows(0).Item("REMARKS")
            Main.Message.AddWarning("There are " & MethodSearch.Tables("SelMethod").Rows.Count & " Coordinate Operation Records with code number: " & SelMethodCode & vbCrLf)
        End If


    End Sub

    Private Sub udRowNo_ValueChanged(sender As Object, e As EventArgs) Handles udRowNo.ValueChanged
        'Select another Coordinate Projection Method in the list.

        SelRowNo = udRowNo.Value
        If SelRowNo = -1 Then

        Else
            If dgvMethodList.Rows.Count > SelRowNo + 1 Then
                SelMethodCode = dgvMethodList.Rows(SelRowNo).Cells(0).Value
                dgvMethodList.ClearSelection()
                dgvMethodList.Rows(SelRowNo).Selected = True
            End If
        End If

    End Sub

    Private Sub Label44_Click(sender As Object, e As EventArgs) Handles Label44.Click

    End Sub

    Private Sub btnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click
        ApplyInputQuery()
    End Sub



#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Events - Events that can be triggered by this form." '==========================================================================================================================

#End Region 'Form Events ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class