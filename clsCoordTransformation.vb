
Imports System.Reflection
Imports System.Runtime.CompilerServices.RuntimeHelpers
Imports System.Text.RegularExpressions
Imports System.Windows.Forms.AxHost
Imports ADVL_Coordinates_2.clsAngle

Public Class clsCoordTransformation
    'Coordinate Transformation Class.

#Region " Variable Declarations - All the variables used in this class." '=====================================================================================================================

    Public WithEvents InputCrs As New CoordRefSystem 'Stores information about the Input coordinate reference system.
    Public WithEvents OutputCrs As New CoordRefSystem

    Public DirectDatumTransOpList As New List(Of DatumTransOp) 'List of Coordinate Operations that can perform a direct Datum Transformation from the Input Datum to the Output Datum.

    Public InputToWgs84TransOpList As New List(Of DatumTransOp) 'List of Coordinate Operations that can perform a Datum Transformation from the Input Datum to the WGS 84 Datum.
    Public OutputFromWgs84TransOpList As New List(Of DatumTransOp) 'List of Coordinate Operations that can ferform a Datum Transformation from the WGS 84 Datum to the Output Datum.

    Public UnitOfMeas As New Dictionary(Of Integer, UnitOfMeasure) 'Dictionary of Unit Of Measures

    Public WithEvents DatumTrans As New clsDatumTrans

    Public WithEvents Angle As New clsAngle 'Used to convert between different angle formats.

#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Properties" '========================================================================================================================================================================

    Private _epsgDatabasePath = "" 'The path of the EPSG Database. This database contains a comprehensive set of coordinate reference system parameters.



    Property EpsgDatabasePath
        Get
            Return _epsgDatabasePath
        End Get
        Set(value)
            _epsgDatabasePath = value
            InputCrs.EpsgDatabasePath = _epsgDatabasePath
            OutputCrs.EpsgDatabasePath = _epsgDatabasePath
            DatumTrans.EpsgDatabasePath = _epsgDatabasePath
            GetUOMs()
        End Set
    End Property

#End Region 'Properties -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods" '===========================================================================================================================================================================

    Public Sub New()
        'DatumTrans needs to access the cartesian coordinates in InputCrs and OutputCrs:
        DatumTrans.SourceCoord = InputCrs.Coord
        DatumTrans.TargetCoord = OutputCrs.Coord
    End Sub

    Private Sub DisplayCrsInfo(CRS As CoordRefSystem)
        'Display the Coordinate Reference System information on the Message Window

        RaiseEvent Message(vbCrLf & "Coordinate Reference System Information:" & vbCrLf)
        RaiseEvent Message("Code: " & CRS.Code & vbCrLf)
        RaiseEvent Message("Name: " & CRS.Name & vbCrLf)
        RaiseEvent Message("Kind: " & CRS.Kind.ToString & vbCrLf)
        RaiseEvent Message("Coordinate system code: " & CRS.CoordSysCode & vbCrLf)
        RaiseEvent Message("Datum code: " & CRS.DatumCode & vbCrLf)
        RaiseEvent Message("Base CRS code: " & CRS.BaseCrsCode & vbCrLf)
        RaiseEvent Message("Projection conversion code: " & CRS.ProjConvCode & vbCrLf)
        RaiseEvent Message("Code of the horizontal component of the Compound CRS: " & CRS.CmpdHorizCrsCode & vbCrLf)
        RaiseEvent Message("Code of the vertical component of the Compound CRS: " & CRS.CmpdVertCrsCode & vbCrLf)
        RaiseEvent Message("Remarks: " & CRS.Remarks & vbCrLf)
        RaiseEvent Message("Information source: " & CRS.InfoSource & vbCrLf)
        RaiseEvent Message("Data source: " & CRS.DataSource & vbCrLf)
        RaiseEvent Message("Revision date: " & CRS.RevisionDate & vbCrLf)
        RaiseEvent Message("Change ID: " & CRS.ChangeID & vbCrLf)
        RaiseEvent Message("Show: " & CRS.Show.ToString & vbCrLf)
        RaiseEvent Message("Deprecated: " & CRS.Deprecated.ToString & vbCrLf)

    End Sub

    Private Sub InputCrs_ErrorMessage(Msg As String) Handles InputCrs.ErrorMessage
        RaiseEvent ErrorMessage(Msg)
    End Sub

    Private Sub InputCrs_Message(Msg As String) Handles InputCrs.Message
        RaiseEvent Message(Msg)
    End Sub


    Private Sub GetUOMs()
        'Get the Unit of Measures.

        If EpsgDatabasePath = "" Then
            RaiseEvent ErrorMessage("No EPSG database has been selected." & vbCrLf)
            Exit Sub
        End If

        If Not System.IO.File.Exists(EpsgDatabasePath) Then
            RaiseEvent ErrorMessage("Selected EPSG database can not be found." & vbCrLf)
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + EpsgDatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter("Select * From [Unit of Measure]", conn)
        Dim ds As DataSet = New DataSet
        da.Fill(ds, "UOM")

        If ds.Tables("UOM").Rows.Count = 0 Then
            RaiseEvent ErrorMessage("No Unit of Measures found." & vbCrLf)
        Else
            UnitOfMeas.Clear()
            For Each Item As DataRow In ds.Tables("UOM").Rows
                Dim NewUom As New UnitOfMeasure
                If IsDBNull(Item("UOM_CODE")) Then
                    'No UomCode - this record can not be added to the UnitOfMeas dictionary.
                    RaiseEvent ErrorMessage("Unit of Measure not found." & vbCrLf)
                Else
                    NewUom.Code = Item("UOM_CODE")
                    If IsDBNull(Item("UNIT_OF_MEAS_NAME")) Then NewUom.Name = "" Else NewUom.Name = Item("UNIT_OF_MEAS_NAME")
                    If IsDBNull(Item("UNIT_OF_MEAS_TYPE")) Then
                        NewUom.Type = UnitOfMeasure.UOMType.Length
                    Else
                        Select Case Item("UNIT_OF_MEAS_TYPE")
                            Case "angle"
                                NewUom.Type = UnitOfMeasure.UOMType.Angle
                            Case "length"
                                NewUom.Type = UnitOfMeasure.UOMType.Length
                            Case "scale"
                                NewUom.Type = UnitOfMeasure.UOMType.Scale
                            Case "time"
                                NewUom.Type = UnitOfMeasure.UOMType.Time
                            Case Else
                                RaiseEvent ErrorMessage("Unknown Unit of Measure: " & Item("UNIT_OF_MEAS_TYPE") & vbCrLf)
                                NewUom.Type = UnitOfMeasure.UOMType.Length
                        End Select
                    End If
                    If IsDBNull(Item("TARGET_UOM_CODE")) Then NewUom.TargetUomCode = -1 Else NewUom.TargetUomCode = Item("TARGET_UOM_CODE")
                    If IsDBNull(Item("FACTOR_B")) Then NewUom.FactorB = Double.NaN Else NewUom.FactorB = Item("FACTOR_B")
                    If IsDBNull(Item("FACTOR_C")) Then NewUom.FactorC = Double.NaN Else NewUom.FactorC = Item("FACTOR_C")
                    If IsDBNull(Item("REMARKS")) Then NewUom.Remarks = "" Else NewUom.Remarks = Item("REMARKS")
                    If IsDBNull(Item("INFORMATION_SOURCE")) Then NewUom.InfoSource = "" Else NewUom.InfoSource = Item("INFORMATION_SOURCE")
                    If IsDBNull(Item("DATA_SOURCE")) Then NewUom.DataSource = "" Else NewUom.DataSource = Item("DATA_SOURCE")
                    If IsDBNull(Item("REVISION_DATE")) Then NewUom.RevisionDate = Date.MinValue Else NewUom.RevisionDate = Item("REVISION_DATE")
                    If IsDBNull(Item("CHANGE_ID")) Then NewUom.ChangeID = "" Else NewUom.ChangeID = Item("CHANGE_ID")
                    If IsDBNull(Item("DEPRECATED")) Then NewUom.Deprecated = False Else NewUom.ChangeID = Item("DEPRECATED")
                    UnitOfMeas.Add(NewUom.Code, NewUom)
                End If
            Next
        End If
    End Sub

    Private Sub OutputCrs_ErrorMessage(Msg As String) Handles OutputCrs.ErrorMessage
        RaiseEvent ErrorMessage(Msg)
    End Sub

    Private Sub OutputCrs_Message(Msg As String) Handles OutputCrs.Message
        RaiseEvent Message(Msg)
    End Sub

    Private Sub GetDirectDatumTransOpList()
        'Get the list of coordinate operations that can be used to transform the datum directly from the Input CRS to the Output CRS.
        DirectDatumTransOpList.Clear()
        ScanInputCrsCodes(InputCrs, 0, OutputCrs, 0, False)
        ScanInputCrsCodes(OutputCrs, 0, InputCrs, 0, True) 'Check coordinate operation with Source and Target Crs codes reversed. If the operation is reversible, it can be used.
    End Sub

    Private Sub ScanInputCrsCodes(InCrs As CoordRefSystem, InLevel As Integer, OutCrs As CoordRefSystem, OutLevel As Integer, Reverse As Boolean)
        'Scan all of the Crs and BaseCrs codes for the Input CRS.
        'Reverse is True if the Source and Target CRS codes are reversed - this is done to check if a reversible coordinate operation can be used for the datum transformation.

        ScanOutputCrsCodes(InCrs, InLevel, OutCrs, OutLevel, Reverse)
        If IsNothing(InCrs.BaseCrs) Then
            'No BaseCRS to process.
        Else
            ScanOutputCrsCodes(InCrs.BaseCrs, InLevel + 1, OutCrs, OutLevel, Reverse)
        End If
    End Sub

    Private Sub ScanOutputCrsCodes(InCrs As CoordRefSystem, InLevel As Integer, OutCrs As CoordRefSystem, OutLevel As Integer, Reverse As Boolean)
        'Scan all of the Crs and BaseCrs codes for the Output CRS.
        'Reverse is True if the Source and Target CRS codes are reversed - this is done to check if a reversible coordinate operation can be used for the datum transformation.

        GetDirectDatumTransOpList(InCrs, InLevel, OutCrs, OutLevel, Reverse)
        If IsNothing(OutCrs.BaseCrs) Then
            'No BaseCRS to process.
        Else
            ScanOutputCrsCodes(InCrs, InLevel, OutCrs.BaseCrs, OutLevel + 1, Reverse)
        End If
    End Sub

    Private Sub GetDirectDatumTransOpList(InCrs As CoordRefSystem, InLevel As Integer, OutCrs As CoordRefSystem, OutLevel As Integer, Reverse As Boolean)
        'Get the list of coordinate operations that can be used to transform the datum directly from the Input CRS to the Output CRS.
        'This method recursively finds the each coordinate operation that uses the Input CRS or any level of Input BaseCrs as the Source CRS and the Output CRS or any level of Output BaseCrs as the Target CRS.
        'Reverse is True if the Source and Target CRS codes are reversed - this is done to check if a reversible coordinate operation can be used for the datum transformation.

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        'Dim dt As DataTable

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + EpsgDatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        'Get list of Input Source Coordinate Operations:
        'Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(
        '    "Select a.COORD_OP_CODE, a.COORD_OP_NAME, a.COORD_OP_TYPE, a.COORD_OP_ACCURACY, a.SOURCE_CRS_CODE, a.TARGET_CRS_CODE, b.REVERSE_OP, a.DEPRECATED, a.COORD_TFM_VERSION, a.COORD_OP_VARIANT, a.REVISION_DATE " &
        '    " From Coordinate_Operation a, [Coordinate_Operation Method] b Where b.COORD_OP_METHOD_CODE = a.COORD_OP_METHOD_CODE " &
        '    " And a.SOURCE_CRS_CODE = " & InCrs.Code &
        '    " And a.TARGET_CRS_CODE = " & OutCrs.Code, conn)
        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(
            "Select a.COORD_OP_CODE, a.COORD_OP_NAME, a.COORD_OP_TYPE, a.COORD_OP_ACCURACY, a.SOURCE_CRS_CODE, a.TARGET_CRS_CODE, b.REVERSE_OP, a.DEPRECATED, a.COORD_TFM_VERSION, a.COORD_OP_VARIANT, a.REVISION_DATE, b.COORD_OP_METHOD_CODE, b.COORD_OP_METHOD_NAME " &
            " From Coordinate_Operation a, [Coordinate_Operation Method] b Where b.COORD_OP_METHOD_CODE = a.COORD_OP_METHOD_CODE " &
            " And a.SOURCE_CRS_CODE = " & InCrs.Code &
            " And a.TARGET_CRS_CODE = " & OutCrs.Code, conn)
        Dim ds As DataSet = New DataSet
        da.Fill(ds, "CoordOps")
        conn.Close()

        'Write acceptable coordinate operations to the DirectDatumTransOpList:
        For Each Row As DataRow In ds.Tables("CoordOps").Rows
            Dim NewCoordOp As New DatumTransOp
            NewCoordOp.Name = Row.Item("COORD_OP_NAME")
            NewCoordOp.Type = Row.Item("COORD_OP_TYPE")
            NewCoordOp.Code = Row.Item("COORD_OP_CODE")
            NewCoordOp.Accuracy = Row.Item("COORD_OP_ACCURACY")
            NewCoordOp.Deprecated = Row.Item("DEPRECATED")
            NewCoordOp.Version = Row.Item("COORD_TFM_VERSION")
            NewCoordOp.RevisionDate = Row.Item("REVISION_DATE")
            'NewCoordOp.SourceCrsLevel = Row.Item("InputLevel")
            NewCoordOp.SourceCrsLevel = InLevel
            NewCoordOp.SourceCrsCode = Row.Item("SOURCE_CRS_CODE")
            'NewCoordOp.TargetCrsLevel = Row.Item("OutputLevel")
            NewCoordOp.TargetCrsLevel = OutLevel
            NewCoordOp.TargetCrsCode = Row.Item("TARGET_CRS_CODE")
            NewCoordOp.Reversible = Row.Item("REVERSE_OP")
            NewCoordOp.ApplyReverse = False 'By default the reverse coordinate transformation is not applied.
            NewCoordOp.MethodCode = Row.Item("COORD_OP_METHOD_CODE")
            NewCoordOp.MethodName = Row.Item("COORD_OP_METHOD_NAME")
            If DirectDatumTransOpList.Contains(NewCoordOp) Then
                'The Coordinate Operation is already in the list
            Else
                If Reverse Then 'Only add the coordinate operation if it is reversible.
                    If NewCoordOp.Reversible = True Then
                        NewCoordOp.ApplyReverse = True 'The reverse coordinate operation is applied.
                        DirectDatumTransOpList.Add(NewCoordOp)
                    End If
                Else
                    DirectDatumTransOpList.Add(NewCoordOp)
                End If
            End If
        Next
    End Sub

    Private Sub GetInputToWgs84DatumTransOpList()
        'Get the list of coordinate operations that can be used to transform the datum from the Input CRS to the WGS 84 datum.
        InputToWgs84TransOpList.Clear()
        ScanInputCrsCodes(InputCrs, 0)
    End Sub

    Private Sub ScanInputCrsCodes(InCrs As CoordRefSystem, InLevel As Integer)
        'Scan all of the Crs and BaseCrs codes for the Input CRS. This version does not use the Output CRS and is used to get the Input To WGS 84 Datum Trans Op list.
        'This method recursively finds the each coordinate operation that uses the Input CRS or any level of Input BaseCrs as the Source CRS and WGS 84 as the Target CRS.
        GetInputToWgs84DatumTransOpList(InCrs, InLevel)
        If IsNothing(InCrs.BaseCrs) Then
            'No BaseCRS to process.
        Else
            ScanInputCrsCodes(InCrs.BaseCrs, InLevel + 1)
        End If
    End Sub

    Private Sub GetInputToWgs84DatumTransOpList(InCrs As CoordRefSystem, InLevel As Integer)
        'Get the list of coordinate operations that can be used to transform the datum from the Input CRS to WGS 84.

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        'Dim dt As DataTable

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + EpsgDatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        'Get list of Input Source Coordinate Operations:
        'Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(
        '    "Select a.COORD_OP_CODE, a.COORD_OP_NAME, a.COORD_OP_TYPE, a.COORD_OP_ACCURACY, a.SOURCE_CRS_CODE, a.TARGET_CRS_CODE, b.REVERSE_OP, a.DEPRECATED, a.COORD_TFM_VERSION, a.COORD_OP_VARIANT, a.REVISION_DATE " &
        '    " From Coordinate_Operation a, [Coordinate_Operation Method] b Where b.COORD_OP_METHOD_CODE = a.COORD_OP_METHOD_CODE " &
        '    " And a.SOURCE_CRS_CODE = " & InCrs.Code &
        '    " And a.COORD_OP_NAME Like '%WGS 84%'", conn)
        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(
            "Select a.COORD_OP_CODE, a.COORD_OP_NAME, a.COORD_OP_TYPE, a.COORD_OP_ACCURACY, a.SOURCE_CRS_CODE, a.TARGET_CRS_CODE, b.REVERSE_OP, a.DEPRECATED, a.COORD_TFM_VERSION, a.COORD_OP_VARIANT, a.REVISION_DATE, b.COORD_OP_METHOD_CODE, b.COORD_OP_METHOD_NAME " &
            " From Coordinate_Operation a, [Coordinate_Operation Method] b Where b.COORD_OP_METHOD_CODE = a.COORD_OP_METHOD_CODE " &
            " And a.SOURCE_CRS_CODE = " & InCrs.Code &
            " And a.COORD_OP_NAME Like '%WGS 84%'", conn)
        Dim ds As DataSet = New DataSet
        da.Fill(ds, "CoordOps")

        'Write acceptable coordinate operations to the InputToWgs84TransOpList:
        For Each Row As DataRow In ds.Tables("CoordOps").Rows
            Dim NewCoordOp As New DatumTransOp
            NewCoordOp.Name = Row.Item("COORD_OP_NAME")
            NewCoordOp.Type = Row.Item("COORD_OP_TYPE")
            NewCoordOp.Code = Row.Item("COORD_OP_CODE")
            NewCoordOp.Accuracy = Row.Item("COORD_OP_ACCURACY")
            NewCoordOp.Deprecated = Row.Item("DEPRECATED")
            NewCoordOp.Version = Row.Item("COORD_TFM_VERSION")
            NewCoordOp.RevisionDate = Row.Item("REVISION_DATE")
            NewCoordOp.SourceCrsLevel = InLevel
            NewCoordOp.SourceCrsCode = Row.Item("SOURCE_CRS_CODE")
            NewCoordOp.TargetCrsLevel = 0
            NewCoordOp.TargetCrsCode = Row.Item("TARGET_CRS_CODE")
            NewCoordOp.Reversible = Row.Item("REVERSE_OP")
            NewCoordOp.ApplyReverse = False  'By default the reverse coordinate transformation is not applied.
            NewCoordOp.MethodCode = Row.Item("COORD_OP_METHOD_CODE")
            NewCoordOp.MethodName = Row.Item("COORD_OP_METHOD_NAME")
            If InputToWgs84TransOpList.Contains(NewCoordOp) Then
                'The Coordinate Operation is already in the list
            Else
                InputToWgs84TransOpList.Add(NewCoordOp)
            End If
        Next

        ''ds.Tables("CoordOp").Rows.Clear()
        'If ds.Tables.Contains("CoordOps") Then
        '    'If IsNothing(ds.Tables("CoordOps").Rows) Then
        '    If IsNothing(ds.Tables("CoordOps")) Then

        '    Else
        '        ds.Tables("CoordOp").Rows.Clear()
        '    End If
        'End If
        ds = New DataSet

        'Find suitable reversible coordinate operations:
        'da.SelectCommand.CommandText = "Select a.COORD_OP_CODE, a.COORD_OP_NAME, a.COORD_OP_TYPE, a.COORD_OP_ACCURACY, a.SOURCE_CRS_CODE, a.TARGET_CRS_CODE, b.REVERSE_OP, a.DEPRECATED, a.COORD_TFM_VERSION, a.COORD_OP_VARIANT, a.REVISION_DATE " &
        '    " From Coordinate_Operation a, [Coordinate_Operation Method] b Where b.COORD_OP_METHOD_CODE = a.COORD_OP_METHOD_CODE " &
        '    " And a.TARGET_CRS_CODE = " & InCrs.Code &
        '    " And a.COORD_OP_NAME Like '%WGS 84%'"
        da.SelectCommand.CommandText = "Select a.COORD_OP_CODE, a.COORD_OP_NAME, a.COORD_OP_TYPE, a.COORD_OP_ACCURACY, a.SOURCE_CRS_CODE, a.TARGET_CRS_CODE, b.REVERSE_OP, a.DEPRECATED, a.COORD_TFM_VERSION, a.COORD_OP_VARIANT, a.REVISION_DATE, b.COORD_OP_METHOD_CODE, b.COORD_OP_METHOD_NAME " &
            " From Coordinate_Operation a, [Coordinate_Operation Method] b Where b.COORD_OP_METHOD_CODE = a.COORD_OP_METHOD_CODE " &
            " And a.TARGET_CRS_CODE = " & InCrs.Code &
            " And a.COORD_OP_NAME Like '%WGS 84%'"
        da.Fill(ds, "CoordOps")

        'Write acceptable coordinate operations to the InputToWgs84TransOpList:
        For Each Row As DataRow In ds.Tables("CoordOps").Rows
            Dim NewCoordOp As New DatumTransOp
            NewCoordOp.Name = Row.Item("COORD_OP_NAME")
            NewCoordOp.Type = Row.Item("COORD_OP_TYPE")
            NewCoordOp.Code = Row.Item("COORD_OP_CODE")
            NewCoordOp.Accuracy = Row.Item("COORD_OP_ACCURACY")
            NewCoordOp.Deprecated = Row.Item("DEPRECATED")
            NewCoordOp.Version = Row.Item("COORD_TFM_VERSION")
            NewCoordOp.RevisionDate = Row.Item("REVISION_DATE")
            NewCoordOp.SourceCrsLevel = 0
            NewCoordOp.SourceCrsCode = Row.Item("SOURCE_CRS_CODE")
            NewCoordOp.TargetCrsLevel = InLevel
            NewCoordOp.TargetCrsCode = Row.Item("TARGET_CRS_CODE")
            NewCoordOp.Reversible = Row.Item("REVERSE_OP")
            NewCoordOp.ApplyReverse = True  'If valid the reverse coordinate transformation is applied.
            NewCoordOp.MethodCode = Row.Item("COORD_OP_METHOD_CODE")
            NewCoordOp.MethodName = Row.Item("COORD_OP_METHOD_NAME")
            If InputToWgs84TransOpList.Contains(NewCoordOp) Then
                'The Coordinate Operation is already in the list
            Else
                If NewCoordOp.Reversible = True Then InputToWgs84TransOpList.Add(NewCoordOp) 'This operation can only be used if it is reversible.
            End If
        Next

        conn.Close()

    End Sub

    Private Sub GetWgs84ToOutoutCrsDatumTransOpList()
        'Get the list of coordinate operations that can be used to transform the datum from the Input CRS to the WGS 84 datum.
        OutputFromWgs84TransOpList.Clear()
        ScanOutputCrsCodes(OutputCrs, 0)
    End Sub

    Private Sub ScanOutputCrsCodes(OutCrs As CoordRefSystem, OutLevel As Integer)
        'Scan all of the Crs and BaseCrs codes for the Output CRS. This version does not use the Input CRS and is used to get the WGS 84 to Outout CRS Datum Trans Op list.
        'This method recursively finds the each coordinate operation that uses the WGS 84 as the Source CRS and the Output CRS or any level of Output BaseCrs as the Target CRS.
        GetWgs84DatumToOutputTransOpList(OutCrs, OutLevel)
        If IsNothing(OutCrs.BaseCrs) Then
            'No BaseCRS to process.
        Else
            ScanOutputCrsCodes(OutCrs.BaseCrs, OutLevel + 1)
        End If
    End Sub

    Private Sub GetWgs84DatumToOutputTransOpList(OutCrs As CoordRefSystem, OutLevel As Integer)
        'Get the list of coordinate operations that can be used to transform the datum from WGS 84 to the Output CRS.

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        'Dim dt As DataTable

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + EpsgDatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        'Get list of Input Source Coordinate Operations:
        'Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(
        '    "Select a.COORD_OP_CODE, a.COORD_OP_NAME, a.COORD_OP_TYPE, a.COORD_OP_ACCURACY, a.SOURCE_CRS_CODE, a.TARGET_CRS_CODE, b.REVERSE_OP, a.DEPRECATED, a.COORD_TFM_VERSION, a.COORD_OP_VARIANT, a.REVISION_DATE " &
        '    " From Coordinate_Operation a, [Coordinate_Operation Method] b Where b.COORD_OP_METHOD_CODE = a.COORD_OP_METHOD_CODE " &
        '    " And a.TARGET_CRS_CODE = " & OutCrs.Code &
        '    " And a.COORD_OP_NAME Like '%WGS 84%'", conn)
        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(
            "Select a.COORD_OP_CODE, a.COORD_OP_NAME, a.COORD_OP_TYPE, a.COORD_OP_ACCURACY, a.SOURCE_CRS_CODE, a.TARGET_CRS_CODE, b.REVERSE_OP, a.DEPRECATED, a.COORD_TFM_VERSION, a.COORD_OP_VARIANT, a.REVISION_DATE, b.COORD_OP_METHOD_CODE, b.COORD_OP_METHOD_NAME " &
            " From Coordinate_Operation a, [Coordinate_Operation Method] b Where b.COORD_OP_METHOD_CODE = a.COORD_OP_METHOD_CODE " &
            " And a.TARGET_CRS_CODE = " & OutCrs.Code &
            " And a.COORD_OP_NAME Like '%WGS 84%'", conn)
        Dim ds As DataSet = New DataSet
        da.Fill(ds, "CoordOps")

        'Write acceptable coordinate operations to the OutputFromWgs84TransOpList:
        For Each Row As DataRow In ds.Tables("CoordOps").Rows
            Dim NewCoordOp As New DatumTransOp
            NewCoordOp.Name = Row.Item("COORD_OP_NAME")
            NewCoordOp.Type = Row.Item("COORD_OP_TYPE")
            NewCoordOp.Code = Row.Item("COORD_OP_CODE")
            NewCoordOp.Accuracy = Row.Item("COORD_OP_ACCURACY")
            NewCoordOp.Deprecated = Row.Item("DEPRECATED")
            NewCoordOp.Version = Row.Item("COORD_TFM_VERSION")
            NewCoordOp.RevisionDate = Row.Item("REVISION_DATE")
            NewCoordOp.SourceCrsLevel = 0
            NewCoordOp.SourceCrsCode = Row.Item("SOURCE_CRS_CODE")
            NewCoordOp.TargetCrsLevel = OutLevel
            NewCoordOp.TargetCrsCode = Row.Item("TARGET_CRS_CODE")
            NewCoordOp.Reversible = Row.Item("REVERSE_OP")
            NewCoordOp.ApplyReverse = False  'By default the reverse coordinate transformation is not applied.
            NewCoordOp.MethodCode = Row.Item("COORD_OP_METHOD_CODE")
            NewCoordOp.MethodName = Row.Item("COORD_OP_METHOD_NAME")
            If OutputFromWgs84TransOpList.Contains(NewCoordOp) Then
                'The Coordinate Operation is already in the list
            Else
                OutputFromWgs84TransOpList.Add(NewCoordOp)
            End If
        Next

        ''ds.Tables("CoordOp").Rows.Clear()
        ''If ds.Tables.Contains("CoordOps") Then ds.Tables("CoordOp").Rows.Clear()
        'If ds.Tables.Contains("CoordOps") Then
        '    'If IsNothing(ds.Tables("CoordOps").Rows) Then
        '    If IsNothing(ds.Tables("CoordOps")) Then

        '    Else
        '        ds.Tables("CoordOp").Rows.Clear()
        '    End If
        'End If
        ds = New DataSet

        'Find suitable reversible coordinate operations:
        'da.SelectCommand.CommandText = "Select a.COORD_OP_CODE, a.COORD_OP_NAME, a.COORD_OP_TYPE, a.COORD_OP_ACCURACY, a.SOURCE_CRS_CODE, a.TARGET_CRS_CODE, b.REVERSE_OP, a.DEPRECATED, a.COORD_TFM_VERSION, a.COORD_OP_VARIANT, a.REVISION_DATE " &
        '    " From Coordinate_Operation a, [Coordinate_Operation Method] b Where b.COORD_OP_METHOD_CODE = a.COORD_OP_METHOD_CODE " &
        '    " And a.SOURCE_CRS_CODE = " & OutCrs.Code &
        '    " And a.COORD_OP_NAME Like '%WGS 84%'"
        da.SelectCommand.CommandText = "Select a.COORD_OP_CODE, a.COORD_OP_NAME, a.COORD_OP_TYPE, a.COORD_OP_ACCURACY, a.SOURCE_CRS_CODE, a.TARGET_CRS_CODE, b.REVERSE_OP, a.DEPRECATED, a.COORD_TFM_VERSION, a.COORD_OP_VARIANT, a.REVISION_DATE, b.COORD_OP_METHOD_CODE, b.COORD_OP_METHOD_NAME " &
            " From Coordinate_Operation a, [Coordinate_Operation Method] b Where b.COORD_OP_METHOD_CODE = a.COORD_OP_METHOD_CODE " &
            " And a.SOURCE_CRS_CODE = " & OutCrs.Code &
            " And a.COORD_OP_NAME Like '%WGS 84%'"
        da.Fill(ds, "CoordOps")

        'Write acceptable coordinate operations to the OutputFromWgs84TransOpList:
        For Each Row As DataRow In ds.Tables("CoordOps").Rows
            Dim NewCoordOp As New DatumTransOp
            NewCoordOp.Name = Row.Item("COORD_OP_NAME")
            NewCoordOp.Type = Row.Item("COORD_OP_TYPE")
            NewCoordOp.Code = Row.Item("COORD_OP_CODE")
            NewCoordOp.Accuracy = Row.Item("COORD_OP_ACCURACY")
            NewCoordOp.Deprecated = Row.Item("DEPRECATED")
            NewCoordOp.Version = Row.Item("COORD_TFM_VERSION")
            NewCoordOp.RevisionDate = Row.Item("REVISION_DATE")
            NewCoordOp.SourceCrsLevel = OutLevel
            NewCoordOp.SourceCrsCode = Row.Item("SOURCE_CRS_CODE")
            NewCoordOp.TargetCrsLevel = 0
            NewCoordOp.TargetCrsCode = Row.Item("TARGET_CRS_CODE")
            NewCoordOp.Reversible = Row.Item("REVERSE_OP")
            NewCoordOp.ApplyReverse = True  'If valid the reverse coordinate transformation is applied.
            NewCoordOp.MethodCode = Row.Item("COORD_OP_METHOD_CODE")
            NewCoordOp.MethodName = Row.Item("COORD_OP_METHOD_NAME")
            If OutputFromWgs84TransOpList.Contains(NewCoordOp) Then
                'The Coordinate Operation is already in the list
            Else
                If NewCoordOp.Reversible = True Then OutputFromWgs84TransOpList.Add(NewCoordOp) 'This operation can only be used if it is reversible.
            End If
        Next

        conn.Close()

    End Sub


    Private Sub GetInputToWgs84TransOpList()
        'Get the list of coordanate operations that can be used to transform the datum from the Input CRS to the WGS 84 CRS.


    End Sub

    Private Sub GetOutputFromWgs84TransOpList()
        'Get the list of coordanate operations that can be used to transform the datum from the WGS 84 CRS to the Output CRS.


    End Sub

    Private Sub InputCrs_Updated() Handles InputCrs.Updated
        'The InputCrs has been updated.
        'If the InputCrs and OutputCrs contain data then calculate the DirectDatumTransOpList'
        If InputCrs.Code > -1 Then
            If OutputCrs.Code > -1 Then
                GetDirectDatumTransOpList()
                GetInputToWgs84DatumTransOpList()
                GetWgs84ToOutoutCrsDatumTransOpList()
            Else
                'The OutputCrs does not contain data.
            End If
        Else
            'The InputCrs does not contain data.
        End If
    End Sub

    Private Sub OutputCrs_Updated() Handles OutputCrs.Updated
        'The OutputCrs has been updated.
        'If the InputCrs and OutputCrs contain data then calculate the DirectDatumTransOpList'
        If InputCrs.Code > -1 Then
            If OutputCrs.Code > -1 Then
                GetDirectDatumTransOpList()
                GetInputToWgs84DatumTransOpList()
                GetWgs84ToOutoutCrsDatumTransOpList()
            Else
                'The OutputCrs does not contain data.
            End If
        Else
            'The InputCrs does not contain data.
        End If
    End Sub

    Private Sub InputCrs_Update(Mode As Coordinate.UpdateMode, From As Coordinate.CoordType) Handles InputCrs.Update
        'Update coordinates from the specified coordinate type.
        'A coordinate in the Input Coordinate Reference System has been updated.
        'Update the coordinates specified by Mode using the coordinate specified by From.
        ' For Example: Mode = EastNorth, From = LongLat    : Update Input Easting, Northing From Input Longitude, Latitude
        '              Mode = TransLongLat, From = LongLat : Update Output Longitude, Latitude From Input Longitude, Latitude

        'Enum UpdateMode
        '    None
        '    InputOutputAll
        '    XYZ
        '    LongLat
        '    EastNorth
        '    All
        '    TransXYZ
        '    TransLongLat
        '    TransEastNorth
        '    TransAll
        'End Enum

        Select Case Mode
            Case Coordinate.UpdateMode.None 'No updates required.
                'No coordinates to update.

            Case Coordinate.UpdateMode.InputOutputAll 'Update all Input and Output coordinate types
                Select Case From
                    Case Coordinate.CoordType.XYZ 'From the Input X, Y, Z
                        InputCrs.XYZToLongLatEllHt() 'Calculate Longitude, Latitude and Ellipsoidal Height from X, Y, Z.
                        If InputCrs.Kind = CoordRefSystem.CrsKind.projected Then InputCrs.LongLatToEastNorth() 'Calculate Easting and Northing from Latitude and Longitude.
                        DatumTrans.InputToOutput() 'Calculate Output X, Y, Z from Input X, Y, Z
                        OutputCrs.XYZToLongLatEllHt() 'Calculate Output Longitude, Latitude and Ellipsoidal Height from Output X, Y, Z.
                        If OutputCrs.Kind = CoordRefSystem.CrsKind.projected Then OutputCrs.LongLatToEastNorth() 'Calculate Output Easting and Northing from Output Latitude and Longitude.

                    Case Coordinate.CoordType.LongLat 'From the Input Latitude, Longitude
                        InputCrs.LongLatEllHtToXYZ()
                        If InputCrs.Kind = CoordRefSystem.CrsKind.projected Then InputCrs.LongLatToEastNorth()
                        DatumTrans.InputToOutput()
                        OutputCrs.XYZToLongLatEllHt()
                        If OutputCrs.Kind = CoordRefSystem.CrsKind.projected Then OutputCrs.LongLatToEastNorth()

                    Case Coordinate.CoordType.EastNorth 'From the Input Easting, Northing
                        InputCrs.EastNorthToLongLat()
                        InputCrs.LongLatEllHtToXYZ()
                        DatumTrans.InputToOutput()
                        OutputCrs.XYZToLongLatEllHt()
                        If OutputCrs.Kind = CoordRefSystem.CrsKind.projected Then OutputCrs.LongLatToEastNorth()

                End Select

            Case Coordinate.UpdateMode.EastNorth 'Update Input Easting, Northing
                Select Case From
                    Case Coordinate.CoordType.XYZ 'From Input X, Y, Z
                        InputCrs.XYZToLongLatEllHt()
                        If InputCrs.Kind = CoordRefSystem.CrsKind.projected Then InputCrs.LongLatToEastNorth()

                    Case Coordinate.CoordType.LongLat 'From Input Longitude, Latitude
                        If InputCrs.Kind = CoordRefSystem.CrsKind.projected Then InputCrs.LongLatToEastNorth()

                End Select

            Case Coordinate.UpdateMode.LongLat  'Update Input Longitude, Latitude
                Select Case From
                    Case Coordinate.CoordType.XYZ 'From Input X, Y, Z
                        InputCrs.XYZToLongLatEllHt()

                    Case Coordinate.CoordType.EastNorth 'From Input Easting, Northing
                        InputCrs.EastNorthToLongLat()

                End Select

            Case Coordinate.UpdateMode.XYZ  'Update Input X, Y, Z
                Select Case From
                    Case Coordinate.CoordType.LongLat 'From Input Longitude, Latitude
                        InputCrs.LongLatEllHtToXYZ()

                    Case Coordinate.CoordType.EastNorth 'From Input Easting, Northing
                        InputCrs.EastNorthToLongLat()
                        InputCrs.LongLatEllHtToXYZ()

                End Select

            Case Coordinate.UpdateMode.All  'Update Input Coordinate types
                Select Case From
                    Case Coordinate.CoordType.XYZ 'From Input X, Y, Z
                        InputCrs.XYZToLongLatEllHt()
                        If InputCrs.Kind = CoordRefSystem.CrsKind.projected Then InputCrs.LongLatToEastNorth()

                    Case Coordinate.CoordType.LongLat 'From Input Longitude, Latitude
                        InputCrs.LongLatEllHtToXYZ()
                        If InputCrs.Kind = CoordRefSystem.CrsKind.projected Then InputCrs.LongLatToEastNorth()

                    Case Coordinate.CoordType.EastNorth 'From Input Easting, Northing
                        InputCrs.EastNorthToLongLat()
                        InputCrs.LongLatEllHtToXYZ()

                End Select

            Case Coordinate.UpdateMode.TransXYZ   'Update Output X, Y, Z
                Select Case From
                    Case Coordinate.CoordType.XYZ 'From Input X, Y, Z
                        DatumTrans.InputToOutput()

                    Case Coordinate.CoordType.LongLat 'From Input Longitude, Latitude
                        InputCrs.LongLatEllHtToXYZ()
                        DatumTrans.InputToOutput()

                    Case Coordinate.CoordType.EastNorth 'From Input Easting, Northing
                        If InputCrs.Kind = CoordRefSystem.CrsKind.projected Then
                            InputCrs.EastNorthToLongLat()
                            InputCrs.LongLatEllHtToXYZ()
                            DatumTrans.InputToOutput()
                        End If

                End Select

            Case Coordinate.UpdateMode.TransLongLat 'Update Output Longitude, Latitude
                Select Case From
                    Case Coordinate.CoordType.XYZ 'From Input X, Y, Z
                        DatumTrans.InputToOutput()
                        OutputCrs.XYZToLongLatEllHt()

                    Case Coordinate.CoordType.EastNorth 'From Input Easting, Northing
                        If InputCrs.Kind = CoordRefSystem.CrsKind.projected Then
                            InputCrs.EastNorthToLongLat()
                            InputCrs.LongLatEllHtToXYZ()
                            DatumTrans.InputToOutput()
                            OutputCrs.XYZToLongLatEllHt()
                        End If
                    Case Coordinate.CoordType.LongLat 'From Input Latitude, Longitude, Ellipsoidal Height
                        InputCrs.LongLatEllHtToXYZ()
                        DatumTrans.InputToOutput()
                        OutputCrs.XYZToLongLatEllHt()

                End Select

            Case Coordinate.UpdateMode.TransEastNorth 'Update Output Easting, Northing
                Select Case From
                    Case Coordinate.CoordType.XYZ 'From Input X, Y, Z
                        DatumTrans.InputToOutput()
                        OutputCrs.XYZToLongLatEllHt()
                        If OutputCrs.Kind = CoordRefSystem.CrsKind.projected Then OutputCrs.LongLatToEastNorth()

                    Case Coordinate.CoordType.LongLat 'From Input Longitude, Latitude
                        InputCrs.LongLatEllHtToXYZ()
                        DatumTrans.InputToOutput()
                        OutputCrs.XYZToLongLatEllHt()
                        If OutputCrs.Kind = CoordRefSystem.CrsKind.projected Then OutputCrs.LongLatToEastNorth()

                    Case Coordinate.CoordType.EastNorth 'From Input Easting Northing
                        If InputCrs.Kind = CoordRefSystem.CrsKind.projected Then
                            InputCrs.EastNorthToLongLat()
                            InputCrs.LongLatEllHtToXYZ()
                            DatumTrans.InputToOutput()
                            OutputCrs.XYZToLongLatEllHt()
                            If OutputCrs.Kind = CoordRefSystem.CrsKind.projected Then OutputCrs.LongLatToEastNorth()
                        End If

                End Select

            Case Coordinate.UpdateMode.TransAll 'Update All Output coordinate types
                Select Case From
                    Case Coordinate.CoordType.XYZ 'From Input X, Y, Z
                        DatumTrans.InputToOutput()
                        OutputCrs.XYZToLongLatEllHt()
                        If OutputCrs.Kind = CoordRefSystem.CrsKind.projected Then OutputCrs.LongLatToEastNorth()

                    Case Coordinate.CoordType.LongLat 'From Input Longitude, Latitude
                        InputCrs.LongLatEllHtToXYZ()
                        DatumTrans.InputToOutput()
                        OutputCrs.XYZToLongLatEllHt()
                        If OutputCrs.Kind = CoordRefSystem.CrsKind.projected Then OutputCrs.LongLatToEastNorth()

                    Case Coordinate.CoordType.EastNorth 'From Input Easting Northing
                        If InputCrs.Kind = CoordRefSystem.CrsKind.projected Then
                            InputCrs.EastNorthToLongLat()
                            InputCrs.LongLatEllHtToXYZ()
                            DatumTrans.InputToOutput()
                            OutputCrs.XYZToLongLatEllHt()
                            If OutputCrs.Kind = CoordRefSystem.CrsKind.projected Then OutputCrs.LongLatToEastNorth()
                        End If

                End Select

        End Select
    End Sub

    Private Sub OutputCrs_Update(Mode As Coordinate.UpdateMode, From As Coordinate.CoordType) Handles OutputCrs.Update
        'Update coordinates from the specified coordinate type.
        'A coordinate in the Output Coordinate Reference System has been updated.
        'Update the coordinates specified by Mode using the coordinate specified by From.
        ' For Example: Mode = EastNorth, From = LongLat    : Update Ouput Easting, Northing From Output Longitude, Latitude
        '              Mode = TransLongLat, From = LongLat : Update Input Longitude, Latitude From Output Longitude, Latitude

        'Enum UpdateMode
        '    None
        '    InputOutputAll
        '    XYZ
        '    LongLat
        '    EastNorth
        '    All
        '    TransXYZ
        '    TransLongLat
        '    TransEastNorth
        '    TransAll
        'End Enum

        Select Case Mode
            Case Coordinate.UpdateMode.None 'No updates required.
                'No coordinates to update.

            Case Coordinate.UpdateMode.InputOutputAll 'Update all Input and Output coordinate types
                Select Case From
                    Case Coordinate.CoordType.XYZ 'Update all Input and Output coordinate types from the OutputCrs X, Y, Z coordinate:
                        OutputCrs.XYZToLongLatEllHt() 'Calculate Longitude, Latitude and Ellipsoidal Height from X, Y, Z.
                        If OutputCrs.Kind = CoordRefSystem.CrsKind.projected Then OutputCrs.LongLatToEastNorth() 'Calculate Easting and Northing from Latitude and Longitude.
                        DatumTrans.OutputToInput() 'Calculate Input X, Y, Z from Output X, Y, Z
                        InputCrs.XYZToLongLatEllHt() 'Calculate Input Longitude, Latitude and Ellipsoidal Height from Input X, Y, Z.
                        If InputCrs.Kind = CoordRefSystem.CrsKind.projected Then InputCrs.LongLatToEastNorth() 'Calculate Input Easting and Northing from Input Latitude and Longitude.

                    Case Coordinate.CoordType.LongLat 'Update all Input and Output coordinate types from the OutputCrs Latitude, Longitude coordinate:
                        OutputCrs.LongLatEllHtToXYZ()
                        If OutputCrs.Kind = CoordRefSystem.CrsKind.projected Then OutputCrs.LongLatToEastNorth()
                        DatumTrans.OutputToInput()
                        InputCrs.XYZToLongLatEllHt()
                        If InputCrs.Kind = CoordRefSystem.CrsKind.projected Then InputCrs.LongLatToEastNorth()

                    Case Coordinate.CoordType.EastNorth 'Update all Input and Output coordinate types from the OutputCrs Easting, Northing coordinate:
                        OutputCrs.EastNorthToLongLat()
                        OutputCrs.LongLatEllHtToXYZ()
                        DatumTrans.OutputToInput()
                        InputCrs.XYZToLongLatEllHt()
                        If InputCrs.Kind = CoordRefSystem.CrsKind.projected Then InputCrs.LongLatToEastNorth()

                End Select

            Case Coordinate.UpdateMode.EastNorth 'Update Output Easting, Northing
                Select Case From
                    Case Coordinate.CoordType.XYZ 'From Output X, Y, Z
                        OutputCrs.XYZToLongLatEllHt()
                        If OutputCrs.Kind = CoordRefSystem.CrsKind.projected Then OutputCrs.LongLatToEastNorth()

                    Case Coordinate.CoordType.LongLat 'From Output Longitude, Latitude
                        If OutputCrs.Kind = CoordRefSystem.CrsKind.projected Then OutputCrs.LongLatToEastNorth()

                End Select

            Case Coordinate.UpdateMode.LongLat  'Update Output Longitude, Latitude
                Select Case From
                    Case Coordinate.CoordType.XYZ 'From Output X, Y, Z
                        OutputCrs.XYZToLongLatEllHt()

                    Case Coordinate.CoordType.EastNorth 'From Output Easting, Northing
                        OutputCrs.EastNorthToLongLat()

                End Select

            Case Coordinate.UpdateMode.XYZ  'Update Output X, Y, Z
                Select Case From
                    Case Coordinate.CoordType.LongLat 'From Output Longitude, Latitude
                        OutputCrs.LongLatEllHtToXYZ()

                    Case Coordinate.CoordType.EastNorth 'From Output Easting, Northing
                        OutputCrs.EastNorthToLongLat()
                        OutputCrs.LongLatEllHtToXYZ()

                End Select

            Case Coordinate.UpdateMode.All  'Update Output Coordinate types
                Select Case From
                    Case Coordinate.CoordType.XYZ
                        OutputCrs.XYZToLongLatEllHt()
                        If OutputCrs.Kind = CoordRefSystem.CrsKind.projected Then OutputCrs.LongLatToEastNorth()

                    Case Coordinate.CoordType.LongLat
                        OutputCrs.LongLatEllHtToXYZ()
                        If OutputCrs.Kind = CoordRefSystem.CrsKind.projected Then OutputCrs.LongLatToEastNorth()

                    Case Coordinate.CoordType.EastNorth
                        OutputCrs.EastNorthToLongLat()
                        OutputCrs.LongLatEllHtToXYZ()

                End Select

            Case Coordinate.UpdateMode.TransXYZ   'Update Input X, Y, Z
                Select Case From
                    Case Coordinate.CoordType.XYZ 'From Output X, Y, Z
                        DatumTrans.OutputToInput()

                    Case Coordinate.CoordType.LongLat 'From Output Longitude, Latitude
                        OutputCrs.LongLatEllHtToXYZ()
                        DatumTrans.OutputToInput()

                    Case Coordinate.CoordType.EastNorth 'From Output Easting, Northing
                        If OutputCrs.Kind = CoordRefSystem.CrsKind.projected Then
                            OutputCrs.EastNorthToLongLat()
                            OutputCrs.LongLatEllHtToXYZ()
                            DatumTrans.OutputToInput()
                        End If

                End Select

            Case Coordinate.UpdateMode.TransLongLat 'Update Input Longitude, Latitude
                Select Case From
                    Case Coordinate.CoordType.XYZ 'From Output X, Y, Z
                        DatumTrans.OutputToInput()
                        InputCrs.XYZToLongLatEllHt()

                    Case Coordinate.CoordType.EastNorth 'From Output Easting, Northing
                        If OutputCrs.Kind = CoordRefSystem.CrsKind.projected Then
                            OutputCrs.EastNorthToLongLat()
                            OutputCrs.LongLatEllHtToXYZ()
                            DatumTrans.OutputToInput()
                            InputCrs.XYZToLongLatEllHt()
                        End If

                End Select

            Case Coordinate.UpdateMode.TransEastNorth 'Update Input Easting, Northing
                Select Case From
                    Case Coordinate.CoordType.XYZ 'From Output X, Y, Z
                        DatumTrans.OutputToInput()
                        InputCrs.XYZToLongLatEllHt()
                        If InputCrs.Kind = CoordRefSystem.CrsKind.projected Then InputCrs.LongLatToEastNorth()

                    Case Coordinate.CoordType.LongLat 'From Output Longitude, Latitude
                        OutputCrs.LongLatEllHtToXYZ()
                        DatumTrans.OutputToInput()
                        InputCrs.XYZToLongLatEllHt()
                        If InputCrs.Kind = CoordRefSystem.CrsKind.projected Then InputCrs.LongLatToEastNorth()

                    Case Coordinate.CoordType.EastNorth 'From Output Easting Northing
                        If OutputCrs.Kind = CoordRefSystem.CrsKind.projected Then
                            OutputCrs.EastNorthToLongLat()
                            OutputCrs.LongLatEllHtToXYZ()
                            DatumTrans.OutputToInput()
                            InputCrs.XYZToLongLatEllHt()
                            If InputCrs.Kind = CoordRefSystem.CrsKind.projected Then InputCrs.LongLatToEastNorth()
                        End If

                End Select

            Case Coordinate.UpdateMode.TransAll 'Update All Input coordinate types
                Select Case From
                    Case Coordinate.CoordType.XYZ 'From Output X, Y, Z
                        DatumTrans.OutputToInput()
                        InputCrs.XYZToLongLatEllHt()
                        If InputCrs.Kind = CoordRefSystem.CrsKind.projected Then InputCrs.LongLatToEastNorth()

                    Case Coordinate.CoordType.LongLat 'From Output Longitude, Latitude
                        OutputCrs.LongLatEllHtToXYZ()
                        DatumTrans.OutputToInput()
                        InputCrs.XYZToLongLatEllHt()
                        If InputCrs.Kind = CoordRefSystem.CrsKind.projected Then InputCrs.LongLatToEastNorth()

                    Case Coordinate.CoordType.EastNorth 'From Output Easting Northing
                        If OutputCrs.Kind = CoordRefSystem.CrsKind.projected Then
                            OutputCrs.EastNorthToLongLat()
                            OutputCrs.LongLatEllHtToXYZ()
                            DatumTrans.OutputToInput()
                            InputCrs.XYZToLongLatEllHt()
                            If InputCrs.Kind = CoordRefSystem.CrsKind.projected Then InputCrs.LongLatToEastNorth()
                        End If

                End Select

        End Select
    End Sub


#End Region 'Methods --------------------------------------------------------------------------------------------------------------------------------------------------------------------------



#Region " Events - Events raised by this class." '=============================================================================================================================================
    Event ErrorMessage(ByVal Msg As String) 'Send an error message.
    Event Message(ByVal Msg As String) 'Send a normal message.
#End Region 'Events ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------





End Class 'clsConversion


'==========================================================================================================================================================================================
'Classes used to store table data:
'CoordinateReferenceSystem          CoordRefSystem
'Extent                             Extent
'Usage                              Usage
'Scope                              Scope
'DefiningOperation                  DefiningOperation
'CoordinateOperation                CoordinateOperation
'CoordinateOperationPath            CoordOperationPath
'CoordinateOperationParameterValue  CoordOpParamValue
'CoordinateOperationParameterUsage  CoordOpParamUsage
'CoordinateOperationParameter       CoordOpParameter
'UnitOfMeasure                      UnitOfMeasure
'CoordinateOperationMethod          CoordOpMethod
'CoordinateSystem                   CoordSystem
'CoordinateAxis                     CoordAxis
'CoordinateAxisName                 CoordAxisName
'Datum                              Datum
'DatumEnsembleMember                DatumEnsembleMember
'DatumEnsemble                      DatumEnsemble
'Ellipsoid                          Ellipsoid
'PrimeMeridian                      PrimeMeridian
'ConventionalRS                     ConventionalRS
'Change                             Change
'Deprecation                        Deprecation
'NamingSystem                       NamingSystem                 
'Alias                              AliasName


Public Class CoordRefSystem
    'Properties and methods of a Coordinate Reference System.

#Region " Variable Declarations - All the variables And class objects used in this class." '===============================================================================================

    Public WithEvents Coord As New Coordinate 'Stores the geographic coordinates: Latitude, Longitude, EllipsoidalHeight, the cartesian coordinates: X, Y, Z, and the projected coordinates Easting, Northing of a location. Also displays Geographic coordinates using different formats.

    Public Extent As New Extent 'The CRS Extent
    Public Scope As New Scope 'The CRS Scope

    Public DefiningCoordOpList As New List(Of CoordinateOperation) 'A CRS can have multiple defining Coordinate Operation.
    Public DefiningCoordOp As New CoordinateOperation 'The selected DefiningCoordinate Operation
    Public DefininfCoordOpPathList As New List(Of CoordOperationPath) 'The Defining Coordinate Operation Path - if it exists

    Public SourceCoordOpList As New List(Of CoordinateOperation) 'Coordinate Operations that use the CRS as the Source.
    Public TargetCoordOpList As New List(Of CoordinateOperation) 'Coordinate Operations that use the CRS as the Target.

    Public CoordSystem As New CoordSystem 'CRS Coordinate System
    Public CoordAxisList As New List(Of CoordAxis) 'The set of axes used by the CRS
    Public CoordAxisNameList As New List(Of CoordAxisName) 'The corresponding set of axis names.

    Public Datum As New Datum 'The CRS datum.
    Public DatumEnsembleList As New List(Of DatumEnsemble)
    Public DatumEnsembleMemberList As New List(Of DatumEnsembleMember)
    Public Ellipsoid As New Ellipsoid 'The CRS ellipsoid.
    Public PrimeMeridian As New PrimeMeridian 'The CRS prime meridian.
    Public ConventionalRS As New ConventionalRS 'Conventional Reference System

    Public WithEvents BaseCrs As CoordRefSystem 'If required, this will store information about the base coordinate reference system.

    Public ProjectionCoordOp As New CoordinateOperation 'The coordinate operation used to convert between the Derived CRS and the Base CRS
    Public ProjectionCoordOpMethod As New CoordOpMethod 'The coordinate operation method used to convert between the Derived CRS and the Base CRS
    Public ProjectionCoordOpParamUseList As New List(Of CoordOpParamUsage)
    Public ProjectionCoordOpParamList As New List(Of CoordOpParameter)
    Public ProjectionCoordOpParamValList As New List(Of CoordOpParamValue)

    Public WithEvents Projection As New Projection 'Used to convert between the projected Northing and Easting and the geodetic Latitude and Longitude.

    'Variables calculated from the ellipsoid - used to calculate cartesian coordinates from Latitude, Longitude, Ellipsoidal Height:
    Dim E2 As Double
    Dim N As Double


#End Region 'Variable Declarations --------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Properties - All the properties used in this class." '===========================================================================================================================

    Private _epsgDatabasePath = "" 'The path of the EPSG Database. This database contains a comprehensive set of coordinate reference system parameters.
    Property EpsgDatabasePath
        Get
            Return _epsgDatabasePath
        End Get
        Set(value)
            _epsgDatabasePath = value
        End Set
    End Property

    Private _name As String = "" 'The name of the Coordinate Reference System.
    Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Private _code As Integer = -1 'Unique code (integer) of the Coordinate Reference System (CRS).
    Property Code As Integer
        Get
            Return _code
        End Get
        Set(value As Integer)
            _code = value
            GetCrsInfo(_code)
            RaiseEvent Updated()
        End Set
    End Property

    Enum CrsKind
        compound
        engineering
        geocentric
        geodetic
        geographic2D
        geographic3D
        projected
        vertical
        derived
    End Enum

    Private _kind As CrsKind = CrsKind.projected 'The type of CRS: "compound";"engineering";"geocentric";"geodetic";"geographic 2D";"geographic 3D";"projected";"vertical";"derived".
    Property Kind As CrsKind
        Get
            Return _kind
        End Get
        Set(value As CrsKind)
            _kind = value
        End Set
    End Property

    Private _coordSysCode As Integer = -1 'The code of the Coordinate System (=set of re-usable axes) this CRS uses.
    Property CoordSysCode As Integer
        Get
            Return _coordSysCode
        End Get
        Set(value As Integer)
            _coordSysCode = value
        End Set
    End Property

    Private _datumCode As Integer = -1 'The code of the datum on which this CRS is based. Not used for projected or compound CRSs, or CRSs having a datum ensemble.
    Property DatumCode As Integer
        Get
            Return _datumCode
        End Get
        Set(value As Integer)
            _datumCode = value
        End Set
    End Property

    Private _baseCrsCode As Integer = -1 'For derived CRSs only, the code for the associated base CRS.
    Property BaseCrsCode As Integer
        Get
            Return _baseCrsCode
        End Get
        Set(value As Integer)
            _baseCrsCode = value
        End Set
    End Property

    Private _projConvCode As Integer = -1 'For derived CRSs only, the code of the conversion that is used to convert the Derived CRS from and to the Base CRS. Map projections are derived CRSs.
    Property ProjConvCode As Integer
        Get
            Return _projConvCode
        End Get
        Set(value As Integer)
            _projConvCode = value
        End Set
    End Property

    Private _CmpdHorizCrsCode As Integer = -1 'For compound CRSs only, the code of the horizontal component of the Compound CRS.
    Property CmpdHorizCrsCode As Integer
        Get
            Return _CmpdHorizCrsCode
        End Get
        Set(value As Integer)
            _CmpdHorizCrsCode = value
        End Set
    End Property

    Private _cmpdVertCrsCode As Integer = -1 'For compound CRSs only, the code of the vertical component of the Compound CRS.
    Property CmpdVertCrsCode As Integer
        Get
            Return _cmpdVertCrsCode
        End Get
        Set(value As Integer)
            _cmpdVertCrsCode = value
        End Set
    End Property

    Private _remarks As String = ""
    Property Remarks As String
        Get
            Return _remarks
        End Get
        Set(value As String)
            _remarks = value
        End Set
    End Property

    Private _infoSource As String = "" 'Reference describing the origin of the information to populate this record; e.g. an authoritative publication.
    Property InfoSource As String
        Get
            Return _infoSource
        End Get
        Set(value As String)
            _infoSource = value
        End Set
    End Property

    Private _dataSource As String = "" 'The organisation, body or person who populated this record; for EPSG supplied reference data: "EPSG".
    Property DataSource As String
        Get
            Return _dataSource
        End Get
        Set(value As String)
            _dataSource = value
        End Set
    End Property

    Private _revisionDate As Date = Date.MinValue 'Field used to record the date of creation or modification of this record. Not used if record is deprecated - see deprecation date field.
    'Private _revisionDate As String = "" 'Field used to record the date of creation or modification of this record. Not used if record is deprecated - see deprecation date field.
    Property RevisionDate As Date
        Get
            Return _revisionDate
        End Get
        Set(value As Date)
            _revisionDate = value
        End Set
    End Property

    Private _changeID As String = "" 'Unlinked reference to change table.
    Property ChangeID As String
        Get
            Return _changeID
        End Get
        Set(value As String)
            _changeID = value
        End Set
    End Property

    Private _show As Boolean = True 'Switch to indicate whether operation data can be made public.  "Yes" or "No". Default is Yes.
    Property Show As Boolean
        Get
            Return _show
        End Get
        Set(value As Boolean)
            _show = value
        End Set
    End Property

    Private _deprecated As Boolean = False '"Yes" = data is deprecated; "No" =  data is current and valid.  Default is No.



    Property Deprecated As Boolean
        Get
            Return _deprecated
        End Get
        Set(value As Boolean)
            _deprecated = value
        End Set
    End Property

    'NOTE: Now using the Coord class to store location coordinates.
    ''---------------------------------------------------------------------------------------------------------------------------------------------------------------
    ''Coordinate Properties:
    ''  Used to define a location using
    ''    Cartesian Z, Y, Z coordinates.
    ''    Geodetic (geographic) Latitide, Longitude, Ellipsoidal Height coordinates.
    ''    Projected Easting, Northing coordinates.

    'Private _x As Double 'The Cartesian X coordinate.
    'Property X As Double
    '    Get
    '        Return _x
    '    End Get
    '    Set(value As Double)
    '        _x = value
    '    End Set
    'End Property

    'Private _y As Double 'The Cartesian Y coordinate.
    'Property Y As Double
    '    Get
    '        Return _y
    '    End Get
    '    Set(value As Double)
    '        _y = value
    '    End Set
    'End Property

    'Private _z As Double 'The Cartesian Z coordinate
    'Property Z As Double
    '    Get
    '        Return _z
    '    End Get
    '    Set(value As Double)
    '        _z = value
    '    End Set
    'End Property

    'Private _longitude As Double 'The Longitude of a location.
    'Property Longitude As Double
    '    Get
    '        Return _longitude
    '    End Get
    '    Set(value As Double)
    '        _longitude = value
    '    End Set
    'End Property

    'Private _latitude As Double 'The Latitude of a location.
    'Property Latitude As Double
    '    Get
    '        Return _latitude
    '    End Get
    '    Set(value As Double)
    '        _latitude = value
    '    End Set
    'End Property

    'Private _ellipsoidalHeight As Double = 0 'The Ellipsoidal Height of a location.
    'Property EllipsoidalHeight As Double
    '    Get
    '        Return _ellipsoidalHeight
    '    End Get
    '    Set(value As Double)
    '        _ellipsoidalHeight = value
    '    End Set
    'End Property

    'Private _easting As Double 'The projected Easting of a location.
    'Property Easting As Double
    '    Get
    '        Return _easting
    '    End Get
    '    Set(value As Double)
    '        _easting = value
    '    End Set
    'End Property

    'Private _northing As Double 'The projected Northing of a location.
    'Property Northing As Double
    '    Get
    '        Return _northing
    '    End Get
    '    Set(value As Double)
    '        _northing = value
    '    End Set
    'End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Public Sub New()
        Projection.Coord = Coord 'Projection can now reference the location coordinates directly.
    End Sub

    Public Sub Clear()
        'Clear the properties.
        Name = ""
        'Code = -1 'Don't set Code to -1 - this will trigger GetCrsInfo(-1)
        _code = -1
        Kind = CrsKind.projected
        CoordSysCode = -1
        DatumCode = -1
        BaseCrsCode = -1
        ProjConvCode = -1
        CmpdHorizCrsCode = -1
        CmpdVertCrsCode = -1
        Remarks = ""
        InfoSource = ""
        DataSource = ""
        RevisionDate = Date.MinValue
        ChangeID = ""
        Show = True
        Deprecated = False

        Extent.Clear()
        Scope.Clear()
        DefiningCoordOpList.Clear()
        DefiningCoordOp.Clear()
        DefininfCoordOpPathList.Clear()
        SourceCoordOpList.Clear()
        TargetCoordOpList.Clear()
        CoordSystem.Clear()
        CoordAxisList.Clear()
        CoordAxisNameList.Clear()
        Datum.Clear()
        DatumEnsembleList.Clear()
        DatumEnsembleMemberList.Clear()
        Ellipsoid.Clear()
        PrimeMeridian.Clear()
        ConventionalRS.Clear()

        BaseCrs = Nothing

        'Classes used to store the coordinate operation used to covert between the Derived CRS and the Base CRS:
        ProjectionCoordOp.Clear()
        ProjectionCoordOpMethod.Clear()
        ProjectionCoordOpParamUseList.Clear()
        ProjectionCoordOpParamList.Clear()
        ProjectionCoordOpParamValList.Clear()

        Projection.Clear

    End Sub


    Private Sub GetCrsInfo(CrsCode As Integer)
        'Get the Input Coordinate Reference System information that corresponds to the CrsCode.
        'CrsCode contains the code for the new CRS.
        'RaiseEvent Message("GetCrsInfo() For CrsCode = " & CrsCode & vbCrLf)

        If EpsgDatabasePath = "" Then
            RaiseEvent ErrorMessage("No EPSG database has been selected." & vbCrLf)
            Exit Sub
        End If

        If Not System.IO.File.Exists(EpsgDatabasePath) Then
            RaiseEvent ErrorMessage("Selected EPSG database can Not be found." & vbCrLf)
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + EpsgDatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter("Select * From [Coordinate Reference System] Where COORD_REF_SYS_CODE = " & CrsCode.ToString, conn)
        Dim ds As DataSet = New DataSet
        da.Fill(ds, "CRS")
        conn.Close()

        If ds.Tables("CRS").Rows.Count = 0 Then
            RaiseEvent ErrorMessage("There Is no Coordinate Reference System With the code: " & CrsCode & vbCrLf)
        ElseIf ds.Tables("CRS").Rows.Count = 1 Then
            Clear()
            _code = CrsCode 'Do not use: Code = CrsCode 'This restarts GetCrsInfo(CrsCode)

            If IsDBNull(ds.Tables("CRS").Rows(0).Item("COORD_REF_SYS_NAME")) Then Name = "" Else Name = ds.Tables("CRS").Rows(0).Item("COORD_REF_SYS_NAME")
            'AREA_OF_USE_CODE has been deprecated
            Select Case ds.Tables("CRS").Rows(0).Item("COORD_REF_SYS_KIND")
                Case "compound"
                    Kind = CoordRefSystem.CrsKind.compound
                Case "engineering"
                    Kind = CoordRefSystem.CrsKind.engineering
                Case "geocentric"
                    Kind = CoordRefSystem.CrsKind.geocentric
                Case "geodetic"
                    Kind = CoordRefSystem.CrsKind.geodetic
                Case "geographic 2D"
                    Kind = CoordRefSystem.CrsKind.geographic2D
                Case "geographic 3D"
                    Kind = CoordRefSystem.CrsKind.geographic3D
                Case "projected"
                    Kind = CoordRefSystem.CrsKind.projected
                Case "vertical"
                    Kind = CoordRefSystem.CrsKind.vertical
                Case "derived"
                    Kind = CoordRefSystem.CrsKind.derived
                Case Else
                    RaiseEvent ErrorMessage("Unknown Coordinate Reference System Kind: " & ds.Tables("CRS").Rows(0).Item("COORD_REF_SYS_KIND") & vbCrLf)
            End Select

            If IsDBNull(ds.Tables("CRS").Rows(0).Item("COORD_SYS_CODE")) Then CoordSysCode = -1 Else CoordSysCode = ds.Tables("CRS").Rows(0).Item("COORD_SYS_CODE")
            If IsDBNull(ds.Tables("CRS").Rows(0).Item("DATUM_CODE")) Then DatumCode = -1 Else DatumCode = ds.Tables("CRS").Rows(0).Item("DATUM_CODE")
            If IsDBNull(ds.Tables("CRS").Rows(0).Item("BASE_CRS_CODE")) Then BaseCrsCode = -1 Else BaseCrsCode = ds.Tables("CRS").Rows(0).Item("BASE_CRS_CODE")
            If IsDBNull(ds.Tables("CRS").Rows(0).Item("PROJECTION_CONV_CODE")) Then ProjConvCode = -1 Else ProjConvCode = ds.Tables("CRS").Rows(0).Item("PROJECTION_CONV_CODE")
            If IsDBNull(ds.Tables("CRS").Rows(0).Item("CMPD_HORIZCRS_CODE")) Then CmpdHorizCrsCode = -1 Else CmpdHorizCrsCode = ds.Tables("CRS").Rows(0).Item("CMPD_HORIZCRS_CODE")
            If IsDBNull(ds.Tables("CRS").Rows(0).Item("CMPD_VERTCRS_CODE")) Then CmpdVertCrsCode = -1 Else CmpdVertCrsCode = ds.Tables("CRS").Rows(0).Item("CMPD_VERTCRS_CODE")
            'CRS_SCOPE has been deprecated
            If IsDBNull(ds.Tables("CRS").Rows(0).Item("REMARKS")) Then Remarks = "" Else Remarks = ds.Tables("CRS").Rows(0).Item("REMARKS")
            If IsDBNull(ds.Tables("CRS").Rows(0).Item("INFORMATION_SOURCE")) Then InfoSource = "" Else InfoSource = ds.Tables("CRS").Rows(0).Item("INFORMATION_SOURCE")
            If IsDBNull(ds.Tables("CRS").Rows(0).Item("DATA_SOURCE")) Then DataSource = "" Else DataSource = ds.Tables("CRS").Rows(0).Item("DATA_SOURCE")
            If IsDBNull(ds.Tables("CRS").Rows(0).Item("REVISION_DATE")) Then RevisionDate = Date.MinValue Else RevisionDate = ds.Tables("CRS").Rows(0).Item("REVISION_DATE")
            If IsDBNull(ds.Tables("CRS").Rows(0).Item("CHANGE_ID")) Then ChangeID = "" Else ChangeID = ds.Tables("CRS").Rows(0).Item("CHANGE_ID")
            If IsDBNull(ds.Tables("CRS").Rows(0).Item("SHOW_CRS")) Then Show = True Else Show = ds.Tables("CRS").Rows(0).Item("SHOW_CRS")
            If IsDBNull(ds.Tables("CRS").Rows(0).Item("DEPRECATED")) Then Deprecated = False Else Deprecated = ds.Tables("CRS").Rows(0).Item("DEPRECATED")

            'RaiseEvent Message(vbCrLf & "------------------------------------------------------------------------------------------" & vbCrLf)
            'RaiseEvent Message("Coordinate reference system name: " & Name & "  CRS kind: " & Kind.ToString & vbCrLf)

            'If BaseCrsCode > -1 Then
            '    RaiseEvent Message("CRS Code " & CrsCode & " has BaseCrsCode = " & BaseCrsCode & vbCrLf)
            '    BaseCrs = New CoordRefSystem
            '    BaseCrs.EpsgDatabasePath = EpsgDatabasePath
            '    RaiseEvent Message(vbCrLf & "Getting Base CRS infor for BaseCrsCode = " & BaseCrsCode & "-----------------------" & vbCrLf)
            '    BaseCrs.GetCrsInfo(BaseCrsCode)
            'End If

            GetCrsExtent(CrsCode) 'Gets the Extent and Scope of the Input CRS.
            GetDefiningCoordOps(CrsCode) 'Get the Defining Coordinate Operation(s) corresponding to the Input CRS (if exists).

            'NOTE: The Coord Op lists will be found after the BaseCRS(s) are instantiated. The lists are not generated correctly when the method is run here.
            'GetSourceTargetCoordOps(CrsCode) 'Get any Coord Ops using the Input CRS as the Source or the Target.

            'RaiseEvent Message("CrsCode " & CrsCode & " has DatumCode " & DatumCode & vbCrLf)
            GetDatum(DatumCode)
            GetCoordSystem(CoordSysCode) 'Get the coordinate system information.


            If BaseCrsCode > -1 Then
                'RaiseEvent Message("CRS Code " & CrsCode & " has BaseCrsCode = " & BaseCrsCode & vbCrLf)
                BaseCrs = New CoordRefSystem
                'Application.DoEvents()
                BaseCrs.EpsgDatabasePath = EpsgDatabasePath
                'RaiseEvent Message(vbCrLf & "Getting Base CRS infor for BaseCrsCode = " & BaseCrsCode & "-----------------------" & vbCrLf)
                BaseCrs.GetCrsInfo(BaseCrsCode)
                'Application.DoEvents()
            End If


            If ProjConvCode = -1 Then
                'There is no projection conversion used to convert between the Derived CRS and the Base CRS.
            Else
                GetProjConvInfo(ProjConvCode) 'Get the information about the Projection Conversion used to convert between the Derived CRS and the Base CRS.
            End If
        Else
            RaiseEvent ErrorMessage("There are " & ds.Tables("CRS").Rows.Count & " Coordinate Reference Systems with the code: " & CrsCode & vbCrLf)
        End If
    End Sub

    Public Sub GetAllSourceTargetCoordOps()
        'Gets the Source and Target Coord Ops and those of the BaseCRS(s)
        GetSourceTargetCoordOps(Code) 'Get any Coord Ops using the Input CRS as the Source or the Target.
        If IsNothing(BaseCrs) Then
            'This CRS does not have a BaseCRS
        Else
            BaseCrs.GetAllSourceTargetCoordOps()
        End If
    End Sub

    Private Sub GetCrsExtent(CrsCode As Integer)
        'Get the Extent and Scope of the CRS

        If EpsgDatabasePath = "" Then
            RaiseEvent ErrorMessage("No EPSG database has been selected." & vbCrLf)
            Exit Sub
        End If

        If Not System.IO.File.Exists(EpsgDatabasePath) Then
            RaiseEvent ErrorMessage("Selected EPSG database can not be found." & vbCrLf)
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + EpsgDatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter("Select * From [Usage] Where  OBJECT_TABLE_NAME = 'Coordinate Reference System' And OBJECT_CODE = " & CrsCode.ToString, conn)
        Dim ds As DataSet = New DataSet
        da.Fill(ds, "Usage")

        If ds.Tables("Usage").Rows.Count = 0 Then
            RaiseEvent ErrorMessage("There are no Usage records for CRS code number: " & CrsCode & vbCrLf)
        ElseIf ds.Tables("Usage").Rows.Count = 1 Then
            Dim ExtentCode As Integer = ds.Tables("Usage").Rows(0).Item("EXTENT_CODE")
            Dim ScopeCode As Integer = ds.Tables("Usage").Rows(0).Item("SCOPE_CODE")

            'Get the Extent record:
            'Extent.Clear()
            da.SelectCommand.CommandText = "Select * From Extent Where EXTENT_CODE = " & ExtentCode
            da.Fill(ds, "Extent")
            'RaiseEvent Message("Getting the extent of CrsCode " & CrsCode & vbCrLf)
            If ds.Tables("Extent").Rows.Count = 0 Then
                RaiseEvent ErrorMessage("There are no Extent records for Extent code number: " & ExtentCode & vbCrLf)
            ElseIf ds.Tables("Extent").Rows.Count = 1 Then
                Extent.Code = ExtentCode
                If IsDBNull(ds.Tables("Extent").Rows(0).Item("EXTENT_NAME")) Then Extent.Name = "" Else Extent.Name = ds.Tables("Extent").Rows(0).Item("EXTENT_NAME")
                If IsDBNull(ds.Tables("Extent").Rows(0).Item("EXTENT_DESCRIPTION")) Then Extent.Description = "" Else Extent.Description = ds.Tables("Extent").Rows(0).Item("EXTENT_DESCRIPTION")
                If IsDBNull(ds.Tables("Extent").Rows(0).Item("BBOX_SOUTH_BOUND_LAT")) Then Extent.SouthBoundLat = Double.NaN Else Extent.SouthBoundLat = ds.Tables("Extent").Rows(0).Item("BBOX_SOUTH_BOUND_LAT")
                If IsDBNull(ds.Tables("Extent").Rows(0).Item("BBOX_WEST_BOUND_LON")) Then Extent.WestBoundLon = Double.NaN Else Extent.WestBoundLon = ds.Tables("Extent").Rows(0).Item("BBOX_WEST_BOUND_LON")
                If IsDBNull(ds.Tables("Extent").Rows(0).Item("BBOX_NORTH_BOUND_LAT")) Then Extent.NorthBoundLat = Double.NaN Else Extent.NorthBoundLat = ds.Tables("Extent").Rows(0).Item("BBOX_NORTH_BOUND_LAT")
                If IsDBNull(ds.Tables("Extent").Rows(0).Item("BBOX_EAST_BOUND_LON")) Then Extent.EastBoundLon = Double.NaN Else Extent.EastBoundLon = ds.Tables("Extent").Rows(0).Item("BBOX_EAST_BOUND_LON")
                If IsDBNull(ds.Tables("Extent").Rows(0).Item("ISO_A2_CODE")) Then Extent.IsoA2Code = "" Else Extent.IsoA2Code = ds.Tables("Extent").Rows(0).Item("ISO_A2_CODE")
                If IsDBNull(ds.Tables("Extent").Rows(0).Item("ISO_A3_CODE")) Then Extent.IsoA3Code = "" Else Extent.IsoA3Code = ds.Tables("Extent").Rows(0).Item("ISO_A3_CODE")
                If IsDBNull(ds.Tables("Extent").Rows(0).Item("ISO_N_CODE")) Then Extent.IsoNCode = -1 Else Extent.IsoNCode = ds.Tables("Extent").Rows(0).Item("ISO_N_CODE")
                If IsDBNull(ds.Tables("Extent").Rows(0).Item("VERTICAL_EXTENT_MIN")) Then Extent.VertExtentMin = Double.NaN Else Extent.VertExtentMin = ds.Tables("Extent").Rows(0).Item("VERTICAL_EXTENT_MIN")
                If IsDBNull(ds.Tables("Extent").Rows(0).Item("VERTICAL_EXTENT_MAX")) Then Extent.VertExtentMax = Double.NaN Else Extent.VertExtentMax = ds.Tables("Extent").Rows(0).Item("VERTICAL_EXTENT_MAX")
                If IsDBNull(ds.Tables("Extent").Rows(0).Item("VERTICAL_EXTENT_CRS_CODE")) Then Extent.VertExtentCrsCode = -1 Else Extent.VertExtentCrsCode = ds.Tables("Extent").Rows(0).Item("VERTICAL_EXTENT_CRS_CODE")
                If IsDBNull(ds.Tables("Extent").Rows(0).Item("TEMPORAL_EXTENT_BEGIN")) Then Extent.TemporalExtentBegin = "" Else Extent.TemporalExtentBegin = ds.Tables("Extent").Rows(0).Item("TEMPORAL_EXTENT_BEGIN")
                If IsDBNull(ds.Tables("Extent").Rows(0).Item("TEMPORAL_EXTENT_END")) Then Extent.TemporalExtentEnd = "" Else Extent.TemporalExtentEnd = ds.Tables("Extent").Rows(0).Item("TEMPORAL_EXTENT_END")
                If IsDBNull(ds.Tables("Extent").Rows(0).Item("REMARKS")) Then Extent.Remarks = "" Else Extent.Remarks = ds.Tables("Extent").Rows(0).Item("REMARKS")
                If IsDBNull(ds.Tables("Extent").Rows(0).Item("INFORMATION_SOURCE")) Then Extent.InfoSource = "" Else Extent.InfoSource = ds.Tables("Extent").Rows(0).Item("INFORMATION_SOURCE")
                If IsDBNull(ds.Tables("Extent").Rows(0).Item("DATA_SOURCE")) Then Extent.DataSource = "" Else Extent.DataSource = ds.Tables("Extent").Rows(0).Item("DATA_SOURCE")
                If IsDBNull(ds.Tables("Extent").Rows(0).Item("REVISION_DATE")) Then Extent.RevisionDate = Date.MinValue Else Extent.RevisionDate = ds.Tables("Extent").Rows(0).Item("REVISION_DATE")
                If IsDBNull(ds.Tables("Extent").Rows(0).Item("CHANGE_ID")) Then Extent.ChangeID = "" Else Extent.ChangeID = ds.Tables("Extent").Rows(0).Item("CHANGE_ID")
                If IsDBNull(ds.Tables("Extent").Rows(0).Item("DEPRECATED")) Then Extent.Deprecated = False Else Extent.Deprecated = ds.Tables("Extent").Rows(0).Item("DEPRECATED")
                'RaiseEvent Message("Extent name: " & Extent.Name & vbCrLf)
            Else
                RaiseEvent ErrorMessage("There are " & ds.Tables("Extent").Rows.Count & " Extent records for Extent code number: " & ExtentCode & vbCrLf)
            End If

            'Get the Scope record:
            'Scope.Clear()
            da.SelectCommand.CommandText = "Select * From Scope Where SCOPE_CODE = " & ScopeCode
            da.Fill(ds, "Scope")
            If ds.Tables("Scope").Rows.Count = 0 Then
                RaiseEvent ErrorMessage("There are no Scope records for CRS code number: " & CrsCode & vbCrLf)
            ElseIf ds.Tables("Scope").Rows.Count = 1 Then
                Scope.Code = ScopeCode
                If IsDBNull(ds.Tables("Scope").Rows(0).Item("SCOPE")) Then Scope.Scope = "" Else Scope.Scope = ds.Tables("Scope").Rows(0).Item("SCOPE")
                If IsDBNull(ds.Tables("Scope").Rows(0).Item("REMARKS")) Then Scope.Remarks = "" Else Scope.Remarks = ds.Tables("Scope").Rows(0).Item("REMARKS")
                If IsDBNull(ds.Tables("Scope").Rows(0).Item("INFORMATION_SOURCE")) Then Scope.InfoSource = "" Else Scope.InfoSource = ds.Tables("Scope").Rows(0).Item("INFORMATION_SOURCE")
                If IsDBNull(ds.Tables("Scope").Rows(0).Item("DATA_SOURCE")) Then Scope.DataSource = "" Else Scope.DataSource = ds.Tables("Scope").Rows(0).Item("DATA_SOURCE")
                If IsDBNull(ds.Tables("Scope").Rows(0).Item("REVISION_DATE")) Then Scope.RevisionDate = Date.MinValue Else Scope.RevisionDate = ds.Tables("Scope").Rows(0).Item("REVISION_DATE")
                If IsDBNull(ds.Tables("Scope").Rows(0).Item("CHANGE_ID")) Then Scope.ChangeID = "" Else Scope.ChangeID = ds.Tables("Scope").Rows(0).Item("CHANGE_ID")
                If IsDBNull(ds.Tables("Scope").Rows(0).Item("DEPRECATED")) Then Scope.Deprecated = "" Else Scope.Deprecated = ds.Tables("Scope").Rows(0).Item("DEPRECATED")
                'RaiseEvent Message("Scope: " & Scope.Scope & vbCrLf)
            Else
                RaiseEvent ErrorMessage("There are " & ds.Tables("Scope").Rows.Count & " Scope records for Scope code number: " & ScopeCode & vbCrLf)
            End If
        Else
            'NOTE: There should not be multiple Usage records for a Coordinate Reference System.
            RaiseEvent ErrorMessage("There are " & ds.Tables("Usage").Rows.Count & " Usage records for CRS code number: " & CrsCode & vbCrLf)
        End If
        conn.Close()
    End Sub

    Private Sub GetDefiningCoordOps(CrsCode As Integer)
        'Get the Coordinate Operation(s) corresponding to the CRS

        If EpsgDatabasePath = "" Then
            'Main.Message.AddWarning("No EPSG database has been selected." & vbCrLf)
            RaiseEvent ErrorMessage("No EPSG database has been selected." & vbCrLf)
            Exit Sub
        End If

        If Not System.IO.File.Exists(EpsgDatabasePath) Then
            'Main.Message.AddWarning("Selected EPSG database can not be found." & vbCrLf)
            RaiseEvent ErrorMessage("Selected EPSG database can not be found." & vbCrLf)
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        'Specify the connection string:
        'Access 2003
        'connectionString = "provider=Microsoft.Jet.OLEDB.4.0;" + _
        '"data source = " + txtDatabase.Text

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + EpsgDatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter("Select * From DefiningOperation Where  CRS_CODE = " & CrsCode.ToString, conn)
        Dim ds As DataSet = New DataSet
        da.Fill(ds, "DefOp")

        'InputDefiningCoordOpList.Clear()
        'DefiningCoordOpList.Clear()

        If ds.Tables("DefOp").Rows.Count = 0 Then
            'RaiseEvent ErrorMessage("There are no Defining Coordinate Operation records for CRS code number: " & CrsCode & vbCrLf)
        Else
            For Each Item As DataRow In ds.Tables("DefOp").Rows
                Dim CtCode As Integer = Item("CT_CODE")
                da.SelectCommand.CommandText = "Select * From Coordinate_Operation Where COORD_OP_CODE = " & CtCode
                da.Fill(ds, "CoordOp")
                If ds.Tables("CoordOp").Rows.Count = 0 Then
                    RaiseEvent ErrorMessage("There are no Coordinate Operation records for Coordinate Transformation code number: " & CtCode & vbCrLf)
                Else
                    Dim NewCoordOp As New CoordinateOperation
                    If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("COORD_OP_CODE")) Then NewCoordOp.Code = -1 Else NewCoordOp.Code = ds.Tables("CoordOp").Rows(0).Item("COORD_OP_CODE")
                    If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("COORD_OP_NAME")) Then NewCoordOp.Name = "" Else NewCoordOp.Name = ds.Tables("CoordOp").Rows(0).Item("COORD_OP_NAME")

                    Select Case ds.Tables("CoordOp").Rows(0).Item("COORD_OP_TYPE")
                        Case "conversion"
                            NewCoordOp.Type = CoordinateOperation.OperationType.conversion
                        Case "transformation"
                            NewCoordOp.Type = CoordinateOperation.OperationType.transformation
                        Case "point motion operation"
                            NewCoordOp.Type = CoordinateOperation.OperationType.pointMotionOperation
                        Case "concatenated operation"
                            NewCoordOp.Type = CoordinateOperation.OperationType.concatenatedOperation
                        Case Else
                            RaiseEvent ErrorMessage("Unknown coordinate operation type: " & ds.Tables("CoordOp").Rows(0).Item("COORD_OP_TYPE") & vbCrLf)
                            NewCoordOp.Type = CoordinateOperation.OperationType.conversion
                    End Select

                    If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("SOURCE_CRS_CODE")) Then NewCoordOp.SourceCrsCode = -1 Else NewCoordOp.SourceCrsCode = ds.Tables("CoordOp").Rows(0).Item("SOURCE_CRS_CODE")
                    If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("TARGET_CRS_CODE")) Then NewCoordOp.TargetCrsCode = -1 Else NewCoordOp.TargetCrsCode = ds.Tables("CoordOp").Rows(0).Item("TARGET_CRS_CODE")
                    If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("COORD_OP_VARIANT")) Then NewCoordOp.OpVariant = -1 Else NewCoordOp.OpVariant = ds.Tables("CoordOp").Rows(0).Item("COORD_OP_VARIANT")
                    'AREA_OF_USE_CODE has been deprecated.
                    'COORD_OP_SCOPE has been deprecated.
                    If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("COORD_OP_ACCURACY")) Then NewCoordOp.Accuracy = Single.NaN Else NewCoordOp.Accuracy = ds.Tables("CoordOp").Rows(0).Item("COORD_OP_ACCURACY")
                    If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("COORD_OP_METHOD_CODE")) Then NewCoordOp.MethodCode = -1 Else NewCoordOp.MethodCode = ds.Tables("CoordOp").Rows(0).Item("COORD_OP_METHOD_CODE")
                    If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("UOM_CODE_SOURCE_COORD_DIFF")) Then NewCoordOp.UomSourceCoordDiffCode = -1 Else NewCoordOp.UomSourceCoordDiffCode = ds.Tables("CoordOp").Rows(0).Item("UOM_CODE_SOURCE_COORD_DIFF")
                    If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("UOM_CODE_TARGET_COORD_DIFF")) Then NewCoordOp.UomTargetCoordDiffCode = -1 Else NewCoordOp.UomTargetCoordDiffCode = ds.Tables("CoordOp").Rows(0).Item("UOM_CODE_TARGET_COORD_DIFF")
                    If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("REMARKS")) Then NewCoordOp.Remarks = "" Else NewCoordOp.Remarks = ds.Tables("CoordOp").Rows(0).Item("REMARKS")
                    If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("INFORMATION_SOURCE")) Then NewCoordOp.InfoSource = -1 Else NewCoordOp.InfoSource = ds.Tables("CoordOp").Rows(0).Item("INFORMATION_SOURCE")
                    If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("DATA_SOURCE")) Then NewCoordOp.DataSource = -1 Else NewCoordOp.DataSource = ds.Tables("CoordOp").Rows(0).Item("DATA_SOURCE")
                    If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("REVISION_DATE")) Then NewCoordOp.RevisionDate = Date.MinValue Else NewCoordOp.RevisionDate = ds.Tables("CoordOp").Rows(0).Item("REVISION_DATE")
                    If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("CHANGE_ID")) Then NewCoordOp.ChangeID = "" Else NewCoordOp.ChangeID = ds.Tables("CoordOp").Rows(0).Item("CHANGE_ID")
                    If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("SHOW_OPERATION")) Then NewCoordOp.Show = True Else NewCoordOp.Show = ds.Tables("CoordOp").Rows(0).Item("SHOW_OPERATION")
                    If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("DEPRECATED")) Then NewCoordOp.Deprecated = False Else NewCoordOp.Deprecated = ds.Tables("CoordOp").Rows(0).Item("DEPRECATED")
                    'InputDefiningCoordOpList.Add(NewCoordOp)
                    DefiningCoordOpList.Add(NewCoordOp)

                    'Main.Message.Add("Coordinate operation name: " & NewCoordOp.Name & "  Accuracy: " & NewCoordOp.Accuracy & vbCrLf)
                    RaiseEvent Message("Coordinate operation name: " & NewCoordOp.Name & "  Accuracy: " & NewCoordOp.Accuracy & vbCrLf)
                    ds.Tables("CoordOp").Rows.Clear()
                End If
            Next
        End If
        conn.Close()
    End Sub

    Public Sub GetSourceTargetCoordOps(CrsCode As Integer)
        'Get a list of Coordinate Operations that use the Input CRS as the Source.
        'Get a list of Coordinate Operations that use the Input CRS as the Target.

        'NOTE: This method is now public so that it can be run again after the CRS has been set up.
        'Currently the SourceCoordOpList and TargetCoordOpList do not store all the relevant Coord Operations when GetCrsInfo() is run  - reason unknown.
        'The method will be re-run later in an attempt to get the Source Coord Op List and the Target Coord Op List.

        If EpsgDatabasePath = "" Then
            RaiseEvent ErrorMessage("No EPSG database has been selected." & vbCrLf)
            Exit Sub
        End If

        If Not System.IO.File.Exists(EpsgDatabasePath) Then
            RaiseEvent ErrorMessage("Selected EPSG database can not be found." & vbCrLf)
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + EpsgDatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        'Get list of Input Source Coordinate Operations:
        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter("Select * From Coordinate_Operation Where  SOURCE_CRS_CODE = " & CrsCode.ToString, conn)
        Dim ds As DataSet = New DataSet
        da.Fill(ds, "CrsSource")
        SourceCoordOpList.Clear()

        'RaiseEvent Message(vbCrLf & "Getting coord ops that use CrsCode " & CrsCode & " for the source or target. -----------------------------" & vbCrLf)

        If ds.Tables("CrsSource").Rows.Count = 0 Then
            'RaiseEvent ErrorMessage("There are no Coordinate Operation records for Source CRS code number: " & CrsCode & vbCrLf)
            'RaiseEvent Message("There are no coord ops using source CrsCode " & CrsCode & vbCrLf)
        Else
            'RaiseEvent Message("There are " & ds.Tables("CrsSource").Rows.Count & " coord ops using source CrsCode " & CrsCode & vbCrLf)
            For Each Item As DataRow In ds.Tables("CrsSource").Rows
                'Dim NewInputSourceCoordOp As New CoordinateOperation
                Dim NewSourceCoordOp As New CoordinateOperation
                If IsDBNull(Item("COORD_OP_CODE")) Then NewSourceCoordOp.Code = -1 Else NewSourceCoordOp.Code = Item("COORD_OP_CODE")
                If IsDBNull(Item("COORD_OP_NAME")) Then NewSourceCoordOp.Name = "" Else NewSourceCoordOp.Name = Item("COORD_OP_NAME")

                Select Case Item("COORD_OP_TYPE")
                    Case "conversion"
                        NewSourceCoordOp.Type = CoordinateOperation.OperationType.conversion
                    Case "transformation"
                        NewSourceCoordOp.Type = CoordinateOperation.OperationType.transformation
                    Case "point motion operation"
                        NewSourceCoordOp.Type = CoordinateOperation.OperationType.pointMotionOperation
                    Case "concatenated operation"
                        NewSourceCoordOp.Type = CoordinateOperation.OperationType.concatenatedOperation
                    Case Else
                        RaiseEvent ErrorMessage("Unknown Input Source coordinate operation type: " & Item("COORD_OP_TYPE") & vbCrLf)
                        NewSourceCoordOp.Type = CoordinateOperation.OperationType.conversion
                End Select

                NewSourceCoordOp.SourceCrsCode = CrsCode
                If IsDBNull(Item("TARGET_CRS_CODE")) Then NewSourceCoordOp.TargetCrsCode = -1 Else NewSourceCoordOp.TargetCrsCode = Item("TARGET_CRS_CODE")
                If IsDBNull(Item("COORD_TFM_VERSION")) Then NewSourceCoordOp.Version = "" Else NewSourceCoordOp.Version = Item("COORD_TFM_VERSION")
                If IsDBNull(Item("COORD_OP_VARIANT")) Then NewSourceCoordOp.OpVariant = -1 Else NewSourceCoordOp.OpVariant = Item("COORD_OP_VARIANT")
                'AREA_OF_USE_CODE has been deprecated.
                'COORD_OP_SCOPE has been deprecated.
                If IsDBNull(Item("COORD_OP_ACCURACY")) Then NewSourceCoordOp.Accuracy = Single.NaN Else NewSourceCoordOp.Accuracy = Item("COORD_OP_ACCURACY")
                If IsDBNull(Item("COORD_OP_METHOD_CODE")) Then NewSourceCoordOp.MethodCode = -1 Else NewSourceCoordOp.MethodCode = Item("COORD_OP_METHOD_CODE")
                If IsDBNull(Item("UOM_CODE_SOURCE_COORD_DIFF")) Then NewSourceCoordOp.UomSourceCoordDiffCode = -1 Else NewSourceCoordOp.UomSourceCoordDiffCode = Item("UOM_CODE_SOURCE_COORD_DIFF")
                If IsDBNull(Item("UOM_CODE_TARGET_COORD_DIFF")) Then NewSourceCoordOp.UomTargetCoordDiffCode = -1 Else NewSourceCoordOp.UomTargetCoordDiffCode = Item("UOM_CODE_TARGET_COORD_DIFF")
                If IsDBNull(Item("REMARKS")) Then NewSourceCoordOp.Remarks = "" Else NewSourceCoordOp.Remarks = Item("REMARKS")
                If IsDBNull(Item("INFORMATION_SOURCE")) Then NewSourceCoordOp.InfoSource = "" Else NewSourceCoordOp.InfoSource = Item("INFORMATION_SOURCE")
                If IsDBNull(Item("DATA_SOURCE")) Then NewSourceCoordOp.DataSource = "" Else NewSourceCoordOp.DataSource = Item("DATA_SOURCE")
                If IsDBNull(Item("REVISION_DATE")) Then NewSourceCoordOp.RevisionDate = Date.MinValue Else NewSourceCoordOp.RevisionDate = Item("REVISION_DATE")
                If IsDBNull(Item("CHANGE_ID")) Then NewSourceCoordOp.ChangeID = "" Else NewSourceCoordOp.ChangeID = Item("CHANGE_ID")
                If IsDBNull(Item("SHOW_OPERATION")) Then NewSourceCoordOp.Show = True Else NewSourceCoordOp.Show = Item("SHOW_OPERATION")
                If IsDBNull(Item("DEPRECATED")) Then NewSourceCoordOp.Deprecated = False Else NewSourceCoordOp.Deprecated = Item("DEPRECATED")
                'InputSourceCoordOpList.Add(NewInputSourceCoordOp)
                SourceCoordOpList.Add(NewSourceCoordOp)
                'Application.DoEvents()
                'RaiseEvent Message("Source Coordinate operation name: " & NewSourceCoordOp.Name & " with code: " & NewSourceCoordOp.Code & "  Source CRS code: " & NewSourceCoordOp.SourceCrsCode & "  Target CRS code: " & NewSourceCoordOp.TargetCrsCode & vbCrLf)
            Next
            'RaiseEvent Message("CRS Code " & CrsCode & " contains " & SourceCoordOpList.Count & " coordinate operations that use the CRS as the source." & vbCrLf)
        End If

        'Get list of Input Target Coordinate Operations:
        da.SelectCommand.CommandText = "Select * From Coordinate_Operation Where  TARGET_CRS_CODE = " & CrsCode.ToString
        da.Fill(ds, "CrsTarget")
        TargetCoordOpList.Clear()

        If ds.Tables("CrsTarget").Rows.Count = 0 Then
            'RaiseEvent ErrorMessage("There are no Coordinate Operation records for Target CRS code number: " & CrsCode & vbCrLf)
            'RaiseEvent Message("There are no coord ops using target CrsCode " & CrsCode & vbCrLf)
        Else
            'RaiseEvent Message("There are " & ds.Tables("CrsTarget").Rows.Count & " coord ops using target CrsCode " & CrsCode & vbCrLf)
            For Each Item As DataRow In ds.Tables("CrsTarget").Rows
                Dim NewTargetCoordOp As New CoordinateOperation
                If IsDBNull(Item("COORD_OP_CODE")) Then NewTargetCoordOp.Code = -1 Else NewTargetCoordOp.Code = Item("COORD_OP_CODE")
                If IsDBNull(Item("COORD_OP_NAME")) Then NewTargetCoordOp.Name = "" Else NewTargetCoordOp.Name = Item("COORD_OP_NAME")

                Select Case Item("COORD_OP_TYPE")
                    Case "conversion"
                        NewTargetCoordOp.Type = CoordinateOperation.OperationType.conversion
                    Case "transformation"
                        NewTargetCoordOp.Type = CoordinateOperation.OperationType.transformation
                    Case "point motion operation"
                        NewTargetCoordOp.Type = CoordinateOperation.OperationType.pointMotionOperation
                    Case "concatenated operation"
                        NewTargetCoordOp.Type = CoordinateOperation.OperationType.concatenatedOperation
                    Case Else
                        RaiseEvent ErrorMessage("Unknown Target coordinate operation type: " & Item("COORD_OP_TYPE") & vbCrLf)
                        NewTargetCoordOp.Type = CoordinateOperation.OperationType.conversion
                End Select

                'NewTargetCoordOp.SourceCrsCode = CrsCode
                NewTargetCoordOp.TargetCrsCode = CrsCode
                'If IsDBNull(Item("TARGET_CRS_CODE")) Then NewTargetCoordOp.TargetCrsCode = -1 Else NewTargetCoordOp.TargetCrsCode = Item("TARGET_CRS_CODE")
                If IsDBNull(Item("SOURCE_CRS_CODE")) Then NewTargetCoordOp.SourceCrsCode = -1 Else NewTargetCoordOp.SourceCrsCode = Item("SOURCE_CRS_CODE")

                If IsDBNull(Item("COORD_TFM_VERSION")) Then NewTargetCoordOp.Version = "" Else NewTargetCoordOp.Version = Item("COORD_TFM_VERSION")
                If IsDBNull(Item("COORD_OP_VARIANT")) Then NewTargetCoordOp.OpVariant = -1 Else NewTargetCoordOp.OpVariant = Item("COORD_OP_VARIANT")
                'AREA_OF_USE_CODE has been deprecated.
                'COORD_OP_SCOPE has been deprecated.
                If IsDBNull(Item("COORD_OP_ACCURACY")) Then NewTargetCoordOp.Accuracy = Single.NaN Else NewTargetCoordOp.Accuracy = Item("COORD_OP_ACCURACY")
                If IsDBNull(Item("COORD_OP_METHOD_CODE")) Then NewTargetCoordOp.MethodCode = -1 Else NewTargetCoordOp.MethodCode = Item("COORD_OP_METHOD_CODE")
                If IsDBNull(Item("UOM_CODE_SOURCE_COORD_DIFF")) Then NewTargetCoordOp.UomSourceCoordDiffCode = -1 Else NewTargetCoordOp.UomSourceCoordDiffCode = Item("UOM_CODE_SOURCE_COORD_DIFF")
                If IsDBNull(Item("UOM_CODE_TARGET_COORD_DIFF")) Then NewTargetCoordOp.UomTargetCoordDiffCode = -1 Else NewTargetCoordOp.UomTargetCoordDiffCode = Item("UOM_CODE_TARGET_COORD_DIFF")
                If IsDBNull(Item("REMARKS")) Then NewTargetCoordOp.Remarks = "" Else NewTargetCoordOp.Remarks = Item("REMARKS")
                If IsDBNull(Item("INFORMATION_SOURCE")) Then NewTargetCoordOp.InfoSource = "" Else NewTargetCoordOp.InfoSource = Item("INFORMATION_SOURCE")
                If IsDBNull(Item("DATA_SOURCE")) Then NewTargetCoordOp.DataSource = "" Else NewTargetCoordOp.DataSource = Item("DATA_SOURCE")
                If IsDBNull(Item("REVISION_DATE")) Then NewTargetCoordOp.RevisionDate = Date.MinValue Else NewTargetCoordOp.RevisionDate = Item("REVISION_DATE")
                If IsDBNull(Item("CHANGE_ID")) Then NewTargetCoordOp.ChangeID = "" Else NewTargetCoordOp.ChangeID = Item("CHANGE_ID")
                If IsDBNull(Item("SHOW_OPERATION")) Then NewTargetCoordOp.Show = True Else NewTargetCoordOp.Show = Item("SHOW_OPERATION")
                If IsDBNull(Item("DEPRECATED")) Then NewTargetCoordOp.Deprecated = False Else NewTargetCoordOp.Deprecated = Item("DEPRECATED")
                TargetCoordOpList.Add(NewTargetCoordOp)
                'Application.DoEvents()
                'RaiseEvent Message("Target Coordinate operation name: " & NewTargetCoordOp.Name & " with code: " & NewTargetCoordOp.Code & "  Source CRS code: " & NewTargetCoordOp.SourceCrsCode & "  Target CRS code: " & NewTargetCoordOp.TargetCrsCode & vbCrLf)
            Next
            'RaiseEvent Message("CRS Code " & CrsCode & " contains " & TargetCoordOpList.Count & " coordinate operations that use the CRS as the target." & vbCrLf)
        End If
        conn.Close()
    End Sub

    Private Sub GetDatum(DatumCode As Integer)
        'Get the CRS Datum information.

        If EpsgDatabasePath = "" Then
            'Main.Message.AddWarning("No EPSG database has been selected." & vbCrLf)
            RaiseEvent ErrorMessage("No EPSG database has been selected." & vbCrLf)
            Exit Sub
        End If

        If Not System.IO.File.Exists(EpsgDatabasePath) Then
            'Main.Message.AddWarning("Selected EPSG database can not be found." & vbCrLf)
            RaiseEvent ErrorMessage("Selected EPSG database can not be found." & vbCrLf)
            Exit Sub
        End If

        If DatumCode = -1 Then
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + EpsgDatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter("Select * From Datum Where DATUM_CODE = " & DatumCode.ToString, conn)
        Dim ds As DataSet = New DataSet
        da.Fill(ds, "Datum")
        conn.Close()

        If ds.Tables("Datum").Rows.Count = 0 Then
            RaiseEvent ErrorMessage("There is no Datum with the code: " & DatumCode & vbCrLf)
        ElseIf ds.Tables("Datum").Rows.Count = 1 Then
            Datum.Code = DatumCode
            If IsDBNull(ds.Tables("Datum").Rows(0).Item("DATUM_NAME")) Then Datum.Name = "" Else Datum.Name = ds.Tables("Datum").Rows(0).Item("DATUM_NAME")

            Select Case ds.Tables("Datum").Rows(0).Item("DATUM_Type")
                Case "geodetic"
                    Datum.Type = Datum.DatumType.geodetic
                Case "dynamic geodetic"
                    Datum.Type = Datum.DatumType.dynamicGeodetic
                Case "vertical"
                    Datum.Type = Datum.DatumType.vertical
                Case "dynamic vertical"
                    Datum.Type = Datum.DatumType.dynamicVertical
                Case "engineering"
                    Datum.Type = Datum.DatumType.engineering
                Case "ensemble"
                    Datum.Type = Datum.DatumType.ensemble
                Case Else
                    RaiseEvent ErrorMessage("Unknown datum type: " & ds.Tables("Datum").Rows(0).Item("DATUM_Type") & vbCrLf)
            End Select

            If IsDBNull(ds.Tables("Datum").Rows(0).Item("ORIGIN_DESCRIPTION")) Then Datum.OriginDescr = "" Else Datum.OriginDescr = ds.Tables("Datum").Rows(0).Item("ORIGIN_DESCRIPTION")
            'Realization Epoch has been deprecated.
            If IsDBNull(ds.Tables("Datum").Rows(0).Item("ELLIPSOID_CODE")) Then Datum.EllipsoidCode = -1 Else Datum.EllipsoidCode = ds.Tables("Datum").Rows(0).Item("ELLIPSOID_CODE")
            If IsDBNull(ds.Tables("Datum").Rows(0).Item("PRIME_MERIDIAN_CODE")) Then Datum.PrimeMeridianCode = -1 Else Datum.PrimeMeridianCode = ds.Tables("Datum").Rows(0).Item("PRIME_MERIDIAN_CODE")
            'Area of Use code has been deprecated.
            'Datum Scope has been deprecated.
            If IsDBNull(ds.Tables("Datum").Rows(0).Item("REMARKS")) Then Datum.Remarks = "" Else Datum.Remarks = ds.Tables("Datum").Rows(0).Item("REMARKS")
            If IsDBNull(ds.Tables("Datum").Rows(0).Item("INFORMATION_SOURCE")) Then Datum.InfoSource = "" Else Datum.InfoSource = ds.Tables("Datum").Rows(0).Item("INFORMATION_SOURCE")
            If IsDBNull(ds.Tables("Datum").Rows(0).Item("DATA_SOURCE")) Then Datum.DataSource = "" Else Datum.DataSource = ds.Tables("Datum").Rows(0).Item("DATA_SOURCE")
            If IsDBNull(ds.Tables("Datum").Rows(0).Item("REVISION_DATE")) Then Datum.RevisionDate = Date.MinValue Else Datum.RevisionDate = ds.Tables("Datum").Rows(0).Item("REVISION_DATE")
            If IsDBNull(ds.Tables("Datum").Rows(0).Item("CHANGE_ID")) Then Datum.ChangeID = "" Else Datum.ChangeID = ds.Tables("Datum").Rows(0).Item("CHANGE_ID")
            If IsDBNull(ds.Tables("Datum").Rows(0).Item("DEPRECATED")) Then Datum.Deprecated = False Else Datum.Deprecated = ds.Tables("Datum").Rows(0).Item("DEPRECATED")
            If IsDBNull(ds.Tables("Datum").Rows(0).Item("CONVENTIONAL_RS_CODE")) Then Datum.ConventionalRSCode = -1 Else Datum.ConventionalRSCode = ds.Tables("Datum").Rows(0).Item("CONVENTIONAL_RS_CODE")
            'If IsDBNull(ds.Tables("Datum").Rows(0).Item("PUBLICATION_DATE")) Then Datum.PublicationDate = Date.MinValue Else Datum.PublicationDate = ds.Tables("Datum").Rows(0).Item("PUBLICATION_DATE") '1999 not valid
            If IsDBNull(ds.Tables("Datum").Rows(0).Item("PUBLICATION_DATE")) Then Datum.PublicationDate = "" Else Datum.PublicationDate = ds.Tables("Datum").Rows(0).Item("PUBLICATION_DATE")
            If IsDBNull(ds.Tables("Datum").Rows(0).Item("FRAME_REFERENCE_EPOCH")) Then Datum.FrameReferenceEpoch = Double.NaN Else Datum.FrameReferenceEpoch = ds.Tables("Datum").Rows(0).Item("FRAME_REFERENCE_EPOCH")
            If IsDBNull(ds.Tables("Datum").Rows(0).Item("REALIZATION_METHOD_CODE")) Then Datum.RealizationMethodCode = -1 Else Datum.RealizationMethodCode = ds.Tables("Datum").Rows(0).Item("REALIZATION_METHOD_CODE")

            GetEllipsoid(Datum.EllipsoidCode)
            GetPrimeMeridian(Datum.PrimeMeridianCode)
            GetConventionalRS(Datum.ConventionalRSCode)
        Else
            RaiseEvent ErrorMessage("There are " & ds.Tables("Datum").Rows.Count & " Datums with the code: " & DatumCode & vbCrLf)
        End If
    End Sub

    Private Sub GetEllipsoid(EllipsoidCode As Integer)
        'Get the CRS Ellipsoid information.

        If EpsgDatabasePath = "" Then
            RaiseEvent ErrorMessage("No EPSG database has been selected." & vbCrLf)
            Exit Sub
        End If

        If Not System.IO.File.Exists(EpsgDatabasePath) Then
            RaiseEvent ErrorMessage("Selected EPSG database can not be found." & vbCrLf)
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + EpsgDatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter("Select * From Ellipsoid Where ELLIPSOID_CODE = " & EllipsoidCode.ToString, conn)
        Dim ds As DataSet = New DataSet
        da.Fill(ds, "Ellipsoid")
        conn.Close()

        If ds.Tables("Ellipsoid").Rows.Count = 0 Then
            RaiseEvent ErrorMessage("There is no Ellipsoid with the code: " & EllipsoidCode & vbCrLf)
        ElseIf ds.Tables("Ellipsoid").Rows.Count = 1 Then
            Ellipsoid.Code = EllipsoidCode
            If IsDBNull(ds.Tables("Ellipsoid").Rows(0).Item("ELLIPSOID_NAME")) Then Ellipsoid.Name = "" Else Ellipsoid.Name = ds.Tables("Ellipsoid").Rows(0).Item("ELLIPSOID_NAME")
            If IsDBNull(ds.Tables("Ellipsoid").Rows(0).Item("SEMI_MAJOR_AXIS")) Then Ellipsoid.SemiMajorAxis = Double.NaN Else Ellipsoid.SemiMajorAxis = ds.Tables("Ellipsoid").Rows(0).Item("SEMI_MAJOR_AXIS")
            If IsDBNull(ds.Tables("Ellipsoid").Rows(0).Item("UOM_CODE")) Then Ellipsoid.UomCode = -1 Else Ellipsoid.UomCode = ds.Tables("Ellipsoid").Rows(0).Item("UOM_CODE")
            If IsDBNull(ds.Tables("Ellipsoid").Rows(0).Item("INV_FLATTENING")) Then Ellipsoid.InvFlattening = Double.NaN Else Ellipsoid.InvFlattening = ds.Tables("Ellipsoid").Rows(0).Item("INV_FLATTENING")
            If IsDBNull(ds.Tables("Ellipsoid").Rows(0).Item("SEMI_MINOR_AXIS")) Then Ellipsoid.SemiMinorAxis = Double.NaN Else Ellipsoid.SemiMinorAxis = ds.Tables("Ellipsoid").Rows(0).Item("SEMI_MINOR_AXIS")
            If IsDBNull(ds.Tables("Ellipsoid").Rows(0).Item("ELLIPSOID_SHAPE")) Then Ellipsoid.EllipsoidShape = True Else Ellipsoid.EllipsoidShape = ds.Tables("Ellipsoid").Rows(0).Item("ELLIPSOID_SHAPE")
            If IsDBNull(ds.Tables("Ellipsoid").Rows(0).Item("REMARKS")) Then Ellipsoid.Remarks = "" Else Ellipsoid.Remarks = ds.Tables("Ellipsoid").Rows(0).Item("REMARKS")
            If IsDBNull(ds.Tables("Ellipsoid").Rows(0).Item("INFORMATION_SOURCE")) Then Ellipsoid.InfoSource = "" Else Ellipsoid.InfoSource = ds.Tables("Ellipsoid").Rows(0).Item("INFORMATION_SOURCE")
            If IsDBNull(ds.Tables("Ellipsoid").Rows(0).Item("DATA_SOURCE")) Then Ellipsoid.DataSource = "" Else Ellipsoid.DataSource = ds.Tables("Ellipsoid").Rows(0).Item("DATA_SOURCE")
            If IsDBNull(ds.Tables("Ellipsoid").Rows(0).Item("REVISION_DATE")) Then Ellipsoid.RevisionDate = Date.MinValue Else Ellipsoid.RevisionDate = ds.Tables("Ellipsoid").Rows(0).Item("REVISION_DATE")
            If IsDBNull(ds.Tables("Ellipsoid").Rows(0).Item("CHANGE_ID")) Then Ellipsoid.ChangeID = "" Else Ellipsoid.ChangeID = ds.Tables("Ellipsoid").Rows(0).Item("CHANGE_ID")
            If IsDBNull(ds.Tables("Ellipsoid").Rows(0).Item("DEPRECATED")) Then Ellipsoid.Deprecated = False Else Ellipsoid.Deprecated = ds.Tables("Ellipsoid").Rows(0).Item("DEPRECATED")
            CalcEllipsoidESquared() 'Calculate the E2 parameter required to calculate X, Y, Z from Lat, Long, Ellipsoidal Height.
        Else
            RaiseEvent ErrorMessage("There are " & ds.Tables("Ellipsoid").Rows.Count & " Ellipsoids with the code: " & EllipsoidCode & vbCrLf)
        End If
    End Sub

    Private Sub GetPrimeMeridian(PrimeMeridianCode As Integer)
        'Get the CRS Prime Meridian information.

        If EpsgDatabasePath = "" Then
            RaiseEvent ErrorMessage("No EPSG database has been selected." & vbCrLf)
            Exit Sub
        End If

        If Not System.IO.File.Exists(EpsgDatabasePath) Then
            RaiseEvent ErrorMessage("Selected EPSG database can not be found." & vbCrLf)
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + EpsgDatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter("Select * From [Prime Meridian] Where PRIME_MERIDIAN_CODE = " & PrimeMeridianCode.ToString, conn)
        Dim ds As DataSet = New DataSet
        da.Fill(ds, "PrimeMeridian")

        If ds.Tables("PrimeMeridian").Rows.Count = 0 Then
            RaiseEvent ErrorMessage("There is no Prime Meridian with the code: " & PrimeMeridianCode & vbCrLf)
        ElseIf ds.Tables("PrimeMeridian").Rows.Count = 1 Then
            PrimeMeridian.Code = PrimeMeridianCode
            If IsDBNull(ds.Tables("PrimeMeridian").Rows(0).Item("PRIME_MERIDIAN_NAME")) Then PrimeMeridian.Name = "" Else PrimeMeridian.Name = ds.Tables("PrimeMeridian").Rows(0).Item("PRIME_MERIDIAN_NAME")
            If IsDBNull(ds.Tables("PrimeMeridian").Rows(0).Item("GREENWICH_LONGITUDE")) Then PrimeMeridian.GreenwichLongitude = Double.NaN Else PrimeMeridian.GreenwichLongitude = ds.Tables("PrimeMeridian").Rows(0).Item("GREENWICH_LONGITUDE")
            If IsDBNull(ds.Tables("PrimeMeridian").Rows(0).Item("UOM_CODE")) Then PrimeMeridian.UomCode = -1 Else PrimeMeridian.UomCode = ds.Tables("PrimeMeridian").Rows(0).Item("UOM_CODE")
            If IsDBNull(ds.Tables("PrimeMeridian").Rows(0).Item("REMARKS")) Then PrimeMeridian.Remarks = "" Else PrimeMeridian.Remarks = ds.Tables("PrimeMeridian").Rows(0).Item("REMARKS")
            If IsDBNull(ds.Tables("PrimeMeridian").Rows(0).Item("INFORMATION_SOURCE")) Then PrimeMeridian.InfoSource = "" Else PrimeMeridian.InfoSource = ds.Tables("PrimeMeridian").Rows(0).Item("INFORMATION_SOURCE")
            If IsDBNull(ds.Tables("PrimeMeridian").Rows(0).Item("DATA_SOURCE")) Then PrimeMeridian.DataSource = "" Else PrimeMeridian.DataSource = ds.Tables("PrimeMeridian").Rows(0).Item("DATA_SOURCE")
            If IsDBNull(ds.Tables("PrimeMeridian").Rows(0).Item("REVISION_DATE")) Then PrimeMeridian.RevisionDate = Date.MinValue Else PrimeMeridian.RevisionDate = ds.Tables("PrimeMeridian").Rows(0).Item("REVISION_DATE")

        Else
            RaiseEvent ErrorMessage("There are " & ds.Tables("PrimeMeridian").Rows.Count & " Prime Meridians with the code: " & PrimeMeridianCode & vbCrLf)
        End If
    End Sub

    Private Sub GetConventionalRS(RSCode As Integer)
        'The the Converntional Reference System information corresponding to RSCode.

        If RSCode = -1 Then
            'There is no Conventional Reference System.
            Exit Sub
        End If

        If EpsgDatabasePath = "" Then
            RaiseEvent ErrorMessage("No EPSG database has been selected." & vbCrLf)
            Exit Sub
        End If

        If Not System.IO.File.Exists(EpsgDatabasePath) Then
            RaiseEvent ErrorMessage("Selected EPSG database can not be found." & vbCrLf)
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + EpsgDatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter("Select * From [ConventionalRS] Where CONVENTIONAL_RS_CODE = " & RSCode.ToString, conn)
        Dim ds As DataSet = New DataSet
        da.Fill(ds, "ConvRS")

        If ds.Tables("ConvRS").Rows.Count = 0 Then

        ElseIf ds.Tables("ConvRS").Rows.Count = 1 Then
            ConventionalRS.Code = RSCode
            If IsDBNull(ds.Tables("ConvRS").Rows(0).Item("CONVENTIONAL_RS_NAME")) Then ConventionalRS.Name = "" Else ConventionalRS.Name = ds.Tables("ConvRS").Rows(0).Item("CONVENTIONAL_RS_NAME")
            If IsDBNull(ds.Tables("ConvRS").Rows(0).Item("REMARKS")) Then ConventionalRS.Remarks = "" Else ConventionalRS.Remarks = ds.Tables("ConvRS").Rows(0).Item("REMARKS")
            If IsDBNull(ds.Tables("ConvRS").Rows(0).Item("INFORMATION_SOURCE")) Then ConventionalRS.InfoSource = "" Else ConventionalRS.InfoSource = ds.Tables("ConvRS").Rows(0).Item("INFORMATION_SOURCE")
            If IsDBNull(ds.Tables("ConvRS").Rows(0).Item("DATA_SOURCE")) Then ConventionalRS.DataSource = "" Else ConventionalRS.DataSource = ds.Tables("ConvRS").Rows(0).Item("DATA_SOURCE")
            If IsDBNull(ds.Tables("ConvRS").Rows(0).Item("REVISION_DATE")) Then ConventionalRS.RevisionDate = Date.MinValue Else ConventionalRS.RevisionDate = ds.Tables("ConvRS").Rows(0).Item("REVISION_DATE")
            If IsDBNull(ds.Tables("ConvRS").Rows(0).Item("CHANGE_ID")) Then ConventionalRS.ChangeID = "" Else ConventionalRS.ChangeID = ds.Tables("ConvRS").Rows(0).Item("CHANGE_ID")
            If IsDBNull(ds.Tables("ConvRS").Rows(0).Item("DEPRECATED")) Then ConventionalRS.Deprecated = False Else ConventionalRS.Deprecated = ds.Tables("ConvRS").Rows(0).Item("DEPRECATED")
        Else
            RaiseEvent ErrorMessage("There are " & ds.Tables("ConvRS").Rows.Count & " Conventional Reference Systems with the code: " & RSCode & vbCrLf)
        End If

    End Sub

    Private Sub GetCoordSystem(CoordSysCode As Integer)
        'Get the coordinate system information.

        If EpsgDatabasePath = "" Then
            RaiseEvent ErrorMessage("No EPSG database has been selected." & vbCrLf)
            Exit Sub
        End If

        If Not System.IO.File.Exists(EpsgDatabasePath) Then
            RaiseEvent ErrorMessage("Selected EPSG database can not be found." & vbCrLf)
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + EpsgDatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter("Select * From [Coordinate System] Where  COORD_SYS_CODE = " & CoordSysCode.ToString, conn)
        Dim ds As DataSet = New DataSet
        da.Fill(ds, "CoordSys")

        If ds.Tables("CoordSys").Rows.Count = 0 Then
            RaiseEvent ErrorMessage("There is no Coordinate System with the code: " & CoordSysCode & vbCrLf)
        ElseIf ds.Tables("CoordSys").Rows.Count = 1 Then
            CoordSystem.Code = CoordSysCode
            If IsDBNull(ds.Tables("CoordSys").Rows(0).Item("COORD_SYS_NAME")) Then CoordSystem.Name = "" Else CoordSystem.Name = ds.Tables("CoordSys").Rows(0).Item("COORD_SYS_NAME")

            Select Case ds.Tables("CoordSys").Rows(0).Item("COORD_SYS_Type")
                Case "affine"
                    CoordSystem.Type = CoordSystem.CsType.affine
                Case "Cartesian"
                    CoordSystem.Type = CoordSystem.CsType.Cartesian
                Case "cylindrical"
                    CoordSystem.Type = CoordSystem.CsType.cylindrical
                Case "ellipsoidal"
                    CoordSystem.Type = CoordSystem.CsType.ellipsoidal
                Case "linear"
                    CoordSystem.Type = CoordSystem.CsType.linear
                Case "ordinal"
                    CoordSystem.Type = CoordSystem.CsType.ordinal
                Case "polar"
                    CoordSystem.Type = CoordSystem.CsType.polar
                Case "spherical"
                    CoordSystem.Type = CoordSystem.CsType.spherical
                Case "vertical"
                    CoordSystem.Type = CoordSystem.CsType.vertical
                Case "parametric"
                    CoordSystem.Type = CoordSystem.CsType.parametric
                Case "temporal count"
                    CoordSystem.Type = CoordSystem.CsType.temporalCount
                Case "temporal datetime"
                    CoordSystem.Type = CoordSystem.CsType.temporalDatetime
                Case "temporal measure"
                    CoordSystem.Type = CoordSystem.CsType.temporalMeasure
                Case Else
                    RaiseEvent ErrorMessage("Unknown coordinate system type: " & ds.Tables("CoordSys").Rows(0).Item("COORD_SYS_Type") & vbCrLf)
            End Select

            If IsDBNull(ds.Tables("CoordSys").Rows(0).Item("DIMENSION")) Then CoordSystem.Dimension = 2 Else CoordSystem.Dimension = ds.Tables("CoordSys").Rows(0).Item("DIMENSION")
            If IsDBNull(ds.Tables("CoordSys").Rows(0).Item("REMARKS")) Then CoordSystem.Remarks = "" Else CoordSystem.Remarks = ds.Tables("CoordSys").Rows(0).Item("REMARKS")
            If IsDBNull(ds.Tables("CoordSys").Rows(0).Item("INFORMATION_SOURCE")) Then CoordSystem.InfoSource = "" Else CoordSystem.InfoSource = ds.Tables("CoordSys").Rows(0).Item("INFORMATION_SOURCE")
            If IsDBNull(ds.Tables("CoordSys").Rows(0).Item("DATA_SOURCE")) Then CoordSystem.DataSource = "" Else CoordSystem.DataSource = ds.Tables("CoordSys").Rows(0).Item("DATA_SOURCE")
            If IsDBNull(ds.Tables("CoordSys").Rows(0).Item("REVISION_DATE")) Then CoordSystem.RevisionDate = Date.MinValue Else CoordSystem.RevisionDate = ds.Tables("CoordSys").Rows(0).Item("REVISION_DATE")
            If IsDBNull(ds.Tables("CoordSys").Rows(0).Item("CHANGE_ID")) Then CoordSystem.ChangeID = "" Else CoordSystem.ChangeID = ds.Tables("CoordSys").Rows(0).Item("CHANGE_ID")
            If IsDBNull(ds.Tables("CoordSys").Rows(0).Item("DEPRECATED")) Then CoordSystem.Deprecated = False Else CoordSystem.Deprecated = ds.Tables("CoordSys").Rows(0).Item("DEPRECATED")
            GetCoordAxisInfo(CoordSysCode)
        Else
            RaiseEvent ErrorMessage("There are " & ds.Tables("CoordSys").Rows.Count & " Coordinate Systems with the code: " & CoordSysCode & vbCrLf)
        End If
        conn.Close()
    End Sub

    Private Sub GetCoordAxisInfo(CoordSysCode As Integer)
        'Get the Coondinate Axis information correponding to CoordSysCode.

        If EpsgDatabasePath = "" Then
            RaiseEvent ErrorMessage("No EPSG database has been selected." & vbCrLf)
            Exit Sub
        End If

        If Not System.IO.File.Exists(EpsgDatabasePath) Then
            RaiseEvent ErrorMessage("Selected EPSG database can not be found." & vbCrLf)
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + EpsgDatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter("Select * From [Coordinate Axis] Where  COORD_SYS_CODE = " & CoordSysCode.ToString & " Order By 'ORDER'", conn)
        Dim ds As DataSet = New DataSet
        da.Fill(ds, "CoordAxis")

        If ds.Tables("CoordAxis").Rows.Count = 0 Then
            RaiseEvent ErrorMessage("There are no Coordinate Axes with the Coordinate System code: " & CoordSysCode & vbCrLf)
        Else
            For Each Item As DataRow In ds.Tables("CoordAxis").Rows
                Dim NewAxis As New CoordAxis
                If IsDBNull(Item("COORD_AXIS_CODE")) Then NewAxis.Code = -1 Else NewAxis.Code = Item("COORD_AXIS_CODE")
                NewAxis.SystemCode = CoordSysCode
                If IsDBNull(Item("COORD_AXIS_NAME_CODE")) Then NewAxis.NameCode = "" Else NewAxis.NameCode = Item("COORD_AXIS_NAME_CODE")
                If IsDBNull(Item("COORD_AXIS_ORIENTATION")) Then NewAxis.Orientation = "" Else NewAxis.Orientation = Item("COORD_AXIS_ORIENTATION")
                If IsDBNull(Item("COORD_AXIS_ABBREVIATION")) Then NewAxis.Abbreviation = "" Else NewAxis.Abbreviation = Item("COORD_AXIS_ABBREVIATION")
                If IsDBNull(Item("UOM_CODE")) Then NewAxis.UomCode = -1 Else NewAxis.UomCode = ds.Tables("CoordAxis").Rows(0).Item("UOM_CODE")
                If IsDBNull(Item("ORDER")) Then NewAxis.Order = 1 Else NewAxis.Order = Item("ORDER")
                CoordAxisList.Add(NewAxis)

                'Get list of Input Target Coordinate Operations:
                da.SelectCommand.CommandText = "Select * From [Coordinate Axis Name] Where  COORD_AXIS_NAME_CODE = " & NewAxis.NameCode.ToString
                da.Fill(ds, "AxisName")
                Dim NewAxisName As New CoordAxisName
                If ds.Tables("AxisName").Rows.Count = 0 Then
                    RaiseEvent ErrorMessage("There are no Coordinate Axis Name with the Name code: " & NewAxis.NameCode.ToString & vbCrLf)
                    CoordAxisNameList.Add(NewAxisName) 'Add a blank entry.
                ElseIf ds.Tables("AxisName").Rows.Count = 1 Then
                    NewAxisName.Code = NewAxis.NameCode
                    If IsDBNull(ds.Tables("AxisName").Rows(0).Item("COORD_AXIS_NAME")) Then NewAxisName.Name = "" Else NewAxisName.Name = ds.Tables("AxisName").Rows(0).Item("COORD_AXIS_NAME")
                    If IsDBNull(ds.Tables("AxisName").Rows(0).Item("DESCRIPTION")) Then NewAxisName.Description = "" Else NewAxisName.Description = ds.Tables("AxisName").Rows(0).Item("DESCRIPTION")
                    If IsDBNull(ds.Tables("AxisName").Rows(0).Item("REMARKS")) Then NewAxisName.Remarks = "" Else NewAxisName.Remarks = ds.Tables("AxisName").Rows(0).Item("REMARKS")
                    If IsDBNull(ds.Tables("AxisName").Rows(0).Item("INFORMATION_SOURCE")) Then NewAxisName.InfoSource = "" Else NewAxisName.InfoSource = ds.Tables("AxisName").Rows(0).Item("INFORMATION_SOURCE")
                    If IsDBNull(ds.Tables("AxisName").Rows(0).Item("DATA_SOURCE")) Then NewAxisName.DataSource = "" Else NewAxisName.DataSource = ds.Tables("AxisName").Rows(0).Item("DATA_SOURCE")
                    If IsDBNull(ds.Tables("AxisName").Rows(0).Item("REVISION_DATE")) Then NewAxisName.RevisionDate = Date.MinValue Else NewAxisName.RevisionDate = ds.Tables("AxisName").Rows(0).Item("REVISION_DATE")
                    If IsDBNull(ds.Tables("AxisName").Rows(0).Item("CHANGE_ID")) Then NewAxisName.ChangeID = "" Else NewAxisName.ChangeID = ds.Tables("AxisName").Rows(0).Item("CHANGE_ID")
                    If IsDBNull(ds.Tables("AxisName").Rows(0).Item("DEPRECATED")) Then NewAxisName.Deprecated = False Else NewAxisName.Deprecated = ds.Tables("AxisName").Rows(0).Item("DEPRECATED")
                    CoordAxisNameList.Add(NewAxisName)
                Else
                    CoordAxisNameList.Add(NewAxisName) 'Add a blank entry.
                End If
                ds.Tables("AxisName").Clear()
            Next
        End If
    End Sub

    Private Sub GetProjConvInfo(ProjConvCode As Integer)
        'Get the information about the Projection Conversion used to convert between the Derived CRS and the Base CRS.

        If EpsgDatabasePath = "" Then
            RaiseEvent ErrorMessage("No EPSG database has been selected." & vbCrLf)
            Exit Sub
        End If

        If Not System.IO.File.Exists(EpsgDatabasePath) Then
            RaiseEvent ErrorMessage("Selected EPSG database can not be found." & vbCrLf)
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + EpsgDatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        'Get list of Input Source Coordinate Operations:
        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter("Select * From Coordinate_Operation Where COORD_OP_CODE = " & ProjConvCode.ToString, conn)
        Dim ds As DataSet = New DataSet
        da.Fill(ds, "CoordOp")
        SourceCoordOpList.Clear()

        If ds.Tables("CoordOp").Rows.Count = 0 Then
            RaiseEvent ErrorMessage("There are no Coordinate Operation records for Projection Conversion code number: " & ProjConvCode & vbCrLf)
        ElseIf ds.Tables("CoordOp").Rows.Count = 1 Then
            ProjectionCoordOp.Code = ProjConvCode
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("COORD_OP_NAME")) Then ProjectionCoordOp.Name = "" Else ProjectionCoordOp.Name = ds.Tables("CoordOp").Rows(0).Item("COORD_OP_NAME")

            Select Case ds.Tables("CoordOp").Rows(0).Item("COORD_OP_TYPE")
                Case "conversion"
                    ProjectionCoordOp.Type = CoordinateOperation.OperationType.conversion
                Case "transformation"
                    ProjectionCoordOp.Type = CoordinateOperation.OperationType.transformation
                Case "point motion operation"
                    ProjectionCoordOp.Type = CoordinateOperation.OperationType.pointMotionOperation
                Case "concatenated operation"
                    ProjectionCoordOp.Type = CoordinateOperation.OperationType.concatenatedOperation
                Case Else
                    RaiseEvent ErrorMessage("Unknown Target coordinate operation type: " & ds.Tables("CoordOp").Rows(0).Item("COORD_OP_TYPE") & vbCrLf)
                    ProjectionCoordOp.Type = CoordinateOperation.OperationType.conversion
            End Select

            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("SOURCE_CRS_CODE")) Then ProjectionCoordOp.SourceCrsCode = -1 Else ProjectionCoordOp.SourceCrsCode = ds.Tables("CoordOp").Rows(0).Item("SOURCE_CRS_CODE")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("TARGET_CRS_CODE")) Then ProjectionCoordOp.TargetCrsCode = -1 Else ProjectionCoordOp.TargetCrsCode = ds.Tables("CoordOp").Rows(0).Item("TARGET_CRS_CODE")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("COORD_TFM_VERSION")) Then ProjectionCoordOp.Version = "" Else ProjectionCoordOp.Version = ds.Tables("CoordOp").Rows(0).Item("COORD_TFM_VERSION")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("COORD_OP_VARIANT")) Then ProjectionCoordOp.OpVariant = -1 Else ProjectionCoordOp.OpVariant = ds.Tables("CoordOp").Rows(0).Item("COORD_OP_VARIANT")
            'AREA_OF_USE_CODE has been deprecated.
            'COORD_OP_SCOPE has been deprecated.
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("COORD_OP_METHOD_CODE")) Then ProjectionCoordOp.MethodCode = -1 Else ProjectionCoordOp.MethodCode = ds.Tables("CoordOp").Rows(0).Item("COORD_OP_METHOD_CODE")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("UOM_CODE_SOURCE_COORD_DIFF")) Then ProjectionCoordOp.UomSourceCoordDiffCode = -1 Else ProjectionCoordOp.UomSourceCoordDiffCode = ds.Tables("CoordOp").Rows(0).Item("UOM_CODE_SOURCE_COORD_DIFF")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("UOM_CODE_TARGET_COORD_DIFF")) Then ProjectionCoordOp.UomTargetCoordDiffCode = -1 Else ProjectionCoordOp.UomTargetCoordDiffCode = ds.Tables("CoordOp").Rows(0).Item("UOM_CODE_TARGET_COORD_DIFF")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("REMARKS")) Then ProjectionCoordOp.Remarks = "" Else ProjectionCoordOp.Remarks = ds.Tables("CoordOp").Rows(0).Item("REMARKS")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("INFORMATION_SOURCE")) Then ProjectionCoordOp.InfoSource = "" Else ProjectionCoordOp.InfoSource = ds.Tables("CoordOp").Rows(0).Item("INFORMATION_SOURCE")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("DATA_SOURCE")) Then ProjectionCoordOp.DataSource = "" Else ProjectionCoordOp.DataSource = ds.Tables("CoordOp").Rows(0).Item("DATA_SOURCE")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("REVISION_DATE")) Then ProjectionCoordOp.RevisionDate = Date.MinValue Else ProjectionCoordOp.RevisionDate = ds.Tables("CoordOp").Rows(0).Item("REVISION_DATE")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("CHANGE_ID")) Then ProjectionCoordOp.ChangeID = "" Else ProjectionCoordOp.ChangeID = ds.Tables("CoordOp").Rows(0).Item("CHANGE_ID")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("SHOW_OPERATION")) Then ProjectionCoordOp.Show = True Else ProjectionCoordOp.Show = ds.Tables("CoordOp").Rows(0).Item("SHOW_OPERATION")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("DEPRECATED")) Then ProjectionCoordOp.Deprecated = False Else ProjectionCoordOp.Deprecated = ds.Tables("CoordOp").Rows(0).Item("DEPRECATED")

            'Get Coordinate Operation Method:
            Dim MethodCode As Integer = ProjectionCoordOp.MethodCode
            da.SelectCommand.CommandText = "Select * From [Coordinate_Operation Method] Where COORD_OP_METHOD_CODE = " & MethodCode.ToString
            da.Fill(ds, "Method")
            If ds.Tables("Method").Rows.Count = 0 Then
                RaiseEvent ErrorMessage("There are no Coordinate Operation Methods for Method code number: " & MethodCode.ToString & vbCrLf)
            ElseIf ds.Tables("Method").Rows.Count = 1 Then
                ProjectionCoordOpMethod.Code = MethodCode
                If IsDBNull(ds.Tables("Method").Rows(0).Item("COORD_OP_METHOD_NAME")) Then ProjectionCoordOpMethod.Name = "" Else ProjectionCoordOpMethod.Name = ds.Tables("Method").Rows(0).Item("COORD_OP_METHOD_NAME")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("REVERSE_OP")) Then ProjectionCoordOpMethod.ReverseOp = False Else ProjectionCoordOpMethod.ReverseOp = ds.Tables("Method").Rows(0).Item("REVERSE_OP")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("FORMULA")) Then ProjectionCoordOpMethod.Formula = "" Else ProjectionCoordOpMethod.Formula = ds.Tables("Method").Rows(0).Item("FORMULA")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("EXAMPLE")) Then ProjectionCoordOpMethod.Example = "" Else ProjectionCoordOpMethod.Example = ds.Tables("Method").Rows(0).Item("EXAMPLE")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("REMARKS")) Then ProjectionCoordOpMethod.Remarks = "" Else ProjectionCoordOpMethod.Remarks = ds.Tables("Method").Rows(0).Item("REMARKS")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("INFORMATION_SOURCE")) Then ProjectionCoordOpMethod.InfoSource = "" Else ProjectionCoordOpMethod.InfoSource = ds.Tables("Method").Rows(0).Item("INFORMATION_SOURCE")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("DATA_SOURCE")) Then ProjectionCoordOpMethod.DataSource = "" Else ProjectionCoordOpMethod.DataSource = ds.Tables("Method").Rows(0).Item("DATA_SOURCE")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("REVISION_DATE")) Then ProjectionCoordOpMethod.RevisionDate = Date.MinValue Else ProjectionCoordOpMethod.RevisionDate = ds.Tables("Method").Rows(0).Item("REVISION_DATE")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("CHANGE_ID")) Then ProjectionCoordOpMethod.ChangeID = "" Else ProjectionCoordOpMethod.ChangeID = ds.Tables("Method").Rows(0).Item("CHANGE_ID")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("DEPRECATED")) Then ProjectionCoordOpMethod.Deprecated = False Else ProjectionCoordOpMethod.Deprecated = ds.Tables("Method").Rows(0).Item("DEPRECATED")
                Projection.MethodName = ProjectionCoordOpMethod.Name
            Else
                RaiseEvent ErrorMessage("There are " & ds.Tables("Method").Rows.Count & " Coordinate Operation Methods with the code: " & MethodCode.ToString & vbCrLf)
            End If

            'Get the Coordinate Operation Parameter Usage:
            da.SelectCommand.CommandText = "Select * From [Coordinate_Operation Parameter Usage] Where COORD_OP_METHOD_CODE = " & MethodCode.ToString & " Order By SORT_ORDER"
            da.Fill(ds, "Usage")
            If ds.Tables("Usage").Rows.Count = 0 Then
                RaiseEvent ErrorMessage("There are no Coordinate Operation Parameter Usage records for Method code number: " & MethodCode.ToString & vbCrLf)
            Else
                Dim ParamCode As Integer
                Dim ParamVal As Double
                For Each Item As DataRow In ds.Tables("Usage").Rows
                    Dim NewUsage As New CoordOpParamUsage
                    NewUsage.MethodCode = MethodCode
                    'If IsDBNull(ds.Tables("Usage").Rows(0).Item("PARAMETER_CODE")) Then NewUsage.ParameterCode = -1 Else NewUsage.ParameterCode = ds.Tables("Usage").Rows(0).Item("PARAMETER_CODE")
                    If IsDBNull(Item("PARAMETER_CODE")) Then NewUsage.ParameterCode = -1 Else NewUsage.ParameterCode = Item("PARAMETER_CODE")
                    If IsDBNull(Item("SORT_ORDER")) Then NewUsage.SortOrder = 0 Else NewUsage.SortOrder = Item("SORT_ORDER")
                    If IsDBNull(Item("PARAM_SIGN_REVERSAL")) Then NewUsage.SignReversal = "" Else NewUsage.SignReversal = Item("PARAM_SIGN_REVERSAL")
                    ProjectionCoordOpParamUseList.Add(NewUsage)

                    'Get the corresponding Coordinate Operation Parameter information:
                    da.SelectCommand.CommandText = "Select * From [Coordinate_Operation Parameter] Where PARAMETER_CODE = " & NewUsage.ParameterCode.ToString
                    da.Fill(ds, "OpParameter")
                    If ds.Tables("OpParameter").Rows.Count = 0 Then
                        RaiseEvent ErrorMessage("There are no Coordinate Operation Parameter records for Parameter code number: " & NewUsage.ParameterCode.ToString & vbCrLf)
                    ElseIf ds.Tables("OpParameter").Rows.Count = 1 Then
                        Dim NewParameter As New CoordOpParameter
                        NewParameter.Code = NewUsage.ParameterCode
                        ParamCode = NewParameter.Code
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("PARAMETER_NAME")) Then NewParameter.Name = "" Else NewParameter.Name = ds.Tables("OpParameter").Rows(0).Item("PARAMETER_NAME")
                        'ParamName = NewParameter.Name
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("DESCRIPTION")) Then NewParameter.Description = "" Else NewParameter.Description = ds.Tables("OpParameter").Rows(0).Item("DESCRIPTION")
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("INFORMATION_SOURCE")) Then NewParameter.InfoSource = "" Else NewParameter.InfoSource = ds.Tables("OpParameter").Rows(0).Item("INFORMATION_SOURCE")
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("DATA_SOURCE")) Then NewParameter.DataSource = "" Else NewParameter.DataSource = ds.Tables("OpParameter").Rows(0).Item("DATA_SOURCE")
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("REVISION_DATE")) Then NewParameter.RevisionDate = Date.MinValue Else NewParameter.RevisionDate = ds.Tables("OpParameter").Rows(0).Item("REVISION_DATE")
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("CHANGE_ID")) Then NewParameter.ChangeID = "" Else NewParameter.ChangeID = ds.Tables("OpParameter").Rows(0).Item("CHANGE_ID")
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("DEPRECATED")) Then NewParameter.Deprecated = False Else NewParameter.Deprecated = ds.Tables("OpParameter").Rows(0).Item("DEPRECATED")
                        ProjectionCoordOpParamList.Add(NewParameter)
                    Else
                        RaiseEvent ErrorMessage("There are " & ds.Tables("OpParameter").Rows.Count & "Parameters with the code: " & NewUsage.ParameterCode.ToString & vbCrLf)
                    End If
                    ds.Tables("OpParameter").Clear()

                    'Get the corresponding Coordinate Operation Parameter Value:
                    da.SelectCommand.CommandText = "Select * From [Coordinate_Operation Parameter Value] Where PARAMETER_CODE = " & NewUsage.ParameterCode.ToString & " And COORD_OP_METHOD_CODE = " & MethodCode.ToString & " And COORD_OP_CODE = " & ProjConvCode.ToString
                    da.Fill(ds, "Value")
                    If ds.Tables("Value").Rows.Count = 0 Then
                        RaiseEvent ErrorMessage("There are no Coordinate Operation Parameter Value records for Parameter code number: " & NewUsage.ParameterCode.ToString & " and COORD_OP_METHOD_CODE = " & MethodCode.ToString & " and COORD_OP_CODE = " & ProjConvCode.ToString & vbCrLf)
                    ElseIf ds.Tables("Value").Rows.Count = 1 Then
                        Dim NewParamValue As New CoordOpParamValue
                        NewParamValue.OpCode = ProjConvCode
                        NewParamValue.MethodCode = MethodCode
                        NewParamValue.ParameterCode = NewUsage.ParameterCode
                        If IsDBNull(ds.Tables("Value").Rows(0).Item("PARAMETER_VALUE")) Then NewParamValue.ParameterValue = Double.NaN Else NewParamValue.ParameterValue = ds.Tables("Value").Rows(0).Item("PARAMETER_VALUE")
                        ParamVal = NewParamValue.ParameterValue
                        If IsDBNull(ds.Tables("Value").Rows(0).Item("PARAM_VALUE_FILE_REF")) Then NewParamValue.ParamValueFileRef = "" Else NewParamValue.ParamValueFileRef = ds.Tables("Value").Rows(0).Item("PARAM_VALUE_FILE_REF")
                        If IsDBNull(ds.Tables("Value").Rows(0).Item("UOM_CODE")) Then NewParamValue.UomCode = -1 Else NewParamValue.UomCode = ds.Tables("Value").Rows(0).Item("UOM_CODE")
                        ProjectionCoordOpParamValList.Add(NewParamValue)
                    Else
                        RaiseEvent ErrorMessage("There are " & ds.Tables("Value").Rows.Count & "Parameter Values with the code: " & NewUsage.ParameterCode.ToString & " and COORD_OP_METHOD_CODE = " & MethodCode.ToString & " and COORD_OP_CODE = " & ProjConvCode.ToString & vbCrLf)
                    End If
                    ds.Tables("Value").Clear()
                    If IsNothing(Projection.Method) Then
                        'The projection method has not been defined.
                        'RaiseEvent ErrorMessage("The projection method has not been defined. " & vbCrLf)
                    Else
                        Projection.Method.SetParameter(ParamCode, ParamVal)
                    End If
                Next
                If IsNothing(Projection.Method) Then
                    'The projection method has not been defined.
                    RaiseEvent ErrorMessage("The projection method named: " & ProjectionCoordOpMethod.Name & " has not been coded. " & vbCrLf)

                Else
                    'Update the Ellipsoid settings:
                    Projection.Method.SemiMajorAxis = BaseCrs.Ellipsoid.SemiMajorAxis
                    Projection.Method.InverseFlattening = BaseCrs.Ellipsoid.InvFlattening
                    Projection.Method.UpdateVariables
                End If
            End If
        Else
            RaiseEvent ErrorMessage("There are " & ds.Tables("CoordOp").Rows.Count & " Coordinate Operations with the code: " & ProjConvCode & vbCrLf)
        End If
        conn.Close()
    End Sub

    Private Sub BaseCrs_ErrorMessage(Msg As String) Handles BaseCrs.ErrorMessage
        RaiseEvent ErrorMessage(Msg)
    End Sub

    Private Sub BaseCrs_Message(Msg As String) Handles BaseCrs.Message
        RaiseEvent Message(Msg)
    End Sub

    'Public Sub LatLongToNorthEast()
    Public Sub LongLatToEastNorth()
        'Convert the Latitude, Longitude coordinates to Northing, Easting coordinates.
        If IsNothing(Projection) Then
            RaiseEvent ErrorMessage("The projection method has not been defined." & vbCrLf)
        Else
            'Projection.Method.Latitude = Latitude
            'Projection.Method.Latitude = Coord.Latitude
            Projection.Method.Coord.Latitude = Coord.Latitude
            'Projection.Method.Longitude = Longitude
            'Projection.Method.Longitude = Coord.Longitude
            Projection.Method.Coord.Longitude = Coord.Longitude
            'Projection.Method.LatLongToEastNorth
            'Projection.Method.LongLatToNorthEast
            Projection.Method.LongLatToEastNorth
            'Easting = Projection.Method.Easting
            'Coord.Easting = Projection.Method.Easting
            Coord.Easting = Projection.Method.Coord.Easting
            'Northing = Projection.Method.Northing
            'Coord.Northing = Projection.Method.Northing
            Coord.Northing = Projection.Method.Coord.Northing
        End If
    End Sub

    'Public Sub NorthEastToLatLong()
    Public Sub EastNorthToLongLat()
        'Convert the Northing Easting coordinates to Latitude, Longitude coordinates.
        If IsNothing(Projection) Then
            RaiseEvent ErrorMessage("The projection method has not been defined." & vbCrLf)
        Else
            'Projection.Method.Northing = Northing
            'Projection.Method.Northing = Coord.Northing
            Projection.Method.Coord.Northing = Coord.Northing
            'Projection.Method.Easting = Easting
            'Projection.Method.Easting = Coord.Easting
            Projection.Method.Coord.Easting = Coord.Easting
            'Projection.Method.EastNorthToLatLong
            Projection.Method.EastNorthToLongLat
            'Latitude = Projection.Method.Latitude
            'Coord.Latitude = Projection.Method.Latitude
            Coord.Latitude = Projection.Method.Coord.Latitude
            'Longitude = Projection.Method.Longitude
            'Coord.Longitude = Projection.Method.Longitude
            Coord.Longitude = Projection.Method.Coord.Longitude
        End If
    End Sub

    Private Sub CalcEllipsoidESquared()
        'Calculate the E squared parameter used for Geodetic Transformation calculations.

        If Ellipsoid.InvFlattening = Double.NaN Then
            If Ellipsoid.SemiMajorAxis = Double.NaN Then
                RaiseEvent ErrorMessage("Please define the Ellipsoid used int he Coordinate Reference System." & vbCrLf)
            Else
                If Ellipsoid.SemiMinorAxis = Double.NaN Then
                    RaiseEvent ErrorMessage("Please define the Ellipsoid used int he Coordinate Reference System." & vbCrLf)
                Else
                    E2 = (Ellipsoid.SemiMajorAxis ^ 2 - Ellipsoid.SemiMinorAxis ^ 2) / Ellipsoid.SemiMajorAxis ^ 2
                End If
            End If
        Else
            Dim Flattening As Double = 1 / Ellipsoid.InvFlattening
            E2 = (2.0# * Flattening) - (Flattening * Flattening)
        End If
    End Sub

    Public Sub LongLatEllHtToXYZ()
        'Convert the geodetic Latitude, Longitude, Ellipsoidal Height coordinates to geocentric cartesian X, Y, Z coordinates.

        If Ellipsoid.Code = -1 Then
            LongLatEllHtToXYZ(BaseCrs)
        Else
            'Dim LatRad As Double = (Latitude / 180.0#) * Math.PI   'Latitude in radians
            Dim LatRad As Double = (Coord.Latitude / 180.0#) * Math.PI   'Latitude in radians
            'Dim LongRad As Double = (Longitude / 180.0#) * Math.PI 'Longitude in radians
            Dim LongRad As Double = (Coord.Longitude / 180.0#) * Math.PI 'Longitude in radians
            Dim N As Double = Ellipsoid.SemiMajorAxis / Math.Sqrt(1.0# - E2 * Math.Sin(LatRad) ^ 2) 'Radius of curvature in the prime vertical.

            'X = (N + EllipsoidalHeight) * Math.Cos(LatRad) * Math.Cos(LongRad)
            Coord.X = (N + Coord.EllipsoidalHeight) * Math.Cos(LatRad) * Math.Cos(LongRad)
            'Y = (N + EllipsoidalHeight) * Math.Cos(LatRad) * Math.Sin(LongRad)
            Coord.Y = (N + Coord.EllipsoidalHeight) * Math.Cos(LatRad) * Math.Sin(LongRad)
            'Z = (N * (1 - E2) + Coord.EllipsoidalHeight) * Math.Sin(LatRad)
            Coord.Z = (N * (1 - E2) + Coord.EllipsoidalHeight) * Math.Sin(LatRad)
        End If
    End Sub

    Public Sub LongLatEllHtToXYZ(BaseCrs As CoordRefSystem)
        'Convert the geodetic Latitude, Longitude, Ellipsoidal Height coordinates to geocentric cartesian X, Y, Z coordinates.
        'Dim LatRad As Double = (Latitude / 180.0#) * Math.PI   'Latitude in radians
        Dim LatRad As Double = (Coord.Latitude / 180.0#) * Math.PI   'Latitude in radians
        'Dim LongRad As Double = (Longitude / 180.0#) * Math.PI 'Longitude in radians
        Dim LongRad As Double = (Coord.Longitude / 180.0#) * Math.PI 'Longitude in radians
        'Dim N As Double = BaseCrs.Ellipsoid.SemiMajorAxis / Math.Sqrt(1.0# - E2 * Math.Sin(LatRad) ^ 2) 'Radius of curvature in the prime vertical.
        Dim N As Double = BaseCrs.Ellipsoid.SemiMajorAxis / Math.Sqrt(1.0# - BaseCrs.E2 * Math.Sin(LatRad) ^ 2) 'Radius of curvature in the prime vertical.

        'X = (N + EllipsoidalHeight) * Math.Cos(LatRad) * Math.Cos(LongRad)
        Coord.X = (N + Coord.EllipsoidalHeight) * Math.Cos(LatRad) * Math.Cos(LongRad)
        'Y = (N + EllipsoidalHeight) * Math.Cos(LatRad) * Math.Sin(LongRad)
        Coord.Y = (N + Coord.EllipsoidalHeight) * Math.Cos(LatRad) * Math.Sin(LongRad)
        'Z = (N * (1 - BaseCrs.E2) + EllipsoidalHeight) * Math.Sin(LatRad)
        Coord.Z = (N * (1 - BaseCrs.E2) + Coord.EllipsoidalHeight) * Math.Sin(LatRad)
    End Sub

    Public Sub XYZToLongLatEllHt()
        'Convert the geocentric cartesian X, Y, Z coordinates to geodetic Latitude, Longitude, Ellipsoidal Height coordinates.

        If Ellipsoid.Code = -1 Then
            XYZToLongLatEllHt(BaseCrs)
        Else
            'Dim Radius As Double = Math.Sqrt(X ^ 2 + Y ^ 2 + Z ^ 2)
            Dim Radius As Double = Math.Sqrt(Coord.X ^ 2 + Coord.Y ^ 2 + Coord.Z ^ 2)
            'Dim XYRadius As Double = Math.Sqrt(X ^ 2 + Y ^ 2)
            Dim XYRadius As Double = Math.Sqrt(Coord.X ^ 2 + Coord.Y ^ 2)
            'Dim LongRad As Double = 2.0# * Math.Atan(Y / (X + XYRadius))
            Dim LongRad As Double = 2.0# * Math.Atan(Coord.Y / (Coord.X + XYRadius))
            'Dim GeocentricLat As Double = Math.Atan(XYRadius / Z)
            Dim GeocentricLat As Double = Math.Atan(XYRadius / Coord.Z)

            'Dim LatEst As Double = GeocentricLat 'The estimated latitude (LatEst) is initially set to GeocentricLat.
            Dim LatEst As Double = GeocentricLat 'The estimated latitude (LatEst) is initially set to GeocentricLat.
            'Dim N As Double = Ellipsoid.SemiMajorAxis / Math.Sqrt(1.0# - E2 * Math.Sin(LatEst) ^ 2) 'Radius of curvature in the prime vertical.
            Dim N As Double = Ellipsoid.SemiMajorAxis / Math.Sqrt(1.0# - E2 * Math.Sin(LatEst) ^ 2) 'Radius of curvature in the prime vertical.

            Dim EllHeightEst As Double = XYRadius / Math.Cos(LatEst) - N

            Dim LatRef As Double = Math.Atan(Coord.Z / XYRadius / (1.0# - E2 * N / (N + EllHeightEst))) 'The refined latitude

            'Do While Math.Abs(LatEst - LatRef) > 0.0000000001#
            Do While Math.Abs(LatEst - LatRef) > 0.00000000001#
                'Do While Math.Abs(LatEst - LatRef) > 0.000000000001# 'No improvement
                LatEst = LatRef 'Update the latitude estimate
                N = Ellipsoid.SemiMajorAxis / Math.Sqrt(1.0# - E2 * Math.Sin(LatEst) ^ 2) 'The refined Radius of curvature in the prime vertical.
                EllHeightEst = XYRadius / Math.Cos(LatEst) - N 'The refined Ellipsoidal Height estimate
                LatRef = Math.Atan(Coord.Z / XYRadius / (1.0# - E2 * N / (N + EllHeightEst))) 'The refined latitude
            Loop

            'Longitude = LongRad / Math.PI * 180.0#
            Coord.Longitude = LongRad / Math.PI * 180.0#
            'Latitude = LatRef / Math.PI * 180.0#
            Coord.Latitude = LatRef / Math.PI * 180.0#
            'EllipsoidalHeight = XYRadius / Math.Cos(LatRef) - N
            Coord.EllipsoidalHeight = XYRadius / Math.Cos(LatRef) - N
        End If
    End Sub

    Public Sub XYZToLongLatEllHt(BaseCrs As CoordRefSystem)
        'Convert the geocentric cartesian X, Y, Z coordinates to geodetic Latitude, Longitude, Ellipsoidal Height coordinates.

        'Dim Radius As Double = Math.Sqrt(X ^ 2 + Y ^ 2 + Z ^ 2)
        Dim Radius As Double = Math.Sqrt(Coord.X ^ 2 + Coord.Y ^ 2 + Coord.Z ^ 2)
        'Dim XYRadius As Double = Math.Sqrt(X ^ 2 + Y ^ 2)
        Dim XYRadius As Double = Math.Sqrt(Coord.X ^ 2 + Coord.Y ^ 2)
        'Dim LongRad As Double = 2.0# * Math.Atan(Y / (X + XYRadius))
        Dim LongRad As Double = 2.0# * Math.Atan(Coord.Y / (Coord.X + XYRadius))
        'Dim GeocentricLat As Double = Math.Atan(XYRadius / Z)
        Dim GeocentricLat As Double = Math.Atan(XYRadius / Coord.Z)

        'Dim LatEst As Double = GeocentricLat 'The estimated latitude (LatEst) is initially set to GeocentricLat.
        Dim LatEst As Double = GeocentricLat 'The estimated latitude (LatEst) is initially set to GeocentricLat.
        Dim N As Double = BaseCrs.Ellipsoid.SemiMajorAxis / Math.Sqrt(1.0# - BaseCrs.E2 * Math.Sin(LatEst) ^ 2) 'Radius of curvature in the prime vertical.

        Dim EllHeightEst As Double = XYRadius / Math.Cos(LatEst) - N

        Dim LatRef As Double = Math.Atan(Coord.Z / XYRadius / (1.0# - BaseCrs.E2 * N / (N + EllHeightEst))) 'The refined latitude

        'Do While Math.Abs(LatEst - LatRef) > 0.0000000001#
        Do While Math.Abs(LatEst - LatRef) > 0.00000000001#
            'Do While Math.Abs(LatEst - LatRef) > 0.000000000001# 'No improvement
            LatEst = LatRef 'Update the latitude estimate
            N = BaseCrs.Ellipsoid.SemiMajorAxis / Math.Sqrt(1.0# - BaseCrs.E2 * Math.Sin(LatEst) ^ 2) 'The refined Radius of curvature in the prime vertical.
            EllHeightEst = XYRadius / Math.Cos(LatEst) - N 'The refined Ellipsoidal Height estimate
            'LatRef = Math.Atan(Z / XYRadius / (1.0# - BaseCrs.E2 * N / (N + EllHeightEst))) 'The refined latitude
            LatRef = Math.Atan(Coord.Z / XYRadius / (1.0# - BaseCrs.E2 * N / (N + EllHeightEst))) 'The refined latitude
        Loop

        'Longitude = LongRad / Math.PI * 180.0#
        Coord.Longitude = LongRad / Math.PI * 180.0#
        'Latitude = LatRef / Math.PI * 180.0#
        Coord.Latitude = LatRef / Math.PI * 180.0#
        'EllipsoidalHeight = XYRadius / Math.Cos(LatRef) - N
        Coord.EllipsoidalHeight = XYRadius / Math.Cos(LatRef) - N
    End Sub

    Private Sub Coord_ErrorMessage(Msg As String) Handles Coord.ErrorMessage
        RaiseEvent ErrorMessage(Msg)
    End Sub

    Private Sub Coord_Message(Msg As String) Handles Coord.Message
        RaiseEvent Message(Msg)
    End Sub

    Private Sub Coord_Update(Mode As Coordinate.UpdateMode, From As Coordinate.CoordType) Handles Coord.Update
        RaiseEvent Update(Mode, From)
    End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------



#Region " Events - Events that can be triggered by this class." '==========================================================================================================================
    Event ErrorMessage(ByVal Msg As String) 'Send an error message.
    Event Message(ByVal Msg As String) 'Send a normal message.
    Event Updated() 'Indicates the CRS has been updated.
    Event Update(ByVal Mode As Coordinate.UpdateMode, ByVal From As Coordinate.CoordType)
#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class 'CoordRefSystem


Public Class Extent
    'Properties and methods of an Extent.

#Region " Properties - All the properties used in this class." '===========================================================================================================================

    Private _name As String = ""  'Name of the extent, used only for extent look-up and not to be confused with extent description ('area of use').
    Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Private _code As Integer = -1 'Unique code (integer) of the extent; primary key.
    Property Code As Integer
        Get
            Return _code
        End Get
        Set(value As Integer)
            _code = value
        End Set
    End Property

    Private _description As String = ""  'Description of the extent. Limited to 4000 characters.
    Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property

    Private _southBoundLat As Double = 0 'The southern latitude of a bounding arc-rectangle, in decimal degrees positive north referenced to WGS 84. -90<value<90. Should be less than northern latitude.
    Property SouthBoundLat As Double
        Get
            Return _southBoundLat
        End Get
        Set(value As Double)
            _southBoundLat = value
        End Set
    End Property

    Private _westBoundLon As Double = 0 'Longitude (WGS 84) of the left side of a bounding arc-rectangle in decimal degrees positive east of Greenwich. -180<value<180. Generally should be less than the right longitude but for areas crossing the 180 degree meridian the left value will be greater.
    Property WestBoundLon As Double
        Get
            Return _westBoundLon
        End Get
        Set(value As Double)
            _westBoundLon = value
        End Set
    End Property

    Private _northBoundLat As Double = 0 'The northern latitude of a bounding arc-rectangle, in decimal degrees positive north referenced to WGS 84. -90<value<90. Should be greater than southern latitude.
    Property NorthBoundLat As Double
        Get
            Return _northBoundLat
        End Get
        Set(value As Double)
            _northBoundLat = value
        End Set
    End Property

    Private _eastBoundLon As Double = 0 'The longitude (WGS 84) of the right side of a bounding arc-rectangle in decimal degrees positive east of Greenwich. -180<value<180. Generally should be greater than the left longitude but for areas crossing the 180 degree meridian the value will be less.
    Property EastBoundLon As Double
        Get
            Return _eastBoundLon
        End Get
        Set(value As Double)
            _eastBoundLon = value
        End Set
    End Property

    Private _isoA2Code As String = "" 'ISO 3166 2-digit alpha country code
    Property IsoA2Code As String
        Get
            Return _isoA2Code
        End Get
        Set(value As String)
            _isoA2Code = value
        End Set
    End Property

    Private _isoA3Code As String = "" 'ISO 3166 3-digit alpha country code
    Property IsoA3Code As String
        Get
            Return _isoA3Code
        End Get
        Set(value As String)
            _isoA3Code = value
        End Set
    End Property

    Private _isoNCode As Integer = -1 'ISO 3166 3-digit numeric country code
    Property IsoNCode As Integer
        Get
            Return _isoNCode
        End Get
        Set(value As Integer)
            _isoNCode = value
        End Set
    End Property

    Private _vertExtentMin As Double = 0 'The minimum vertical value, in metres positive referenced to the vertical extent CRS.
    Property VertExtentMin As Double
        Get
            Return _vertExtentMin
        End Get
        Set(value As Double)
            _vertExtentMin = value
        End Set
    End Property

    Private _vertExtentMax As Double = 0 'The maximum vertical value, in metres positive referenced to the vertical extent CRS.
    Property VertExtentMax As Double
        Get
            Return _vertExtentMax
        End Get
        Set(value As Double)
            _vertExtentMax = value
        End Set
    End Property

    Private _vertExtentCrsCode As Integer 'The code of the vertical CRS to which the vertical extent values are referenced. Usually a height system but may be a depth system. If either vertical_extent_minimum or vertical_extent_maximum is given, this field is mandatory.
    Property VertExtentCrsCode As Integer
        Get
            Return _vertExtentCrsCode
        End Get
        Set(value As Integer)
            _vertExtentCrsCode = value
        End Set
    End Property

    Private _temporalExtentBegin As String 'The instant when the time period starts. Must be less than (earlier than) the end time. For date-time use format yyyy-mm-ddThh:mm:ss.sZ to required precision.
    Property TemporalExtentBegin As String
        Get
            Return _temporalExtentBegin
        End Get
        Set(value As String)
            _temporalExtentBegin = value
        End Set
    End Property

    Private _temporalExtentEnd As String 'The instant when the time period ends. Must be greater than (after) the begin time. For date-time use format yyyy-mm-ddThh:mm:ss.sZ to required precision.
    Property TemporalExtentEnd As String
        Get
            Return _temporalExtentEnd
        End Get
        Set(value As String)
            _temporalExtentEnd = value
        End Set
    End Property

    Private _remarks As String = "" '
    Property Remarks As String
        Get
            Return _remarks
        End Get
        Set(value As String)
            _remarks = value
        End Set
    End Property

    Private _infoSource As String = "" 'Reference describing the origin of the information to populate this record; e.g. an authoritative publication.
    Property InfoSource As String
        Get
            Return _infoSource
        End Get
        Set(value As String)
            _infoSource = value
        End Set
    End Property

    Private _dataSource As String = "" 'The organisation, body or person who populated this record; for EPSG supplied reference data: "EPSG".
    Property DataSource As String
        Get
            Return _dataSource
        End Get
        Set(value As String)
            _dataSource = value
        End Set
    End Property

    Private _revisionDate As Date = Date.MinValue 'Field used to record the date of creation or modification of this record. Not used if record is deprecated - see deprecation date field.
    'Private _revisionDate As String = "" 'Field used to record the date of creation or modification of this record. Not used if record is deprecated - see deprecation date field.
    Property RevisionDate As Date
        Get
            Return _revisionDate
        End Get
        Set(value As Date)
            _revisionDate = value
        End Set
    End Property

    Private _changeID As String = "" 'Unlinked reference to change table.
    Property ChangeID As String
        Get
            Return _changeID
        End Get
        Set(value As String)
            _changeID = value
        End Set
    End Property

    Private _deprecated As Boolean '"Yes" = data is deprecated; "No" =  data is current and valid.  Default is No.
    Property Deprecated As Boolean
        Get
            Return _deprecated
        End Get
        Set(value As Boolean)
            _deprecated = value
        End Set
    End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Public Sub Clear()
        'Clear the properties.
        Name = ""
        Code = -1
        Description = ""
        SouthBoundLat = Double.NaN
        WestBoundLon = Double.NaN
        NorthBoundLat = Double.NaN
        EastBoundLon = Double.NaN
        IsoA2Code = ""
        IsoA3Code = ""
        IsoNCode = -1
        VertExtentMin = Double.NaN
        VertExtentMax = Double.NaN
        VertExtentCrsCode = -1
        TemporalExtentBegin = ""
        TemporalExtentEnd = ""
        Remarks = ""
        InfoSource = ""
        DataSource = ""
        'RevisionDate = "1-Jan-1900 12:00:00"
        RevisionDate = Date.MinValue
        'RevisionDate = ""
        ChangeID = ""
        Deprecated = False
    End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================

#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class 'Extent

Public Class Usage
    'Properties and methods of a Usage.


#Region " Properties - All the properties used in this class." '===========================================================================================================================

    Private _code As Integer = -1 'Unique code (integer) of the Usage.
    Property Code As Integer
        Get
            Return _code
        End Get
        Set(value As Integer)
            _code = value
        End Set
    End Property

    Private _objectTableName As String = ""
    Property ObjectTableName As String
        Get
            Return _objectTableName
        End Get
        Set(value As String)
            _objectTableName = value
        End Set
    End Property

    Private _objectCode As Integer = -1
    Property ObjectCode As Integer
        Get
            Return _objectCode
        End Get
        Set(value As Integer)
            _objectCode = value
        End Set
    End Property

    Private _extentCode As Integer = -1
    Property ExtentCode As Integer
        Get
            Return _extentCode
        End Get
        Set(value As Integer)
            _extentCode = value
        End Set
    End Property

    Private _scopeCode As Integer = -1
    Property ScopeCode As Integer
        Get
            Return _scopeCode
        End Get
        Set(value As Integer)
            _scopeCode = value
        End Set
    End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Public Sub Clear()
        'Clear the properties
        Code = -1
        ObjectTableName = ""
        ObjectCode = -1
        ExtentCode = -1
        ScopeCode = -1
    End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================

#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class 'Usage

Public Class Scope
    'Properties and methods of a Scope.

#Region " Properties - All the properties used in this class." '===========================================================================================================================

    Private _code As Integer = -1 'Unique code (integer) of the Scope (purpose or use) of a CRS or coordinate operation; primary key.
    Property Code As Integer
        Get
            Return _code
        End Get
        Set(value As Integer)
            _code = value
        End Set
    End Property

    Private _scope As String = "" 'Description of the purpose(s) for which a CRS or Coordinate Operation is applied.
    Property Scope As String
        Get
            Return _scope
        End Get
        Set(value As String)
            _scope = value
        End Set
    End Property

    Private _remarks As String = "" '
    Property Remarks As String
        Get
            Return _remarks
        End Get
        Set(value As String)
            _remarks = value
        End Set
    End Property

    Private _infoSource As String = "" 'Reference describing the origin of the information to populate this record; e.g. an authoritative publication.
    Property InfoSource As String
        Get
            Return _infoSource
        End Get
        Set(value As String)
            _infoSource = value
        End Set
    End Property

    Private _dataSource As String = "" 'The organisation, body or person who populated this record; for EPSG supplied reference data: "EPSG".
    Property DataSource As String
        Get
            Return _dataSource
        End Get
        Set(value As String)
            _dataSource = value
        End Set
    End Property

    Private _revisionDate As Date = Date.MinValue 'Field used to record the date of creation or modification of this record. Not used if record is deprecated - see deprecation date field.
    'Private _revisionDate As String = "" 'Field used to record the date of creation or modification of this record. Not used if record is deprecated - see deprecation date field.
    Property RevisionDate As Date
        Get
            Return _revisionDate
        End Get
        Set(value As Date)
            _revisionDate = value
        End Set
    End Property

    Private _changeID As String = "" 'Unlinked reference to change table.
    Property ChangeID As String
        Get
            Return _changeID
        End Get
        Set(value As String)
            _changeID = value
        End Set
    End Property

    Private _deprecated As Boolean '"Yes" = data is deprecated; "No" =  data is current and valid.  Default is No.
    Property Deprecated As Boolean
        Get
            Return _deprecated
        End Get
        Set(value As Boolean)
            _deprecated = value
        End Set
    End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Public Sub Clear()
        'Clear the properties
        Code = -1
        Scope = ""
        Remarks = ""
        InfoSource = ""
        DataSource = ""
        RevisionDate = Date.MinValue
        ChangeID = ""
        Deprecated = False
    End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================

#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class 'Scope

Public Class DefiningOperation
    'Properties of a Defining Operation of a CRS

#Region " Properties - All the properties used in this class." '===========================================================================================================================

    Private _crsCode As Integer 'Foreign key to CRS table; points to the coordinate reference that is defined by the coordinate operation.

    Property CrsCode As Integer
        Get
            Return _crsCode
        End Get
        Set(value As Integer)
            _crsCode = value
        End Set
    End Property

    Private _ctCode As Integer 'Foreign key to COORD_OP table; points to the coordinate transformation that defines the CRS.

    Property CTCode As Integer
        Get
            Return _ctCode
        End Get
        Set(value As Integer)
            _ctCode = value
        End Set
    End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Public Sub Clear()
        'Clear the properties
        CrsCode = -1
        CTCode = -1
    End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================

#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class 'DefiningOperation

Public Class CoordinateOperation
    'Properties of a Coordinate Operation.

#Region " Properties - All the properties used in this class." '===========================================================================================================================

    Private _name As String = ""  'Unique name of the coordinate operation
    Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Private _code As Integer = -1 'Unique code (integer) of the coordinate operation, which is unique over this table; primary key
    Property Code As Integer
        Get
            Return _code
        End Get
        Set(value As Integer)
            _code = value
        End Set
    End Property

    Enum OperationType
        conversion
        transformation
        pointMotionOperation
        concatenatedOperation
    End Enum

    Private _type As OperationType = OperationType.conversion  'The type of coordinate operation: "conversion", "transformation", "point motion operation" or "concatenated operation". A Map Projection is a "Conversion". Concatenated operations consist of several single operations in a prescribed order.
    Property Type As OperationType
        Get
            Return _type
        End Get
        Set(value As OperationType)
            _type = value
        End Set
    End Property

    Private _methodCode As Integer = -1 'Foreign key to coordinate operation method
    Property MethodCode As Integer
        Get
            Return _methodCode
        End Get
        Set(value As Integer)
            _methodCode = value
        End Set
    End Property

    Private _remarks As String = "" '
    Property Remarks As String
        Get
            Return _remarks
        End Get
        Set(value As String)
            _remarks = value
        End Set
    End Property

    Private _infoSource As String = "" 'Reference describing the origin of the information to populate this record; e.g. an authoritative publication
    Property InfoSource As String
        Get
            Return _infoSource
        End Get
        Set(value As String)
            _infoSource = value
        End Set
    End Property

    Private _dataSource As String = "" 'The organisation, body or person who populated this record; for EPSG supplied reference data: "EPSG".
    Property DataSource As String
        Get
            Return _dataSource
        End Get
        Set(value As String)
            _dataSource = value
        End Set
    End Property

    Private _revisionDate As Date = Date.MinValue 'Field used to record the date of creation or modification of this record. Not used if record is deprecated - see deprecation date field.
    'Private _revisionDate As String = "" 'Field used to record the date of creation or modification of this record. Not used if record is deprecated - see deprecation date field.
    Property RevisionDate As Date
        Get
            Return _revisionDate
        End Get
        Set(value As Date)
            _revisionDate = value
        End Set
    End Property

    Private _changeID As String = "" 'Unlinked reference to change table.
    Property ChangeID As String
        Get
            Return _changeID
        End Get
        Set(value As String)
            _changeID = value
        End Set
    End Property

    Private _deprecated As Boolean '"Yes" = data is deprecated; "No" =  data is current and valid.  Default is No.
    Property Deprecated As Boolean
        Get
            Return _deprecated
        End Get
        Set(value As Boolean)
            _deprecated = value
        End Set
    End Property

    Private _sourceCrsCode As Integer = -1 'Source (or input) coordinate reference system for an operation; system to be transformed or converted. Mandatory for transformations and concatenated operations, not required for conversions.
    Property SourceCrsCode As Integer
        Get
            Return _sourceCrsCode
        End Get
        Set(value As Integer)
            _sourceCrsCode = value
        End Set
    End Property

    Private _targetCrsCode As Integer = -1 'Target (or output) coordinate reference system for  an operation;  system for end result. Mandatory for transformations and concatenated operations, not required for conversions.
    Property TargetCrsCode As Integer
        Get
            Return _targetCrsCode
        End Get
        Set(value As Integer)
            _targetCrsCode = value
        End Set
    End Property

    Property _version As String = "" 'The version of the  transformation between these source and target coordinate reference systems.  Not required for conversions. For  transformations (single or concatenated) may act as a secondary triple key with source and target coordinate ref systems.
    Property Version As String
        Get
            Return _version
        End Get
        Set(value As String)
            _version = value
        End Set
    End Property

    Private _opVariant As Short 'The counter for the transformation between this source and this target coordinate systems.  Not required for conversions.  In EPSG prior to v5.0 acted as the version identifier.  Retained only for purposes of backward compatibility.
    Property OpVariant As Short
        Get
            Return _opVariant
        End Get
        Set(value As Short)
            _opVariant = value
        End Set
    End Property

    Private _accuracy As Single = -1 'An indicative number indicating the loss of accuracy in metres that applying the transformation might bring to target coordinates. For conversions, which are considered exact by definition, the value is 0.
    Property Accuracy As Single
        Get
            Return _accuracy
        End Get
        Set(value As Single)
            _accuracy = value
        End Set
    End Property

    Private _uomSourceCoordDiffCode As Integer = -1 'Unit of measure of the input or source coordinate differences in a polynomial operation.  Often different from the UOM of the coordinate reference system.
    Property UomSourceCoordDiffCode As Integer
        Get
            Return _uomSourceCoordDiffCode
        End Get
        Set(value As Integer)
            _uomSourceCoordDiffCode = value
        End Set
    End Property

    Private _uomTargetCoordDiffCode As Integer = -1 'Unit of measure of the output or target coordinate differences in a polynomial operation.  Often different from the UOM of the coordinate reference system.
    Property UomTargetCoordDiffCode As Integer
        Get
            Return _uomTargetCoordDiffCode
        End Get
        Set(value As Integer)
            _uomTargetCoordDiffCode = value
        End Set
    End Property

    Private _show As Boolean = True 'Switch to indicate whether operation data can be made public.  "Yes" or "No". Default is Yes.
    Property Show As Boolean
        Get
            Return _show
        End Get
        Set(value As Boolean)
            _show = value
        End Set
    End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Public Sub Clear()
        'Clear the properties
        Name = ""
        Code = -1
        Type = OperationType.conversion
        MethodCode = -1
        Remarks = ""
        InfoSource = ""
        DataSource = ""
        RevisionDate = Date.MinValue
        ChangeID = ""
        Deprecated = False
        SourceCrsCode = -1
        TargetCrsCode = -1
        Version = -1
        OpVariant = -1
        Accuracy = -1
        UomSourceCoordDiffCode = -1
        UomTargetCoordDiffCode = -1
        Show = True
    End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================

#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class 'CoordinateOperation

Public Class CoordOperationPath
    'Properties of a Coordinate Operation Path for a multi step operation

#Region " Properties - All the properties used in this class." '===========================================================================================================================

    Private _concatOpCode As Integer 'Foreign key to COORD_OP table; points to the concatenated coordinate operation that contains one or more single operation steps.
    Property ConcatOpCode As Integer
        Get
            Return _concatOpCode
        End Get
        Set(value As Integer)
            _concatOpCode = value
        End Set
    End Property

    Private _singleOpCode As Integer 'Foreign key to COORD_OP table; points to the single coord operation that is contained as a step in a concatenated operation.
    Property SingleOpCode As Integer
        Get
            Return _singleOpCode
        End Get
        Set(value As Integer)
            _singleOpCode = value
        End Set
    End Property

    Private _opPathStep As Short 'The sequence number of this step within this concatenated operation.
    Property OpPathStep As Short
        Get
            Return _opPathStep
        End Get
        Set(value As Short)
            _opPathStep = value
        End Set
    End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Public Sub Clear()
        'Clear the properties
        ConcatOpCode = -1
        SingleOpCode = -1
        OpPathStep = -1
    End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================

#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class 'CoordOperationPath

Public Class CoordOpParamValue
    'Properties of a Coordinate Operation Parameter Value

#Region " Properties - All the properties used in this class." '===========================================================================================================================

    Private _opCode As Integer = -1 'Code (integer) of the coordinate operation, part of concatenated (triple) primary key.
    Property OpCode As Integer
        Get
            Return _opCode
        End Get
        Set(value As Integer)
            _opCode = value
        End Set
    End Property

    Private _methodCode As Integer = -1 'Unique code (integer) of the coordinate operation; part of concatenated (triple) primary key.
    Property MethodCode As Integer
        Get
            Return _methodCode
        End Get
        Set(value As Integer)
            _methodCode = value
        End Set
    End Property

    Private _parameterCode As Integer = -1 'Code (integer) of the reference coordinate operation parameter; part of concatenated (triple) primary key.
    Property ParameterCode As Integer
        Get
            Return _parameterCode
        End Get
        Set(value As Integer)
            _parameterCode = value
        End Set
    End Property

    Private _parameterValue As Double = 0 'Numeric value of the transformation parametery.
    Property ParameterValue As Double
        Get
            Return _parameterValue
        End Get
        Set(value As Double)
            _parameterValue = value
        End Set
    End Property

    Private _paramValueFileRef As String = "" 'File name and path, in cases when the transformation parameter is not a numeric value but points to a file, eg. NADCON transformations.
    Property ParamValueFileRef As String
        Get
            Return _paramValueFileRef
        End Get
        Set(value As String)
            _paramValueFileRef = value
        End Set
    End Property

    Private _uomCode As Integer = -1 'Foreign key to the name of the unit for the transformation parameter value.
    Property UomCode As Integer
        Get
            Return _uomCode
        End Get
        Set(value As Integer)
            _uomCode = value
        End Set
    End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Public Sub Clear()
        'Clear the properties
        OpCode = -1
        MethodCode = -1
        ParameterCode = -1
        ParameterValue = 0
        ParamValueFileRef = ""
        UomCode = -1
    End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================

#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class 'CoordOpParamValue

Public Class CoordOpParamUsage
    'Properties of a Coordinate Operation Parameter Usage

#Region " Properties - All the properties used in this class." '===========================================================================================================================

    Private _methodCode As Integer = -1 'Unique code (integer) of the coordinate operation method; together with "PARAMETER_CODE" primary key.
    Property MethodCode As Integer
        Get
            Return _methodCode
        End Get
        Set(value As Integer)
            _methodCode = value
        End Set
    End Property

    Private _parameterCode As Integer = -1 'EPSG assigned code (integer) of the reference coordinate operation  parameter; together with "COORD_OP_METHOD_CODE" primary key.
    Property ParameterCode As Integer
        Get
            Return _parameterCode
        End Get
        Set(value As Integer)
            _parameterCode = value
        End Set
    End Property

    Private _sortOrder As Short = -1 'The sequence number indicating the order in which the parameters are shown.
    Property SortOrder As Short
        Get
            Return _sortOrder
        End Get
        Set(value As Short)
            _sortOrder = value
        End Set
    End Property

    Private _signReversal As String = "" 'Indicates if the sign of the parameters should be reversed in the reverse operation; "Yes" or "No".  Only valid if field Coord Operation Method.REVERSE_OP = "Yes"; if that field ="No" leave this field blank.
    Property SignReversal As String
        Get
            Return _signReversal
        End Get
        Set(value As String)
            _signReversal = value
        End Set
    End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Public Sub Clear()
        'Clear the properties
        MethodCode = -1
        ParameterCode = -1
        SortOrder = -1
        SignReversal = ""
    End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================

#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class 'CoordOpParamUsage

Public Class CoordOpParameter
    'Properties of a Coordinate Operation Parameter

#Region " Properties - All the properties used in this class." '===========================================================================================================================

    Private _name As String = ""  'Name of the parameter
    Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Private _code As Integer = -1 'Code (integer) of the parameter, which is unique over this table; primary key
    Property Code As Integer
        Get
            Return _code
        End Get
        Set(value As Integer)
            _code = value
        End Set
    End Property

    Private _description As String = "" 'A decription of the parameter. Limited to 4000 characters.
    Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property

    Private _infoSource As String = "" 'Reference describing the origin of the information to populate this record; e.g. an authoritative publication
    Property InfoSource As String
        Get
            Return _infoSource
        End Get
        Set(value As String)
            _infoSource = value
        End Set
    End Property

    Private _dataSource As String = "" 'The organisation, body or person who populated this record; for EPSG supplied reference data: "EPSG".
    Property DataSource As String
        Get
            Return _dataSource
        End Get
        Set(value As String)
            _dataSource = value
        End Set
    End Property

    Private _revisionDate As Date = Date.MinValue 'Field used to record the date of creation or modification of this record. Not used if record is deprecated - see deprecation date field.
    Property RevisionDate As Date
        Get
            Return _revisionDate
        End Get
        Set(value As Date)
            _revisionDate = value
        End Set
    End Property

    Private _changeID As String = "" 'Unlinked reference to change table.
    Property ChangeID As String
        Get
            Return _changeID
        End Get
        Set(value As String)
            _changeID = value
        End Set
    End Property

    Private _deprecated As Boolean '"Yes" = data is deprecated; "No" =  data is current and valid.  Default is No.
    Property Deprecated As Boolean
        Get
            Return _deprecated
        End Get
        Set(value As Boolean)
            _deprecated = value
        End Set
    End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Public Sub Clear()
        'Clear the properties
        Name = ""
        Code = -1
        Description = ""
        InfoSource = ""
        DataSource = ""
        RevisionDate = Date.MinValue
        ChangeID = ""
        Deprecated = False
    End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================

#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class 'CoordOpParameter

Public Class UnitOfMeasure
    'Properties of a Unit of Measure.

#Region " Properties - All the properties used in this class." '===========================================================================================================================

    Private _name As String = ""  'Name of the unit of measure
    Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Private _code As Integer = -1 'Unique code (integer) of the unit of measure; primary key
    Property Code As Integer
        Get
            Return _code
        End Get
        Set(value As Integer)
            _code = value
        End Set
    End Property

    Enum UOMType
        Length
        Angle
        Scale
        Time
    End Enum

    Private _type As UOMType = UOMType.Length  'The type of Unit of Measure: Length, Angle and Scale are the only types allowed in the EPSG database
    Property Type As UOMType
        Get
            Return _type
        End Get
        Set(value As UOMType)
            _type = value
        End Set
    End Property

    Private _targetUomCode As Integer 'Other UOM of the same type into which the current UOM can be converted using the formula (POSC); POSC factors A and D always equal zero for EPSG supplied units of measure
    Property TargetUomCode As Integer
        Get
            Return _targetUomCode
        End Get
        Set(value As Integer)
            _targetUomCode = value
        End Set
    End Property

    Private _factorB As Double = 1 'A quantity in the target UOM (y) is obtained from a quantity in the current UOM (x) through the conversion: y =  (B/C).x
    Property FactorB As Double
        Get
            Return _factorB
        End Get
        Set(value As Double)
            _factorB = value
        End Set
    End Property

    Private _factorC As Double = 1 'A quantity in the target UOM (y) is obtained from a quantity in the current UOM (x) through the conversion: y =  (B/C).x
    Property FactorC As Double
        Get
            Return _factorC
        End Get
        Set(value As Double)
            _factorC = value
        End Set
    End Property

    Private _remarks As String = "" '
    Property Remarks As String
        Get
            Return _remarks
        End Get
        Set(value As String)
            _remarks = value
        End Set
    End Property

    Private _infoSource As String = "" 'Reference describing the origin of the information to populate this record; e.g. an authoritative publication
    Property InfoSource As String
        Get
            Return _infoSource
        End Get
        Set(value As String)
            _infoSource = value
        End Set
    End Property

    Private _dataSource As String = "" 'The organisation, body or person who populated this record; for EPSG supplied reference data: "EPSG".
    Property DataSource As String
        Get
            Return _dataSource
        End Get
        Set(value As String)
            _dataSource = value
        End Set
    End Property

    Private _revisionDate As Date = Date.MinValue 'Field used to record the date of creation or modification of this record. Not used if record is deprecated - see deprecation date field.
    Property RevisionDate As Date
        Get
            Return _revisionDate
        End Get
        Set(value As Date)
            _revisionDate = value
        End Set
    End Property

    Private _changeID As String = "" 'Unlinked reference to change table.
    Property ChangeID As String
        Get
            Return _changeID
        End Get
        Set(value As String)
            _changeID = value
        End Set
    End Property

    Private _deprecated As Boolean '"Yes" = data is deprecated; "No" =  data is current and valid.  Default is No.
    Property Deprecated As Boolean
        Get
            Return _deprecated
        End Get
        Set(value As Boolean)
            _deprecated = value
        End Set
    End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Public Sub Clear()
        'Clear the properties
        Name = ""
        Code = -1
        Type = UOMType.Length
        TargetUomCode = -1
        FactorB = 1
        FactorC = 1
        Remarks = ""
        InfoSource = -1
        DataSource = -1
        RevisionDate = Date.MinValue
        ChangeID = -1
        Deprecated = False
    End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================

#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class 'UnitOfMeasure

Public Class CoordOpMethod
    'Properties of a Coordinate Operation Method.

#Region " Properties - All the properties used in this class." '===========================================================================================================================

    Private _name As String = ""  'UName of the coordinate transformation
    Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Private _code As Integer = -1 'Unique code (integer) of the coordinate transformation; primary key
    Property Code As Integer
        Get
            Return _code
        End Get
        Set(value As Integer)
            _code = value
        End Set
    End Property

    Private _reverseOp As Boolean = False 'Indication of the validity of the transformation parameters for the reverse operation; if "No", search for an explicit definition of the reverse transformation.
    Property ReverseOp As Boolean
        Get
            Return _reverseOp
        End Get
        Set(value As Boolean)
            _reverseOp = value
        End Set
    End Property

    Private _formula As String = "" 'The formulas associated with this method or algorithm.  Limited to 4000 characters.
    Property Formula As String
        Get
            Return _formula
        End Get
        Set(value As String)
            _formula = value
        End Set
    End Property

    Private _example As String = "" 'Worked example of this transformation method. Limited to 4000 characters.
    Property Example As String
        Get
            Return _example
        End Get
        Set(value As String)
            _example = value
        End Set
    End Property

    Private _remarks As String = "" '
    Property Remarks As String
        Get
            Return _remarks
        End Get
        Set(value As String)
            _remarks = value
        End Set
    End Property

    Private _infoSource As String = "" 'Reference describing the origin of the information to populate this record; e.g. an authoritative publication
    Property InfoSource As String
        Get
            Return _infoSource
        End Get
        Set(value As String)
            _infoSource = value
        End Set
    End Property

    Private _dataSource As String = "" 'The organisation, body or person who populated this record; for EPSG supplied reference data: "EPSG".
    Property DataSource As String
        Get
            Return _dataSource
        End Get
        Set(value As String)
            _dataSource = value
        End Set
    End Property

    Private _revisionDate As Date = Date.MinValue 'Field used to record the date of creation or modification of this record. Not used if record is deprecated - see deprecation date field.
    Property RevisionDate As Date
        Get
            Return _revisionDate
        End Get
        Set(value As Date)
            _revisionDate = value
        End Set
    End Property

    Private _changeID As String = "" 'Unlinked reference to change table.
    Property ChangeID As String
        Get
            Return _changeID
        End Get
        Set(value As String)
            _changeID = value
        End Set
    End Property

    Private _deprecated As Boolean '"Yes" = data is deprecated; "No" =  data is current and valid.  Default is No.
    Property Deprecated As Boolean
        Get
            Return _deprecated
        End Get
        Set(value As Boolean)
            _deprecated = value
        End Set
    End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Public Sub Clear()
        'Clear the properties
        Name = ""
        Code = -1
        ReverseOp = False
        Formula = ""
        Example = ""
        Remarks = ""
        InfoSource = -1
        DataSource = -1
        RevisionDate = Date.MinValue
        ChangeID = -1
        Deprecated = False
    End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================

#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class 'CoordOpMethod

Public Class CoordSystem
    'Properties of a Coordinate System.

#Region " Properties - All the properties used in this class." '===========================================================================================================================


    Private _name As String = ""  'A description of the coordinate system type, axes, axis orientations and units
    Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Private _code As Integer = -1 'Unique code (integer) of the Coordinate System (CS); primary key.
    Property Code As Integer
        Get
            Return _code
        End Get
        Set(value As Integer)
            _code = value
        End Set
    End Property

    Enum CsType
        affine
        Cartesian
        cylindrical
        ellipsoidal
        linear
        ordinal
        polar
        spherical
        vertical
        parametric
        temporalCount
        temporalDatetime
        temporalMeasure
    End Enum

    Private _type As CsType = CsType.Cartesian  'Type of the CS: "affine", "Cartesian", "cylindrical", "ellipsoidal", "linear", "ordinal", "polar", "spherical", "vertical", "parametric", "temporal count", "temporal datetime", "temporal measure".
    Property Type As CsType
        Get
            Return _type
        End Get
        Set(value As CsType)
            _type = value
        End Set
    End Property

    Property _dimension As Short = 2 'The number of dimensions of the CS: 1, 2 or 3.
    Property Dimension As Short
        Get
            Return _dimension
        End Get
        Set(value As Short)
            _dimension = value
        End Set
    End Property

    Private _remarks As String = "" '
    Property Remarks As String
        Get
            Return _remarks
        End Get
        Set(value As String)
            _remarks = value
        End Set
    End Property

    Private _infoSource As String = "" 'Reference describing the origin of the information to populate this record; e.g. an authoritative publication.
    Property InfoSource As String
        Get
            Return _infoSource
        End Get
        Set(value As String)
            _infoSource = value
        End Set
    End Property

    Private _dataSource As String = "" 'The organisation, body or person who populated this record; for EPSG supplied reference data: "EPSG".
    Property DataSource As String
        Get
            Return _dataSource
        End Get
        Set(value As String)
            _dataSource = value
        End Set
    End Property

    Private _revisionDate As Date = Date.MinValue 'Field used to record the date of creation or modification of this record. Not used if record is deprecated - see deprecation date field.
    Property RevisionDate As Date
        Get
            Return _revisionDate
        End Get
        Set(value As Date)
            _revisionDate = value
        End Set
    End Property

    Private _changeID As String = "" 'Unlinked reference to change table.
    Property ChangeID As String
        Get
            Return _changeID
        End Get
        Set(value As String)
            _changeID = value
        End Set
    End Property

    Private _deprecated As Boolean '"Yes" = data is deprecated; "No" =  data is current and valid.  Default is No.
    Property Deprecated As Boolean
        Get
            Return _deprecated
        End Get
        Set(value As Boolean)
            _deprecated = value
        End Set
    End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Public Sub Clear()
        'Clear the properties
        Name = ""
        Code = -1
        Type = CsType.Cartesian
        Dimension = 2
        Remarks = ""
        InfoSource = -1
        DataSource = -1
        RevisionDate = Date.MinValue
        ChangeID = -1
        Deprecated = False
    End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================

#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class 'CoordSystem

Public Class CoordAxis
    'Propertries of a Coordinate Axis.

#Region " Properties - All the properties used in this class." '===========================================================================================================================


    Private _code As Integer = -1 'Unique code for records in Coordinate Axis table. Not required for EPSG relational implementation, but provided to assist other implementations.
    Property Code As Integer
        Get
            Return _code
        End Get
        Set(value As Integer)
            _code = value
        End Set
    End Property

    Private _systemCode As Integer = -1 'Foreign key to the Coordinate System this axis is used in. A coordinate system uses a unique set of coordinate axes.
    Property SystemCode As Integer
        Get
            Return _systemCode
        End Get
        Set(value As Integer)
            _systemCode = value
        End Set
    End Property

    Private _nameCode As Integer = -1 'Foreign key to Coordinate Axis Name table; part of dual primary key.
    Property NameCode As Integer
        Get
            Return _nameCode
        End Get
        Set(value As Integer)
            _nameCode = value
        End Set
    End Property

    Private _orientation As String = "" 'The direction of the POSITIVE increments of the of the coordinate axis: north, east, south, west, up, down or the APPROXIMATE orientation in case of an oblique orientation: ~NE, ~SE, ~SW, ~NW.
    Property Orientation As String
        Get
            Return _orientation
        End Get
        Set(value As String)
            _orientation = value
        End Set
    End Property

    Private _abbreviation As String = "" 'Abbreviation for the coordinate axis.
    Property Abbreviation As String
        Get
            Return _abbreviation
        End Get
        Set(value As String)
            _abbreviation = value
        End Set
    End Property

    Private _uomCode As Integer = -1 'ID of the unit measure of the coordinate axis.
    Property UomCode As Integer
        Get
            Return _uomCode
        End Get
        Set(value As Integer)
            _uomCode = value
        End Set
    End Property

    Private _order As Short = 1 'The position of this axis within a Coordinate System: 1, 2 or 3.
    Property Order As Short
        Get
            Return _order
        End Get
        Set(value As Short)
            _order = value
        End Set
    End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Public Sub Clear()
        'Clear the properties
        Code = -1
        SystemCode = -1
        NameCode = -1
        Orientation = ""
        Abbreviation = ""
        UomCode = -1
        Order = 1
    End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================

#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class 'CoordAxis

Public Class CoordAxisName
    'Properties of a Coordinate Axis name.

#Region " Properties - All the properties used in this class." '===========================================================================================================================

    Private _name As String = ""  'Unique name for the coordinate axis, e.g. "latitude"; reference coordinate axes are re-utilized by multiple coordinate systems.
    Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Private _code As Integer = -1 'Unique code for the coordinate axis name.
    Property Code As Integer
        Get
            Return _code
        End Get
        Set(value As Integer)
            _code = value
        End Set
    End Property

    Private _description As String = "" '
    Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property

    Private _remarks As String = "" '
    Property Remarks As String
        Get
            Return _remarks
        End Get
        Set(value As String)
            _remarks = value
        End Set
    End Property

    Private _infoSource As String = "" 'Reference describing the origin of the information to populate this record; e.g. an authoritative publication
    Property InfoSource As String
        Get
            Return _infoSource
        End Get
        Set(value As String)
            _infoSource = value
        End Set
    End Property

    Private _dataSource As String = "" 'The organisation, body or person who populated this record; for EPSG supplied reference data: "EPSG".
    Property DataSource As String
        Get
            Return _dataSource
        End Get
        Set(value As String)
            _dataSource = value
        End Set
    End Property

    Private _revisionDate As Date = Date.MinValue 'Field used to record the date of creation or modification of this record. Not used if record is deprecated - see deprecation date field.
    Property RevisionDate As Date
        Get
            Return _revisionDate
        End Get
        Set(value As Date)
            _revisionDate = value
        End Set
    End Property

    Private _changeID As String = "" 'Unlinked reference to change table.
    Property ChangeID As String
        Get
            Return _changeID
        End Get
        Set(value As String)
            _changeID = value
        End Set
    End Property

    Private _deprecated As Boolean '"Yes" = data is deprecated; "No" =  data is current and valid.  Default is No.
    Property Deprecated As Boolean
        Get
            Return _deprecated
        End Get
        Set(value As Boolean)
            _deprecated = value
        End Set
    End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Public Sub Clear()
        'Clear the properties
        Name = ""
        Code = -1
        Description = ""
        Remarks = ""
        InfoSource = -1
        DataSource = -1
        RevisionDate = Date.MinValue
        ChangeID = -1
        Deprecated = False
    End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================

#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class 'CoordAxisName

Public Class Datum
    'Properties of a Datum.

#Region " Properties - All the properties used in this class." '===========================================================================================================================

    Private _name As String = ""  'Unique name of the geodetic datum
    Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Private _code As Integer = -1 'Unique code (integer) of the geodetic datum; primary key
    Property Code As Integer
        Get
            Return _code
        End Get
        Set(value As Integer)
            _code = value
        End Set
    End Property

    Enum DatumType
        geodetic
        dynamicGeodetic
        vertical
        dynamicVertical
        engineering
        ensemble
    End Enum

    Private _type As DatumType = DatumType.geodetic  'The type of datum: "geodetic", "dynamic geodetic","vertical"; "dynamic vertical"; "engineering" or "ensemble".
    Property Type As DatumType
        Get
            Return _type
        End Get
        Set(value As DatumType)
            _type = value
        End Set
    End Property

    Private _remarks As String = "" '
    Property Remarks As String
        Get
            Return _remarks
        End Get
        Set(value As String)
            _remarks = value
        End Set
    End Property

    Private _infoSource As String = "" 'Reference describing the origin of the information to populate this record; e.g. an authoritative publication
    Property InfoSource As String
        Get
            Return _infoSource
        End Get
        Set(value As String)
            _infoSource = value
        End Set
    End Property

    Private _dataSource As String = "" 'The organisation, body or person who populated this record; for EPSG supplied reference data: "EPSG".
    Property DataSource As String
        Get
            Return _dataSource
        End Get
        Set(value As String)
            _dataSource = value
        End Set
    End Property

    Private _revisionDate As Date = Date.MinValue 'Field used to record the date of creation or modification of this record. Not used if record is deprecated - see deprecation date field.
    'Private _revisionDate As String = "" 'Field used to record the date of creation or modification of this record. Not used if record is deprecated - see deprecation date field.
    'Property RevisionDate As Date
    Property RevisionDate As Date
        Get
            Return _revisionDate
        End Get
        Set(value As Date)
            _revisionDate = value
        End Set
    End Property

    Private _changeID As String = "" 'Unlinked reference to change table.
    Property ChangeID As String
        Get
            Return _changeID
        End Get
        Set(value As String)
            _changeID = value
        End Set
    End Property

    Private _deprecated As Boolean '"Yes" = data is deprecated; "No" =  data is current and valid.  Default is No.
    Property Deprecated As Boolean
        Get
            Return _deprecated
        End Get
        Set(value As Boolean)
            _deprecated = value
        End Set
    End Property

    Private _originDescr As String = "" 'A description of the anchor point, origin or datum definition. In ISO19111 called "anchor definition".
    Property OriginDescr As String
        Get
            Return _originDescr
        End Get
        Set(value As String)
            _originDescr = value
        End Set
    End Property

    Private _ellipsoidCode As Integer 'Ellipsoid used in the definition of a datum where DATUM_TYPE="geodetic"
    Property EllipsoidCode As Integer
        Get
            Return _ellipsoidCode
        End Get
        Set(value As Integer)
            _ellipsoidCode = value
        End Set
    End Property

    Private _primeMeridianCode As Integer = -1 'Prime Meridian used in the definition of a datum where DATUM_TYPE="geodetic"
    Property PrimeMeridianCode As Integer
        Get
            Return _primeMeridianCode
        End Get
        Set(value As Integer)
            _primeMeridianCode = value
        End Set
    End Property

    Private _conventionalRSCode As Integer = -1 'For datums that are members of datum ensembles, the code for the Conventional RS for that ensemble.
    Property ConventionalRSCode As Integer
        Get
            Return _conventionalRSCode
        End Get
        Set(value As Integer)
            _conventionalRSCode = value
        End Set
    End Property

    'Private _publicationDate As Date 'The date on which the adjustment or realization was published.
    Private _publicationDate As String = "" 'The date on which the adjustment or realization was published.
    Property PublicationDate As String
        Get
            Return _publicationDate
        End Get
        Set(value As String)
            _publicationDate = value
        End Set
    End Property

    Private _frameReferenceEpoch As Double 'For dynamic datums, the epoch which the coordinates of stations defining the dynamic datum (reference frame) are referenced, given as a decimal year in the Gregorian calendar, usually to two decimal places.
    Property FrameReferenceEpoch As Double
        Get
            Return _frameReferenceEpoch
        End Get
        Set(value As Double)
            _frameReferenceEpoch = value
        End Set
    End Property

    Private _realizationMethodCode As Integer 'For vertical datums, the code of the method by which the datum was promulgated.
    Property RealizationMethodCode As Integer
        Get
            Return _realizationMethodCode
        End Get
        Set(value As Integer)
            _realizationMethodCode = value
        End Set
    End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Public Sub Clear()
        'Clear the properties
        Name = ""
        Code = -1
        Type = DatumType.dynamicGeodetic
        Remarks = ""
        InfoSource = -1
        DataSource = -1
        RevisionDate = Date.MinValue
        'RevisionDate = ""
        ChangeID = -1
        Deprecated = False
        OriginDescr = ""
        EllipsoidCode = -1
        PrimeMeridianCode = -1
        ConventionalRSCode = -1
        'PublicationDate = Date.MinValue
        PublicationDate = ""
        FrameReferenceEpoch = Double.NaN
        RealizationMethodCode = -1
    End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================

#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class 'Datum

Public Class DatumEnsembleMember
    'Properties of a Datum Ensemble Member.

#Region " Properties - All the properties used in this class." '===========================================================================================================================

    Private _code As Integer = -1 'Foreign key to DATUM ENSEMBLE table; points to the ensemble that contains two or more datums (reference frames).
    Property Code As Integer
        Get
            Return _code
        End Get
        Set(value As Integer)
            _code = value
        End Set
    End Property

    Private _datumCode As Integer = -1 'Foreign key to DATUM table; points to the single datum (reference frame) that is contained in this ensemble.
    Property DatumCode As Integer
        Get
            Return _datumCode
        End Get
        Set(value As Integer)
            _datumCode = value
        End Set
    End Property

    Private _sequence As Short 'The sequence number of this datum (reference frame) within this ensemble.
    Property Sequence As Short
        Get
            Return _sequence
        End Get
        Set(value As Short)
            _sequence = value
        End Set
    End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Public Sub Clear()
        'Clear the properties
        Code = -1
        DatumCode = -1
        Sequence = -1
    End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================

#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class 'DatumEnsembleMember

Public Class DatumEnsemble
    'Properties of a Datum Ensemble

#Region " Properties - All the properties used in this class." '===========================================================================================================================

    Private _code As Integer = -1 'Foreign key to DATUM ENSEMBLE table; points to the ensemble that contains two or more datums (reference frames).
    Property Code As Integer
        Get
            Return _code
        End Get
        Set(value As Integer)
            _code = value
        End Set
    End Property

    Private _accuracy As Single 'An indicative number describing the loss of accuracy in metres that using the datum ensemble may bring to coordinates.
    Property Accuracy As Single
        Get
            Return _accuracy
        End Get
        Set(value As Single)
            _accuracy = value
        End Set
    End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Public Sub Clear()
        'Clear the properties
        Code = -1
        Accuracy = Single.NaN
    End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================

#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class 'DatumEnsemble

Public Class Ellipsoid
    'Properties of an Ellipsoid.

#Region " Properties - All the properties used in this class." '===========================================================================================================================

    Private _name As String = ""  'Name of the ellipsoid
    Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Private _code As Integer = -1 'Unique code (integer) of the ellipsoid; primary key.
    Property Code As Integer
        Get
            Return _code
        End Get
        Set(value As Integer)
            _code = value
        End Set
    End Property

    Private _semiMajorAxis As Double = Double.NaN 'Length of the semi-major axis of the ellipsoid
    Property SemiMajorAxis As Double
        Get
            Return _semiMajorAxis
        End Get
        Set(value As Double)
            _semiMajorAxis = value
        End Set
    End Property

    Private _uomCode As Integer = -1 'ID of the unit measure of the ellipsoid's axes, as defined in this record.
    Property UomCode As Integer
        Get
            Return _uomCode
        End Get
        Set(value As Integer)
            _uomCode = value
        End Set
    End Property

    Private _invFlattening As Double = Double.NaN 'Preferred second defining parameter of the ellipsoid
    Property InvFlattening As Double
        Get
            Return _invFlattening
        End Get
        Set(value As Double)
            _invFlattening = value
        End Set
    End Property

    Private _semiMinorAxis As Double = Double.NaN 'Alternative to inverse flattening; it is preferred not to supply both inverse flattening and semi-minor axis
    Property SemiMinorAxis As Double
        Get
            Return _semiMinorAxis
        End Get
        Set(value As Double)
            _semiMinorAxis = value
        End Set
    End Property

    Private _ellipsoidShape As Boolean = True 'Indicator of the shape of the ellipsoid: Yes = "Ellipsoid"; No = "Sphere".  Default is Yes. No equates to OGC second ellipsoid parameter being "isSphere" with the value of isSphere beinging "sphere".
    Property EllipsoidShape As Boolean
        Get
            Return _ellipsoidShape
        End Get
        Set(value As Boolean)
            _ellipsoidShape = value
        End Set
    End Property

    Private _remarks As String = "" '
    Property Remarks As String
        Get
            Return _remarks
        End Get
        Set(value As String)
            _remarks = value
        End Set
    End Property

    Private _infoSource As String = "" 'Reference describing the origin of the information to populate this record; e.g. an authoritative publication.
    Property InfoSource As String
        Get
            Return _infoSource
        End Get
        Set(value As String)
            _infoSource = value
        End Set
    End Property

    Private _dataSource As String = "" 'The organisation, body or person who populated this record; for EPSG supplied reference data: "EPSG".
    Property DataSource As String
        Get
            Return _dataSource
        End Get
        Set(value As String)
            _dataSource = value
        End Set
    End Property

    Private _revisionDate As Date = Date.MinValue 'Field used to record the date of creation or modification of this record. Not used if record is deprecated - see deprecation date field.
    'Private _revisionDate As String = "" 'Field used to record the date of creation or modification of this record. Not used if record is deprecated - see deprecation date field.
    Property RevisionDate As Date
        Get
            Return _revisionDate
        End Get
        Set(value As Date)
            _revisionDate = value
        End Set
    End Property

    Private _changeID As String = "" 'Unlinked reference to change table.
    Property ChangeID As String
        Get
            Return _changeID
        End Get
        Set(value As String)
            _changeID = value
        End Set
    End Property

    Private _deprecated As Boolean = False '"Yes" = data is deprecated; "No" =  data is current and valid.  Default is No.
    Property Deprecated As Boolean
        Get
            Return _deprecated
        End Get
        Set(value As Boolean)
            _deprecated = value
        End Set
    End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Public Sub Clear()
        'Clear the properties
        Name = ""
        Code = -1
        SemiMajorAxis = Double.NaN
        UomCode = -1
        InvFlattening = Double.NaN
        SemiMinorAxis = Double.NaN
        EllipsoidShape = True
        Remarks = ""
        InfoSource = -1
        DataSource = -1
        RevisionDate = Date.MinValue
        'RevisionDate = ""
        ChangeID = -1
        Deprecated = False
    End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================

#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class 'Ellipsoid

Public Class PrimeMeridian
    'Properties of a Prime Meridian.

#Region " Properties - All the properties used in this class." '===========================================================================================================================

    Private _name As String = ""  'Name of the prime meridian
    Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Private _code As Integer = -1 'Unique code (integer) of the prime meridian; primary key.
    Property Code As Integer
        Get
            Return _code
        End Get
        Set(value As Integer)
            _code = value
        End Set
    End Property

    Private _greenwichLongitude As Double = Double.NaN 'The longitude of the prime meridian in the Greenwich system.
    Property GreenwichLongitude As Double
        Get
            Return _greenwichLongitude
        End Get
        Set(value As Double)
            _greenwichLongitude = value
        End Set
    End Property

    Private _uomCode As Integer = -1 'ID of the unit of measure in which the longitude in the previous field is expressed.
    Property UomCode As Integer
        Get
            Return _uomCode
        End Get
        Set(value As Integer)
            _uomCode = value
        End Set
    End Property


    Private _remarks As String = "" '
    Property Remarks As String
        Get
            Return _remarks
        End Get
        Set(value As String)
            _remarks = value
        End Set
    End Property

    Private _infoSource As String = "" 'Reference describing the origin of the information to populate this record; e.g. an authoritative publication.
    Property InfoSource As String
        Get
            Return _infoSource
        End Get
        Set(value As String)
            _infoSource = value
        End Set
    End Property

    Private _dataSource As String = "" 'The organisation, body or person who populated this record; for EPSG supplied reference data: "EPSG".
    Property DataSource As String
        Get
            Return _dataSource
        End Get
        Set(value As String)
            _dataSource = value
        End Set
    End Property

    Private _revisionDate As Date = Date.MinValue 'Field used to record the date of creation or modification of this record. Not used if record is deprecated - see deprecation date field.
    'Private _revisionDate As String = "" 'Field used to record the date of creation or modification of this record. Not used if record is deprecated - see deprecation date field.
    Property RevisionDate As Date
        Get
            Return _revisionDate
        End Get
        Set(value As Date)
            _revisionDate = value
        End Set
    End Property

    Private _changeID As String = "" 'Unlinked reference to change table.
    Property ChangeID As String
        Get
            Return _changeID
        End Get
        Set(value As String)
            _changeID = value
        End Set
    End Property

    Private _deprecated As Boolean = False '"Yes" = data is deprecated; "No" =  data is current and valid.  Default is No.
    Property Deprecated As Boolean
        Get
            Return _deprecated
        End Get
        Set(value As Boolean)
            _deprecated = value
        End Set
    End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Public Sub Clear()
        'Clear the properties
        Name = ""
        Code = -1
        GreenwichLongitude = Double.NaN
        UomCode = -1
        Remarks = ""
        InfoSource = -1
        DataSource = -1
        RevisionDate = Date.MinValue
        ChangeID = -1
        Deprecated = False
    End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================

#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class 'PrimeMeridian

Public Class ConventionalRS
    'Properties of a Conventional Reference System.

#Region " Properties - All the properties used in this class." '===========================================================================================================================

    Private _name As String = ""  'Unique name for the conventional (terrestrial or vertical) reference system.
    Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Private _code As Integer = -1 'Unique code for the conventional (terrestrial or vertical) reference system.
    Property Code As Integer
        Get
            Return _code
        End Get
        Set(value As Integer)
            _code = value
        End Set
    End Property

    Private _remarks As String = "" '
    Property Remarks As String
        Get
            Return _remarks
        End Get
        Set(value As String)
            _remarks = value
        End Set
    End Property

    Private _infoSource As String = "" 'Reference describing the origin of the information to populate this record; e.g. an authoritative publication.
    Property InfoSource As String
        Get
            Return _infoSource
        End Get
        Set(value As String)
            _infoSource = value
        End Set
    End Property

    Private _dataSource As String = "" 'The organisation, body or person who populated this record; for EPSG supplied reference data: "EPSG".
    Property DataSource As String
        Get
            Return _dataSource
        End Get
        Set(value As String)
            _dataSource = value
        End Set
    End Property

    Private _revisionDate As Date 'Field used to record the date of creation or modification of this record. Not used if record is deprecated - see deprecation date field.
    'Private _revisionDate As String = "" 'Field used to record the date of creation or modification of this record. Not used if record is deprecated - see deprecation date field.
    Property RevisionDate As Date
        Get
            Return _revisionDate
        End Get
        Set(value As Date)
            _revisionDate = value
        End Set
    End Property

    Private _changeID As String = "" 'Unlinked reference to change table.
    Property ChangeID As String
        Get
            Return _changeID
        End Get
        Set(value As String)
            _changeID = value
        End Set
    End Property

    Private _deprecated As Boolean '"Yes" = data is deprecated; "No" =  data is current and valid.  Default is No.
    Property Deprecated As Boolean
        Get
            Return _deprecated
        End Get
        Set(value As Boolean)
            _deprecated = value
        End Set
    End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Public Sub Clear()
        'Clear the properties
        Name = ""
        Code = -1
        Remarks = ""
        InfoSource = -1
        DataSource = -1
        RevisionDate = Date.MinValue
        'RevisionDate = ""
        ChangeID = -1
        Deprecated = False
    End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================

#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class 'ConventionalRS

Public Class Change
    'Properties of a Change.

#Region " Properties - All the properties used in this class." '===========================================================================================================================

    Private _changeID As Double 'The unique sequential key
    Property ChangeID As Double
        Get
            Return _changeID
        End Get
        Set(value As Double)
            _changeID = value
        End Set
    End Property

    Private _reportDate As Date 'Date that change request created.
    Property ReportDate As Date
        Get
            Return _reportDate
        End Get
        Set(value As Date)
            _reportDate = value
        End Set
    End Property

    Private _dateClosed As Date 'The date at which action on the change request is considered complete.
    Property DateClosed As Date
        Get
            Return _dateClosed
        End Get
        Set(value As Date)
            _dateClosed = value
        End Set
    End Property

    Private _reporter As String = "" 'Name and affiliation of person requesting change.
    Property Reporter As String
        Get
            Return _reporter
        End Get
        Set(value As String)
            _reporter = value
        End Set
    End Property

    Private _request As String = "" 'A description of the change request.
    Property Request As String
        Get
            Return _request
        End Get
        Set(value As String)
            _request = value
        End Set
    End Property


    Private _tablesAffected As String = "" 'A list of tables affected.  GeogCS and ProjCS are subsets of HorizCS.
    Property TablesAffected As String
        Get
            Return _tablesAffected
        End Get
        Set(value As String)
            _tablesAffected = value
        End Set
    End Property

    Private _codesAffected As String = "" 'A list of existing codes that are changed.  New codes not entered here.
    Property CodesAffected As String
        Get
            Return _codesAffected
        End Get
        Set(value As String)
            _codesAffected = value
        End Set
    End Property

    Private _comment As String = "" 'Supplementary remarks about the request or action taken.
    Property Comment As String
        Get
            Return _comment
        End Get
        Set(value As String)
            _comment = value
        End Set
    End Property

    Private _action As String = "" 'Description of changes made to data. Limited to 4000 characters.
    Property Action As String
        Get
            Return _action
        End Get
        Set(value As String)
            _action = value
        End Set
    End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Public Sub Clear()
        'Clear the properties
        ChangeID = Double.NaN
        ReportDate = Date.MinValue
        DateClosed = Date.MinValue
        Reporter = ""
        Request = ""
        TablesAffected = ""
        CodesAffected = ""
        Comment = ""
        Action = ""
    End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================

#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class 'Change

Public Class Deprecation
    'Properties of a Deprecation.

#Region " Properties - All the properties used in this class." '===========================================================================================================================

    Private _code As Integer = -1 'Unique primary key of this table.
    Property Code As Integer
        Get
            Return _code
        End Get
        Set(value As Integer)
            _code = value
        End Set
    End Property

    Private _depDate As Date 'Field used to record the date of deprecation of this record. If blank the record remains valid.
    Property DepDate As Date
        Get
            Return _depDate
        End Get
        Set(value As Date)
            _depDate = value
        End Set
    End Property

    Private _changeID As Double 'The unique sequential key
    Property ChangeID As Double
        Get
            Return _changeID
        End Get
        Set(value As Double)
            _changeID = value
        End Set
    End Property

    Private _objectTableName As String = "" 'Name of the table in which object which has been deprecated is found: Area, Coordinate Axis Name, CRS, CS, Coordinate_Operation, Coordinate_Operation Method, Coordinate_Operation Parameter, Datum, Ellipsoid, Naming System, Prime Meridian or UoM
    Property ObjectTableName As String
        Get
            Return _objectTableName
        End Get
        Set(value As String)
            _objectTableName = value
        End Set
    End Property

    Private _objectCode As Integer = -1 'Code of the object for which has been deprecated.
    Property ObjectCode As Integer
        Get
            Return _objectCode
        End Get
        Set(value As Integer)
            _objectCode = value
        End Set
    End Property

    Private _replacedBy As Integer = -1 'The code of the record which replaces the deprecated record.
    Property ReplacedBy As Integer
        Get
            Return _replacedBy
        End Get
        Set(value As Integer)
            _replacedBy = value
        End Set
    End Property

    Private _reason As String = "" 'Comment on the reason for deprecation.
    Property Reason As String
        Get
            Return _reason
        End Get
        Set(value As String)
            _reason = value
        End Set
    End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Public Sub Clear()
        'Clear the properties
        Code = -1
        DepDate = Date.MinValue
        ChangeID = Double.NaN
        ObjectTableName = ""
        ObjectCode = -1
        ReplacedBy = -1
        Reason = ""
    End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================

#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class 'Deprecation

Public Class NamingSystem
    'Properties of a Naming System.

#Region " Properties - All the properties used in this class." '===========================================================================================================================

    Private _name As String = ""  'Assigned name for the Naming System.
    Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Private _code As Integer = -1 'EPSG assigned unique code for the Naming System; primary key.
    Property Code As Integer
        Get
            Return _code
        End Get
        Set(value As Integer)
            _code = value
        End Set
    End Property

    Private _remarks As String = "" '
    Property Remarks As String
        Get
            Return _remarks
        End Get
        Set(value As String)
            _remarks = value
        End Set
    End Property

    Private _infoSource As String = "" 'Reference describing the origin of the information to populate this record; e.g. an authoritative publication.
    Property InfoSource As String
        Get
            Return _infoSource
        End Get
        Set(value As String)
            _infoSource = value
        End Set
    End Property

    Private _dataSource As String = "" 'The organisation, body or person who populated this record; for EPSG supplied reference data: "EPSG".
    Property DataSource As String
        Get
            Return _dataSource
        End Get
        Set(value As String)
            _dataSource = value
        End Set
    End Property

    Private _revisionDate As Date 'Field used to record the date of creation or modification of this record. Not used if record is deprecated - see deprecation date field.
    'Private _revisionDate As String = "" 'Field used to record the date of creation or modification of this record. Not used if record is deprecated - see deprecation date field.
    Property RevisionDate As Date
        Get
            Return _revisionDate
        End Get
        Set(value As Date)
            _revisionDate = value
        End Set
    End Property

    Private _changeID As String = "" 'Unlinked reference to change table.
    Property ChangeID As String
        Get
            Return _changeID
        End Get
        Set(value As String)
            _changeID = value
        End Set
    End Property

    Private _deprecated As Boolean '"Yes" = data is deprecated; "No" =  data is current and valid.  Default is No.
    Property Deprecated As Boolean
        Get
            Return _deprecated
        End Get
        Set(value As Boolean)
            _deprecated = value
        End Set
    End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Public Sub Clear()
        'Clear the properties
        Name = ""
        Code = -1
        Remarks = ""
        InfoSource = -1
        DataSource = -1
        RevisionDate = Date.MinValue
        'RevisionDate = ""
        ChangeID = -1
        Deprecated = False
    End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================

#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class 'NamingSystem

Public Class AliasName
    'Properties of an Alias.

#Region " Properties - All the properties used in this class." '===========================================================================================================================

    Private _code As Integer = -1 'Unique alias record code.
    Property Code As Integer
        Get
            Return _code
        End Get
        Set(value As Integer)
            _code = value
        End Set
    End Property

    Private _objectTableName As String = "" 'Name of the table in which object to be aliased is found: Area, Coordinate Axis Name, Coordinate Reference System, Coordinate_Operation, Coordinate_Operation Method, Coordinate_Operation Parameter, Datum, Ellipsoid, Naming System, Prime Meridian or UoM
    Property ObjectTableName As String
        Get
            Return _objectTableName
        End Get
        Set(value As String)
            _objectTableName = value
        End Set
    End Property

    Private _objectCode As Integer = -1 'Code of the object for which an alias is defined in this table; part of dual primary key.
    Property ObjectCode As Integer
        Get
            Return _objectCode
        End Get
        Set(value As Integer)
            _objectCode = value
        End Set
    End Property

    Private _namingSystemCode As Integer = -1 'Assigned unique code for the naming convention; part of dual primary key.
    Property NamingSystemCode As Integer
        Get
            Return _namingSystemCode
        End Get
        Set(value As Integer)
            _namingSystemCode = value
        End Set
    End Property

    Private _aliasText As String = "" 'Alias of the object.
    Property AliasText As String
        Get
            Return _aliasText
        End Get
        Set(value As String)
            _aliasText = value
        End Set
    End Property

    Private _remarks As String = "" '
    Property Remarks As String
        Get
            Return _remarks
        End Get
        Set(value As String)
            _remarks = value
        End Set
    End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Public Sub Clear()
        'Clear the properties
        Code = -1
        ObjectTableName = ""
        ObjectCode = -1
        NamingSystemCode = -1
        AliasText = ""
        Remarks = ""
    End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================

#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class 'AliasName

Public Class Projection
    'Projection calculations.

    'Usage example:
    '  Set up the projection:
    '    Projection.MethodName = "Transverse Mercator"
    '    Projection.Method.SemiMajorAxis = 
    '    Projection.Method.InverseFlattening = 
    '    Projection.Method.LatitudeOfnaturalOrigin = 
    '    Projection.Method.LongitudeOfNaturalOrigin = 
    '    Projection.Method.ScaleFactorAtNaturalOrigin = 
    '    Projection.Method.FalseEasting = 
    '    Projection.Method.FalseNorthing = 
    '    Projection.Latitide =
    '    Projection.Longitude = 
    '    Projection.CalcNorthEast
    '    Northing = Projection.Northing
    '    Easting = Projection.Easting


#Region " Variable Declarations - All the variables used in this class." '=====================================================================================================================

    Public Coord As Coordinate 'This class reference is set before the method is used. For example InputCrs.Projection.Method.Coord = InputCrs.Coord - The Projection class now accesses the Coord values directly.
    Public Method As Object

#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Properties - All the properties used in this class." '===========================================================================================================================

    Private _methodName As String = "" 'The name of the projection method.
    Property MethodName As String
        Get
            Return _methodName
        End Get
        Set(value As String)
            _methodName = value
            ApplyMethodName()
        End Set
    End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Public Sub Clear()
        'Clear the Projection
        _methodName = ""
        Method = Nothing
    End Sub

    Private Sub ApplyMethodName()
        'Apply the Method Name.

        Select Case MethodName
            Case "Transverse Mercator"
                'Method = New TransverseMercator
                'Method = New TransverseMercatorRedfearn
                Method = New TransverseMercatorKruegerNSeries 'More accurate than the Redfearn method.
                Method.Coord = Coord

            Case Else
                RaiseEvent ErrorMessage("Unknown projection method name: " & MethodName & vbCrLf)
        End Select

    End Sub

    Public Sub ShowAvailableMethods()
        'Return a list of the available projection methods.
        RaiseEvent Message(vbCrLf & "List of available projection methods:" & vbCrLf)
        RaiseEvent Message("Transverse Mercator" & vbCrLf)
    End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================
    Event ErrorMessage(ByVal Msg As String) 'Send an error message.
    Event Message(ByVal Msg As String) 'Send a normal message.
#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


    Public Class TransverseMercatorKruegerNSeries
        'Transverse Mercator projection calculations.
        'Uses the Krueger N Series equations
        'More accurate than the Redfearn formula.
        '
        'The calculation method is from the document:
        'TRANSVERSE MERCATOR PROJECTION
        'Karney-Krueger equations
        'R.E. Deakin
        '02-Sep-2014


#Region " Variable Declarations - All the variables used in this class." '=====================================================================================================================

        Public Coord As Coordinate 'This class reference is set before the method is used. 

        Dim E As Double 'Ellipsoid constant: Eccentricity
        Dim E2 As Double 'Ellipsoid constant: Eccentricity squared
        Dim N As Double 'Ellipsoid constant: 3rd flattening of ellipsoid
        Dim N2 As Double 'N squared
        Dim N3 As Double 'N cubed
        Dim N4 As Double
        Dim N5 As Double
        Dim N6 As Double
        Dim N7 As Double
        Dim N8 As Double
        Dim A As Double 'Rectifying radius

        'Alpha coefficients:
        Dim A2 As Double
        Dim A4 As Double
        Dim A6 As Double
        Dim A8 As Double
        Dim A10 As Double
        Dim A12 As Double
        Dim A14 As Double
        Dim A16 As Double

        Dim LatRad As Double 'The Latitude in Radians.
        Dim TanLat As Double '= Tan(LatRad) where LatRad is the Latitude in radians.
        Dim Temp As Double '= E*TanLatRad/Sqrt(1+TanLatRad^2)
        Dim Sigma As Double '= Sinh(E*ATanh*(E*TanLatRad/Sqrt(1 + TanLatRad^2))) Note: ATanh(x) = (Log(1+x)-Log(1-x))/2 so Sigma = Sinh(E*(Log(1+Temp)-Log(1-Temp))/2)
        Dim TanConfLat As Double 'Tan(Conformal Latitude) = TanLat * Sqrt(1 + Sigma^2) - Sigma * Sqrt(1 + TanLat^2)
        Dim ConfLat As Double 'Conformal Latitude in radians
        Dim ConfLatDeg As Double 'Conformal Latitude in degrees

        Dim DifLon As Double 'The difference between the Longitide and the Central Maeridian Longitude in Radians
        Dim u As Double 'Gauss-Schreiber coordinate (north)
        Dim v As Double 'Gauss-Schreiber coordinate (east)

        Dim Xi_Prime As Double ' ξ'  Xi_Prime
        Dim Eta_Prime As Double ' η'  Eta_Prime

        Dim Xi_Prime1 As Double
        Dim Xi_Prime2 As Double
        Dim Xi_Prime3 As Double
        Dim Xi_Prime4 As Double
        Dim Xi_Prime5 As Double
        Dim Xi_Prime6 As Double
        Dim Xi_Prime7 As Double
        Dim Xi_Prime8 As Double
        Dim Xi As Double

        Dim Eta_Prime1 As Double
        Dim Eta_Prime2 As Double
        Dim Eta_Prime3 As Double
        Dim Eta_Prime4 As Double
        Dim Eta_Prime5 As Double
        Dim Eta_Prime6 As Double
        Dim Eta_Prime7 As Double
        Dim Eta_Prime8 As Double
        Dim Eta As Double

        Dim X As Double
        Dim Y As Double

        'For Scale Factor calculations:
        Dim Q1 As Double
        Dim Q2 As Double
        Dim Q3 As Double
        Dim Q4 As Double
        Dim Q5 As Double
        Dim Q6 As Double
        Dim Q7 As Double
        Dim Q8 As Double
        Dim Q As Double

        Dim P1 As Double
        Dim P2 As Double
        Dim P3 As Double
        Dim P4 As Double
        Dim P5 As Double
        Dim P6 As Double
        Dim P7 As Double
        Dim P8 As Double
        Dim P As Double

        Dim EorWofCM As Integer 'East or West of Central Meridian (-1 or 1)
        Dim NorSofEq As Integer 'North or South of Equator (-1 or 1)
        Dim GCMult As Integer 'Grid Convergence multiplier.

        Dim GridConvRad As Double 'The Grid Convergence in radians

        'Beta coefficients:
        Dim B2 As Double
        Dim B4 As Double
        Dim B6 As Double
        Dim B8 As Double
        Dim B10 As Double
        Dim B12 As Double
        Dim B14 As Double
        Dim B16 As Double

        Dim Xi1 As Double
        Dim Xi2 As Double
        Dim Xi3 As Double
        Dim Xi4 As Double
        Dim Xi5 As Double
        Dim Xi6 As Double
        Dim Xi7 As Double
        Dim Xi8 As Double

        Dim Eta1 As Double
        Dim Eta2 As Double
        Dim Eta3 As Double
        Dim Eta4 As Double
        Dim Eta5 As Double
        Dim Eta6 As Double
        Dim Eta7 As Double
        Dim Eta8 As Double

        Dim I As Integer 'Iteration loop

        Dim FnT As Double
        Dim FnPrimeT As Double
        Dim NewTanLat As Double 'New TanLat - calculated in the Newton-Raphson method to refine TanLat.

        Dim DifLonDeg As Double 'Longitude difference in degrees

#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Properties - All the properties used in this class." '===========================================================================================================================

        Private _semiMajorAxis As Double = Double.NaN 'The semi major axis of the reference ellipsoid
        Property SemiMajorAxis As Double
            Get
                Return _semiMajorAxis
            End Get
            Set(value As Double)
                _semiMajorAxis = value
            End Set
        End Property

        Private _inverseFlattening As Double = Double.NaN 'The inverse flattening of the reference ellipsoid.
        Property InverseFlattening As Double
            Get
                Return _inverseFlattening
            End Get
            Set(value As Double)
                _inverseFlattening = value
                _flattening = 1.0# / _inverseFlattening
                _semiMinorAxis = _semiMajorAxis * (1.0# - _flattening)
            End Set
        End Property

        Private _flattening As Double = Double.NaN 'The flattening of the ellipsoid.
        Property Flattening As Double
            Get
                Return _flattening
            End Get
            Set(value As Double)
                _flattening = value
                _semiMinorAxis = _semiMajorAxis * (1.0# - _flattening)
            End Set
        End Property

        Private _semiMinorAxis As Double = Double.NaN 'The semi minor axis of the reference ellipsoid. If this value is entered, the InverseFlattening value is calculated.
        Property SemiMinorAxis As Double
            Get
                Return _semiMinorAxis
            End Get
            Set(value As Double)
                _semiMinorAxis = value
                If _semiMajorAxis = Double.NaN Then
                    'Cannot calculate inverse flattening
                Else
                    _inverseFlattening = _semiMajorAxis / (_semiMajorAxis - _semiMinorAxis)
                    _flattening = (_semiMajorAxis - _semiMinorAxis) / _semiMajorAxis
                End If
            End Set
        End Property


        Private _latitudeOfNaturalOrigin As Double = Double.NaN 'The Latitude of the natural origin (in degrees).
        Property LatitudeOfNaturalOrigin As Double
            Get
                Return _latitudeOfNaturalOrigin
            End Get
            Set(value As Double)
                _latitudeOfNaturalOrigin = value
            End Set
        End Property

        Private _longitudeOfNaturalOrigin As Double = Double.NaN 'The Longitude of the natural origin (in degrees).
        Property LongitudeOfNaturalOrigin As Double
            Get
                Return _longitudeOfNaturalOrigin
            End Get
            Set(value As Double)
                _longitudeOfNaturalOrigin = value
            End Set
        End Property

        Private _scaleFactorAtNaturalOrigin As Double = Double.NaN 'The scale factor at the natural origin.
        Property ScaleFactorAtNaturalOrigin As Double
            Get
                Return _scaleFactorAtNaturalOrigin
            End Get
            Set(value As Double)
                _scaleFactorAtNaturalOrigin = value
            End Set
        End Property


        Private _FalseEasting As Double = Double.NaN
        Property FalseEasting As Double
            Get
                Return _FalseEasting
            End Get
            Set(value As Double)
                _FalseEasting = value
            End Set
        End Property

        Private _FalseNorthing As Double = Double.NaN
        Property FalseNorthing As Double
            Get
                Return _FalseNorthing
            End Get
            Set(value As Double)
                _FalseNorthing = value
            End Set
        End Property

        Private _distanceCode As Integer = -1 'The UOM code
        Property DistanceCode As Integer
            Get
                Return _distanceCode
            End Get
            Set(value As Integer)
                _distanceCode = value
            End Set
        End Property


        Private _distanceUnits As String = "" 'The units used to measure eastings, northings, FalseEasting and FalseNorthing
        Property DistanceUnits As String
            Get
                Return _distanceUnits
            End Get
            Set(value As String)
                _distanceUnits = value
            End Set
        End Property

        'NOTE: The Coord class is now used to store Latitude, Longitude, Easting and Northing.
        'Property _latitude As Double = Double.NaN 'The decimal Latitude of the location.
        'Property Latitude As Double
        '    Get
        '        Return _latitude
        '    End Get
        '    Set(value As Double)
        '        _latitude = value
        '    End Set
        'End Property

        'Private _longitude As Double = Double.NaN 'The decimal Longtude of the location.
        'Property Longitude As Double
        '    Get
        '        Return _longitude
        '    End Get
        '    Set(value As Double)
        '        _longitude = value
        '    End Set
        'End Property

        'Private _northing As Double = Double.NaN 'The Northing of the location.
        'Property Northing As Double
        '    Get
        '        Return _northing
        '    End Get
        '    Set(value As Double)
        '        _northing = value
        '    End Set
        'End Property

        'Private _easting As Double = Double.NaN 'The Easting of the location.
        'Property Easting As Double
        '    Get
        '        Return _easting
        '    End Get
        '    Set(value As Double)
        '        _easting = value
        '    End Set
        'End Property

        Private _gridConvergence As Double = Double.NaN 'The Grid Convergence at the location.
        Property GridConvergence As Double
            Get
                Return _gridConvergence
            End Get
            Set(value As Double)
                _gridConvergence = value
            End Set
        End Property

        Private _pointScale As Double = Double.NaN 'The Point Scale at the location.
        Property PointScale As Double
            Get
                Return _pointScale
            End Get
            Set(value As Double)
                _pointScale = value
            End Set
        End Property


#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Methods - The main actions performed by this class." '===========================================================================================================================

        Public Sub Clear()
            'Clear the properties
            SemiMajorAxis = Double.NaN
            InverseFlattening = Double.NaN
            Flattening = Double.NaN
            SemiMinorAxis = Double.NaN
            LatitudeOfNaturalOrigin = Double.NaN
            LongitudeOfNaturalOrigin = Double.NaN
            ScaleFactorAtNaturalOrigin = Double.NaN
            FalseEasting = Double.NaN
            FalseNorthing = Double.NaN
            DistanceCode = -1
            DistanceUnits = ""
        End Sub

        Public Sub SetParameter(ParameterCode As Integer, ParameterValue As Double)
            'Sets the Transverse Mercator property from a PropertyCode and a PropertyValue.
            Select Case ParameterCode
                Case 8801 'LatitudeOfNaturalOrigin
                    LatitudeOfNaturalOrigin = ParameterValue
                Case 8802 'LongitudeOfNaturalOrigin
                    LongitudeOfNaturalOrigin = ParameterValue
                Case 8805 'ScaleFactorAtNaturalOrigin
                    ScaleFactorAtNaturalOrigin = ParameterValue
                Case 8806 'FalseEasting
                    FalseEasting = ParameterValue
                Case 8807 'FalseNorthing
                    FalseNorthing = ParameterValue
                Case Else
                    RaiseEvent ErrorMessage("Unknown property code: " & ParameterCode & vbCrLf)
            End Select
        End Sub

        Public Function ValidProperty(PropertyName As String) As Boolean
            'Return True if PropertyName is a valid property of this projection method.

            Select Case PropertyName
                Case "SemiMajorAxis"
                    Return True
                Case "InverseFlattening"
                    Return True
                Case "Flattening"
                    Return True
                Case "SemiMinorAxis"
                    Return True
                Case "LatitudeOfNaturalOrigin"
                    Return True
                Case "LongitudeOfNaturalOrigin"
                    Return True
                Case "ScaleFactorAtNaturalOrigin"
                    Return True
                Case "FalseEasting"
                    Return True
                Case "FalseNorthing"
                    Return True
                Case Else
                    Return False
            End Select

        End Function

        Public Sub UpdateVariables()
            'Update the variables used in the projection calculations.

            'Ellipsoid constants:
            E2 = (2.0# * Flattening) - (Flattening * Flattening)
            E = Math.Sqrt(E2)
            N = Flattening / (2.0# - Flattening)
            N2 = N * N
            N3 = N2 * N
            N4 = N3 * N
            N5 = N4 * N
            N6 = N5 * N
            N7 = N6 * N
            N8 = N7 * N

            'Rectifying radius:
            A = (SemiMajorAxis / (1.0# + N)) * (N2 * (N2 * (N2 * (25.0# * N2 + 64.0) + 256.0#) + 4096.0#) + 16384.0#) / 16384.0# 'Expressed in Horner form for efficient computation.

            'Alpha coefficients (forward transformation):
            'The coefficients are expressed in Horner form for efficient computation:
            A2 = (N * (N * (N * (N * (N * (N * ((37884525.0# - 75900428.0# * N) * N + 42422016.0#) - 89611200) + 46287360.0#) + 64504000.0#) - 135475200.0#) + 101606400.0#)) / 203212800.0#
            A4 = (N2 * (N * (N * (N * (N * (N * (14800388.0# * N + 83274912.0#) - 178508970.0#) + 77690880.0#) + 67374720.0#) - 104509440.0#) + 47174400.0#)) / 174182400.0#
            A6 = (N3 * (N * (N * (N * (N * (318729724 * N - 738126169.0#) + 294981280.0#) + 178924680.0#) - 234938880.0#) + 81164160.0#)) / 319334400.0#
            A8 = (N4 * (N * (N * ((14967552000.0# - 40176129013.0# * N) * N + 6971354016.0#) - 8165836800.0#) + 2355138720.0#)) / 7664025600.0#
            A10 = (N5 * (N * (N * (10421654396.0# * N + 3997835751.0#) - 4266773472.0#) + 1072709352.0#)) / 2490808320.0#
            A12 = (N6 * (N * (175214326799.0# * N - 171950693600.0#) + 38652967262.0#)) / 58118860800.0#
            A14 = (13700311101.0# - 67039739596.0# * N) * N7 / 12454041600.0#
            A16 = 1424729850961.0# * N8 / 743921418240.0#

            'Beta coefficients (inverse transformation):
            'The coefficients are expressed in Horner form for efficient computation:
            B2 = (N * (N * (N * (N * (N * (N * ((37845269.0# - 31777436.0# * N) * N - 43097152.0#) + 42865200.0#) + 752640.0#) - 104428800) + 180633600) - 135475200.0#)) / 270950400.0#
            B4 = (N ^ 2 * (N * (N * (N * (N * ((-24749483.0# * N - 14930208.0#) * N + 100683990.0#) - 152616960.0#) + 105719040.0#) - 23224320.0#) - 7257600.0#)) / 348364800.0#
            B6 = (N ^ 3 * (N * (N * (N * (N * (232468668.0# * N - 101880889.0#) - 39205760.0#) + 297950040.0#) + 28131840.0#) - 22619520.0#)) / 638668800.0#
            B8 = (N ^ 4 * (N * (N * ((-324154477.0# * N - 1433121792.0#) * N + 876745056.0#) + 167270400.0#) - 208945440.0#)) / 7664025600.0#
            B10 = N ^ 5 * (N * ((312227409.0# - 457888660.0# * N) * N + 67920528.0#) - 70779852) / 2490808320.0#
            B12 = N ^ 6 * (N * (19841813847.0# * N + 3665348512.0#) - 3758062126.0#) / 116237721600.0#
            B14 = N ^ 7 * (1989295244.0# * N - 1979471673.0#) / 49816166400.0#
            B16 = -191773887257.0# * N ^ 8 / 3719607091200.0#
        End Sub

        'Public Sub LatLongToNorthEast()
        Public Sub LongLatToEastNorth()
            'Calculate the Easting and Northing values from the Latitude and Longitude.

            'LatRad = (Latitude / 180.0#) * Math.PI
            LatRad = (Coord.Latitude / 180.0#) * Math.PI
            'TanLat = Math.Tan((Latitude / 180.0#) * Math.PI)
            TanLat = Math.Tan(LatRad)
            Temp = E * TanLat / Math.Sqrt(1 + TanLat ^ 2)
            Sigma = Math.Sinh(E * Atanh(E * TanLat / Math.Sqrt(1.0# + TanLat ^ 2)))

            TanConfLat = TanLat * Math.Sqrt(1.0# + Sigma ^ 2) - Sigma * Math.Sqrt(1.0# + TanLat ^ 2)

            'Conformal latitude:
            ConfLat = Math.Atan(TanConfLat)
            ConfLatDeg = ConfLat * 180.0# / Math.PI

            'Longitude difference:
            'DifLon = ((Longitude - LongitudeOfNaturalOrigin) / 180.0#) * Math.PI
            DifLon = ((Coord.Longitude - LongitudeOfNaturalOrigin) / 180.0#) * Math.PI

            'Gauss-Schreiber coords:
            u = SemiMajorAxis * Math.Atan(TanConfLat / Math.Cos(DifLon))
            v = SemiMajorAxis * Asinh(Math.Sin(DifLon) / Math.Sqrt(TanConfLat ^ 2 + Math.Cos(DifLon) ^ 2))

            'Gauss-Schreiber ratios:
            Xi_Prime = u / SemiMajorAxis ' ξ'  Xi_Prime
            Eta_Prime = v / SemiMajorAxis ' η'  Eta_Prime

            Eta_Prime1 = A2 * Math.Cos(2.0# * Xi_Prime) * Math.Sinh(2.0# * Eta_Prime)
            Eta_Prime2 = A4 * Math.Cos(4.0# * Xi_Prime) * Math.Sinh(4.0# * Eta_Prime)
            Eta_Prime3 = A6 * Math.Cos(6.0# * Xi_Prime) * Math.Sinh(6.0# * Eta_Prime)
            Eta_Prime4 = A8 * Math.Cos(8.0# * Xi_Prime) * Math.Sinh(8.0# * Eta_Prime)
            Eta_Prime5 = A10 * Math.Cos(10.0# * Xi_Prime) * Math.Sinh(10.0# * Eta_Prime)
            Eta_Prime6 = A12 * Math.Cos(12.0# * Xi_Prime) * Math.Sinh(12.0# * Eta_Prime)
            Eta_Prime7 = A14 * Math.Cos(14.0# * Xi_Prime) * Math.Sinh(14.0# * Eta_Prime)
            Eta_Prime8 = A16 * Math.Cos(16.0# * Xi_Prime) * Math.Sinh(16.0# * Eta_Prime)
            Eta = Eta_Prime + Eta_Prime1 + Eta_Prime2 + Eta_Prime3 + Eta_Prime4 + Eta_Prime5 + Eta_Prime6 + Eta_Prime7 + Eta_Prime8


            Xi_Prime1 = A2 * Math.Sin(2.0# * Xi_Prime) * Math.Cosh(2.0# * Eta_Prime)
            Xi_Prime2 = A4 * Math.Sin(4.0# * Xi_Prime) * Math.Cosh(4.0# * Eta_Prime)
            Xi_Prime3 = A6 * Math.Sin(6.0# * Xi_Prime) * Math.Cosh(6.0# * Eta_Prime)
            Xi_Prime4 = A8 * Math.Sin(8.0# * Xi_Prime) * Math.Cosh(8.0# * Eta_Prime)
            Xi_Prime5 = A10 * Math.Sin(10.0# * Xi_Prime) * Math.Cosh(10.0# * Eta_Prime)
            Xi_Prime6 = A12 * Math.Sin(12.0# * Xi_Prime) * Math.Cosh(12.0# * Eta_Prime)
            Xi_Prime7 = A14 * Math.Sin(14.0# * Xi_Prime) * Math.Cosh(14.0# * Eta_Prime)
            Xi_Prime8 = A16 * Math.Sin(16.0# * Xi_Prime) * Math.Cosh(16.0# * Eta_Prime)
            Xi = Xi_Prime + Xi_Prime1 + Xi_Prime2 + Xi_Prime3 + Xi_Prime4 + Xi_Prime5 + Xi_Prime6 + Xi_Prime7 + Xi_Prime8

            X = Eta * A
            Y = Xi * A

            'Easting = ScaleFactorAtNaturalOrigin * X + FalseEasting
            Coord.Easting = ScaleFactorAtNaturalOrigin * X + FalseEasting
            'Northing = ScaleFactorAtNaturalOrigin * Y + FalseNorthing
            Coord.Northing = ScaleFactorAtNaturalOrigin * Y + FalseNorthing

            'Point Scale Factor calculations:
            Q1 = 2.0# * A2 * Math.Sin(2.0# * Xi_Prime) * Math.Sinh(2.0# * Eta_Prime)
            Q2 = 4.0# * A4 * Math.Sin(4.0# * Xi_Prime) * Math.Sinh(4.0# * Eta_Prime)
            Q3 = 6.0# * A6 * Math.Sin(6.0# * Xi_Prime) * Math.Sinh(6.0# * Eta_Prime)
            Q4 = 8.0# * A8 * Math.Sin(8.0# * Xi_Prime) * Math.Sinh(8.0# * Eta_Prime)
            Q5 = 10.0# * A10 * Math.Sin(10.0# * Xi_Prime) * Math.Sinh(10.0# * Eta_Prime)
            Q6 = 12.0# * A12 * Math.Sin(12.0# * Xi_Prime) * Math.Sinh(12.0# * Eta_Prime)
            Q7 = 14.0# * A14 * Math.Sin(14.0# * Xi_Prime) * Math.Sinh(14.0# * Eta_Prime)
            Q8 = 16.0# * A16 * Math.Sin(16.0# * Xi_Prime) * Math.Sinh(16.0# * Eta_Prime)
            Q = 0.0# - (Q1 + Q2 + Q3 + Q4 + Q5 + Q6 + Q7 + Q8)

            P1 = 2.0# * A2 * Math.Cos(2.0# * Xi_Prime) * Math.Cosh(2.0# * Eta_Prime)
            P2 = 4.0# * A4 * Math.Cos(4.0# * Xi_Prime) * Math.Cosh(4.0# * Eta_Prime)
            P3 = 6.0# * A6 * Math.Cos(6.0# * Xi_Prime) * Math.Cosh(6.0# * Eta_Prime)
            P4 = 8.0# * A8 * Math.Cos(8.0# * Xi_Prime) * Math.Cosh(8.0# * Eta_Prime)
            P5 = 10.0# * A10 * Math.Cos(10.0# * Xi_Prime) * Math.Cosh(10.0# * Eta_Prime)
            P6 = 12.0# * A12 * Math.Cos(12.0# * Xi_Prime) * Math.Cosh(12.0# * Eta_Prime)
            P7 = 14.0# * A14 * Math.Cos(14.0# * Xi_Prime) * Math.Cosh(14.0# * Eta_Prime)
            P8 = 16.0# * A16 * Math.Cos(16.0# * Xi_Prime) * Math.Cosh(16.0# * Eta_Prime)
            P = 1.0# + P1 + P2 + P3 + P4 + P5 + P6 + P7 + P8

            PointScale = ScaleFactorAtNaturalOrigin * A / SemiMajorAxis * Math.Sqrt(P ^ 2 + Q ^ 2) * Math.Sqrt(1.0# + TanLat ^ 2) * Math.Sqrt(1.0# - E2 * Math.Sin(LatRad) ^ 2) / Math.Sqrt(TanConfLat ^ 2 + Math.Cos(DifLon) ^ 2)

            'If Longitude > LongitudeOfNaturalOrigin Then EorWofCM = 1 Else EorWofCM = -1
            If Coord.Longitude > LongitudeOfNaturalOrigin Then EorWofCM = 1 Else EorWofCM = -1
            'If Latitude < 0 Then NorSofEq = -1 Else NorSofEq = 1
            If Coord.Latitude < 0 Then NorSofEq = -1 Else NorSofEq = 1
            If EorWofCM = -1 And NorSofEq = -1 Then
                GCMult = -1
            ElseIf EorWofCM = 1 And NorSofEq = -1 Then
                GCMult = 1
            ElseIf EorWofCM = 1 And NorSofEq = 1 Then
                GCMult = -1
            Else
                GCMult = 1
            End If

            GridConvRad = Math.Atan(Math.Abs(Q / P)) + Math.Atan(Math.Abs(TanConfLat * Math.Tan(DifLon)) / Math.Sqrt(1.0# + TanConfLat ^ 2))
            GridConvergence = GridConvRad * 180.0# / Math.PI * GCMult

        End Sub

        'Public Sub NorthEastToLatLong()
        Public Sub EastNorthToLongLat()
            'Calculate the Latitude and Longitude values from the Easting and Northing.

            'X = (Easting - FalseEasting) / ScaleFactorAtNaturalOrigin
            X = (Coord.Easting - FalseEasting) / ScaleFactorAtNaturalOrigin
            'Y = (Northing - FalseNorthing) / ScaleFactorAtNaturalOrigin
            Y = (Coord.Northing - FalseNorthing) / ScaleFactorAtNaturalOrigin

            Xi = Y / A
            Eta = X / A

            Xi1 = B2 * Math.Sin(2.0# * Xi) * Math.Cosh(2.0# * Eta)
            Xi2 = B4 * Math.Sin(4.0# * Xi) * Math.Cosh(4.0# * Eta)
            Xi3 = B6 * Math.Sin(6.0# * Xi) * Math.Cosh(6.0# * Eta)
            Xi4 = B8 * Math.Sin(8.0# * Xi) * Math.Cosh(8.0# * Eta)
            Xi5 = B10 * Math.Sin(10.0# * Xi) * Math.Cosh(10.0# * Eta)
            Xi6 = B12 * Math.Sin(12.0# * Xi) * Math.Cosh(12.0# * Eta)
            Xi7 = B14 * Math.Sin(14.0# * Xi) * Math.Cosh(14.0# * Eta)
            Xi8 = B16 * Math.Sin(16.0# * Xi) * Math.Cosh(16.0# * Eta)
            Xi_Prime = Xi + Xi1 + Xi2 + Xi3 + Xi4 + Xi5 + Xi6 + Xi7 + Xi8

            Eta1 = B2 * Math.Cos(2.0# * Xi) * Math.Sinh(2.0# * Eta)
            Eta2 = B4 * Math.Cos(4.0# * Xi) * Math.Sinh(4.0# * Eta)
            Eta3 = B6 * Math.Cos(6.0# * Xi) * Math.Sinh(6.0# * Eta)
            Eta4 = B8 * Math.Cos(8.0# * Xi) * Math.Sinh(8.0# * Eta)
            Eta5 = B10 * Math.Cos(10.0# * Xi) * Math.Sinh(10.0# * Eta)
            Eta6 = B12 * Math.Cos(12.0# * Xi) * Math.Sinh(12.0# * Eta)
            Eta7 = B14 * Math.Cos(14.0# * Xi) * Math.Sinh(14.0# * Eta)
            Eta8 = B16 * Math.Cos(16.0# * Xi) * Math.Sinh(16.0# * Eta)
            Eta_Prime = Eta + Eta1 + Eta2 + Eta3 + Eta4 + Eta5 + Eta6 + Eta7 + Eta8

            TanConfLat = Math.Sin(Xi_Prime) / Math.Sqrt(Math.Sinh(Eta_Prime) ^ 2 + Math.Cos(Xi_Prime) ^ 2)
            TanLat = TanConfLat
            Sigma = Math.Sinh(E * Atanh(E * TanLat / Math.Sqrt(1.0# + TanLat ^ 2)))
            FnT = TanLat * Math.Sqrt(1.0# + Sigma ^ 2) - Sigma * Math.Sqrt(1.0# + TanLat ^ 2) - TanConfLat
            FnPrimeT = (Math.Sqrt(1.0# + Sigma ^ 2) * Math.Sqrt(1.0# + TanLat ^ 2) - Sigma * TanLat) * (1.0# - E2) * Math.Sqrt(1.0# + TanLat ^ 2) / (1.0# + (1.0# - E2) * TanLat ^ 2)

            'Newton-Raphson iteration used to refine TanLat value. Maximum of 20 iterations. Iteration stop if the TanLat difference is less than 0.000000001.
            For I = 1 To 20
                NewTanLat = TanLat - FnT / FnPrimeT
                If Math.Abs(NewTanLat - TanLat) < 0.0000000000001 Then
                    TanLat = NewTanLat
                    Exit For
                Else
                    TanLat = NewTanLat
                    Sigma = Math.Sinh(E * Atanh(E * TanLat / Math.Sqrt(1.0# + TanLat ^ 2)))
                    FnT = TanLat * Math.Sqrt(1.0# + Sigma ^ 2) - Sigma * Math.Sqrt(1.0# + TanLat ^ 2) - TanConfLat
                    FnPrimeT = (Math.Sqrt(1.0# + Sigma ^ 2) * Math.Sqrt(1.0# + TanLat ^ 2) - Sigma * TanLat) * (1.0# - E2) * Math.Sqrt(1.0# + TanLat ^ 2) / (1.0# + (1.0# - E2) * TanLat ^ 2)
                End If
            Next

            LatRad = Math.Atan(TanLat)
            'Latitude = LatRad * 180.0# / Math.PI
            Coord.Latitude = LatRad * 180.0# / Math.PI

            DifLon = Math.Atan(Math.Sinh(Eta_Prime) / Math.Cos(Xi_Prime))
            DifLonDeg = DifLon * 180.0# / Math.PI
            'Longitude = LongitudeOfNaturalOrigin + DifLonDeg
            Coord.Longitude = LongitudeOfNaturalOrigin + DifLonDeg

            'Point Scale Factor calculations:
            Q1 = 2.0# * A2 * Math.Sin(2.0# * Xi_Prime) * Math.Sinh(2.0# * Eta_Prime)
            Q2 = 4.0# * A4 * Math.Sin(4.0# * Xi_Prime) * Math.Sinh(4.0# * Eta_Prime)
            Q3 = 6.0# * A6 * Math.Sin(6.0# * Xi_Prime) * Math.Sinh(6.0# * Eta_Prime)
            Q4 = 8.0# * A8 * Math.Sin(8.0# * Xi_Prime) * Math.Sinh(8.0# * Eta_Prime)
            Q5 = 10.0# * A10 * Math.Sin(10.0# * Xi_Prime) * Math.Sinh(10.0# * Eta_Prime)
            Q6 = 12.0# * A12 * Math.Sin(12.0# * Xi_Prime) * Math.Sinh(12.0# * Eta_Prime)
            Q7 = 14.0# * A14 * Math.Sin(14.0# * Xi_Prime) * Math.Sinh(14.0# * Eta_Prime)
            Q8 = 16.0# * A16 * Math.Sin(16.0# * Xi_Prime) * Math.Sinh(16.0# * Eta_Prime)
            Q = 0.0# - (Q1 + Q2 + Q3 + Q4 + Q5 + Q6 + Q7 + Q8)

            P1 = 2.0# * A2 * Math.Cos(2.0# * Xi_Prime) * Math.Cosh(2.0# * Eta_Prime)
            P2 = 4.0# * A4 * Math.Cos(4.0# * Xi_Prime) * Math.Cosh(4.0# * Eta_Prime)
            P3 = 6.0# * A6 * Math.Cos(6.0# * Xi_Prime) * Math.Cosh(6.0# * Eta_Prime)
            P4 = 8.0# * A8 * Math.Cos(8.0# * Xi_Prime) * Math.Cosh(8.0# * Eta_Prime)
            P5 = 10.0# * A10 * Math.Cos(10.0# * Xi_Prime) * Math.Cosh(10.0# * Eta_Prime)
            P6 = 12.0# * A12 * Math.Cos(12.0# * Xi_Prime) * Math.Cosh(12.0# * Eta_Prime)
            P7 = 14.0# * A14 * Math.Cos(14.0# * Xi_Prime) * Math.Cosh(14.0# * Eta_Prime)
            P8 = 16.0# * A16 * Math.Cos(16.0# * Xi_Prime) * Math.Cosh(16.0# * Eta_Prime)
            P = 1.0# + P1 + P2 + P3 + P4 + P5 + P6 + P7 + P8

            PointScale = ScaleFactorAtNaturalOrigin * A / SemiMajorAxis * Math.Sqrt(P ^ 2 + Q ^ 2) * (Math.Sqrt(1.0# + TanLat ^ 2) * Math.Sqrt(1.0# - E2 * Math.Sin(LatRad) ^ 2) / Math.Sqrt(TanConfLat ^ 2 + Math.Cos(DifLon) ^ 2))

            'If Longitude > LongitudeOfNaturalOrigin Then EorWofCM = 1 Else EorWofCM = -1
            If Coord.Longitude > LongitudeOfNaturalOrigin Then EorWofCM = 1 Else EorWofCM = -1
            'If Latitude < 0 Then NorSofEq = -1 Else NorSofEq = 1
            If Coord.Latitude < 0 Then NorSofEq = -1 Else NorSofEq = 1
            If EorWofCM = -1 And NorSofEq = -1 Then
                GCMult = -1
            ElseIf EorWofCM = 1 And NorSofEq = -1 Then
                GCMult = 1
            ElseIf EorWofCM = 1 And NorSofEq = 1 Then
                GCMult = -1
            Else
                GCMult = 1
            End If

            GridConvRad = Math.Atan(Math.Abs(Q / P)) + Math.Atan(Math.Abs(TanConfLat * Math.Tan(DifLon)) / Math.Sqrt(1.0# + TanConfLat ^ 2))
            GridConvergence = GridConvRad * 180.0# / Math.PI * GCMult

        End Sub

        Private Function Asinh(X As Double) As Double
            Return Math.Log(X + Math.Sqrt(X ^ 2 + 1.0#))
        End Function

        Private Function Acosh(X As Double) As Double
            Return Math.Log(X + Math.Sqrt(X ^ 2 - 1.0#))
        End Function

        Private Function Atanh(X As Double) As Double
            Return (Math.Log(1.0# + X) - Math.Log(1.0# - X)) / 2.0#
        End Function


#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================
        Event ErrorMessage(ByVal Msg As String) 'Send an error message.
        Event Message(ByVal Msg As String) 'Send a normal message.
#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


    End Class 'TransverseMercatorKruegerNSeries


    Public Class TransverseMercatorRedfearn
        'Transverse Mercator projection calculations.
        'Uses the Redfearn formula.

#Region " Variable Declarations - All the variables used in this class." '=====================================================================================================================

        Public Coord As Coordinate  'This class reference is set before the method is used. For example InputCrs.Projection.Method.Coord = InputCrs.Coord - The Projection class now accesses the Coord values directly.

        'Variables used in the projection calculations:
        Dim E2 As Double 'Eccentricity squared
        Dim E4 As Double
        Dim E6 As Double
        Dim A0 As Double
        Dim A2 As Double
        Dim A4 As Double
        Dim A6 As Double

        'Variables updated for each projection calculation:
        Dim LatRad As Double 'The Latitude in radians
        Dim SinLat As Double
        Dim Sin2Lat As Double
        Dim Sin4Lat As Double
        Dim Sin6Lat As Double

        Dim Term1 As Double
        Dim Term2 As Double
        Dim Term3 As Double
        Dim Term4 As Double
        Dim Mdist As Double 'The Meridian distance.

        Dim Rho As Double
        Dim Nu As Double

        Dim CosLat As Double
        Dim CosLat2 As Double
        Dim CosLat3 As Double
        Dim CosLat4 As Double
        Dim CosLat5 As Double
        Dim CosLat6 As Double
        Dim CosLat7 As Double
        Dim CosLat8 As Double

        Dim DifLonRad As Double
        Dim DifLonRad2 As Double
        Dim DifLonRad3 As Double
        Dim DifLonRad4 As Double
        Dim DifLonRad5 As Double
        Dim DifLonRad6 As Double
        Dim DifLonRad7 As Double
        Dim DifLonRad8 As Double

        Dim Psi As Double
        Dim Psi2 As Double
        Dim Psi3 As Double
        Dim Psi4 As Double

        Dim TanLat As Double
        Dim TanLat2 As Double
        Dim TanLat4 As Double
        Dim TanLat6 As Double

        'These are Properties of the class:
        'Dim GridConv As Double 'Grid Convergence
        'Dim PointScale As Double 'Point Scale

        'Variables used in the Lat Long calculations:
        Dim N As Double
        Dim N2 As Double
        Dim N3 As Double
        Dim N4 As Double
        Dim G As Double
        Dim Sigma As Double

        'Variables updated for each Lat Long calculation:
        Dim Edash As Double
        Dim EdashOnK0 As Double
        Dim Ndash As Double
        Dim m As Double

        'These are also used in the Projection calculations:
        'Dim Term1 As Double
        'Dim Term2 As Double
        'Dim Term3 As Double
        'Dim Term4 As Double

        Dim Term5 As Double
        Dim FPLat As Double 'Foot Point Latitude
        Dim SinFPLat As Double
        Dim SecFPLat As Double
        Dim TanFPLat As Double
        Dim TanFPLat2 As Double
        Dim TanFPLat3 As Double
        Dim TanFPLat4 As Double
        Dim TanFPLat5 As Double
        Dim TanFPLat6 As Double

        'These are also used inthe Projection calculations:
        'Dim E2 As Double
        'Dim Rho As Double
        'Dim Nu As Double
        'Dim Psi As Double 'Nu / Rho
        'Dim Psi2 As Double
        'Dim Psi3 As Double
        'Dim Psi4 As Double

        Dim TonK0NuRho As Double
        Dim X As Double
        Dim X2 As Double
        Dim X3 As Double
        Dim X4 As Double
        Dim X5 As Double
        Dim X6 As Double
        Dim X7 As Double
        Dim E2onK02NuRho As Double
        Dim E2onK02NuRho2 As Double
        Dim E2onK02NuRho3 As Double
        Dim CMRad As Double


#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Properties - All the properties used in this class." '===========================================================================================================================

        Private _semiMajorAxis As Double = Double.NaN 'The semi major axis of the reference ellipsoid
        Property SemiMajorAxis As Double
            Get
                Return _semiMajorAxis
            End Get
            Set(value As Double)
                _semiMajorAxis = value
            End Set
        End Property

        Private _inverseFlattening As Double = Double.NaN 'The inverse flattening of the reference ellipsoid.
        Property InverseFlattening As Double
            Get
                Return _inverseFlattening
            End Get
            Set(value As Double)
                _inverseFlattening = value
                _flattening = 1.0# / _inverseFlattening
                _semiMinorAxis = _semiMajorAxis * (1.0# - _flattening)
            End Set
        End Property

        Private _flattening As Double = Double.NaN 'The flattening of the ellipsoid.
        Property Flattening As Double
            Get
                Return _flattening
            End Get
            Set(value As Double)
                _flattening = value
                _semiMinorAxis = _semiMajorAxis * (1.0# - _flattening)
            End Set
        End Property

        Private _semiMinorAxis As Double = Double.NaN 'The semi minor axis of the reference ellipsoid. If this value is entered, the InverseFlattening value is calculated.
        Property SemiMinorAxis As Double
            Get
                Return _semiMinorAxis
            End Get
            Set(value As Double)
                _semiMinorAxis = value
                If _semiMajorAxis = Double.NaN Then
                    'Cannot calculate inverse flattening
                Else
                    _inverseFlattening = _semiMajorAxis / (_semiMajorAxis - _semiMinorAxis)
                    _flattening = (_semiMajorAxis - _semiMinorAxis) / _semiMajorAxis
                End If
            End Set
        End Property


        Private _latitudeOfNaturalOrigin As Double = Double.NaN 'The Latitude of the natural origin (in degrees).
        Property LatitudeOfNaturalOrigin As Double
            Get
                Return _latitudeOfNaturalOrigin
            End Get
            Set(value As Double)
                _latitudeOfNaturalOrigin = value
            End Set
        End Property

        Private _longitudeOfNaturalOrigin As Double = Double.NaN 'The Longitude of the natural origin (in degrees).
        Property LongitudeOfNaturalOrigin As Double
            Get
                Return _longitudeOfNaturalOrigin
            End Get
            Set(value As Double)
                _longitudeOfNaturalOrigin = value
            End Set
        End Property

        Private _scaleFactorAtNaturalOrigin As Double = Double.NaN 'The scale factor at the natural origin.
        Property ScaleFactorAtNaturalOrigin As Double
            Get
                Return _scaleFactorAtNaturalOrigin
            End Get
            Set(value As Double)
                _scaleFactorAtNaturalOrigin = value
            End Set
        End Property


        Private _FalseEasting As Double = Double.NaN
        Property FalseEasting As Double
            Get
                Return _FalseEasting
            End Get
            Set(value As Double)
                _FalseEasting = value
            End Set
        End Property

        Private _FalseNorthing As Double = Double.NaN
        Property FalseNorthing As Double
            Get
                Return _FalseNorthing
            End Get
            Set(value As Double)
                _FalseNorthing = value
            End Set
        End Property

        Private _distanceCode As Integer = -1 'The UOM code
        Property DistanceCode As Integer
            Get
                Return _distanceCode
            End Get
            Set(value As Integer)
                _distanceCode = value
            End Set
        End Property

        'The units used to measure eastings, northings, FalseEasting and FalseNorthing
        Private _distanceUnits As String = ""
        Property DistanceUnits As String
            Get
                Return _distanceUnits
            End Get
            Set(value As String)
                _distanceUnits = value
            End Set
        End Property

        'NOTE: The Coord class is now used to access the Latitude, Longitude, Easting and Northing values.
        'Property _latitude As Double = Double.NaN 'The decimal Latitude of the location.
        'Property Latitude As Double
        '    Get
        '        Return _latitude
        '    End Get
        '    Set(value As Double)
        '        _latitude = value
        '    End Set
        'End Property

        'Private _longitude As Double = Double.NaN 'The decimal Longtude of the location.
        'Property Longitude As Double
        '    Get
        '        Return _longitude
        '    End Get
        '    Set(value As Double)
        '        _longitude = value
        '    End Set
        'End Property

        'Private _northing As Double = Double.NaN 'The Northing of the location.
        'Property Northing As Double
        '    Get
        '        Return _northing
        '    End Get
        '    Set(value As Double)
        '        _northing = value
        '    End Set
        'End Property

        'Private _easting As Double = Double.NaN 'The Easting of the location.
        'Property Easting As Double
        '    Get
        '        Return _easting
        '    End Get
        '    Set(value As Double)
        '        _easting = value
        '    End Set
        'End Property

        Private _gridConvergence As Double = Double.NaN 'The Grid Convergence at the location.
        Property GridConvergence As Double
            Get
                Return _gridConvergence
            End Get
            Set(value As Double)
                _gridConvergence = value
            End Set
        End Property

        Private _pointScale As Double = Double.NaN 'The Point Scale at the location.
        Property PointScale As Double
            Get
                Return _pointScale
            End Get
            Set(value As Double)
                _pointScale = value
            End Set
        End Property


#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Methods - The main actions performed by this class." '===========================================================================================================================

        Public Sub Clear()
            'Clear the properties
            SemiMajorAxis = Double.NaN
            InverseFlattening = Double.NaN
            Flattening = Double.NaN
            SemiMinorAxis = Double.NaN
            LatitudeOfNaturalOrigin = Double.NaN
            LongitudeOfNaturalOrigin = Double.NaN
            ScaleFactorAtNaturalOrigin = Double.NaN
            FalseEasting = Double.NaN
            FalseNorthing = Double.NaN
            DistanceCode = -1
            DistanceUnits = ""
        End Sub

        Public Sub SetParameter(ParameterCode As Integer, ParameterValue As Double)
            'Sets the Transverse Mercator property from a PropertyCode and a PropertyValue.
            Select Case ParameterCode
                Case 8801 'LatitudeOfNaturalOrigin
                    LatitudeOfNaturalOrigin = ParameterValue
                Case 8802 'LongitudeOfNaturalOrigin
                    LongitudeOfNaturalOrigin = ParameterValue
                Case 8805 'ScaleFactorAtNaturalOrigin
                    ScaleFactorAtNaturalOrigin = ParameterValue
                Case 8806 'FalseEasting
                    FalseEasting = ParameterValue
                Case 8807 'FalseNorthing
                    FalseNorthing = ParameterValue
                Case Else
                    RaiseEvent ErrorMessage("Unknown property code: " & ParameterCode & vbCrLf)
            End Select
        End Sub

        Public Function ValidProperty(PropertyName As String) As Boolean
            'Return True if PropertyName is a valid property of this projection method.

            Select Case PropertyName
                Case "SemiMajorAxis"
                    Return True
                Case "InverseFlattening"
                    Return True
                Case "Flattening"
                    Return True
                Case "SemiMinorAxis"
                    Return True
                Case "LatitudeOfNaturalOrigin"
                    Return True
                Case "LongitudeOfNaturalOrigin"
                    Return True
                Case "ScaleFactorAtNaturalOrigin"
                    Return True
                Case "FalseEasting"
                    Return True
                Case "FalseNorthing"
                    Return True
                Case Else
                    Return False
            End Select

        End Function

        'Private Sub UpdateVariables()
        Public Sub UpdateVariables()
            'Update the variables used in the projection calculations.
            E2 = (2.0# * Flattening) - (Flattening * Flattening)
            E4 = E2 * E2
            E6 = E2 * E4
            A0 = 1.0# - (E2 / 4.0#) - ((3.0# * E4) / 64.0#) - ((5.0# * E6) / 256.0#)
            A2 = (3.0# / 8.0#) * (E2 + (E4 / 4.0#) + ((15.0# * E6) / 128.0#))
            A4 = (15.0# / 256.0#) * (E4 + ((3.0# * E6) / 4.0#))
            A6 = (35.0# * E6) / 3072.0#

            'Update the variables used in the Lat Long calculations:
            N = (SemiMajorAxis - SemiMinorAxis) / (SemiMajorAxis + SemiMinorAxis)
            N2 = N * N
            N3 = N2 * N
            N4 = N2 * N2
            G = SemiMajorAxis * (1.0# - N) * (1 - N2) * (1.0# + (9.0# * N2) / 4.0# + (225.0# * N4) / 64.0#) * Math.PI / 180.0#
        End Sub

        'Public Sub LatLongToNorthEast()
        Public Sub LongLatToEastNorth()
            'Calculate the Easting and Northing values from the Latitude and Longitude.

            'LatRad = (Latitude / 180.0#) * Math.PI
            LatRad = (Coord.Latitude / 180.0#) * Math.PI
            SinLat = System.Math.Sin(LatRad)
            Sin2Lat = System.Math.Sin(2.0# * LatRad)
            Sin4Lat = System.Math.Sin(4.0# * LatRad)
            Sin6Lat = System.Math.Sin(6.0# * LatRad)

            Term1 = SemiMajorAxis * A0 * LatRad
            Term2 = -SemiMajorAxis * A2 * Sin2Lat
            Term3 = SemiMajorAxis * A4 * Sin4Lat
            Term4 = -SemiMajorAxis * A6 * Sin6Lat
            Mdist = Term1 + Term2 + Term3 + Term4 'The Meridian distance.

            Rho = SemiMajorAxis * (1.0# - E2) / (1.0# - (E2 - SinLat * SinLat)) ^ 1.5
            Nu = SemiMajorAxis / (1.0# - (E2 * SinLat * SinLat)) ^ 0.5

            'DifLonRad = ((Longitude - LongitudeOfNaturalOrigin) / 180.0#) * Math.PI
            DifLonRad = ((Coord.Longitude - LongitudeOfNaturalOrigin) / 180.0#) * Math.PI
            DifLonRad2 = DifLonRad * DifLonRad
            DifLonRad3 = DifLonRad2 * DifLonRad
            DifLonRad4 = DifLonRad2 * DifLonRad2
            DifLonRad5 = DifLonRad4 * DifLonRad
            DifLonRad6 = DifLonRad3 * DifLonRad3
            DifLonRad7 = DifLonRad3 * DifLonRad4
            DifLonRad8 = DifLonRad4 * DifLonRad4

            CosLat = System.Math.Cos(LatRad)
            CosLat2 = CosLat * CosLat
            CosLat3 = CosLat2 * CosLat
            CosLat4 = CosLat2 * CosLat2
            CosLat5 = CosLat3 * CosLat2
            CosLat6 = CosLat3 * CosLat3
            CosLat7 = CosLat4 * CosLat3
            CosLat8 = CosLat4 * CosLat4

            Psi = Nu / Rho
            Psi2 = Psi * Psi
            Psi3 = Psi2 * Psi
            Psi4 = Psi2 * Psi2

            TanLat = System.Math.Tan(LatRad)
            TanLat2 = TanLat * TanLat
            TanLat4 = TanLat2 * TanLat2
            TanLat6 = TanLat4 * TanLat2

            'Calculate Easting:
            Term1 = Nu * DifLonRad * CosLat
            Term2 = Nu * DifLonRad3 * CosLat3 * (Psi - TanLat2) / 6.0#
            Term3 = Nu * DifLonRad5 * CosLat5 * (4.0# * Psi3 * (1.0# - 6.0# * TanLat2) + Psi2 * (1.0# + 8.0# * TanLat2) - Psi * (2.0# * TanLat2) + TanLat4) / 120.0#
            Term4 = Nu * DifLonRad7 * CosLat7 * (61.0# - 479.0# * TanLat2 * 179.0# * TanLat4 - TanLat6) / 5040.0#
            'Easting = (Term1 + Term2 + Term3 + Term4) * ScaleFactorAtNaturalOrigin + FalseEasting
            Coord.Easting = (Term1 + Term2 + Term3 + Term4) * ScaleFactorAtNaturalOrigin + FalseEasting

            'Calculate Northing:
            Term1 = Nu * SinLat * DifLonRad2 * CosLat / 2.0#
            Term2 = Nu * SinLat * DifLonRad4 * CosLat3 * (4.0# * Psi2 + Psi - TanLat2) / 24.0#
            Term3 = Nu * SinLat * DifLonRad6 * CosLat5 * (8.0# * Psi4 * (11.0# - 24.0# * TanLat2) - 28.0# * Psi3 * (1.0# - 6.0# * TanLat2) + Psi2 * (1.0# - 32.0# * TanLat2) - Psi * (2.0# * TanLat2) + TanLat4) / 720.0#
            Term4 = Nu * SinLat * DifLonRad8 * CosLat7 * (1385.0# - 3111.0# * TanLat2 + 543.0# * TanLat4 - TanLat6) / 40320.0#
            'Northing = (Mdist + Term1 + Term2 + Term3 + Term4) * ScaleFactorAtNaturalOrigin + FalseNorthing
            Coord.Northing = (Mdist + Term1 + Term2 + Term3 + Term4) * ScaleFactorAtNaturalOrigin + FalseNorthing

            'Calculate Grid Convergence:
            Term1 = -SinLat * DifLonRad
            Term2 = -SinLat * DifLonRad3 * CosLat2 * (2.0# * Psi2 - Psi) / 3.0#
            Term3 = -SinLat * DifLonRad5 * CosLat4 * (Psi4 * (11.0# - 24.0# * TanLat2) - Psi3 * (11.0# - 36.0# * TanLat2) + 2.0# * Psi2 * (1.0# - 7.0# * TanLat2) + Psi * TanLat2) / 15.0#
            Term4 = SinLat * DifLonRad7 * CosLat6 * (17.0# - 26.0# * TanLat2 + 2.0# * TanLat4) / 315.0#
            GridConvergence = (Term1 + Term2 + Term3 + Term4) / Math.PI * 180.0# 'in degrees.

            'Calculate the Point Scale
            Term1 = 1.0# + (DifLonRad2 * CosLat2 * Psi) / 2.0#
            Term2 = DifLonRad4 * CosLat4 * (4.0# * Psi3 * (1.0# - 6.0# * TanLat2) + Psi2 * (1.0# + 24.0# * TanLat2) - 4.0# * Psi * TanLat2) / 24.0#
            Term3 = DifLonRad6 * CosLat6 * (61.0# - 148.0# * TanLat2 + 16.0# * TanLat4) / 720.0#
            PointScale = Term1 + Term2 + Term3

        End Sub

        'Public Sub NorthEastToLatLong()
        Public Sub EastNorthToLongLat()
            'Calculate the Latitude and Longitude values from the Easting and Northing.

            'Edash = Easting - FalseEasting
            Edash = Coord.Easting - FalseEasting
            EdashOnK0 = Edash / ScaleFactorAtNaturalOrigin

            'Ndash = Northing - FalseNorthing
            Ndash = Coord.Northing - FalseNorthing
            m = Ndash / ScaleFactorAtNaturalOrigin

            Sigma = (m * Math.PI) / (G * 180.0#)

            Term1 = Sigma
            Term2 = ((3.0# * N / 2.0#) - (27.0# * N3 / 32.0#)) * System.Math.Sin(Sigma * 2.0#)
            Term3 = ((21.0# * N2 / 16.0#) - (55.0# * N4 / 32.0#)) * System.Math.Sin(Sigma * 4.0#)
            Term4 = (151.0# * N3) * System.Math.Sin(Sigma * 6.0#) / 96.0#
            Term5 = 1097.0# * N4 * System.Math.Sin(Sigma * 8.0#) / 512.0#
            FPLat = Term1 + Term2 + Term3 + Term4 + Term5 'The Foot point latitude.

            SinFPLat = System.Math.Sin(FPLat)
            SecFPLat = 1.0# / System.Math.Cos(FPLat)

            TanFPLat = System.Math.Tan(FPLat)
            TanFPLat2 = TanFPLat * TanFPLat
            TanFPLat3 = TanFPLat2 * TanFPLat
            TanFPLat4 = TanFPLat2 * TanFPLat2
            TanFPLat5 = TanFPLat3 * TanFPLat2
            TanFPLat6 = TanFPLat3 * TanFPLat3

            Rho = SemiMajorAxis * (1.0# - E2) / (1.0# - E2 * SinFPLat * SinFPLat) ^ 1.5
            Nu = SemiMajorAxis / (1.0# - E2 * SinFPLat * SinFPLat) ^ 0.5

            TonK0NuRho = TanFPLat / (ScaleFactorAtNaturalOrigin * Nu)

            X = EdashOnK0 / Nu
            X2 = X * X
            X3 = X2 * X
            X4 = X2 * X2
            X5 = X3 * X2
            X6 = X3 * X3
            X7 = X4 * X3

            E2onK02NuRho = (EdashOnK0 * EdashOnK0) / (Rho * Nu)
            E2onK02NuRho2 = E2onK02NuRho * E2onK02NuRho
            E2onK02NuRho3 = E2onK02NuRho2 * E2onK02NuRho

            Psi = Nu / Rho
            Psi2 = Psi * Psi
            Psi3 = Psi2 * Psi
            Psi4 = Psi2 * Psi2

            'Calculate Latitude:
            Term1 = -((TanFPLat / (ScaleFactorAtNaturalOrigin * Rho)) * X * Edash / 2.0#)
            Term2 = (TanFPLat / (ScaleFactorAtNaturalOrigin * Rho)) * (X3 * Edash / 24.0#) * (-4.0# * Psi2 + 9.0# * Psi * (1.0# - TanFPLat2) + 12.0# * TanFPLat2)
            Term3 = -(TanFPLat / (ScaleFactorAtNaturalOrigin * Rho)) * (X5 * Edash / 720.0#) * (8.0# * Psi4 * (11.0# - 24.0# * TanFPLat2) - 12.0# * Psi3 * (21.0# - 71.0# * TanFPLat2) + 15.0# * Psi2 * (15.0# - 98.0# * TanFPLat2 + 15.0# * TanFPLat4) + 180.0# * Psi * (5.0# * TanFPLat2 - 3.0# * TanFPLat4) + 360.0# * TanFPLat4)
            Term4 = (TanFPLat / (ScaleFactorAtNaturalOrigin * Rho)) * (X7 * Edash / 40320.0#) * (1385.0# + 3633.0# * TanFPLat2 + 4095.0# * TanFPLat4 + 1575.0# * TanFPLat6)
            'Latitude = (FPLat + Term1 + Term2 + Term3 + Term4) / Math.PI * 180.0# 'Latitude in degrees.
            Coord.Latitude = (FPLat + Term1 + Term2 + Term3 + Term4) / Math.PI * 180.0# 'Latitude in degrees.

            'Calculate Longitude:
            CMRad = LongitudeOfNaturalOrigin / 180.0# * Math.PI 'Central Meridian in radians.
            Term1 = SecFPLat * X
            Term2 = -SecFPLat * (X3 / 6.0#) * (Psi + 2.0# * TanFPLat2)
            Term3 = SecFPLat * (X5 / 120.0#) * (-4.0# * Psi3 * (1.0# - 6.0# * TanFPLat2) + Psi2 * (9.0# - 68.0# * TanFPLat2) + 72.0# * Psi * TanFPLat2 + 24.0# * TanFPLat4)
            Term4 = -SecFPLat * (X7 / 5040.0#) * (61.0# + 662.0# * TanFPLat2 + 1320.0# * TanFPLat4 + 720.0# * TanFPLat6)
            'Longitude = (CMRad + Term1 + Term2 + Term3 + Term4) / Math.PI * 180.0# 'Longitude in degrees.
            Coord.Longitude = (CMRad + Term1 + Term2 + Term3 + Term4) / Math.PI * 180.0# 'Longitude in degrees.

            'Calculate grid convergence:
            Term1 = -TanFPLat * X
            Term2 = (TanFPLat * X3 / 3.0#) * (-2.0# * Psi2 + 3.0# * Psi + TanFPLat2)
            Term3 = -(TanFPLat * X5 / 15.0#) * (Psi4 * (11.0# - 24.0# * TanFPLat2) - 3.0# * Psi3 * (8.0# - 23.0# * TanFPLat2) + 5.0# * Psi2 * (3.0# - 14.0# * TanFPLat2) + 30.0# * Psi * TanFPLat2 + 3.0# * TanFPLat4)
            Term4 = (TanFPLat * X7 / 315.0#) * (17.0# + 77.0# * TanFPLat2 + 105.0# * TanFPLat4 + 45.0# * TanFPLat6)
            GridConvergence = (Term1 + Term2 + Term3 + Term4) / Math.PI * 180.0# 'Grid Convergence in degrees.

            'Calculate point scale:
            Term1 = 1.0# + E2onK02NuRho / 2.0#
            Term2 = (E2onK02NuRho2 / 24.0#) * (4.0# * Psi * (1.0# - 6.0# * TanFPLat2) - 3.0# * (1.0# - 16.0# * TanFPLat2) - 24.0# * TanFPLat2 / Psi)
            Term3 = E2onK02NuRho3 / 720.0#
            PointScale = (Term1 + Term2 + Term3) * ScaleFactorAtNaturalOrigin

        End Sub


#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================
        Event ErrorMessage(ByVal Msg As String) 'Send an error message.
        Event Message(ByVal Msg As String) 'Send a normal message.
#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------

    End Class 'TransverseMercatorRedfearn

End Class 'Projection

Public Class DatumTransOp
    'Stores information about a Datum Transformation operation.

    Private _name As String = "" 'The name of a the coordinate operation used for the datum transformation.
    Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Private _type As String = "" 'The type of coordinate operation.
    Property Type As String
        Get
            Return _type
        End Get
        Set(value As String)
            _type = value
        End Set
    End Property

    Private _code As Integer = -1 'The EPSG code for the coordinate operation.
    Property Code As Integer
        Get
            Return _code
        End Get
        Set(value As Integer)
            _code = value
        End Set
    End Property

    Private _accuracy As Single = Single.NaN 'The accuracy of the coordinate operation.
    Property Accuracy As Single
        Get
            Return _accuracy
        End Get
        Set(value As Single)
            _accuracy = value
        End Set
    End Property

    Private _deprecated As Boolean = False 'True if the coordinate operation has been deprecated.
    Property Deprecated As Boolean
        Get
            Return _deprecated
        End Get
        Set(value As Boolean)
            _deprecated = value
        End Set
    End Property

    Private _version As String = "" 'The version of the coordinate operation.
    Property Version As String
        Get
            Return _version
        End Get
        Set(value As String)
            _version = value
        End Set
    End Property

    Private _revisionDate As Date = Date.MinValue 'The date of the revision of the coordinate operation.
    Property RevisionDate As Date
        Get
            Return _revisionDate
        End Get
        Set(value As Date)
            _revisionDate = value
        End Set
    End Property

    Private _sourceCrsLevel As Integer = 0 'The BaseCrs level of the Source CRS. The top level is 0. If the CRS has a BaseCrs, that has level 1. If the BaseCrs has a BaseCrs, that has level 2, etc.
    Property SourceCrsLevel As Integer
        Get
            Return _sourceCrsLevel
        End Get
        Set(value As Integer)
            _sourceCrsLevel = value
        End Set
    End Property

    Property _sourceCrsCode As Integer = -1 'The EPSG code for the Source CRS.
    Property SourceCrsCode As Integer
        Get
            Return _sourceCrsCode
        End Get
        Set(value As Integer)
            _sourceCrsCode = value
        End Set
    End Property

    Property _targetCrsLevel As Integer = 0 'The BaseCrs level of the Target CRS. The top level is 0. If the CRS has a BaseCrs, that has level 1. If the BaseCrs has a BaseCrs, that has level 2, etc.
    Property TargetCrsLevel As Integer
        Get
            Return _targetCrsLevel
        End Get
        Set(value As Integer)
            _targetCrsLevel = value
        End Set
    End Property

    Private _targetCrsCode As Integer = -1 'The EPSG code for the Target CRS.
    Property TargetCrsCode As Integer
        Get
            Return _targetCrsCode
        End Get
        Set(value As Integer)
            _targetCrsCode = value
        End Set
    End Property

    Private _reversible As Boolean = True 'If True, the coordinate operation is reversible.
    Property Reversible As Boolean
        Get
            Return _reversible
        End Get
        Set(value As Boolean)
            _reversible = value
        End Set
    End Property

    Private _applyReverse As Boolean = False 'If True the reverse coordinate transformation should be applied.
    Property ApplyReverse As Boolean
        Get
            Return _applyReverse
        End Get
        Set(value As Boolean)
            _applyReverse = value
        End Set
    End Property

    Private _methodName As String = "" 'The name of the Coordinate Operation Method corresponding to the Coordinate Operation.
    Property MethodName As String
        Get
            Return _methodName
        End Get
        Set(value As String)
            _methodName = value
        End Set
    End Property

    Private _methodCode As Integer = -1 'The Code of the Coordinate Operation Method corresponding to the Coordinate Operation.
    Property MethodCode As Integer
        Get
            Return _methodCode
        End Get
        Set(value As Integer)
            _methodCode = value
        End Set
    End Property

End Class

Public Class clsDatumTrans
    'Information used for the Datum Transformation calculation.

#Region " Variable Declarations - All the variables used in this class." '=====================================================================================================================

    Public SourceCoord As Coordinate 'This will be set to InputCrs.Coord
    Public Wgs84Coord As New Coordinate
    Public TargetCoord As Coordinate 'This will be set to OutputCrs.Coord

    'Information about the Direct Datum Transformation operation:
    Public DirectCoordOp As New CoordinateOperation
    Public DirectCoordOpMethod As New CoordOpMethod
    Public DirectCoordOpParamUseList As New List(Of CoordOpParamUsage)
    Public DirectCoordOpParamList As New List(Of CoordOpParameter)
    Public DirectCoordOpParamValList As New List(Of CoordOpParamValue)

    Public DirectMethod As Object

    'Information about the Input to WGS 84 Datum Transformation operation:
    Public InputToWgs84CoordOp As New CoordinateOperation
    Public InputToWgs84CoordOpMethod As New CoordOpMethod
    Public InputToWgs84CoordOpParamUseList As New List(Of CoordOpParamUsage)
    Public InputToWgs84CoordOpParamList As New List(Of CoordOpParameter)
    Public InputToWgs84CoordOpParamValList As New List(Of CoordOpParamValue)

    Public InputToWgs84Method As Object

    'Information about the WGS 84 to Output Datum Transformation operation:
    Public Wgs84ToOutputCoordOp As New CoordinateOperation
    Public Wgs84ToOutputCoordOpMethod As New CoordOpMethod
    Public Wgs84ToOutputCoordOpParamUseList As New List(Of CoordOpParamUsage)
    Public Wgs84ToOutputCoordOpParamList As New List(Of CoordOpParameter)
    Public Wgs84ToOutputCoordOpParamValList As New List(Of CoordOpParamValue)

    Public Wgs84ToOutputMethod As Object


#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Properties - All the properties used in this class." '===============================================================================================================================

    Private _epsgDatabasePath = "" 'The path of the EPSG Database. This database contains a comprehensive set of coordinate reference system parameters.
    Property EpsgDatabasePath
        Get
            Return _epsgDatabasePath
        End Get
        Set(value)
            _epsgDatabasePath = value
        End Set
    End Property

    Enum enumType
        None
        Direct
        ViaWgs84
    End Enum

    Private _type As enumType = enumType.Direct 'The type of of Datum Transformation (None, Direct, ViaWgs84)
    Property Type As enumType
        Get
            Return _type
        End Get
        Set(value As enumType)
            _type = value
        End Set
    End Property

    Private _directMethodName As String = "" 'The name of the direct datum transformation method.
    Property DirectMethodName As String
        Get
            Return _directMethodName
        End Get
        Set(value As String)
            _directMethodName = value
            ApplyDirectMethodName()
        End Set
    End Property

    Private _directMethodApplyReverse As Boolean = False 'If True the reverse method is applied.
    Property DirectMethodApplyReverse As Boolean
        Get
            Return _directMethodApplyReverse
        End Get
        Set(value As Boolean)
            _directMethodApplyReverse = value
        End Set
    End Property

    Private _inputToWgs84MethodName As String = "" 'The name of the Input to WGS84 datum transformation method.
    Property InputToWgs84MethodName As String
        Get
            Return _inputToWgs84MethodName
        End Get
        Set(value As String)
            _inputToWgs84MethodName = value
            ApplyInputToWgs84MethodName()
        End Set
    End Property

    Private _inputToWgs84MethodApplyReverse As Boolean = False 'If True the reverse method is applied.
    Property InputToWgs84MethodApplyReverse As Boolean
        Get
            Return _inputToWgs84MethodApplyReverse
        End Get
        Set(value As Boolean)
            _directMethodApplyReverse = value
        End Set
    End Property

    Private _wgs84ToOutputMethodName As String = "" 'The name of the WGS84 to Output datum transformation method.
    Property Wgs84ToOutputMethodName As String
        Get
            Return _wgs84ToOutputMethodName
        End Get
        Set(value As String)
            _wgs84ToOutputMethodName = value
            ApplyWgs84ToOutputMethodName()
        End Set
    End Property

    Private _wgs84ToOutputMethodApplyReverse As Boolean = False 'If True the reverse method is applied.
    Property Wgs84ToOutputMethodApplyReverse As Boolean
        Get
            Return _wgs84ToOutputMethodApplyReverse
        End Get
        Set(value As Boolean)
            _wgs84ToOutputMethodApplyReverse = value
        End Set
    End Property


#End Region 'Properties -----------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Methods - The main actions performed by this class." '===============================================================================================================================

    Public Sub Clear()
        'Clear the current settings.

        DirectCoordOp.Clear()
        DirectCoordOpMethod.Clear()
        DirectCoordOpParamUseList.Clear()
        DirectCoordOpParamList.Clear()
        DirectCoordOpParamValList.Clear()
        DirectMethod = Nothing

        InputToWgs84CoordOp.Clear()
        InputToWgs84CoordOpMethod.Clear()
        InputToWgs84CoordOpParamUseList.Clear()
        InputToWgs84CoordOpParamList.Clear()
        InputToWgs84CoordOpParamValList.Clear()
        InputToWgs84Method = Nothing

        Wgs84ToOutputCoordOp.Clear()
        Wgs84ToOutputCoordOpMethod.Clear()
        Wgs84ToOutputCoordOpParamUseList.Clear()
        Wgs84ToOutputCoordOpParamList.Clear()
        Wgs84ToOutputCoordOpParamValList.Clear()
        Wgs84ToOutputMethod = Nothing
    End Sub

    Public Sub GetDirectDatumTransCoordOp(OpCode As Integer)
        'Get the Direct Datum Transformation coordinate operation corresponding to OpCode.

        If EpsgDatabasePath = "" Then
            RaiseEvent ErrorMessage("No EPSG database has been selected." & vbCrLf)
            Exit Sub
        End If

        If Not System.IO.File.Exists(EpsgDatabasePath) Then
            RaiseEvent ErrorMessage("Selected EPSG database can not be found." & vbCrLf)
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + EpsgDatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        'Clear the existing data.
        DirectCoordOp.Clear()
        DirectCoordOpMethod.Clear()
        DirectCoordOpParamUseList.Clear()
        DirectCoordOpParamList.Clear()
        DirectCoordOpParamValList.Clear()

        'Get list of Input Source Coordinate Operations:
        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter("Select * From Coordinate_Operation Where COORD_OP_CODE = " & OpCode.ToString, conn)
        Dim ds As DataSet = New DataSet
        da.Fill(ds, "CoordOp")

        If ds.Tables("CoordOp").Rows.Count = 0 Then
            RaiseEvent ErrorMessage("There are no Coordinate Operation records for code number: " & OpCode & vbCrLf)
        ElseIf ds.Tables("CoordOp").Rows.Count = 1 Then
            DirectCoordOp.Code = OpCode
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("COORD_OP_NAME")) Then DirectCoordOp.Name = "" Else DirectCoordOp.Name = ds.Tables("CoordOp").Rows(0).Item("COORD_OP_NAME")

            Select Case ds.Tables("CoordOp").Rows(0).Item("COORD_OP_TYPE")
                Case "conversion"
                    DirectCoordOp.Type = CoordinateOperation.OperationType.conversion
                Case "transformation"
                    DirectCoordOp.Type = CoordinateOperation.OperationType.transformation
                Case "point motion operation"
                    DirectCoordOp.Type = CoordinateOperation.OperationType.pointMotionOperation
                Case "concatenated operation"
                    DirectCoordOp.Type = CoordinateOperation.OperationType.concatenatedOperation
                Case Else
                    RaiseEvent ErrorMessage("Unknown Target coordinate operation type: " & ds.Tables("CoordOp").Rows(0).Item("COORD_OP_TYPE") & vbCrLf)
                    DirectCoordOp.Type = CoordinateOperation.OperationType.conversion
            End Select

            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("SOURCE_CRS_CODE")) Then DirectCoordOp.SourceCrsCode = -1 Else DirectCoordOp.SourceCrsCode = ds.Tables("CoordOp").Rows(0).Item("SOURCE_CRS_CODE")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("TARGET_CRS_CODE")) Then DirectCoordOp.TargetCrsCode = -1 Else DirectCoordOp.TargetCrsCode = ds.Tables("CoordOp").Rows(0).Item("TARGET_CRS_CODE")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("COORD_TFM_VERSION")) Then DirectCoordOp.Version = "" Else DirectCoordOp.Version = ds.Tables("CoordOp").Rows(0).Item("COORD_TFM_VERSION")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("COORD_OP_VARIANT")) Then DirectCoordOp.OpVariant = -1 Else DirectCoordOp.OpVariant = ds.Tables("CoordOp").Rows(0).Item("COORD_OP_VARIANT")
            'AREA_OF_USE_CODE has been deprecated.
            'COORD_OP_SCOPE has been deprecated.
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("COORD_OP_ACCURACY")) Then DirectCoordOp.Accuracy = Single.NaN Else DirectCoordOp.Accuracy = ds.Tables("CoordOp").Rows(0).Item("COORD_OP_ACCURACY")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("COORD_OP_METHOD_CODE")) Then DirectCoordOp.MethodCode = -1 Else DirectCoordOp.MethodCode = ds.Tables("CoordOp").Rows(0).Item("COORD_OP_METHOD_CODE")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("UOM_CODE_SOURCE_COORD_DIFF")) Then DirectCoordOp.UomSourceCoordDiffCode = -1 Else DirectCoordOp.UomSourceCoordDiffCode = ds.Tables("CoordOp").Rows(0).Item("UOM_CODE_SOURCE_COORD_DIFF")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("UOM_CODE_TARGET_COORD_DIFF")) Then DirectCoordOp.UomTargetCoordDiffCode = -1 Else DirectCoordOp.UomTargetCoordDiffCode = ds.Tables("CoordOp").Rows(0).Item("UOM_CODE_TARGET_COORD_DIFF")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("REMARKS")) Then DirectCoordOp.Remarks = "" Else DirectCoordOp.Remarks = ds.Tables("CoordOp").Rows(0).Item("REMARKS")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("INFORMATION_SOURCE")) Then DirectCoordOp.InfoSource = "" Else DirectCoordOp.InfoSource = ds.Tables("CoordOp").Rows(0).Item("INFORMATION_SOURCE")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("DATA_SOURCE")) Then DirectCoordOp.DataSource = "" Else DirectCoordOp.DataSource = ds.Tables("CoordOp").Rows(0).Item("DATA_SOURCE")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("REVISION_DATE")) Then DirectCoordOp.RevisionDate = Date.MinValue Else DirectCoordOp.RevisionDate = ds.Tables("CoordOp").Rows(0).Item("REVISION_DATE")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("CHANGE_ID")) Then DirectCoordOp.ChangeID = "" Else DirectCoordOp.ChangeID = ds.Tables("CoordOp").Rows(0).Item("CHANGE_ID")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("SHOW_OPERATION")) Then DirectCoordOp.Show = True Else DirectCoordOp.Show = ds.Tables("CoordOp").Rows(0).Item("SHOW_OPERATION")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("DEPRECATED")) Then DirectCoordOp.Deprecated = False Else DirectCoordOp.Deprecated = ds.Tables("CoordOp").Rows(0).Item("DEPRECATED")

            'Get Coordinate Operation Method:
            Dim MethodCode As Integer = DirectCoordOp.MethodCode
            da.SelectCommand.CommandText = "Select * From [Coordinate_Operation Method] Where COORD_OP_METHOD_CODE = " & MethodCode.ToString
            da.Fill(ds, "Method")
            If ds.Tables("Method").Rows.Count = 0 Then
                RaiseEvent ErrorMessage("There are no Coordinate Operation Methods for Method code number: " & MethodCode.ToString & vbCrLf)
            ElseIf ds.Tables("Method").Rows.Count = 1 Then
                DirectCoordOpMethod.Code = MethodCode
                If IsDBNull(ds.Tables("Method").Rows(0).Item("COORD_OP_METHOD_NAME")) Then DirectCoordOpMethod.Name = "" Else DirectCoordOpMethod.Name = ds.Tables("Method").Rows(0).Item("COORD_OP_METHOD_NAME")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("REVERSE_OP")) Then DirectCoordOpMethod.ReverseOp = False Else DirectCoordOpMethod.ReverseOp = ds.Tables("Method").Rows(0).Item("REVERSE_OP")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("FORMULA")) Then DirectCoordOpMethod.Formula = "" Else DirectCoordOpMethod.Formula = ds.Tables("Method").Rows(0).Item("FORMULA")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("EXAMPLE")) Then DirectCoordOpMethod.Example = "" Else DirectCoordOpMethod.Example = ds.Tables("Method").Rows(0).Item("EXAMPLE")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("REMARKS")) Then DirectCoordOpMethod.Remarks = "" Else DirectCoordOpMethod.Remarks = ds.Tables("Method").Rows(0).Item("REMARKS")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("INFORMATION_SOURCE")) Then DirectCoordOpMethod.InfoSource = "" Else DirectCoordOpMethod.InfoSource = ds.Tables("Method").Rows(0).Item("INFORMATION_SOURCE")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("DATA_SOURCE")) Then DirectCoordOpMethod.DataSource = "" Else DirectCoordOpMethod.DataSource = ds.Tables("Method").Rows(0).Item("DATA_SOURCE")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("REVISION_DATE")) Then DirectCoordOpMethod.RevisionDate = Date.MinValue Else DirectCoordOpMethod.RevisionDate = ds.Tables("Method").Rows(0).Item("REVISION_DATE")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("CHANGE_ID")) Then DirectCoordOpMethod.ChangeID = "" Else DirectCoordOpMethod.ChangeID = ds.Tables("Method").Rows(0).Item("CHANGE_ID")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("DEPRECATED")) Then DirectCoordOpMethod.Deprecated = False Else DirectCoordOpMethod.Deprecated = ds.Tables("Method").Rows(0).Item("DEPRECATED")
                'Projection.MethodName = ProjectionCoordOpMethod.Name
                DirectMethodName = DirectCoordOpMethod.Name 'This will create corresponding Datum Transformation class in DirectMethod
            Else
                RaiseEvent ErrorMessage("There are " & ds.Tables("Method").Rows.Count & " Coordinate Operation Methods with the code: " & MethodCode.ToString & vbCrLf)
            End If

            'Get the Coordinate Operation Parameter Usage:
            da.SelectCommand.CommandText = "Select * From [Coordinate_Operation Parameter Usage] Where COORD_OP_METHOD_CODE = " & MethodCode.ToString & " Order By SORT_ORDER"
            da.Fill(ds, "Usage")
            If ds.Tables("Usage").Rows.Count = 0 Then
                RaiseEvent ErrorMessage("There are no Coordinate Operation Parameter Usage records for Method code number: " & MethodCode.ToString & vbCrLf)
            Else
                Dim ParamCode As Integer
                Dim ParamVal As Double
                For Each Item As DataRow In ds.Tables("Usage").Rows
                    Dim NewUsage As New CoordOpParamUsage
                    NewUsage.MethodCode = MethodCode
                    'If IsDBNull(ds.Tables("Usage").Rows(0).Item("PARAMETER_CODE")) Then NewUsage.ParameterCode = -1 Else NewUsage.ParameterCode = ds.Tables("Usage").Rows(0).Item("PARAMETER_CODE")
                    If IsDBNull(Item("PARAMETER_CODE")) Then NewUsage.ParameterCode = -1 Else NewUsage.ParameterCode = Item("PARAMETER_CODE")
                    If IsDBNull(Item("SORT_ORDER")) Then NewUsage.SortOrder = 0 Else NewUsage.SortOrder = Item("SORT_ORDER")
                    If IsDBNull(Item("PARAM_SIGN_REVERSAL")) Then NewUsage.SignReversal = "" Else NewUsage.SignReversal = Item("PARAM_SIGN_REVERSAL")
                    DirectCoordOpParamUseList.Add(NewUsage)

                    'Get the corresponding Coordinate Operation Parameter information:
                    da.SelectCommand.CommandText = "Select * From [Coordinate_Operation Parameter] Where PARAMETER_CODE = " & NewUsage.ParameterCode.ToString
                    da.Fill(ds, "OpParameter")
                    If ds.Tables("OpParameter").Rows.Count = 0 Then
                        RaiseEvent ErrorMessage("There are no Coordinate Operation Parameter records for Parameter code number: " & NewUsage.ParameterCode.ToString & vbCrLf)
                    ElseIf ds.Tables("OpParameter").Rows.Count = 1 Then
                        Dim NewParameter As New CoordOpParameter
                        NewParameter.Code = NewUsage.ParameterCode
                        ParamCode = NewParameter.Code
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("PARAMETER_NAME")) Then NewParameter.Name = "" Else NewParameter.Name = ds.Tables("OpParameter").Rows(0).Item("PARAMETER_NAME")
                        'ParamName = NewParameter.Name
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("DESCRIPTION")) Then NewParameter.Description = "" Else NewParameter.Description = ds.Tables("OpParameter").Rows(0).Item("DESCRIPTION")
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("INFORMATION_SOURCE")) Then NewParameter.InfoSource = "" Else NewParameter.InfoSource = ds.Tables("OpParameter").Rows(0).Item("INFORMATION_SOURCE")
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("DATA_SOURCE")) Then NewParameter.DataSource = "" Else NewParameter.DataSource = ds.Tables("OpParameter").Rows(0).Item("DATA_SOURCE")
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("REVISION_DATE")) Then NewParameter.RevisionDate = Date.MinValue Else NewParameter.RevisionDate = ds.Tables("OpParameter").Rows(0).Item("REVISION_DATE")
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("CHANGE_ID")) Then NewParameter.ChangeID = "" Else NewParameter.ChangeID = ds.Tables("OpParameter").Rows(0).Item("CHANGE_ID")
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("DEPRECATED")) Then NewParameter.Deprecated = False Else NewParameter.Deprecated = ds.Tables("OpParameter").Rows(0).Item("DEPRECATED")
                        DirectCoordOpParamList.Add(NewParameter)
                    Else
                        RaiseEvent ErrorMessage("There are " & ds.Tables("OpParameter").Rows.Count & "Parameters with the code: " & NewUsage.ParameterCode.ToString & vbCrLf)
                    End If
                    ds.Tables("OpParameter").Clear()

                    'Get the corresponding Coordinate Operation Parameter Value:
                    da.SelectCommand.CommandText = "Select * From [Coordinate_Operation Parameter Value] Where PARAMETER_CODE = " & NewUsage.ParameterCode.ToString & " And COORD_OP_METHOD_CODE = " & MethodCode.ToString & " And COORD_OP_CODE = " & OpCode.ToString
                    da.Fill(ds, "Value")
                    If ds.Tables("Value").Rows.Count = 0 Then
                        RaiseEvent ErrorMessage("There are no Coordinate Operation Parameter Value records for Parameter code number: " & NewUsage.ParameterCode.ToString & " and COORD_OP_METHOD_CODE = " & MethodCode.ToString & " and COORD_OP_CODE = " & OpCode.ToString & vbCrLf)
                    ElseIf ds.Tables("Value").Rows.Count = 1 Then
                        Dim NewParamValue As New CoordOpParamValue
                        NewParamValue.OpCode = OpCode
                        NewParamValue.MethodCode = MethodCode
                        NewParamValue.ParameterCode = NewUsage.ParameterCode
                        If IsDBNull(ds.Tables("Value").Rows(0).Item("PARAMETER_VALUE")) Then NewParamValue.ParameterValue = Double.NaN Else NewParamValue.ParameterValue = ds.Tables("Value").Rows(0).Item("PARAMETER_VALUE")
                        ParamVal = NewParamValue.ParameterValue
                        If IsDBNull(ds.Tables("Value").Rows(0).Item("PARAM_VALUE_FILE_REF")) Then NewParamValue.ParamValueFileRef = "" Else NewParamValue.ParamValueFileRef = ds.Tables("Value").Rows(0).Item("PARAM_VALUE_FILE_REF")
                        If IsDBNull(ds.Tables("Value").Rows(0).Item("UOM_CODE")) Then NewParamValue.UomCode = -1 Else NewParamValue.UomCode = ds.Tables("Value").Rows(0).Item("UOM_CODE")
                        DirectCoordOpParamValList.Add(NewParamValue)
                    Else
                        RaiseEvent ErrorMessage("There are " & ds.Tables("Value").Rows.Count & "Parameter Values with the code: " & NewUsage.ParameterCode.ToString & " and COORD_OP_METHOD_CODE = " & MethodCode.ToString & " and COORD_OP_CODE = " & OpCode.ToString & vbCrLf)
                    End If
                    ds.Tables("Value").Clear()
                    If IsNothing(DirectMethod) Then
                        'The Direct Datum Transformation method has Not been defined.
                    Else
                        DirectMethod.SetParameter(ParamCode, ParamVal)
                    End If
                Next
                If IsNothing(DirectMethod) Then
                    'The Direct Datum Transformation method has Not been defined.
                Else
                    DirectMethod.ApplyReverse = DirectMethodApplyReverse
                    DirectMethod.UpdateVariables
                End If
            End If
        Else
            RaiseEvent ErrorMessage("There are " & ds.Tables("CoordOp").Rows.Count & " Coordinate Operations with the code: " & OpCode & vbCrLf)
        End If
        conn.Close()
    End Sub

    Public Sub GetInputToWgs84DatumTransCoordOp(OpCode As Integer)
        'Get the Input to WGS 84 Datum Transformation coordinate operation correspondignt to OpCode.

        If EpsgDatabasePath = "" Then
            RaiseEvent ErrorMessage("No EPSG database has been selected." & vbCrLf)
            Exit Sub
        End If

        If Not System.IO.File.Exists(EpsgDatabasePath) Then
            RaiseEvent ErrorMessage("Selected EPSG database can not be found." & vbCrLf)
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + EpsgDatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        'Clear the existing data.
        'DirectCoordOp.Clear()
        'DirectCoordOpMethod.Clear()
        'DirectCoordOpParamUseList.Clear()
        'DirectCoordOpParamList.Clear()
        'DirectCoordOpParamValList.Clear()
        InputToWgs84CoordOp.Clear()
        InputToWgs84CoordOpMethod.Clear()
        InputToWgs84CoordOpParamUseList.Clear()
        InputToWgs84CoordOpParamList.Clear()
        InputToWgs84CoordOpParamValList.Clear()

        'Get list of Input Source Coordinate Operations:
        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter("Select * From Coordinate_Operation Where COORD_OP_CODE = " & OpCode.ToString, conn)
        Dim ds As DataSet = New DataSet
        da.Fill(ds, "CoordOp")

        If ds.Tables("CoordOp").Rows.Count = 0 Then
            RaiseEvent ErrorMessage("There are no Coordinate Operation records for code number: " & OpCode & vbCrLf)
        ElseIf ds.Tables("CoordOp").Rows.Count = 1 Then
            InputToWgs84CoordOp.Code = OpCode
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("COORD_OP_NAME")) Then InputToWgs84CoordOp.Name = "" Else InputToWgs84CoordOp.Name = ds.Tables("CoordOp").Rows(0).Item("COORD_OP_NAME")

            Select Case ds.Tables("CoordOp").Rows(0).Item("COORD_OP_TYPE")
                Case "conversion"
                    InputToWgs84CoordOp.Type = CoordinateOperation.OperationType.conversion
                Case "transformation"
                    InputToWgs84CoordOp.Type = CoordinateOperation.OperationType.transformation
                Case "point motion operation"
                    InputToWgs84CoordOp.Type = CoordinateOperation.OperationType.pointMotionOperation
                Case "concatenated operation"
                    InputToWgs84CoordOp.Type = CoordinateOperation.OperationType.concatenatedOperation
                Case Else
                    RaiseEvent ErrorMessage("Unknown Target coordinate operation type: " & ds.Tables("CoordOp").Rows(0).Item("COORD_OP_TYPE") & vbCrLf)
                    InputToWgs84CoordOp.Type = CoordinateOperation.OperationType.conversion
            End Select

            'If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("SOURCE_CRS_CODE")) Then DirectCoordOp.SourceCrsCode = -1 Else DirectCoordOp.SourceCrsCode = ds.Tables("CoordOp").Rows(0).Item("SOURCE_CRS_CODE")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("SOURCE_CRS_CODE")) Then InputToWgs84CoordOp.SourceCrsCode = -1 Else InputToWgs84CoordOp.SourceCrsCode = ds.Tables("CoordOp").Rows(0).Item("SOURCE_CRS_CODE")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("TARGET_CRS_CODE")) Then InputToWgs84CoordOp.TargetCrsCode = -1 Else InputToWgs84CoordOp.TargetCrsCode = ds.Tables("CoordOp").Rows(0).Item("TARGET_CRS_CODE")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("COORD_TFM_VERSION")) Then InputToWgs84CoordOp.Version = "" Else InputToWgs84CoordOp.Version = ds.Tables("CoordOp").Rows(0).Item("COORD_TFM_VERSION")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("COORD_OP_VARIANT")) Then InputToWgs84CoordOp.OpVariant = -1 Else InputToWgs84CoordOp.OpVariant = ds.Tables("CoordOp").Rows(0).Item("COORD_OP_VARIANT")
            'AREA_OF_USE_CODE has been deprecated.
            'COORD_OP_SCOPE has been deprecated.
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("COORD_OP_ACCURACY")) Then InputToWgs84CoordOp.Accuracy = Single.NaN Else InputToWgs84CoordOp.Accuracy = ds.Tables("CoordOp").Rows(0).Item("COORD_OP_ACCURACY")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("COORD_OP_METHOD_CODE")) Then InputToWgs84CoordOp.MethodCode = -1 Else InputToWgs84CoordOp.MethodCode = ds.Tables("CoordOp").Rows(0).Item("COORD_OP_METHOD_CODE")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("UOM_CODE_SOURCE_COORD_DIFF")) Then InputToWgs84CoordOp.UomSourceCoordDiffCode = -1 Else InputToWgs84CoordOp.UomSourceCoordDiffCode = ds.Tables("CoordOp").Rows(0).Item("UOM_CODE_SOURCE_COORD_DIFF")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("UOM_CODE_TARGET_COORD_DIFF")) Then InputToWgs84CoordOp.UomTargetCoordDiffCode = -1 Else InputToWgs84CoordOp.UomTargetCoordDiffCode = ds.Tables("CoordOp").Rows(0).Item("UOM_CODE_TARGET_COORD_DIFF")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("REMARKS")) Then InputToWgs84CoordOp.Remarks = "" Else InputToWgs84CoordOp.Remarks = ds.Tables("CoordOp").Rows(0).Item("REMARKS")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("INFORMATION_SOURCE")) Then InputToWgs84CoordOp.InfoSource = "" Else InputToWgs84CoordOp.InfoSource = ds.Tables("CoordOp").Rows(0).Item("INFORMATION_SOURCE")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("DATA_SOURCE")) Then InputToWgs84CoordOp.DataSource = "" Else InputToWgs84CoordOp.DataSource = ds.Tables("CoordOp").Rows(0).Item("DATA_SOURCE")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("REVISION_DATE")) Then InputToWgs84CoordOp.RevisionDate = Date.MinValue Else InputToWgs84CoordOp.RevisionDate = ds.Tables("CoordOp").Rows(0).Item("REVISION_DATE")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("CHANGE_ID")) Then InputToWgs84CoordOp.ChangeID = "" Else InputToWgs84CoordOp.ChangeID = ds.Tables("CoordOp").Rows(0).Item("CHANGE_ID")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("SHOW_OPERATION")) Then InputToWgs84CoordOp.Show = True Else InputToWgs84CoordOp.Show = ds.Tables("CoordOp").Rows(0).Item("SHOW_OPERATION")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("DEPRECATED")) Then InputToWgs84CoordOp.Deprecated = False Else InputToWgs84CoordOp.Deprecated = ds.Tables("CoordOp").Rows(0).Item("DEPRECATED")

            'Get Coordinate Operation Method:
            Dim MethodCode As Integer = InputToWgs84CoordOp.MethodCode
            da.SelectCommand.CommandText = "Select * From [Coordinate_Operation Method] Where COORD_OP_METHOD_CODE = " & MethodCode.ToString
            da.Fill(ds, "Method")
            If ds.Tables("Method").Rows.Count = 0 Then
                RaiseEvent ErrorMessage("There are no Coordinate Operation Methods for Method code number: " & MethodCode.ToString & vbCrLf)
            ElseIf ds.Tables("Method").Rows.Count = 1 Then
                InputToWgs84CoordOpMethod.Code = MethodCode
                If IsDBNull(ds.Tables("Method").Rows(0).Item("COORD_OP_METHOD_NAME")) Then InputToWgs84CoordOpMethod.Name = "" Else InputToWgs84CoordOpMethod.Name = ds.Tables("Method").Rows(0).Item("COORD_OP_METHOD_NAME")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("REVERSE_OP")) Then InputToWgs84CoordOpMethod.ReverseOp = False Else InputToWgs84CoordOpMethod.ReverseOp = ds.Tables("Method").Rows(0).Item("REVERSE_OP")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("FORMULA")) Then InputToWgs84CoordOpMethod.Formula = "" Else InputToWgs84CoordOpMethod.Formula = ds.Tables("Method").Rows(0).Item("FORMULA")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("EXAMPLE")) Then InputToWgs84CoordOpMethod.Example = "" Else InputToWgs84CoordOpMethod.Example = ds.Tables("Method").Rows(0).Item("EXAMPLE")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("REMARKS")) Then InputToWgs84CoordOpMethod.Remarks = "" Else InputToWgs84CoordOpMethod.Remarks = ds.Tables("Method").Rows(0).Item("REMARKS")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("INFORMATION_SOURCE")) Then InputToWgs84CoordOpMethod.InfoSource = "" Else InputToWgs84CoordOpMethod.InfoSource = ds.Tables("Method").Rows(0).Item("INFORMATION_SOURCE")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("DATA_SOURCE")) Then InputToWgs84CoordOpMethod.DataSource = "" Else InputToWgs84CoordOpMethod.DataSource = ds.Tables("Method").Rows(0).Item("DATA_SOURCE")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("REVISION_DATE")) Then InputToWgs84CoordOpMethod.RevisionDate = Date.MinValue Else InputToWgs84CoordOpMethod.RevisionDate = ds.Tables("Method").Rows(0).Item("REVISION_DATE")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("CHANGE_ID")) Then InputToWgs84CoordOpMethod.ChangeID = "" Else InputToWgs84CoordOpMethod.ChangeID = ds.Tables("Method").Rows(0).Item("CHANGE_ID")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("DEPRECATED")) Then InputToWgs84CoordOpMethod.Deprecated = False Else InputToWgs84CoordOpMethod.Deprecated = ds.Tables("Method").Rows(0).Item("DEPRECATED")
                'Projection.MethodName = ProjectionCoordOpMethod.Name
            Else
                RaiseEvent ErrorMessage("There are " & ds.Tables("Method").Rows.Count & " Coordinate Operation Methods with the code: " & MethodCode.ToString & vbCrLf)
            End If

            'Get the Coordinate Operation Parameter Usage:
            da.SelectCommand.CommandText = "Select * From [Coordinate_Operation Parameter Usage] Where COORD_OP_METHOD_CODE = " & MethodCode.ToString & " Order By SORT_ORDER"
            da.Fill(ds, "Usage")
            If ds.Tables("Usage").Rows.Count = 0 Then
                RaiseEvent ErrorMessage("There are no Coordinate Operation Parameter Usage records for Method code number: " & MethodCode.ToString & vbCrLf)
            Else
                Dim ParamCode As Integer
                Dim ParamVal As Double
                For Each Item As DataRow In ds.Tables("Usage").Rows
                    Dim NewUsage As New CoordOpParamUsage
                    NewUsage.MethodCode = MethodCode
                    'If IsDBNull(ds.Tables("Usage").Rows(0).Item("PARAMETER_CODE")) Then NewUsage.ParameterCode = -1 Else NewUsage.ParameterCode = ds.Tables("Usage").Rows(0).Item("PARAMETER_CODE")
                    If IsDBNull(Item("PARAMETER_CODE")) Then NewUsage.ParameterCode = -1 Else NewUsage.ParameterCode = Item("PARAMETER_CODE")
                    If IsDBNull(Item("SORT_ORDER")) Then NewUsage.SortOrder = 0 Else NewUsage.SortOrder = Item("SORT_ORDER")
                    If IsDBNull(Item("PARAM_SIGN_REVERSAL")) Then NewUsage.SignReversal = "" Else NewUsage.SignReversal = Item("PARAM_SIGN_REVERSAL")
                    InputToWgs84CoordOpParamUseList.Add(NewUsage)

                    'Get the corresponding Coordinate Operation Parameter information:
                    da.SelectCommand.CommandText = "Select * From [Coordinate_Operation Parameter] Where PARAMETER_CODE = " & NewUsage.ParameterCode.ToString
                    da.Fill(ds, "OpParameter")
                    If ds.Tables("OpParameter").Rows.Count = 0 Then
                        RaiseEvent ErrorMessage("There are no Coordinate Operation Parameter records for Parameter code number: " & NewUsage.ParameterCode.ToString & vbCrLf)
                    ElseIf ds.Tables("OpParameter").Rows.Count = 1 Then
                        Dim NewParameter As New CoordOpParameter
                        NewParameter.Code = NewUsage.ParameterCode
                        ParamCode = NewParameter.Code
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("PARAMETER_NAME")) Then NewParameter.Name = "" Else NewParameter.Name = ds.Tables("OpParameter").Rows(0).Item("PARAMETER_NAME")
                        'ParamName = NewParameter.Name
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("DESCRIPTION")) Then NewParameter.Description = "" Else NewParameter.Description = ds.Tables("OpParameter").Rows(0).Item("DESCRIPTION")
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("INFORMATION_SOURCE")) Then NewParameter.InfoSource = "" Else NewParameter.InfoSource = ds.Tables("OpParameter").Rows(0).Item("INFORMATION_SOURCE")
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("DATA_SOURCE")) Then NewParameter.DataSource = "" Else NewParameter.DataSource = ds.Tables("OpParameter").Rows(0).Item("DATA_SOURCE")
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("REVISION_DATE")) Then NewParameter.RevisionDate = Date.MinValue Else NewParameter.RevisionDate = ds.Tables("OpParameter").Rows(0).Item("REVISION_DATE")
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("CHANGE_ID")) Then NewParameter.ChangeID = "" Else NewParameter.ChangeID = ds.Tables("OpParameter").Rows(0).Item("CHANGE_ID")
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("DEPRECATED")) Then NewParameter.Deprecated = False Else NewParameter.Deprecated = ds.Tables("OpParameter").Rows(0).Item("DEPRECATED")
                        InputToWgs84CoordOpParamList.Add(NewParameter)
                    Else
                        RaiseEvent ErrorMessage("There are " & ds.Tables("OpParameter").Rows.Count & "Parameters with the code: " & NewUsage.ParameterCode.ToString & vbCrLf)
                    End If
                    ds.Tables("OpParameter").Clear()

                    'Get the corresponding Coordinate Operation Parameter Value:
                    da.SelectCommand.CommandText = "Select * From [Coordinate_Operation Parameter Value] Where PARAMETER_CODE = " & NewUsage.ParameterCode.ToString & " And COORD_OP_METHOD_CODE = " & MethodCode.ToString & " And COORD_OP_CODE = " & OpCode.ToString
                    da.Fill(ds, "Value")
                    If ds.Tables("Value").Rows.Count = 0 Then
                        RaiseEvent ErrorMessage("There are no Coordinate Operation Parameter Value records for Parameter code number: " & NewUsage.ParameterCode.ToString & " and COORD_OP_METHOD_CODE = " & MethodCode.ToString & " and COORD_OP_CODE = " & OpCode.ToString & vbCrLf)
                    ElseIf ds.Tables("Value").Rows.Count = 1 Then
                        Dim NewParamValue As New CoordOpParamValue
                        NewParamValue.OpCode = OpCode
                        NewParamValue.MethodCode = MethodCode
                        NewParamValue.ParameterCode = NewUsage.ParameterCode
                        If IsDBNull(ds.Tables("Value").Rows(0).Item("PARAMETER_VALUE")) Then NewParamValue.ParameterValue = Double.NaN Else NewParamValue.ParameterValue = ds.Tables("Value").Rows(0).Item("PARAMETER_VALUE")
                        ParamVal = NewParamValue.ParameterValue
                        If IsDBNull(ds.Tables("Value").Rows(0).Item("PARAM_VALUE_FILE_REF")) Then NewParamValue.ParamValueFileRef = "" Else NewParamValue.ParamValueFileRef = ds.Tables("Value").Rows(0).Item("PARAM_VALUE_FILE_REF")
                        If IsDBNull(ds.Tables("Value").Rows(0).Item("UOM_CODE")) Then NewParamValue.UomCode = -1 Else NewParamValue.UomCode = ds.Tables("Value").Rows(0).Item("UOM_CODE")
                        InputToWgs84CoordOpParamValList.Add(NewParamValue)
                    Else
                        RaiseEvent ErrorMessage("There are " & ds.Tables("Value").Rows.Count & "Parameter Values with the code: " & NewUsage.ParameterCode.ToString & " and COORD_OP_METHOD_CODE = " & MethodCode.ToString & " and COORD_OP_CODE = " & OpCode.ToString & vbCrLf)
                    End If
                    ds.Tables("Value").Clear()
                Next
            End If
        Else
            RaiseEvent ErrorMessage("There are " & ds.Tables("CoordOp").Rows.Count & " Coordinate Operations with the code: " & OpCode & vbCrLf)
        End If
        conn.Close()
    End Sub


    Public Sub GetWgs84ToOutputDatumTransCoordOp(OpCode As Integer)
        'Get the Direct Datum Transformation coordinate operation correspondignt to OpCode.

        If EpsgDatabasePath = "" Then
            RaiseEvent ErrorMessage("No EPSG database has been selected." & vbCrLf)
            Exit Sub
        End If

        If Not System.IO.File.Exists(EpsgDatabasePath) Then
            RaiseEvent ErrorMessage("Selected EPSG database can not be found." & vbCrLf)
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + EpsgDatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        'Clear the existing data.
        'DirectCoordOp.Clear()
        'DirectCoordOpMethod.Clear()
        'DirectCoordOpParamUseList.Clear()
        'DirectCoordOpParamList.Clear()
        'DirectCoordOpParamValList.Clear()
        Wgs84ToOutputCoordOp.Clear()
        Wgs84ToOutputCoordOpMethod.Clear()
        Wgs84ToOutputCoordOpParamUseList.Clear()
        Wgs84ToOutputCoordOpParamList.Clear()
        Wgs84ToOutputCoordOpParamValList.Clear()

        'Get list of Input Source Coordinate Operations:
        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter("Select * From Coordinate_Operation Where COORD_OP_CODE = " & OpCode.ToString, conn)
        Dim ds As DataSet = New DataSet
        da.Fill(ds, "CoordOp")

        If ds.Tables("CoordOp").Rows.Count = 0 Then
            RaiseEvent ErrorMessage("There are no Coordinate Operation records for code number: " & OpCode & vbCrLf)
        ElseIf ds.Tables("CoordOp").Rows.Count = 1 Then
            Wgs84ToOutputCoordOp.Code = OpCode
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("COORD_OP_NAME")) Then Wgs84ToOutputCoordOp.Name = "" Else Wgs84ToOutputCoordOp.Name = ds.Tables("CoordOp").Rows(0).Item("COORD_OP_NAME")

            Select Case ds.Tables("CoordOp").Rows(0).Item("COORD_OP_TYPE")
                Case "conversion"
                    Wgs84ToOutputCoordOp.Type = CoordinateOperation.OperationType.conversion
                Case "transformation"
                    Wgs84ToOutputCoordOp.Type = CoordinateOperation.OperationType.transformation
                Case "point motion operation"
                    Wgs84ToOutputCoordOp.Type = CoordinateOperation.OperationType.pointMotionOperation
                Case "concatenated operation"
                    Wgs84ToOutputCoordOp.Type = CoordinateOperation.OperationType.concatenatedOperation
                Case Else
                    RaiseEvent ErrorMessage("Unknown Target coordinate operation type: " & ds.Tables("CoordOp").Rows(0).Item("COORD_OP_TYPE") & vbCrLf)
                    Wgs84ToOutputCoordOp.Type = CoordinateOperation.OperationType.conversion
            End Select

            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("SOURCE_CRS_CODE")) Then Wgs84ToOutputCoordOp.SourceCrsCode = -1 Else Wgs84ToOutputCoordOp.SourceCrsCode = ds.Tables("CoordOp").Rows(0).Item("SOURCE_CRS_CODE")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("TARGET_CRS_CODE")) Then Wgs84ToOutputCoordOp.TargetCrsCode = -1 Else Wgs84ToOutputCoordOp.TargetCrsCode = ds.Tables("CoordOp").Rows(0).Item("TARGET_CRS_CODE")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("COORD_TFM_VERSION")) Then Wgs84ToOutputCoordOp.Version = "" Else Wgs84ToOutputCoordOp.Version = ds.Tables("CoordOp").Rows(0).Item("COORD_TFM_VERSION")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("COORD_OP_VARIANT")) Then Wgs84ToOutputCoordOp.OpVariant = -1 Else Wgs84ToOutputCoordOp.OpVariant = ds.Tables("CoordOp").Rows(0).Item("COORD_OP_VARIANT")
            'AREA_OF_USE_CODE has been deprecated.
            'COORD_OP_SCOPE has been deprecated.
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("COORD_OP_ACCURACY")) Then Wgs84ToOutputCoordOp.Accuracy = Single.NaN Else Wgs84ToOutputCoordOp.Accuracy = ds.Tables("CoordOp").Rows(0).Item("COORD_OP_ACCURACY")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("COORD_OP_METHOD_CODE")) Then Wgs84ToOutputCoordOp.MethodCode = -1 Else Wgs84ToOutputCoordOp.MethodCode = ds.Tables("CoordOp").Rows(0).Item("COORD_OP_METHOD_CODE")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("UOM_CODE_SOURCE_COORD_DIFF")) Then Wgs84ToOutputCoordOp.UomSourceCoordDiffCode = -1 Else Wgs84ToOutputCoordOp.UomSourceCoordDiffCode = ds.Tables("CoordOp").Rows(0).Item("UOM_CODE_SOURCE_COORD_DIFF")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("UOM_CODE_TARGET_COORD_DIFF")) Then Wgs84ToOutputCoordOp.UomTargetCoordDiffCode = -1 Else Wgs84ToOutputCoordOp.UomTargetCoordDiffCode = ds.Tables("CoordOp").Rows(0).Item("UOM_CODE_TARGET_COORD_DIFF")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("REMARKS")) Then Wgs84ToOutputCoordOp.Remarks = "" Else Wgs84ToOutputCoordOp.Remarks = ds.Tables("CoordOp").Rows(0).Item("REMARKS")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("INFORMATION_SOURCE")) Then Wgs84ToOutputCoordOp.InfoSource = "" Else Wgs84ToOutputCoordOp.InfoSource = ds.Tables("CoordOp").Rows(0).Item("INFORMATION_SOURCE")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("DATA_SOURCE")) Then Wgs84ToOutputCoordOp.DataSource = "" Else Wgs84ToOutputCoordOp.DataSource = ds.Tables("CoordOp").Rows(0).Item("DATA_SOURCE")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("REVISION_DATE")) Then Wgs84ToOutputCoordOp.RevisionDate = Date.MinValue Else Wgs84ToOutputCoordOp.RevisionDate = ds.Tables("CoordOp").Rows(0).Item("REVISION_DATE")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("CHANGE_ID")) Then Wgs84ToOutputCoordOp.ChangeID = "" Else Wgs84ToOutputCoordOp.ChangeID = ds.Tables("CoordOp").Rows(0).Item("CHANGE_ID")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("SHOW_OPERATION")) Then Wgs84ToOutputCoordOp.Show = True Else Wgs84ToOutputCoordOp.Show = ds.Tables("CoordOp").Rows(0).Item("SHOW_OPERATION")
            If IsDBNull(ds.Tables("CoordOp").Rows(0).Item("DEPRECATED")) Then Wgs84ToOutputCoordOp.Deprecated = False Else Wgs84ToOutputCoordOp.Deprecated = ds.Tables("CoordOp").Rows(0).Item("DEPRECATED")

            'Get Coordinate Operation Method:
            Dim MethodCode As Integer = Wgs84ToOutputCoordOp.MethodCode
            da.SelectCommand.CommandText = "Select * From [Coordinate_Operation Method] Where COORD_OP_METHOD_CODE = " & MethodCode.ToString
            da.Fill(ds, "Method")
            If ds.Tables("Method").Rows.Count = 0 Then
                RaiseEvent ErrorMessage("There are no Coordinate Operation Methods for Method code number: " & MethodCode.ToString & vbCrLf)
            ElseIf ds.Tables("Method").Rows.Count = 1 Then
                Wgs84ToOutputCoordOpMethod.Code = MethodCode
                If IsDBNull(ds.Tables("Method").Rows(0).Item("COORD_OP_METHOD_NAME")) Then Wgs84ToOutputCoordOpMethod.Name = "" Else Wgs84ToOutputCoordOpMethod.Name = ds.Tables("Method").Rows(0).Item("COORD_OP_METHOD_NAME")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("REVERSE_OP")) Then Wgs84ToOutputCoordOpMethod.ReverseOp = False Else Wgs84ToOutputCoordOpMethod.ReverseOp = ds.Tables("Method").Rows(0).Item("REVERSE_OP")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("FORMULA")) Then Wgs84ToOutputCoordOpMethod.Formula = "" Else Wgs84ToOutputCoordOpMethod.Formula = ds.Tables("Method").Rows(0).Item("FORMULA")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("EXAMPLE")) Then Wgs84ToOutputCoordOpMethod.Example = "" Else Wgs84ToOutputCoordOpMethod.Example = ds.Tables("Method").Rows(0).Item("EXAMPLE")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("REMARKS")) Then Wgs84ToOutputCoordOpMethod.Remarks = "" Else Wgs84ToOutputCoordOpMethod.Remarks = ds.Tables("Method").Rows(0).Item("REMARKS")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("INFORMATION_SOURCE")) Then Wgs84ToOutputCoordOpMethod.InfoSource = "" Else Wgs84ToOutputCoordOpMethod.InfoSource = ds.Tables("Method").Rows(0).Item("INFORMATION_SOURCE")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("DATA_SOURCE")) Then Wgs84ToOutputCoordOpMethod.DataSource = "" Else Wgs84ToOutputCoordOpMethod.DataSource = ds.Tables("Method").Rows(0).Item("DATA_SOURCE")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("REVISION_DATE")) Then Wgs84ToOutputCoordOpMethod.RevisionDate = Date.MinValue Else Wgs84ToOutputCoordOpMethod.RevisionDate = ds.Tables("Method").Rows(0).Item("REVISION_DATE")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("CHANGE_ID")) Then Wgs84ToOutputCoordOpMethod.ChangeID = "" Else Wgs84ToOutputCoordOpMethod.ChangeID = ds.Tables("Method").Rows(0).Item("CHANGE_ID")
                If IsDBNull(ds.Tables("Method").Rows(0).Item("DEPRECATED")) Then Wgs84ToOutputCoordOpMethod.Deprecated = False Else Wgs84ToOutputCoordOpMethod.Deprecated = ds.Tables("Method").Rows(0).Item("DEPRECATED")
                'Projection.MethodName = ProjectionCoordOpMethod.Name
            Else
                RaiseEvent ErrorMessage("There are " & ds.Tables("Method").Rows.Count & " Coordinate Operation Methods with the code: " & MethodCode.ToString & vbCrLf)
            End If

            'Get the Coordinate Operation Parameter Usage:
            da.SelectCommand.CommandText = "Select * From [Coordinate_Operation Parameter Usage] Where COORD_OP_METHOD_CODE = " & MethodCode.ToString & " Order By SORT_ORDER"
            da.Fill(ds, "Usage")
            If ds.Tables("Usage").Rows.Count = 0 Then
                RaiseEvent ErrorMessage("There are no Coordinate Operation Parameter Usage records for Method code number: " & MethodCode.ToString & vbCrLf)
            Else
                Dim ParamCode As Integer
                Dim ParamVal As Double
                For Each Item As DataRow In ds.Tables("Usage").Rows
                    Dim NewUsage As New CoordOpParamUsage
                    NewUsage.MethodCode = MethodCode
                    'If IsDBNull(ds.Tables("Usage").Rows(0).Item("PARAMETER_CODE")) Then NewUsage.ParameterCode = -1 Else NewUsage.ParameterCode = ds.Tables("Usage").Rows(0).Item("PARAMETER_CODE")
                    If IsDBNull(Item("PARAMETER_CODE")) Then NewUsage.ParameterCode = -1 Else NewUsage.ParameterCode = Item("PARAMETER_CODE")
                    If IsDBNull(Item("SORT_ORDER")) Then NewUsage.SortOrder = 0 Else NewUsage.SortOrder = Item("SORT_ORDER")
                    If IsDBNull(Item("PARAM_SIGN_REVERSAL")) Then NewUsage.SignReversal = "" Else NewUsage.SignReversal = Item("PARAM_SIGN_REVERSAL")
                    Wgs84ToOutputCoordOpParamUseList.Add(NewUsage)

                    'Get the corresponding Coordinate Operation Parameter information:
                    da.SelectCommand.CommandText = "Select * From [Coordinate_Operation Parameter] Where PARAMETER_CODE = " & NewUsage.ParameterCode.ToString
                    da.Fill(ds, "OpParameter")
                    If ds.Tables("OpParameter").Rows.Count = 0 Then
                        RaiseEvent ErrorMessage("There are no Coordinate Operation Parameter records for Parameter code number: " & NewUsage.ParameterCode.ToString & vbCrLf)
                    ElseIf ds.Tables("OpParameter").Rows.Count = 1 Then
                        Dim NewParameter As New CoordOpParameter
                        NewParameter.Code = NewUsage.ParameterCode
                        ParamCode = NewParameter.Code
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("PARAMETER_NAME")) Then NewParameter.Name = "" Else NewParameter.Name = ds.Tables("OpParameter").Rows(0).Item("PARAMETER_NAME")
                        'ParamName = NewParameter.Name
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("DESCRIPTION")) Then NewParameter.Description = "" Else NewParameter.Description = ds.Tables("OpParameter").Rows(0).Item("DESCRIPTION")
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("INFORMATION_SOURCE")) Then NewParameter.InfoSource = "" Else NewParameter.InfoSource = ds.Tables("OpParameter").Rows(0).Item("INFORMATION_SOURCE")
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("DATA_SOURCE")) Then NewParameter.DataSource = "" Else NewParameter.DataSource = ds.Tables("OpParameter").Rows(0).Item("DATA_SOURCE")
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("REVISION_DATE")) Then NewParameter.RevisionDate = Date.MinValue Else NewParameter.RevisionDate = ds.Tables("OpParameter").Rows(0).Item("REVISION_DATE")
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("CHANGE_ID")) Then NewParameter.ChangeID = "" Else NewParameter.ChangeID = ds.Tables("OpParameter").Rows(0).Item("CHANGE_ID")
                        If IsDBNull(ds.Tables("OpParameter").Rows(0).Item("DEPRECATED")) Then NewParameter.Deprecated = False Else NewParameter.Deprecated = ds.Tables("OpParameter").Rows(0).Item("DEPRECATED")
                        Wgs84ToOutputCoordOpParamList.Add(NewParameter)
                    Else
                        RaiseEvent ErrorMessage("There are " & ds.Tables("OpParameter").Rows.Count & "Parameters with the code: " & NewUsage.ParameterCode.ToString & vbCrLf)
                    End If
                    ds.Tables("OpParameter").Clear()

                    'Get the corresponding Coordinate Operation Parameter Value:
                    da.SelectCommand.CommandText = "Select * From [Coordinate_Operation Parameter Value] Where PARAMETER_CODE = " & NewUsage.ParameterCode.ToString & " And COORD_OP_METHOD_CODE = " & MethodCode.ToString & " And COORD_OP_CODE = " & OpCode.ToString
                    da.Fill(ds, "Value")
                    If ds.Tables("Value").Rows.Count = 0 Then
                        RaiseEvent ErrorMessage("There are no Coordinate Operation Parameter Value records for Parameter code number: " & NewUsage.ParameterCode.ToString & " and COORD_OP_METHOD_CODE = " & MethodCode.ToString & " and COORD_OP_CODE = " & OpCode.ToString & vbCrLf)
                    ElseIf ds.Tables("Value").Rows.Count = 1 Then
                        Dim NewParamValue As New CoordOpParamValue
                        NewParamValue.OpCode = OpCode
                        NewParamValue.MethodCode = MethodCode
                        NewParamValue.ParameterCode = NewUsage.ParameterCode
                        If IsDBNull(ds.Tables("Value").Rows(0).Item("PARAMETER_VALUE")) Then NewParamValue.ParameterValue = Double.NaN Else NewParamValue.ParameterValue = ds.Tables("Value").Rows(0).Item("PARAMETER_VALUE")
                        ParamVal = NewParamValue.ParameterValue
                        If IsDBNull(ds.Tables("Value").Rows(0).Item("PARAM_VALUE_FILE_REF")) Then NewParamValue.ParamValueFileRef = "" Else NewParamValue.ParamValueFileRef = ds.Tables("Value").Rows(0).Item("PARAM_VALUE_FILE_REF")
                        If IsDBNull(ds.Tables("Value").Rows(0).Item("UOM_CODE")) Then NewParamValue.UomCode = -1 Else NewParamValue.UomCode = ds.Tables("Value").Rows(0).Item("UOM_CODE")
                        Wgs84ToOutputCoordOpParamValList.Add(NewParamValue)
                    Else
                        RaiseEvent ErrorMessage("There are " & ds.Tables("Value").Rows.Count & "Parameter Values with the code: " & NewUsage.ParameterCode.ToString & " and COORD_OP_METHOD_CODE = " & MethodCode.ToString & " and COORD_OP_CODE = " & OpCode.ToString & vbCrLf)
                    End If
                    ds.Tables("Value").Clear()
                Next
            End If
        Else
            RaiseEvent ErrorMessage("There are " & ds.Tables("CoordOp").Rows.Count & " Coordinate Operations with the code: " & OpCode & vbCrLf)
        End If
        conn.Close()
    End Sub


    Public Sub ApplyDirectMethodName()
        'Apply the Direct Method Name
        Select Case DirectMethodName
            Case "Coordinate Frame rotation (geog2D domain)"
                DirectMethod = New Helmert
                DirectMethod.SourceCoord = SourceCoord
                DirectMethod.TargetCoord = TargetCoord
                'Apply the parameters:



            Case Else
                RaiseEvent ErrorMessage("Unknown Direct Datum Transformation method name: " & DirectMethodName & vbCrLf)
        End Select
    End Sub

    Public Sub ApplyInputToWgs84MethodName()
        'Apply the Input to WGS 84 Method Name
        Select Case InputToWgs84MethodName
            Case "Coordinate Frame rotation (geog2D domain)"
                InputToWgs84Method = New Helmert
                InputToWgs84Method.SourceCoord = SourceCoord
                InputToWgs84Method.TargetCoord = Wgs84Coord
            Case Else
                RaiseEvent ErrorMessage("Unknown Input to WGS 84 Datum Transformation method name: " & InputToWgs84MethodName & vbCrLf)
        End Select
    End Sub

    Public Sub ApplyWgs84ToOutputMethodName()
        'Apply the Direct Method Name
        Select Case Wgs84ToOutputMethodName
            Case "Coordinate Frame rotation (geog2D domain)"
                Wgs84ToOutputMethod = New Helmert
                Wgs84ToOutputMethod.SourceCoord = Wgs84Coord
                Wgs84ToOutputMethod.TargetCoord = TargetCoord
            Case Else
                RaiseEvent ErrorMessage("Unknown Direct Datum Transformation method name: " & Wgs84ToOutputMethodName & vbCrLf)
        End Select
    End Sub

    Public Sub InputToOutput()
        'This converts the Input cartesian coordinates to the corresponding Output cartesian coordinates.

        If Type = enumType.Direct Then
            If IsNothing(DirectMethod) Then

            Else
                DirectMethod.Transform
            End If
        ElseIf Type = enumType.None Then
            TargetCoord.X = SourceCoord.X
            TargetCoord.Y = SourceCoord.Y
            TargetCoord.Z = SourceCoord.Z
        ElseIf Type = enumType.ViaWgs84 Then
            InputToWgs84Method.Transform
            Wgs84ToOutputMethod.Transform
        Else

        End If
    End Sub

    Public Sub OutputToInput()
        'This converts the Ouput cartesian coordinates to the corresponding Input cartesian coordinates.

        If Type = enumType.Direct Then
            DirectMethod.ReverseTransform
        ElseIf Type = enumType.None Then
            TargetCoord.X = SourceCoord.X
            TargetCoord.Y = SourceCoord.Y
            TargetCoord.Z = SourceCoord.Z
        ElseIf Type = enumType.ViaWgs84 Then
            Wgs84ToOutputMethod.ReverseTransform
            InputToWgs84Method.ReverseTransform
        Else

        End If
    End Sub

#End Region 'Methods --------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Events - Events that can be triggered by this class." '==========================================================================================================================
    Event ErrorMessage(ByVal Msg As String) 'Send an error message.
    Event Message(ByVal Msg As String) 'Send a normal message.
#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


    'Public Class HelmertTransformation
    Public Class Helmert
        'Helmert seven parameter transformation - converts a cartesian X, Y, Z point in one datum to corresponding coordinate values (XT, YT, ZT) in another datum

#Region " Variable Declarations - All the variables used in this class." '=====================================================================================================================

        Public SourceCoord As Coordinate 'This class reference is set before the method is used. For example DatumTrans.SourceCoord = InputCrs.Coord - The DatumTrans class now accesses the Coord values directly.
        Public TargetCoord As Coordinate 'This class reference is set before the method is used. For example DatumTrans.TargetCoord = OutputCrs.Coord - The DatumTrans class now accesses the Coord values directly.

        Dim Scale As Double
        Dim XRotRad As Double
        Dim YRotRad As Double
        Dim ZRotRad As Double
        Dim XShiftM As Double
        Dim YShiftM As Double
        Dim ZShiftM As Double

#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Properties - All the properties used in this class." '===========================================================================================================================

        Private _cX As Double 'The X Axis translation value. Millimetres.
        Property CX As Double
            Get
                Return _cX
            End Get
            Set(value As Double)
                _cX = value
            End Set
        End Property

        Private _cY As Double 'The Y Axis translation value. Millimetres.
        Property CY As Double
            Get
                Return _cY
            End Get
            Set(value As Double)
                _cY = value
            End Set
        End Property

        Private _cZ As Double 'The Z Axis translation value. Millimetres.
        Property CZ As Double
            Get
                Return _cZ
            End Get
            Set(value As Double)
                _cZ = value
            End Set
        End Property

        Private _rX As Double 'The X Axis rotation value. MilliArcSeconds.
        Property RX As Double
            Get
                Return _rX
            End Get
            Set(value As Double)
                _rX = value
            End Set
        End Property

        Private _rY As Double 'The Y Axis rotation value. MilliArcSeconds.
        Property RY As Double
            Get
                Return _rY
            End Get
            Set(value As Double)
                _rY = value
            End Set
        End Property

        Private _rZ As Double 'The Z Axis rotation value. MilliArcSeconds.
        Property RZ As Double
            Get
                Return _rZ
            End Get
            Set(value As Double)
                _rZ = value
            End Set
        End Property

        'Private _s As Double 'The Scale difference. Parts per billion.
        ''Private _dS As Double 'The Scale difference. Parts per billion.
        'Property S As Double
        '    'Property dS As Double
        '    Get
        '        Return _s
        '    End Get
        '    Set(value As Double)
        '        _s = value
        '    End Set
        'End Property

        Private _ds As Double 'The Scale difference. Parts per billion.
        'Private _dS As Double 'The Scale difference. Parts per billion.
        Property dS As Double
            'Property dS As Double
            Get
                Return _ds
            End Get
            Set(value As Double)
                _ds = value
            End Set
        End Property

        Private _applyReverse As Boolean = False 'If True, the reverse datum transformation is applied.
        Property ApplyReverse As Boolean
            Get
                Return _applyReverse
            End Get
            Set(value As Boolean)
                _applyReverse = value
            End Set
        End Property

        'NOTE: SourceCoord And TargetCoord are now used to access the Cartesian coordinate values.
        'Private _x As Double 'The X cartesian coordinate in the source datum.
        'Property X As Double
        '    Get
        '        Return _x
        '    End Get
        '    Set(value As Double)
        '        _x = value
        '    End Set
        'End Property

        'Private _y As Double 'The Y cartesian coordinate in the source datum.
        'Property Y As Double
        '    Get
        '        Return _y
        '    End Get
        '    Set(value As Double)
        '        _y = value
        '    End Set
        'End Property

        'Private _z As Double 'The Z cartesian coordinate in the source datum.
        'Property Z As Double
        '    Get
        '        Return _z
        '    End Get
        '    Set(value As Double)
        '        _z = value
        '    End Set
        'End Property

        'Private _xT As Double 'The X cartesian coordinate in the target datum.
        'Property XT As Double
        '    Get
        '        Return _xT
        '    End Get
        '    Set(value As Double)
        '        _xT = value
        '    End Set
        'End Property

        'Private _yT As Double 'The Y cartesian coordinate in the target datum.
        'Property YT As Double
        '    Get
        '        Return _yT
        '    End Get
        '    Set(value As Double)
        '        _yT = value
        '    End Set
        'End Property

        'Private _zT As Double 'The Z cartesian coordinate in the target datum.
        'Property ZT As Double
        '    Get
        '        Return _zT
        '    End Get
        '    Set(value As Double)
        '        _zT = value
        '    End Set
        'End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods - The main actions performed by this class." '===========================================================================================================================

        Public Sub SetParameter(ParameterCode As Integer, ParameterValue As Double)
            'Sets the Helmert Datum Transformation property from a PropertyCode and a PropertyValue.
            Select Case ParameterCode
                Case 8605 'X-axis translation
                    CX = ParameterValue
                Case 8606 'Y-axis translation
                    CY = ParameterValue
                Case 8607 'Z-axis translation
                    CZ = ParameterValue
                Case 8608 'X-axis rotation
                    RX = ParameterValue
                Case 8609 'Y-axis rotation
                    RY = ParameterValue
                Case 8610 'Z-axis rotation
                    RZ = ParameterValue
                Case 8611 'Scale difference
                    'S = ParameterValue
                    dS = ParameterValue
                Case Else
                    RaiseEvent ErrorMessage("Unknown property code: " & ParameterCode & vbCrLf)
            End Select
        End Sub

        Public Sub UpdateVariables()
            'Update the variables used in the Datum Transformation
            If ApplyReverse Then
                'Scale = -S * 0.000000001#
                Scale = 1 - dS * 0.000000001#
                'Scale = -1 - dS * 0.000000001#
                'XRotRad = -(RX / 3.6#) * (Math.PI / 180.0#)
                XRotRad = -(RX / 3600000.0#) * (Math.PI / 180.0#)
                'YRotRad = -(RY / 3.6#) * (Math.PI / 180.0#)
                YRotRad = -(RY / 3600000.0#) * (Math.PI / 180.0#)
                'ZRotRad = -(RZ / 3.6#) * (Math.PI / 180.0#)
                ZRotRad = -(RZ / 3600000.0#) * (Math.PI / 180.0#)
                XShiftM = -CX * 0.001#
                YShiftM = -CY * 0.001#
                ZShiftM = -CZ * 0.001#
            Else
                'Scale = S * 0.000000001#
                Scale = 1 + dS * 0.000000001#
                'Scale = 1 + dS * 0.000000001#
                'XRotRad = (RX / 3.6#) * (Math.PI / 180.0#)
                XRotRad = (RX / 3600000.0#) * (Math.PI / 180.0#)
                'YRotRad = (RY / 3.6#) * (Math.PI / 180.0#)
                YRotRad = (RY / 3600000.0#) * (Math.PI / 180.0#)
                'ZRotRad = (RZ / 3.6#) * (Math.PI / 180.0#)
                ZRotRad = (RZ / 3600000.0#) * (Math.PI / 180.0#)
                XShiftM = CX * 0.001#
                YShiftM = CY * 0.001#
                ZShiftM = CZ * 0.001#
            End If
        End Sub

        Public Sub Transform()
            'Transform the cartesian X, Y, Z coordinates in the source datum into corresponding XT, YT, ZT coordinates in the target datum.
            'NOTE: EPSG parameters have units of parts per billion, milliarc-seconds and millimetres.

            'NOTE: These are now global variables that are only calculated once when the datum treansformation parameters as set.
            'Dim Scale As Double
            'Dim XRotRad As Double
            'Dim YRotRad As Double
            'Dim ZRotRad As Double
            'Dim XShiftM As Double
            'Dim YShiftM As Double
            'Dim ZShiftM As Double

            'If ApplyReverse Then
            '    Scale = -S * 0.000000001#
            '    XRotRad = -(RX / 3.6#) * (Math.PI / 180.0#)
            '    YRotRad = -(RY / 3.6#) * (Math.PI / 180.0#)
            '    ZRotRad = -(RZ / 3.6#) * (Math.PI / 180.0#)
            '    XShiftM = -CX * 0.001#
            '    YShiftM = -CY * 0.001#
            '    ZShiftM = -CZ * 0.001#
            'Else
            '    Scale = S * 0.000000001#
            '    XRotRad = (RX / 3.6#) * (Math.PI / 180.0#)
            '    YRotRad = (RY / 3.6#) * (Math.PI / 180.0#)
            '    ZRotRad = (RZ / 3.6#) * (Math.PI / 180.0#)
            '    XShiftM = CX * 0.001#
            '    YShiftM = CY * 0.001#
            '    ZShiftM = CZ * 0.001#
            'End If

            'XT = X + (X * Scale) - (Y * ZRotRad) + (Z * YRotRad) + XShiftM

            'TargetCoord.X = SourceCoord.X + (SourceCoord.X * Scale) - (SourceCoord.Y * ZRotRad) + (SourceCoord.Z * YRotRad) + XShiftM
            'TargetCoord.Y = (SourceCoord.X * ZRotRad) + SourceCoord.Y + (SourceCoord.Y * Scale) - (SourceCoord.Z * XRotRad) + YShiftM
            'TargetCoord.Z = (-SourceCoord.X * YRotRad) + (SourceCoord.Y * XRotRad) + SourceCoord.Z + (SourceCoord.Z * Scale) + ZShiftM

            'TargetCoord.X = SourceCoord.X + (SourceCoord.X * Scale) - (SourceCoord.Y * -ZRotRad) + (SourceCoord.Z * -YRotRad) + XShiftM
            'TargetCoord.Y = (SourceCoord.X * -ZRotRad) + SourceCoord.Y + (SourceCoord.Y * Scale) - (SourceCoord.Z * -XRotRad) + YShiftM
            'TargetCoord.Z = (-SourceCoord.X * -YRotRad) + (SourceCoord.Y * -XRotRad) + SourceCoord.Z + (SourceCoord.Z * Scale) + ZShiftM

            'TargetCoord.X = XShiftM + (1 + Scale) * (SourceCoord.X + SourceCoord.Y * ZRotRad - SourceCoord.Z * YRotRad)
            'TargetCoord.Y = YShiftM + (1 + Scale) * (-SourceCoord.X * ZRotRad + SourceCoord.Y + SourceCoord.Z * XRotRad)
            'TargetCoord.Z = ZShiftM + (1 + Scale) * (SourceCoord.X * YRotRad - SourceCoord.Y * XRotRad + SourceCoord.Z)

            TargetCoord.X = XShiftM + Scale * (SourceCoord.X + SourceCoord.Y * ZRotRad - SourceCoord.Z * YRotRad)
            TargetCoord.Y = YShiftM + Scale * (-SourceCoord.X * ZRotRad + SourceCoord.Y + SourceCoord.Z * XRotRad)
            TargetCoord.Z = ZShiftM + Scale * (SourceCoord.X * YRotRad - SourceCoord.Y * XRotRad + SourceCoord.Z)



        End Sub

        Public Sub ReverseTransform()
            'Transform the cartesian XT, YT, ZT coordinates in the target datum into corresponding X, Y, Z coordinates in the source datum.
            'NOTE: EPSG parameters have units of parts per billion, milliarc-seconds and millimetres.

            'SourceCoord.X = TargetCoord.X + (TargetCoord.X * -Scale) - (TargetCoord.Y * -ZRotRad) + (TargetCoord.Z * -YRotRad) - XShiftM
            'SourceCoord.Y = (TargetCoord.X * -ZRotRad) + TargetCoord.Y + (TargetCoord.Y * -Scale) - (TargetCoord.Z * -XRotRad) - YShiftM
            'SourceCoord.Z = (-TargetCoord.X * -YRotRad) + (TargetCoord.Y * -XRotRad) + TargetCoord.Z + (TargetCoord.Z * -Scale) - ZShiftM

            'SourceCoord.X = TargetCoord.X + (TargetCoord.X * -Scale) - (TargetCoord.Y * ZRotRad) + (TargetCoord.Z * YRotRad) - XShiftM
            'SourceCoord.Y = (TargetCoord.X * ZRotRad) + TargetCoord.Y + (TargetCoord.Y * -Scale) - (TargetCoord.Z * XRotRad) - YShiftM
            'SourceCoord.Z = (-TargetCoord.X * YRotRad) + (TargetCoord.Y * XRotRad) + TargetCoord.Z + (TargetCoord.Z * -Scale) - ZShiftM

            'TargetCoord.X = -XShiftM + (1 - Scale) * (SourceCoord.X - SourceCoord.Y * ZRotRad + SourceCoord.Z * YRotRad)
            'TargetCoord.Y = -YShiftM + (1 - Scale) * (SourceCoord.X * ZRotRad + SourceCoord.Y - SourceCoord.Z * XRotRad)
            'TargetCoord.Z = -ZShiftM + (1 - Scale) * (-SourceCoord.X * YRotRad + SourceCoord.Y * XRotRad + SourceCoord.Z)

            TargetCoord.X = -XShiftM + Scale * (SourceCoord.X + SourceCoord.Y * ZRotRad - SourceCoord.Z * YRotRad)
            TargetCoord.Y = -YShiftM + Scale * (-SourceCoord.X * ZRotRad + SourceCoord.Y + SourceCoord.Z * XRotRad)
            TargetCoord.Z = -ZShiftM + Scale * (SourceCoord.X * YRotRad - SourceCoord.Y * XRotRad + SourceCoord.Z)


        End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================
        Event ErrorMessage(ByVal Msg As String) 'Send an error message.
        Event Message(ByVal Msg As String) 'Send a normal message.
#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------

    End Class 'HelmertTransformation



End Class 'clsDatumTrans

Public Class Coordinate
    'Stores coordinate information.


#Region " Properties - All the properties used in this class." '===========================================================================================================================

    Private _latitude As Double 'The Latitude of a Geographic coordinate
    Property Latitude As Double
        Get
            Return _latitude
        End Get
        Set(value As Double)
            _latitude = value
            If _latitude < 0 Then
                LatitudeSign = LatLongSign.Negative
            Else
                LatitudeSign = LatLongSign.Positive
            End If
        End Set
    End Property

    Private _longitude As Double 'The Longitude of a Geographic coordinate
    Property Longitude As Double
        Get
            Return _longitude
        End Get
        Set(value As Double)
            _longitude = value
            If _longitude < 0 Then
                LongitudeSign = LatLongSign.Negative
            Else
                LongitudeSign = LatLongSign.Positive
            End If
        End Set
    End Property

    Private _ellipsoidalHeight As Double 'The Ellipsoidal Height of a Geographic coordinate
    Property EllipsoidalHeight As Double
        Get
            Return _ellipsoidalHeight
        End Get
        Set(value As Double)
            _ellipsoidalHeight = value
        End Set
    End Property

    'NOTE: Only the decimal Latitude and Longitude values are used to store the geographic coordinates.
    'Setting the other latitude and longitude formats updates these values.
    'Getting the other latitude and longitude formats uses these values
    Private _latitudeDegrees As Integer 'The Degrees component of a Degrees Minutes Seconds latitude
    Property LatitudeDegrees As Integer
        Get
            'Return _latitudeDegrees
            'Return Math.Floor(Latitude)
            If Double.IsNaN(Latitude) Then

            Else
                Return Math.Floor(Math.Abs(Latitude))
            End If
        End Get
        Set(value As Integer)
            _latitudeDegrees = value
            'Latitude = DMSToDecDeg(_latitudeDegrees, LatitudeMinutes, LatitudeSeconds)
            Latitude = _latitudeDegrees + LatitudeMinutes / 60 + LatitudeSeconds / 3600
        End Set
    End Property

    Private _latitudeMinutes 'The Minutes component of a Degrees Minutes Seconds latitude
    Property LatitudeMinutes As Integer
        Get
            'Return _latitudeMinutes
            'Return Math.Floor((Latitude - LatitudeDegrees) * 60)
            If Double.IsNaN(Latitude) Then

            Else
                Return Math.Floor((Math.Abs(Latitude) - LatitudeDegrees) * 60)
            End If
        End Get
        Set(value As Integer)
            _latitudeMinutes = value
            'Latitude = DMSToDecDeg(LatitudeDegrees, _latitudeMinutes, LatitudeSeconds)
            Latitude = LatitudeDegrees + _latitudeMinutes / 60 + LatitudeSeconds / 3600
        End Set
    End Property

    Private _latitudeSeconds As Double 'The Seconds component of a Degrees Minutes Seconds latitude
    Property LatitudeSeconds As Double
        Get
            'Return _latitudeSeconds
            'Return (Latitude - LatitudeDegrees - LatitudeMinutes / 60) * 3600
            If Double.IsNaN(Latitude) Then

            Else
                Return (Math.Abs(Latitude) - LatitudeDegrees - LatitudeMinutes / 60) * 3600
            End If
        End Get
        Set(value As Double)
            _latitudeSeconds = value
            'Latitude = DMSToDecDeg(LatitudeDegrees, LatitudeMinutes, _latitudeSeconds)
            Latitude = LatitudeDegrees + LatitudeMinutes / 60 + _latitudeSeconds / 3600
        End Set
    End Property

    Private _latitudeSecondsFormat As String = "F4" 'The format of the LatitudeSeconds display - this used used by the LatitudeDMS property.
    Property LatitudeSecondsFormat As String
        Get
            Return _latitudeSecondsFormat
        End Get
        Set(value As String)
            _latitudeSecondsFormat = value
        End Set
    End Property

    Private _latitudeSecondsDecPlaces As Integer = 4 'THe number of decimal places used to display Seconds in DMS Latitude string.
    Property LatitudeSecondsDecPlaces As Integer
        Get
            Return _latitudeSecondsDecPlaces
        End Get
        Set(value As Integer)
            _latitudeSecondsDecPlaces = value
            LatitudeSecondsFormat = "F" & _latitudeSecondsDecPlaces
        End Set
    End Property

    Enum LatLongDirStyle
        N_S_E_W
        North_South_East_West
        Positive_Negative
    End Enum

    Private _latLongDirectionStyle As LatLongDirStyle = LatLongDirStyle.N_S_E_W 'N_S_E_W: Use N/S latitude direction and E/W longitude direction. Positive_Negative: Use +/- characters for latitude and longitude directions. North_South_East_West: Use North/South latitude direction and East/West longitude direction.
    Property LatLongDirectionStyle As LatLongDirStyle
        Get
            Return _latLongDirectionStyle
        End Get
        Set(value As LatLongDirStyle)
            _latLongDirectionStyle = value
        End Set
    End Property

    Private _showDmsSymbols As Boolean = False 'If True, the DMS Longitude and Latitude string will include symbols for degrees, minutes and seconds.
    Property ShowDmsSymbols As Boolean
        Get
            Return _showDmsSymbols
        End Get
        Set(value As Boolean)
            _showDmsSymbols = value
        End Set
    End Property

    Private _degMinSecDecimalPlaces As Integer = 4 'The number of decimal places to used in the Seconds display of a Deg Min Sec latitude or longitude string.
    Property DegMinSecDecimalPlaces As Integer
        Get
            Return _degMinSecDecimalPlaces
        End Get
        Set(value As Integer)
            _degMinSecDecimalPlaces = value
            LatitudeSecondsDecPlaces = value
            LongitudeSecondsDecPlaces = value
        End Set
    End Property


    Private _latitudeDirection As String 'The direction of the Latitude value
    Property LatitudeDirection As String
        Get
            Return _latitudeDirection
        End Get
        Set(value As String)
            _latitudeDirection = value
        End Set
    End Property

    Enum LatLongSign
        Positive
        Negative
    End Enum

    Private _latitudeSign As LatLongSign = LatLongSign.Positive 'The Sign of the Latitude value.
    Property LatitudeSign As LatLongSign
        Get
            Return _latitudeSign
        End Get
        Set(value As LatLongSign)
            _latitudeSign = value
        End Set
    End Property

    Private _longitudeDegrees As Integer 'The Degrees component of a Degrees Minutes Seconds longitude
    Property LongitudeDegrees As Integer
        Get
            'Return _longitudeDegrees
            'Return Math.Floor(Longitude)
            If Double.IsNaN(Longitude) Then

            Else
                Return Math.Floor(Math.Abs(Longitude))
            End If
        End Get
        Set(value As Integer)
            _longitudeDegrees = value
            'Longitude = DMSToDecDeg(_longitudeDegrees, LongitudeMinutes, LongitudeSeconds)
            Longitude = _longitudeDegrees + LongitudeMinutes / 60 + LongitudeSeconds / 3600
        End Set
    End Property

    Private _longitudeMinutes As Integer 'The Minutes component of a Degrees Minutes Seconds longitude
    Property LongitudeMinutes As Integer
        Get
            'Return _longitudeMinutes
            'Return Math.Floor((Longitude - LongitudeDegrees) * 60)
            If Double.IsNaN(Longitude) Then

            Else
                Return Math.Floor((Math.Abs(Longitude) - LongitudeDegrees) * 60)
            End If
        End Get
        Set(value As Integer)
            _longitudeMinutes = value
            'Longitude = DMSToDecDeg(LongitudeDegrees, _longitudeMinutes, LongitudeSeconds)
            Longitude = LongitudeDegrees + _longitudeMinutes / 60 + LongitudeSeconds / 3600
        End Set
    End Property

    Private _longitudeSeconds As Double  'The Seconds component of a Degrees Minutes Seconds longitude
    Property LongitudeSeconds As Double
        Get
            'Return _longitudeSeconds
            'Return (Longitude - LongitudeDegrees - LongitudeMinutes / 60) * 3600
            If Double.IsNaN(Longitude) Then

            Else
                Return (Math.Abs(Longitude) - LongitudeDegrees - LongitudeMinutes / 60) * 3600
            End If
        End Get
        Set(value As Double)
            _longitudeSeconds = value
            'Longitude = DMSToDecDeg(LongitudeDegrees, LongitudeMinutes, _longitudeSeconds)
            Longitude = LongitudeDegrees + LongitudeMinutes / 60 + _longitudeSeconds / 3600
        End Set
    End Property

    Private _longitudeSecondsFormat As String = "F4" 'The format of the LongitudeSeconds display - this used used by the LongitudeDMS property.
    Property LongitudeSecondsFormat As String
        Get
            Return _longitudeSecondsFormat
        End Get
        Set(value As String)
            _longitudeSecondsFormat = value
        End Set
    End Property

    Private _longitudeSecondsDecPlaces As Integer = 4 'THe number of decimal places used to display Seconds in DMS Longitude string.
    Property LongitudeSecondsDecPlaces As Integer
        Get
            Return _longitudeSecondsDecPlaces
        End Get
        Set(value As Integer)
            _longitudeSecondsDecPlaces = value
            LongitudeSecondsFormat = "F" & _longitudeSecondsDecPlaces
        End Set
    End Property

    Private _longitudeDirection As String 'The direction of the Longitude value
    Property LongitudeDirection As String
        Get
            Return _longitudeDirection
        End Get
        Set(value As String)
            _longitudeDirection = value
        End Set
    End Property

    Private _longitudeSign As LatLongSign = LatLongSign.Positive 'The Sign of the Latitude value.
    Property LongitudeSign As LatLongSign
        Get
            Return _longitudeSign
        End Get
        Set(value As LatLongSign)
            _longitudeSign = value
        End Set
    End Property


    'Private _latitudeDMS As String 'The Latitude displayed in Degrees Minutes Seconds N/S format
    'Property LatitudeDMS As String
    '    Get
    '        Return _latitudeDMS
    '    End Get
    '    Set(value As String)
    '        _latitudeDMS = value
    '    End Set
    'End Property

    ReadOnly Property LatitudeDMS As String  'The Latitude displayed in Degrees Minutes Seconds N/S format
        Get
            If ShowDmsSymbols Then
                Select Case LatLongDirectionStyle
                    Case LatLongDirStyle.North_South_East_West
                        If LatitudeSign = LatLongSign.Positive Then
                            'Return LatitudeDegrees & AscW(176) & " " & LatitudeMinutes & "' " & Format(LatitudeSeconds, LatitudeSecondsFormat) & """ " & "North"
                            Return LatitudeDegrees & ChrW(176) & " " & LatitudeMinutes & "' " & Format(LatitudeSeconds, LatitudeSecondsFormat) & """ " & "North"
                        ElseIf LatitudeSign = LatLongSign.Negative Then
                            Return LatitudeDegrees & ChrW(176) & " " & LatitudeMinutes & "' " & Format(LatitudeSeconds, LatitudeSecondsFormat) & """ " & "South"
                        Else

                        End If
                    Case LatLongDirStyle.N_S_E_W
                        If LatitudeSign = LatLongSign.Positive Then
                            Return LatitudeDegrees & ChrW(176) & " " & LatitudeMinutes & "' " & Format(LatitudeSeconds, LatitudeSecondsFormat) & """ " & "N"
                        ElseIf LatitudeSign = LatLongSign.Negative Then
                            Return LatitudeDegrees & ChrW(176) & " " & LatitudeMinutes & "' " & Format(LatitudeSeconds, LatitudeSecondsFormat) & """ " & "S"
                        Else

                        End If
                    Case LatLongDirStyle.Positive_Negative
                        If LatitudeSign = LatLongSign.Positive Then
                            Return LatitudeDegrees & ChrW(176) & " " & LatitudeMinutes & "' " & Format(LatitudeSeconds, LatitudeSecondsFormat) & """"
                        ElseIf LatitudeSign = LatLongSign.Negative Then
                            Return "- " & LatitudeDegrees & ChrW(176) & " " & LatitudeMinutes & "' " & Format(LatitudeSeconds, LatitudeSecondsFormat) & """"
                        Else

                        End If
                    Case Else

                End Select
            Else
                Select Case LatLongDirectionStyle
                    Case LatLongDirStyle.North_South_East_West
                        If LatitudeSign = LatLongSign.Positive Then
                            Return LatitudeDegrees & " " & LatitudeMinutes & " " & Format(LatitudeSeconds, LatitudeSecondsFormat) & " " & "North"
                        ElseIf LatitudeSign = LatLongSign.Negative Then
                            Return LatitudeDegrees & " " & LatitudeMinutes & " " & Format(LatitudeSeconds, LatitudeSecondsFormat) & " " & "South"
                        Else

                        End If
                    Case LatLongDirStyle.N_S_E_W
                        If LatitudeSign = LatLongSign.Positive Then
                            Return LatitudeDegrees & " " & LatitudeMinutes & " " & Format(LatitudeSeconds, LatitudeSecondsFormat) & " " & "N"
                        ElseIf LatitudeSign = LatLongSign.Negative Then
                            Return LatitudeDegrees & " " & LatitudeMinutes & " " & Format(LatitudeSeconds, LatitudeSecondsFormat) & " " & "S"
                        Else

                        End If
                    Case LatLongDirStyle.Positive_Negative
                        If LatitudeSign = LatLongSign.Positive Then
                            Return LatitudeDegrees & " " & LatitudeMinutes & " " & Format(LatitudeSeconds, LatitudeSecondsFormat)
                        ElseIf LatitudeSign = LatLongSign.Negative Then
                            Return "- " & LatitudeDegrees & " " & LatitudeMinutes & " " & Format(LatitudeSeconds, LatitudeSecondsFormat)
                        Else

                        End If
                    Case Else

                End Select
            End If

        End Get
    End Property

    'Private _longitudeDMS 'The Longitude displayed in Degrees Minutes Seconds E/W format
    'Property LongitudeDMS As String
    '    Get
    '        Return _longitudeDMS
    '    End Get
    '    Set(value As String)
    '        _longitudeDMS = value
    '    End Set
    'End Property

    ReadOnly Property LongitudeDMS As String  'The Latitude displayed in Degrees Minutes Seconds N/S format
        Get
            If ShowDmsSymbols Then
                Select Case LatLongDirectionStyle
                    Case LatLongDirStyle.North_South_East_West
                        If LongitudeSign = LatLongSign.Positive Then
                            Return LongitudeDegrees & ChrW(176) & " " & LongitudeMinutes & "' " & Format(LongitudeSeconds, LongitudeSecondsFormat) & """ " & "East"
                        ElseIf LongitudeSign = LatLongSign.Negative Then
                            Return LongitudeDegrees & ChrW(176) & " " & LongitudeMinutes & "' " & Format(LongitudeSeconds, LongitudeSecondsFormat) & """ " & "West"
                        Else

                        End If
                    Case LatLongDirStyle.N_S_E_W
                        If LongitudeSign = LatLongSign.Positive Then
                            Return LongitudeDegrees & ChrW(176) & " " & LongitudeMinutes & "' " & Format(LongitudeSeconds, LongitudeSecondsFormat) & """ " & "E"
                        ElseIf LongitudeSign = LatLongSign.Negative Then
                            Return LongitudeDegrees & ChrW(176) & " " & LongitudeMinutes & "' " & Format(LongitudeSeconds, LongitudeSecondsFormat) & """ " & "W"
                        Else

                        End If
                    Case LatLongDirStyle.Positive_Negative
                        If LongitudeSign = LatLongSign.Positive Then
                            Return LongitudeDegrees & ChrW(176) & " " & LongitudeMinutes & "' " & Format(LongitudeSeconds, LongitudeSecondsFormat) & """"
                        ElseIf LongitudeSign = LatLongSign.Negative Then
                            Return "- " & LongitudeDegrees & ChrW(176) & " " & LongitudeMinutes & "' " & Format(LongitudeSeconds, LongitudeSecondsFormat) & """"
                        Else

                        End If
                    Case Else

                End Select
            Else
                Select Case LatLongDirectionStyle
                    Case LatLongDirStyle.North_South_East_West
                        If LongitudeSign = LatLongSign.Positive Then
                            Return LongitudeDegrees & " " & LongitudeMinutes & " " & Format(LongitudeSeconds, LongitudeSecondsFormat) & " " & "East"
                        ElseIf LongitudeSign = LatLongSign.Negative Then
                            Return LongitudeDegrees & " " & LongitudeMinutes & " " & Format(LongitudeSeconds, LongitudeSecondsFormat) & " " & "West"
                        Else

                        End If
                    Case LatLongDirStyle.N_S_E_W
                        If LongitudeSign = LatLongSign.Positive Then
                            Return LongitudeDegrees & " " & LongitudeMinutes & " " & Format(LongitudeSeconds, LongitudeSecondsFormat) & " " & "E"
                        ElseIf LongitudeSign = LatLongSign.Negative Then
                            Return LongitudeDegrees & " " & LongitudeMinutes & " " & Format(LongitudeSeconds, LongitudeSecondsFormat) & " " & "W"
                        Else

                        End If
                    Case LatLongDirStyle.Positive_Negative
                        If LongitudeSign = LatLongSign.Positive Then
                            Return LongitudeDegrees & " " & LongitudeMinutes & " " & Format(LongitudeSeconds, LongitudeSecondsFormat)
                        ElseIf LongitudeSign = LatLongSign.Negative Then
                            Return "- " & LongitudeDegrees & " " & LongitudeMinutes & " " & Format(LongitudeSeconds, LongitudeSecondsFormat)
                        Else

                        End If
                    Case Else

                End Select
            End If
        End Get
    End Property

    Property _easting As Double 'The Easting of a Projected coordinate
    Property Easting As Double
        Get
            Return _easting
        End Get
        Set(value As Double)
            _easting = value
        End Set
    End Property

    Private _northing As Double 'The Northing of a Projected coordinate.
    Property Northing As Double
        Get
            Return _northing
        End Get
        Set(value As Double)
            _northing = value
        End Set
    End Property

    Private _x As Double 'The X value of a Cartesian coordinate
    Property X As Double
        Get
            Return _x
        End Get
        Set(value As Double)
            _x = value
        End Set
    End Property

    Private _y As Double 'The Y value of a Cartesian coordinate.
    Property Y As Double
        Get
            Return _y
        End Get
        Set(value As Double)
            _y = value
        End Set
    End Property

    Private _z As Double 'The Z value of a Cartessian coordinate.
    Property Z As Double
        Get
            Return _z
        End Get
        Set(value As Double)
            _z = value
        End Set
    End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Private Sub DecLongToDMS(DecLong As Double)
        'Calculate the Longitude Degrees, Minutes and Seconds from the Decimal Degrees value.
        LongitudeDegrees = Math.Floor(DecLong)
        LongitudeMinutes = Math.Floor((DecLong - LongitudeDegrees) * 60)
        LongitudeSeconds = (DecLong - LongitudeDegrees - LongitudeMinutes / 60) * 3600
    End Sub

    Private Sub DecLatToDMS(DecLat As Double)
        'Calculate the Latitude Degrees, Minutes and Seconds from the Decimal Degrees value.
        LatitudeDegrees = Math.Floor(DecLat)
        LatitudeMinutes = Math.Floor((DecLat - LatitudeDegrees) * 60)
        LatitudeSeconds = (DecLat - LatitudeDegrees - LatitudeMinutes / 60) * 3600
    End Sub

    Private Function DMSToDecDeg(Degrees As Integer, Minutes As Integer, Seconds As Double) As Double
        'Calculate the Decimal Degrees from the Degrees, Minutes and Seconds values.
        Return Degrees + Minutes / 60 + Seconds / 3600
    End Function

    Enum UpdateMode
        None
        InputOutputAll
        XYZ
        LongLat
        EastNorth
        All
        TransXYZ
        TransLongLat
        TransEastNorth
        TransAll
    End Enum

    Enum CoordType
        LongLat
        EastNorth
        XYZ
    End Enum

    Public Sub SetLongitude(StringVal As String, Mode As UpdateMode)
        'Read the Longtitude value from the StringVal
        'Use the value to update the specified location data.

        'Longitude examples:
        '123 40 18.455 East
        '98 29 18.333 West
        '-98 29 18.333
        '123.4567321 East
        '-340.328640
        If IsNothing(StringVal) Then

        Else
            Dim MatchFound As Boolean = False

            'Attempt to match Deg-Min-Sec angle value:
            Dim RegExPattern As String = "^\s*(?<LeadingSign>\+|\-|)\s*(?<Degrees>\d{1,3})\s*[\u00B0\u2103\u2109\u00BA\u02DA]{0,1}\s+(?<Minutes>\d{1,2})\s*[\u2019\u0027\u2032]{0,1}\s+(?<Seconds>\d{1,2}\.\d*)\s*[\u201D\u0022\u2033\u02DD\u00A8]{0,1}\s*(?<Direction>(?i)West|Wst|W|East|Est|E|South|Sth|S|North|Nth|N){0,1}"
            Dim myRegEx As New System.Text.RegularExpressions.Regex(RegExPattern)
            Dim myMatch As System.Text.RegularExpressions.Match = myRegEx.Match(StringVal)
            If myMatch.Success Then 'Deg Min Sec Longitude value
                MatchFound = True
                Dim LeadingSign As String = myMatch.Groups("LeadingSign").ToString
                Dim Degrees As Integer = Val(myMatch.Groups("Degrees").ToString)
                Dim Minutes As Integer = Val(myMatch.Groups("Minutes").ToString)
                Dim Seconds As Double = Val(myMatch.Groups("Seconds").ToString)
                Dim Direction As String = myMatch.Groups("Direction").ToString
                If Direction.ToLower = "south" Or Direction.ToLower = "sth" Or Direction.ToLower = "s" Or Direction.ToLower = "north" Or Direction.ToLower = "nth" Or Direction.ToLower = "n" Then
                    'The angle is a Latitude not a Longitude!
                    RaiseEvent ErrorMessage("The angle string represents a Latitude, not a Longitude." & vbCrLf)
                ElseIf LeadingSign = "-" Or Direction.ToLower = "west" Or Direction.ToLower = "wst" Or Direction.ToLower = "w" Then
                    Longitude = -Degrees - Minutes / 60 - Seconds / 3600
                Else
                    Longitude = Degrees + Minutes / 60 + Seconds / 3600
                End If
            Else
                'Attempt to match a Decimal Longitude value:
                RegExPattern = "^\s*(?<LeadingSign>\+|\-|)\s*(?<DecimalDegrees>\d{1,3}\.\d*|\d{1,3})\s*(?<Direction>(?i)West|Wst|W|East|Est|E|)"
                'Named capture LeadingSign : "+", "-" or ""
                'Named capture DecimalDegrees : 1 to 3 digits OR 1 to 3 digits & "." & 0 or more digits
                'Named capture Direction : N or Nth or North or S or Sth or South (case insensitive)
                Dim myRegEx2 As New System.Text.RegularExpressions.Regex(RegExPattern)
                Dim myMatch2 As System.Text.RegularExpressions.Match = myRegEx2.Match(StringVal)
                If myMatch2.Success Then 'Decimal Longitude value
                    MatchFound = True
                    Dim LeadingSign As String = myMatch2.Groups("LeadingSign").ToString
                    Dim DecimalDegrees As Double = Val(myMatch2.Groups("DecimalDegrees").ToString)
                    Dim Direction As String = myMatch2.Groups("Direction").ToString
                    If Direction.ToLower = "south" Or Direction.ToLower = "sth" Or Direction.ToLower = "s" Or Direction.ToLower = "north" Or Direction.ToLower = "nth" Or Direction.ToLower = "n" Then
                        'The angle is a Latitude not a Longitude!
                        RaiseEvent ErrorMessage("The angle string represents a Latitude, not a Longitude." & vbCrLf)
                    ElseIf LeadingSign = "-" Or Direction.ToLower = "west" Or Direction.ToLower = "wst" Or Direction.ToLower = "w" Then
                        Longitude = -DecimalDegrees
                    Else
                        Longitude = DecimalDegrees
                        End If

                    End If
                End If

            ''Attempt to match a Decimal Longitude value:
            'Dim RegExPattern As String = "^\s*(?<LeadingSign>\+|\-|)\s*(?<DecimalDegrees>\d{1,3}\.\d*|\d{1,3})\s*(?<Direction>(?i)West|Wst|W|East|Est|E|)"
            '    'Named capture LeadingSign : "+", "-" or ""
            '    'Named capture DecimalDegrees : 1 to 3 digits OR 1 to 3 digits & "." & 0 or more digits
            '    'Named capture Direction : N or Nth or North or S or Sth or South (case insensitive)
            '    Dim myRegEx As New System.Text.RegularExpressions.Regex(RegExPattern)
            '    Dim myMatch As System.Text.RegularExpressions.Match = myRegEx.Match(StringVal)
            'If myMatch.Success Then 'Decimal Longitude value
            '    MatchFound = True
            '    Dim LeadingSign As String = myMatch.Groups("LeadingSign").ToString
            '    Dim DecimalDegrees As Double = Val(myMatch.Groups("DecimalDegrees").ToString)
            '    Dim Direction As String = myMatch.Groups("Direction").ToString
            '    If LeadingSign = "-" Or Direction.ToLower = "west" Or Direction.ToLower = "wst" Or Direction.ToLower = "w" Then
            '        Longitude = -DecimalDegrees
            '    Else
            '        Longitude = DecimalDegrees
            '    End If

            'Else
            '    RegExPattern = "^\s*(?<LeadingSign>\+|\-|)\s*(?<Degrees>\d{1,3})\s+(?<Minutes>\d{1,2})\s+(?<Seconds>\d{1,2}|\d{1,2}\.\d*)\s*(?<Direction>(?i)West|Wst|W|East|Est|E|)"
            '    'Named capture LeadingSign : "+", "-" or ""
            '    'Named capture Degrees : 1 to 3 digits
            '    'Named capture Minutes : 1 to 2 digits
            '    'Named capture Seconds : 1 to 2 digits OR 1 to 2 digits & "." & 0 or more digits
            '    'Named capture Direction : N or Nth or North or S or Sth or South (case insensitive)

            '    Dim myRegEx2 As New System.Text.RegularExpressions.Regex(RegExPattern)
            '    Dim myMatch2 As System.Text.RegularExpressions.Match = myRegEx2.Match(StringVal)
            '    If myMatch2.Success Then 'Deg Min Sec Longitude value
            '        MatchFound = True
            '        Dim LeadingSign As String = myMatch.Groups("LeadingSign").ToString
            '        Dim Degrees As Integer = Val(myMatch.Groups("Degrees").ToString)
            '        Dim Minutes As Integer = Val(myMatch.Groups("Minutes").ToString)
            '        Dim Seconds As Double = Val(myMatch.Groups("Seconds").ToString)
            '        Dim Direction As String = myMatch.Groups("Direction").ToString
            '        If LeadingSign = "-" Or Direction.ToLower = "west" Or Direction.ToLower = "wst" Or Direction.ToLower = "w" Then
            '            Longitude = -Degrees - Minutes / 60 - Seconds / 3600
            '        Else
            '            Longitude = Degrees + Minutes / 60 + Seconds / 3600
            '        End If

            '    Else
            '        RaiseEvent ErrorMessage("The Longitude string could not be matched: " & StringVal & vbCrLf)
            '    End If
            'End If

            If MatchFound Then
                RaiseEvent Update(Mode, CoordType.LongLat)
            End If
        End If
    End Sub
    Public Sub SetLatitude(StringVal As String, Mode As UpdateMode)
        'Read the Latitude value from the StringVal
        'Use the value to update the specified location data.

        If IsNothing(StringVal) Then

        Else
            Dim MatchFound As Boolean = False

            'Attempt to match Deg-Min-Sec angle value:
            Dim RegExPattern As String = "^\s*(?<LeadingSign>\+|\-|)\s*(?<Degrees>\d{1,3})\s*[\u00B0\u2103\u2109\u00BA\u02DA]{0,1}\s+(?<Minutes>\d{1,2})\s*[\u2019\u0027\u2032]{0,1}\s+(?<Seconds>\d{1,2}\.\d*)\s*[\u201D\u0022\u2033\u02DD\u00A8]{0,1}\s*(?<Direction>(?i)West|Wst|W|East|Est|E|South|Sth|S|North|Nth|N){0,1}"
            Dim myRegEx As New System.Text.RegularExpressions.Regex(RegExPattern)
            Dim myMatch As System.Text.RegularExpressions.Match = myRegEx.Match(StringVal)
            If myMatch.Success Then 'Deg Min Sec Longitude value
                MatchFound = True
                Dim LeadingSign As String = myMatch.Groups("LeadingSign").ToString
                Dim Degrees As Integer = Val(myMatch.Groups("Degrees").ToString)
                Dim Minutes As Integer = Val(myMatch.Groups("Minutes").ToString)
                Dim Seconds As Double = Val(myMatch.Groups("Seconds").ToString)
                Dim Direction As String = myMatch.Groups("Direction").ToString
                If Direction.ToLower = "west" Or Direction.ToLower = "wst" Or Direction.ToLower = "w" Or Direction.ToLower = "east" Or Direction.ToLower = "est" Or Direction.ToLower = "e" Then
                    'The angle is a Longitude not a Latitude!
                    RaiseEvent ErrorMessage("The angle string represents a Longitude, not a Latitude." & vbCrLf)
                ElseIf LeadingSign = "-" Or Direction.ToLower = "south" Or Direction.ToLower = "sth" Or Direction.ToLower = "s" Then
                    Latitude = -Degrees - Minutes / 60 - Seconds / 3600
                Else
                    Latitude = Degrees + Minutes / 60 + Seconds / 3600
                End If
            Else
                'Attempt to match a Decimal Longitude value:
                RegExPattern = "^\s*(?<LeadingSign>\+|\-|)\s*(?<DecimalDegrees>\d{1,3}\.\d*|\d{1,3})\s*(?<Direction>(?i)West|Wst|W|East|Est|E|)"
                'Named capture LeadingSign : "+", "-" or ""
                'Named capture DecimalDegrees : 1 to 3 digits OR 1 to 3 digits & "." & 0 or more digits
                'Named capture Direction : N or Nth or North or S or Sth or South (case insensitive)
                Dim myRegEx2 As New System.Text.RegularExpressions.Regex(RegExPattern)
                Dim myMatch2 As System.Text.RegularExpressions.Match = myRegEx2.Match(StringVal)
                If myMatch2.Success Then 'Decimal Longitude value
                    MatchFound = True
                    Dim LeadingSign As String = myMatch2.Groups("LeadingSign").ToString
                    Dim DecimalDegrees As Double = Val(myMatch2.Groups("DecimalDegrees").ToString)
                    Dim Direction As String = myMatch2.Groups("Direction").ToString
                    If Direction.ToLower = "west" Or Direction.ToLower = "wst" Or Direction.ToLower = "w" Or Direction.ToLower = "east" Or Direction.ToLower = "est" Or Direction.ToLower = "e" Then
                        'The angle is a Longitude not a Latitude!
                        RaiseEvent ErrorMessage("The angle string represents a Longitude, not a Latitude." & vbCrLf)
                    ElseIf LeadingSign = "-" Or Direction.ToLower = "south" Or Direction.ToLower = "sth" Or Direction.ToLower = "s" Then
                        Latitude = -DecimalDegrees
                    Else
                        Latitude = DecimalDegrees
                        End If

                    End If
                End If



            ''Attempt to match a Decimal Latitude value:
            'Dim RegExPattern As String = "^\s*(?<LeadingSign>\+|\-|)\s*(?<DecimalDegrees>\d{1,3}\.\d*|\d{1,3})\s*(?<Direction>(?i)South|Sth|S|North|Nth|N|)"
            'Dim myRegEx As New System.Text.RegularExpressions.Regex(RegExPattern)
            'Dim myMatch As System.Text.RegularExpressions.Match = myRegEx.Match(StringVal)
            'If myMatch.Success Then 'Decimal Latitude value
            '    MatchFound = True
            '    Dim LeadingSign As String = myMatch.Groups("LeadingSign").ToString
            '    Dim DecimalDegrees As Double = Val(myMatch.Groups("DecimalDegrees").ToString)
            '    Dim Direction As String = myMatch.Groups("Direction").ToString
            '    If LeadingSign = "-" Or Direction.ToLower = "south" Or Direction.ToLower = "sth" Or Direction.ToLower = "s" Then
            '        Latitude = -DecimalDegrees
            '    Else
            '        Latitude = DecimalDegrees
            '    End If
            'Else
            '    RegExPattern = "^\s*(?<LeadingSign>\+|\-|)\s*(?<Degrees>\d{1,3})\s+(?<Minutes>\d{1,2})\s+(?<Seconds>\d{1,2}|\d{1,2}\.\d*)\s*(?<Direction>(?i)South|Sth|S|North|Nth|N|)"
            '    Dim myRegEx2 As New System.Text.RegularExpressions.Regex(RegExPattern)
            '    Dim myMatch2 As System.Text.RegularExpressions.Match = myRegEx2.Match(StringVal)
            '    If myMatch2.Success Then 'Deg Min Sec Latitude value
            '        MatchFound = True
            '        Dim LeadingSign As String = myMatch.Groups("LeadingSign").ToString
            '        Dim Degrees As Integer = Val(myMatch.Groups("Degrees").ToString)
            '        Dim Minutes As Integer = Val(myMatch.Groups("Minutes").ToString)
            '        Dim Seconds As Double = Val(myMatch.Groups("Seconds").ToString)
            '        Dim Direction As String = myMatch.Groups("Direction").ToString
            '        If LeadingSign = "-" Or Direction.ToLower = "south" Or Direction.ToLower = "sth" Or Direction.ToLower = "s" Then
            '            Latitude = -Degrees - Minutes / 60 - Seconds / 3600
            '        Else
            '            Latitude = Degrees + Minutes / 60 + Seconds / 3600
            '        End If
            '    Else
            '        RaiseEvent ErrorMessage("The Latitude string could not be matched: " & StringVal & vbCrLf)
            '    End If
            '    'If MatchFound Then
            '    '    RaiseEvent Update(Mode, CoordType.LongLat)
            '    'End If
            'End If


            If MatchFound Then
                RaiseEvent Update(Mode, CoordType.LongLat)
            End If
        End If
    End Sub

    Public Sub SetEllipsoidalHeight(HeightVal As String, Mode As UpdateMode)
        'Read the Ellipsoidal Height value from the StringVal
        If IsNothing(HeightVal) Then

        Else
            EllipsoidalHeight = HeightVal
            RaiseEvent Update(Mode, CoordType.LongLat)
        End If
    End Sub


    Public Sub SetEasting(StringVal As String, Mode As UpdateMode)
        'Read the Easting value from the StringVal
        'Use the value to update the specified location data.

        If IsNothing(StringVal) Then

        Else
            'Attempt to match a Decimal Easting value:
            'Dim RegExPattern As String = "^\s*(?<LeadingSign>\+|\-|)\s*(?<DecimalEasting>\d{1,3}\.\d*|\d{1,3})\s*(?<Direction>(?i)West|Wst|W|East|Est|E|)"
            Dim RegExPattern As String = "^\s*(?<LeadingSign>\+|\-|)\s*(?<DecimalEasting>\d*\.\d*|\d*)\s*(?<Direction>(?i)West|Wst|W|East|Est|E|)"
            Dim myRegEx As New System.Text.RegularExpressions.Regex(RegExPattern)
            Dim myMatch As System.Text.RegularExpressions.Match = myRegEx.Match(StringVal)
            If myMatch.Success Then 'Decimal Easting value
                Dim LeadingSign As String = myMatch.Groups("LeadingSign").ToString
                Dim DecimalEasting As Double = Val(myMatch.Groups("DecimalEasting").ToString)
                Dim Direction As String = myMatch.Groups("Direction").ToString
                If LeadingSign = "-" Or Direction.ToLower = "west" Or Direction.ToLower = "wst" Or Direction.ToLower = "w" Then
                    Easting = -DecimalEasting
                Else
                    Easting = DecimalEasting
                End If
                RaiseEvent Update(Mode, CoordType.EastNorth)
            End If
        End If
    End Sub

    Public Sub SetNorthing(StringVal As String, Mode As UpdateMode)
        'Read the Northing value from the StringVal
        'Use the value to update the specified location data.

        If IsNothing(StringVal) Then

        Else
            'Attempt to match a Decimal Northing value:
            'Dim RegExPattern As String = "^\s*(?<LeadingSign>\+|\-|)\s*(?<DecimalNorthing>\d{1,3}\.\d*|\d{1,3})\s*(?<Direction>(?i)South|Sth|S|North|Nth|N|)"
            Dim RegExPattern As String = "^\s*(?<LeadingSign>\+|\-|)\s*(?<DecimalNorthing>\d*\.\d*|\d*)\s*(?<Direction>(?i)South|Sth|S|North|Nth|N|)"
            Dim myRegEx As New System.Text.RegularExpressions.Regex(RegExPattern)
            Dim myMatch As System.Text.RegularExpressions.Match = myRegEx.Match(StringVal)
            If myMatch.Success Then 'Decimal Northing value
                Dim LeadingSign As String = myMatch.Groups("LeadingSign").ToString
                Dim DecimalNorthing As Double = Val(myMatch.Groups("DecimalNorthing").ToString)
                Dim Direction As String = myMatch.Groups("Direction").ToString
                If LeadingSign = "-" Or Direction.ToLower = "south" Or Direction.ToLower = "sth" Or Direction.ToLower = "s" Then
                    Northing = -DecimalNorthing
                Else
                    Northing = DecimalNorthing
                End If
                RaiseEvent Update(Mode, CoordType.EastNorth)
            End If
        End If
    End Sub

    Public Sub SetX(XVal As String, Mode As UpdateMode)
        'Read the Cartestian X value from the StringVal

        If IsNothing(XVal) Then

        Else
            X = XVal
            RaiseEvent Update(Mode, CoordType.XYZ)
        End If
    End Sub

    Public Sub SetY(YVal As String, Mode As UpdateMode)
        'Read the Cartestian Y value from the StringVal

        If IsNothing(YVal) Then

        Else
            Y = YVal
            RaiseEvent Update(Mode, CoordType.XYZ)
        End If
    End Sub

    Public Sub SetZ(ZVal As String, Mode As UpdateMode)
        'Read the Cartestian Z value from the StringVal

        If IsNothing(ZVal) Then

        Else
            Z = ZVal
            RaiseEvent Update(Mode, CoordType.XYZ)
        End If
    End Sub

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================
    Event ErrorMessage(ByVal Msg As String) 'Send an error message.
    Event Message(ByVal Msg As String) 'Send a normal message.
    Event Update(ByVal Mode As UpdateMode, ByVal From As CoordType)
#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class 'Coordinate

Public Class clsAngle
    'Represents and angle using different formats: Decimal Degree, Degree-Minute-Second +/- N-S/E-W


#Region " Properties - All the properties used in this class." '===========================================================================================================================

    Enum AngleType
        Longitude
        Latitude
        General
    End Enum

    Private _type As AngleType = AngleType.General 'The type of angle: Longitude, Latitude or General angle.
    Property Type As AngleType
        Get
            Return _type
        End Get
        Set(value As AngleType)
            _type = value
        End Set
    End Property

    Private _decimalDegrees As Double = Double.NaN 'The angle in decimal degrees. All other angle formats are generated from this value.
    Property DecimalDegrees As Double
        Get
            Return _decimalDegrees
        End Get
        Set(value As Double)
            _decimalDegrees = value
        End Set
    End Property

    Private _decDegreesFormat As String = "F4" 'The format used to display the Decimal Degrees.
    Property DecDegreesFormat As String
        Get
            Return _decDegreesFormat
        End Get
        Set(value As String)
            _decDegreesFormat = value
        End Set
    End Property

    Private _decDegreeDecPlaces As Integer = 4 'The number of decimal places in the Formatted Decimal Degrees.
    Property DecDegreeDecPlaces As Integer
        Get
            Return _decDegreeDecPlaces
        End Get
        Set(value As Integer)
            _decDegreeDecPlaces = value
            DecDegreesFormat = "F" & _decDegreeDecPlaces
        End Set
    End Property

    ReadOnly Property FormattedDecimalDegrees As String
        Get
            Return Format(DecimalDegrees, DecDegreesFormat)
        End Get
    End Property

    Private _degrees As Integer 'The Degrees component of an angle in Degrees-Minutes-Seconds format.
    Property Degrees As Integer
        Get
            'Return _degrees

            If Double.IsNaN(DecimalDegrees) Then

            Else
                Return Math.Floor(Math.Abs(DecimalDegrees))
            End If

        End Get
        Set(value As Integer)
            _degrees = value
            DecimalDegrees = _degrees + Minutes / 60 + Seconds / 3600
        End Set
    End Property

    Private _minutes As Integer 'The Minutes component of an angle in Degrees-Minutes-Seconds format.
    Property Minutes As Integer
        Get
            'Return _minutes
            If Double.IsNaN(DecimalDegrees) Then

            Else
                Return Math.Floor((Math.Abs(DecimalDegrees) - Degrees) * 60)
            End If
        End Get
        Set(value As Integer)
            _minutes = value
            DecimalDegrees = Degrees + _minutes / 60 + Seconds / 3600
        End Set
    End Property

    Private _seconds As Double 'The Seconds component of an angle in Degrees-Minutes-Seconds format.
    Property Seconds As Double
        Get
            'Return _seconds
            If Double.IsNaN(DecimalDegrees) Then

            Else
                Return (Math.Abs(DecimalDegrees) - Degrees - Minutes / 60) * 3600
            End If

        End Get
        Set(value As Double)
            _seconds = value
            DecimalDegrees = Degrees + Minutes / 60 + _seconds / 3600
        End Set
    End Property

    Private _secondsFormat As String = "F4" 'The format used to display the Seconds - this used used by the DMS property.
    Property SecondsFormat As String
        Get
            Return _secondsFormat
        End Get
        Set(value As String)
            _secondsFormat = value
        End Set
    End Property

    Private _secondsDecPlaces As Integer = 4 'The number of decimal places used to display Seconds in DMS string.
    Property SecondsDecPlaces As Integer
        Get
            Return _secondsDecPlaces
        End Get
        Set(value As Integer)
            _secondsDecPlaces = value
            SecondsFormat = "F" & _secondsDecPlaces
        End Set
    End Property

    ReadOnly Property FormattedSeconds As String
        Get
            Return Format(Seconds, SecondsFormat)
        End Get
    End Property

    Enum LatLongDirStyle
        N_S_E_W
        North_South_East_West
        Positive_Negative
    End Enum

    Private _latLongDirectionStyle As LatLongDirStyle = LatLongDirStyle.N_S_E_W 'N_S_E_W: Use N/S latitude direction and E/W longitude direction. Positive_Negative: Use +/- characters for latitude and longitude directions. North_South_East_West: Use North/South latitude direction and East/West longitude direction.
    Property LatLongDirectionStyle As LatLongDirStyle
        Get
            Return _latLongDirectionStyle
        End Get
        Set(value As LatLongDirStyle)
            _latLongDirectionStyle = value
        End Set
    End Property

    Private _showDmsSymbols As Boolean = False 'If True, the DMS Longitude and Latitude string will include symbols for degrees, minutes and seconds.
    Property ShowDmsSymbols As Boolean
        Get
            Return _showDmsSymbols
        End Get
        Set(value As Boolean)
            _showDmsSymbols = value
        End Set
    End Property

    Private _degMinSecDecimalPlaces As Integer = 4 'The number of decimal places to used in the Seconds display of a Deg Min Sec angle string.
    Property DegMinSecDecimalPlaces As Integer
        Get
            Return _degMinSecDecimalPlaces
        End Get
        Set(value As Integer)
            _degMinSecDecimalPlaces = value
            SecondsDecPlaces = value
        End Set
    End Property

    Enum AngleSign
        Positive
        Negative
    End Enum

    Private _sign As AngleSign = AngleSign.Positive 'The Sign of the angle value.
    Property Sign As AngleSign
        Get
            Return _sign
        End Get
        Set(value As AngleSign)
            _sign = value
        End Set
    End Property

    ReadOnly Property DegMinSec As String
        Get
            If Type = AngleType.General Then
                If ShowDmsSymbols Then
                    If Sign = AngleSign.Positive Then
                        Return Degrees & ChrW(176) & " " & Minutes & "' " & Format(Seconds, SecondsFormat) & """"
            ElseIf Sign = AngleSign.Negative Then
                        Return "- " & Degrees & ChrW(176) & " " & Minutes & "' " & Format(Seconds, SecondsFormat) & """"
                    End If
                Else
                    If Sign = AngleSign.Positive Then
                        'Return Degrees & " " & Minutes & " " & Format(Seconds, SecondsFormat)
                        Return Degrees & " " & Minutes & " " & FormattedSeconds
                    ElseIf Sign = AngleSign.Negative Then
                        Return "- " & Degrees & " " & Minutes & " " & FormattedSeconds
                    End If
                End If
            ElseIf Type = AngleType.Longitude Then
                If ShowDmsSymbols Then
                    Select Case LatLongDirectionStyle
                        Case LatLongDirStyle.North_South_East_West
                            If Sign = AngleSign.Positive Then
                                Return Degrees & ChrW(176) & " " & Minutes & "' " & FormattedSeconds & """ " & "East"
                            ElseIf Sign = AngleSign.Negative Then
                                Return Degrees & ChrW(176) & " " & Minutes & "' " & FormattedSeconds & """ " & "West"
                            Else

                            End If
                        Case LatLongDirStyle.N_S_E_W
                            If Sign = AngleSign.Positive Then
                                Return Degrees & ChrW(176) & " " & Minutes & "' " & FormattedSeconds & """ " & "E"
                            ElseIf Sign = AngleSign.Negative Then
                                Return Degrees & ChrW(176) & " " & Minutes & "' " & FormattedSeconds & """ " & "W"
                            Else

                            End If
                        Case LatLongDirStyle.Positive_Negative
                            If Sign = AngleSign.Positive Then
                                Return Degrees & ChrW(176) & " " & Minutes & "' " & FormattedSeconds & """ "
                            ElseIf Sign = AngleSign.Negative Then
                                Return "- " & Degrees & ChrW(176) & " " & Minutes & "' " & FormattedSeconds & """ "
                            Else

                            End If
                    End Select
                Else
                    Select Case LatLongDirectionStyle
                        Case LatLongDirStyle.North_South_East_West
                            If Sign = AngleSign.Positive Then
                                Return Degrees & " " & Minutes & " " & FormattedSeconds & " East"
                            ElseIf Sign = AngleSign.Negative Then
                                Return Degrees & " " & Minutes & " " & FormattedSeconds & " West"
                            Else

                            End If
                        Case LatLongDirStyle.N_S_E_W
                            If Sign = AngleSign.Positive Then
                                Return Degrees & " " & Minutes & " " & FormattedSeconds & " E"
                            ElseIf Sign = AngleSign.Negative Then
                                Return Degrees & " " & Minutes & " " & FormattedSeconds & " W"
                            Else

                            End If
                        Case LatLongDirStyle.Positive_Negative
                            If Sign = AngleSign.Positive Then
                                Return Degrees & " " & Minutes & " " & FormattedSeconds
                            ElseIf Sign = AngleSign.Negative Then
                                Return "- " & Degrees & " " & Minutes & " " & FormattedSeconds
                            Else

                            End If
                    End Select
                End If
            ElseIf Type = AngleType.Latitude Then
                If ShowDmsSymbols Then
                    Select Case LatLongDirectionStyle
                        Case LatLongDirStyle.North_South_East_West
                            If Sign = AngleSign.Positive Then
                                Return Degrees & ChrW(176) & " " & Minutes & "' " & FormattedSeconds & """ " & "North"
                            ElseIf Sign = AngleSign.Negative Then
                                Return Degrees & ChrW(176) & " " & Minutes & "' " & FormattedSeconds & """ " & "South"
                            Else

                            End If
                        Case LatLongDirStyle.N_S_E_W
                            If Sign = AngleSign.Positive Then
                                Return Degrees & ChrW(176) & " " & Minutes & "' " & FormattedSeconds & """ " & "N"
                            ElseIf Sign = AngleSign.Negative Then
                                Return Degrees & ChrW(176) & " " & Minutes & "' " & FormattedSeconds & """ " & "S"
                            Else

                            End If

                        Case LatLongDirStyle.Positive_Negative
                            If Sign = AngleSign.Positive Then
                                Return Degrees & ChrW(176) & " " & Minutes & "' " & FormattedSeconds & """ "
                            ElseIf Sign = AngleSign.Negative Then
                                Return "- " & Degrees & ChrW(176) & " " & Minutes & "' " & FormattedSeconds & """ "
                            Else

                            End If
                    End Select
                Else
                    Select Case LatLongDirectionStyle
                        Case LatLongDirStyle.North_South_East_West
                            If Sign = AngleSign.Positive Then
                                Return Degrees & " " & Minutes & " " & FormattedSeconds & " North"
                            ElseIf Sign = AngleSign.Negative Then
                                Return Degrees & " " & Minutes & " " & FormattedSeconds & " South"
                            Else

                            End If
                        Case LatLongDirStyle.N_S_E_W
                            If Sign = AngleSign.Positive Then
                                Return Degrees & " " & Minutes & " " & FormattedSeconds & " N"
                            ElseIf Sign = AngleSign.Negative Then
                                Return Degrees & " " & Minutes & " " & FormattedSeconds & " S"
                            Else

                            End If
                        Case LatLongDirStyle.Positive_Negative
                            If Sign = AngleSign.Positive Then
                                Return Degrees & " " & Minutes & " " & FormattedSeconds
                            ElseIf Sign = AngleSign.Negative Then
                                Return "- " & Degrees & " " & Minutes & " " & FormattedSeconds
                            Else

                            End If
                    End Select
                End If

            End If
        End Get
    End Property

    Private _match As Boolean = False
    Property Match As Boolean
        Get
            Return _match
        End Get
        Set(value As Boolean)
            _match = value
        End Set
    End Property

#End Region 'Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods - The main actions performed by this class." '===========================================================================================================================

    Public Sub SetLongitude(StringVal As String)
        'Read the Longtitude value from the StringVal
        'Use the value to update the specified location data.

        'Longitude examples:
        '123 40 18.455 East
        '98 29 18.333 West
        '-98 29 18.333
        '123.4567321 East
        '-340.328640

        'Dim MatchFound As Boolean = False
        Match = False

        'Attempt to match a Decimal Longitude value:
        Dim RegExPattern As String = "^\s*(?<LeadingSign>\+|\-|)\s*(?<DecimalDegrees>\d{1,3}\.\d*|\d{1,3})\s*(?<Direction>(?i)West|Wst|W|East|Est|E|)"
        'Named capture LeadingSign : "+", "-" or ""
        'Named capture DecimalDegrees : 1 to 3 digits OR 1 to 3 digits & "." & 0 or more digits
        'Named capture Direction : N or Nth or North or S or Sth or South (case insensitive)
        Dim myRegEx As New System.Text.RegularExpressions.Regex(RegExPattern)
        Dim myMatch As System.Text.RegularExpressions.Match = myRegEx.Match(StringVal)
        If myMatch.Success Then 'Decimal Longitude value
            'MatchFound = True
            Match = True
            Type = AngleType.Longitude
            Dim LeadingSign As String = myMatch.Groups("LeadingSign").ToString
            Dim DecDegrees As Double = Val(myMatch.Groups("DecimalDegrees").ToString)
            Dim Direction As String = myMatch.Groups("Direction").ToString
            If LeadingSign = "-" Or Direction.ToLower = "west" Or Direction.ToLower = "wst" Or Direction.ToLower = "w" Then
                DecimalDegrees = -DecDegrees
            Else
                DecimalDegrees = DecDegrees
            End If

        Else
            RegExPattern = "^\s*(?<LeadingSign>\+|\-|)\s*(?<Degrees>\d{1,3})\s+(?<Minutes>\d{1,2})\s+(?<Seconds>\d{1,2}|\d{1,2}\.\d*)\s*(?<Direction>(?i)West|Wst|W|East|Est|E|)"
            'Named capture LeadingSign : "+", "-" or ""
            'Named capture Degrees : 1 to 3 digits
            'Named capture Minutes : 1 to 2 digits
            'Named capture Seconds : 1 to 2 digits OR 1 to 2 digits & "." & 0 or more digits
            'Named capture Direction : N or Nth or North or S or Sth or South (case insensitive)

            Dim myRegEx2 As New System.Text.RegularExpressions.Regex(RegExPattern)
            Dim myMatch2 As System.Text.RegularExpressions.Match = myRegEx2.Match(StringVal)
            If myMatch2.Success Then 'Deg Min Sec Longitude value
                Match = True
                Type = AngleType.Longitude
                Dim LeadingSign As String = myMatch.Groups("LeadingSign").ToString
                Dim Degrees As Integer = Val(myMatch.Groups("Degrees").ToString)
                Dim Minutes As Integer = Val(myMatch.Groups("Minutes").ToString)
                Dim Seconds As Double = Val(myMatch.Groups("Seconds").ToString)
                Dim Direction As String = myMatch.Groups("Direction").ToString
                If LeadingSign = "-" Or Direction.ToLower = "west" Or Direction.ToLower = "wst" Or Direction.ToLower = "w" Then
                    DecimalDegrees = -Degrees - Minutes / 60 - Seconds / 3600
                Else
                    DecimalDegrees = Degrees + Minutes / 60 + Seconds / 3600
                End If
            Else
                RaiseEvent ErrorMessage("The Longitude string could not be matched: " & StringVal & vbCrLf)
            End If
        End If
    End Sub

    Public Sub SetLatitude(StringVal As String)
        'Read the Latitude value from the StringVal
        'Use the value to update the specified location data.

        Match = False

        'Attempt to match a Decimal Latitude value:
        Dim RegExPattern As String = "^\s*(?<LeadingSign>\+|\-|)\s*(?<DecimalDegrees>\d{1,3}\.\d*|\d{1,3})\s*(?<Direction>(?i)South|Sth|S|North|Nth|N|)"
        Dim myRegEx As New System.Text.RegularExpressions.Regex(RegExPattern)
        Dim myMatch As System.Text.RegularExpressions.Match = myRegEx.Match(StringVal)
        If myMatch.Success Then 'Decimal Latitude value
            Match = True
            Type = AngleType.Latitude
            Dim LeadingSign As String = myMatch.Groups("LeadingSign").ToString
            Dim DecDegrees As Double = Val(myMatch.Groups("DecimalDegrees").ToString)
            Dim Direction As String = myMatch.Groups("Direction").ToString
            If LeadingSign = "-" Or Direction.ToLower = "south" Or Direction.ToLower = "sth" Or Direction.ToLower = "s" Then
                DecimalDegrees = -DecDegrees
            Else
                DecimalDegrees = DecDegrees
            End If
        Else
            RegExPattern = "^\s*(?<LeadingSign>\+|\-|)\s*(?<Degrees>\d{1,3})\s+(?<Minutes>\d{1,2})\s+(?<Seconds>\d{1,2}|\d{1,2}\.\d*)\s*(?<Direction>(?i)South|Sth|S|North|Nth|N|)"
            Dim myRegEx2 As New System.Text.RegularExpressions.Regex(RegExPattern)
            Dim myMatch2 As System.Text.RegularExpressions.Match = myRegEx2.Match(StringVal)
            If myMatch2.Success Then 'Deg Min Sec Latitude value
                Match = True
                Type = AngleType.Latitude
                Dim LeadingSign As String = myMatch.Groups("LeadingSign").ToString
                Dim Degrees As Integer = Val(myMatch.Groups("Degrees").ToString)
                Dim Minutes As Integer = Val(myMatch.Groups("Minutes").ToString)
                Dim Seconds As Double = Val(myMatch.Groups("Seconds").ToString)
                Dim Direction As String = myMatch.Groups("Direction").ToString
                If LeadingSign = "-" Or Direction.ToLower = "south" Or Direction.ToLower = "sth" Or Direction.ToLower = "s" Then
                    DecimalDegrees = -Degrees - Minutes / 60 - Seconds / 3600
                Else
                    DecimalDegrees = Degrees + Minutes / 60 + Seconds / 3600
                End If
            Else
                RaiseEvent ErrorMessage("The Latitude string could not be matched: " & StringVal & vbCrLf)
            End If
        End If
    End Sub

    'This is now a function
    'Public Sub SetAngle(StringVal As String)
    '    'Read an angle value from the StringVal.
    '    'Use the value to update the DecimalDegrees.

    '    Match = False

    '    'Attempt to match a Decimal angle value:
    '    Dim RegExPattern As String = "^\s*(?<LeadingSign>\+|\-|)\s*(?<DecimalDegrees>\d{1,3}\.\d*|\d{1,3})\s*(?<Direction>(?i)West|Wst|W|East|Est|E|South|Sth|S|North|Nth|N|)"
    '    'Named capture DecimalDegrees : 1 to 3 digits OR 1 to 3 digits & "." & 0 or more digits
    '    'Named capture Direction : N or Nth or North or S or Sth or South (case insensitive)
    '    Dim myRegEx As New System.Text.RegularExpressions.Regex(RegExPattern)
    '    Dim myMatch As System.Text.RegularExpressions.Match = myRegEx.Match(StringVal)
    '    If myMatch.Success Then 'Decimal Longitude value
    '        Match = True
    '        Dim LeadingSign As String = myMatch.Groups("LeadingSign").ToString
    '        Dim DecDegrees As Double = Val(myMatch.Groups("DecimalDegrees").ToString)
    '        Dim Direction As String = myMatch.Groups("Direction").ToString

    '        If Direction.ToLower = "west" Or Direction.ToLower = "wst" Or Direction.ToLower = "w" Then
    '            Type = AngleType.Longitude
    '            Sign = AngleSign.Negative
    '            DecimalDegrees = -DecDegrees

    '        ElseIf Direction.ToLower = "east" Or Direction.ToLower = "est" Or Direction.ToLower = "e" Then
    '            Type = AngleType.Longitude
    '            Sign = AngleSign.Positive
    '            DecimalDegrees = DecDegrees

    '        ElseIf Direction.ToLower = "south" Or Direction.ToLower = "sth" Or Direction.ToLower = "s" Then
    '            Type = AngleType.Latitude
    '            Sign = AngleSign.Negative
    '            DecimalDegrees = -DecDegrees

    '        ElseIf Direction.ToLower = "north" Or Direction.ToLower = "nth" Or Direction.ToLower = "n" Then
    '            Type = AngleType.Latitude
    '            Sign = AngleSign.Positive
    '            DecimalDegrees = DecDegrees

    '        Else
    '            Type = AngleType.General
    '            If LeadingSign = "-" Then
    '                Sign = AngleSign.Negative
    '                DecimalDegrees = -DecDegrees
    '            Else
    '                Sign = AngleSign.Positive
    '                DecimalDegrees = DecDegrees
    '            End If
    '        End If

    '    Else 'Attempt to match Deg-Min-Sec angle value:
    '        RegExPattern = "^\s*(?<LeadingSign>\+|\-|)\s*(?<Degrees>\d{1,3})\s+(?<Minutes>\d{1,2})\s+(?<Seconds>\d{1,2}|\d{1,2}\.\d*)\s*(?<Direction>(?i)West|Wst|W|East|Est|E|South|Sth|S|North|Nth|N)"
    '        Dim myRegEx2 As New System.Text.RegularExpressions.Regex(RegExPattern)
    '        Dim myMatch2 As System.Text.RegularExpressions.Match = myRegEx2.Match(StringVal)
    '        If myMatch2.Success Then 'Deg Min Sec Longitude value
    '            Match = True
    '            Type = AngleType.Longitude
    '            Dim LeadingSign As String = myMatch.Groups("LeadingSign").ToString
    '            Dim Degrees As Integer = Val(myMatch2.Groups("Degrees").ToString)
    '            Dim Minutes As Integer = Val(myMatch2.Groups("Minutes").ToString)
    '            Dim Seconds As Double = Val(myMatch2.Groups("Seconds").ToString)
    '            Dim Direction As String = myMatch2.Groups("Direction").ToString
    '            If Direction.ToLower = "west" Or Direction.ToLower = "wst" Or Direction.ToLower = "w" Then
    '                Type = AngleType.Longitude
    '                Sign = AngleSign.Negative
    '                DecimalDegrees = -Degrees - Minutes / 60 - Seconds / 3600
    '            ElseIf Direction.ToLower = "east" Or Direction.ToLower = "est" Or Direction.ToLower = "e" Then
    '                Type = AngleType.Longitude
    '                Sign = AngleSign.Positive
    '                DecimalDegrees = Degrees + Minutes / 60 + Seconds / 3600
    '            ElseIf Direction.ToLower = "south" Or Direction.ToLower = "sth" Or Direction.ToLower = "s" Then
    '                Type = AngleType.Latitude
    '                Sign = AngleSign.Negative
    '                DecimalDegrees = -Degrees - Minutes / 60 - Seconds / 3600
    '            ElseIf Direction.ToLower = "north" Or Direction.ToLower = "nth" Or Direction.ToLower = "n" Then
    '                Type = AngleType.Latitude
    '                Sign = AngleSign.Positive
    '                DecimalDegrees = Degrees + Minutes / 60 + Seconds / 3600
    '            Else
    '                Type = AngleType.General
    '                If LeadingSign = "-" Then
    '                    Sign = AngleSign.Negative
    '                    DecimalDegrees = -Degrees - Minutes / 60 - Seconds / 3600
    '                Else
    '                    Sign = AngleSign.Positive
    '                    DecimalDegrees = Degrees + Minutes / 60 + Seconds / 3600
    '                End If

    '            End If
    '        End If
    '    End If
    'End Sub

    Public Function SetAngle_OLD(StringVal As String) As Boolean
        'Read an angle value from the StringVal.
        'Use the value to update the DecimalDegrees.

        Match = False

        'Attempt to match a Decimal angle value:
        Dim RegExPattern As String = "^\s*(?<LeadingSign>\+|\-|)\s*(?<DecimalDegrees>\d{1,3}\.\d*|\d{1,3})\s*(?<Direction>(?i)West|Wst|W|East|Est|E|South|Sth|S|North|Nth|N|)"
        'Named capture DecimalDegrees : 1 to 3 digits OR 1 to 3 digits & "." & 0 or more digits
        'Named capture Direction : N or Nth or North or S or Sth or South (case insensitive)
        Dim myRegEx As New System.Text.RegularExpressions.Regex(RegExPattern)
        Dim myMatch As System.Text.RegularExpressions.Match = myRegEx.Match(StringVal)
        If myMatch.Success Then 'Decimal Longitude value
            Match = True
            Dim LeadingSign As String = myMatch.Groups("LeadingSign").ToString
            Dim DecDegrees As Double = Val(myMatch.Groups("DecimalDegrees").ToString)
            Dim Direction As String = myMatch.Groups("Direction").ToString

            If Direction.ToLower = "west" Or Direction.ToLower = "wst" Or Direction.ToLower = "w" Then
                Type = AngleType.Longitude
                Sign = AngleSign.Negative
                DecimalDegrees = -DecDegrees

            ElseIf Direction.ToLower = "east" Or Direction.ToLower = "est" Or Direction.ToLower = "e" Then
                Type = AngleType.Longitude
                Sign = AngleSign.Positive
                DecimalDegrees = DecDegrees

            ElseIf Direction.ToLower = "south" Or Direction.ToLower = "sth" Or Direction.ToLower = "s" Then
                Type = AngleType.Latitude
                Sign = AngleSign.Negative
                DecimalDegrees = -DecDegrees

            ElseIf Direction.ToLower = "north" Or Direction.ToLower = "nth" Or Direction.ToLower = "n" Then
                Type = AngleType.Latitude
                Sign = AngleSign.Positive
                DecimalDegrees = DecDegrees

            Else
                Type = AngleType.General
                If LeadingSign = "-" Then
                    Sign = AngleSign.Negative
                    DecimalDegrees = -DecDegrees
                Else
                    Sign = AngleSign.Positive
                    DecimalDegrees = DecDegrees
                End If
            End If

        Else 'Attempt to match Deg-Min-Sec angle value:
            'RegExPattern = "^\s*(?<LeadingSign>\+|\-|)\s*(?<Degrees>\d{1,3})\s+(?<Minutes>\d{1,2})\s+(?<Seconds>\d{1,2}|\d{1,2}\.\d*)\s*(?<Direction>(?i)West|Wst|W|East|Est|E|South|Sth|S|North|Nth|N)"
            'RegExPattern = "^\s*(?<LeadingSign>\+|\-|)\s*(?<Degrees>\d{1,3})\s*[\uB0]{0,1}\s+(?<Minutes>\d{1,2})\s*[\u2019]{0,1}\s+(?<Seconds>\d{1,2}|\d{1,2}\.\d*)\s*[\u201D]{0,1}\s*(?<Direction>(?i)West|Wst|W|East|Est|E|South|Sth|S|North|Nth|N)"
            RegExPattern = "^\s*(?<LeadingSign>\+|\-|)\s*(?<Degrees>\d{1,3})\s*[\u00B0]{0,1}\s+(?<Minutes>\d{1,2})\s*[\u2019]{0,1}\s+(?<Seconds>\d{1,2}|\d{1,2}\.\d*)\s*[\u201D]{0,1}\s*(?<Direction>(?i)West|Wst|W|East|Est|E|South|Sth|S|North|Nth|N)"
            Dim myRegEx2 As New System.Text.RegularExpressions.Regex(RegExPattern)
            Dim myMatch2 As System.Text.RegularExpressions.Match = myRegEx2.Match(StringVal)
            If myMatch2.Success Then 'Deg Min Sec Longitude value
                Match = True
                Type = AngleType.Longitude
                Dim LeadingSign As String = myMatch.Groups("LeadingSign").ToString
                Dim Degrees As Integer = Val(myMatch2.Groups("Degrees").ToString)
                Dim Minutes As Integer = Val(myMatch2.Groups("Minutes").ToString)
                Dim Seconds As Double = Val(myMatch2.Groups("Seconds").ToString)
                Dim Direction As String = myMatch2.Groups("Direction").ToString
                If Direction.ToLower = "west" Or Direction.ToLower = "wst" Or Direction.ToLower = "w" Then
                    Type = AngleType.Longitude
                    Sign = AngleSign.Negative
                    DecimalDegrees = -Degrees - Minutes / 60 - Seconds / 3600
                ElseIf Direction.ToLower = "east" Or Direction.ToLower = "est" Or Direction.ToLower = "e" Then
                    Type = AngleType.Longitude
                    Sign = AngleSign.Positive
                    DecimalDegrees = Degrees + Minutes / 60 + Seconds / 3600
                ElseIf Direction.ToLower = "south" Or Direction.ToLower = "sth" Or Direction.ToLower = "s" Then
                    Type = AngleType.Latitude
                    Sign = AngleSign.Negative
                    DecimalDegrees = -Degrees - Minutes / 60 - Seconds / 3600
                ElseIf Direction.ToLower = "north" Or Direction.ToLower = "nth" Or Direction.ToLower = "n" Then
                    Type = AngleType.Latitude
                    Sign = AngleSign.Positive
                    DecimalDegrees = Degrees + Minutes / 60 + Seconds / 3600
                Else
                    Type = AngleType.General
                    If LeadingSign = "-" Then
                        Sign = AngleSign.Negative
                        DecimalDegrees = -Degrees - Minutes / 60 - Seconds / 3600
                    Else
                        Sign = AngleSign.Positive
                        DecimalDegrees = Degrees + Minutes / 60 + Seconds / 3600
                    End If

                End If
            End If
        End If
        Return Match
    End Function

    Public Function SetAngle(StringVal As String) As Boolean
        'Read an angle value from the StringVal.
        'Use the value to update the DecimalDegrees.

        If IsNothing(StringVal) Then
            Return False
        Else
            Match = False

            'Attempt to match Deg-Min-Sec angle value:
            'Dim RegExPattern As String = "^\s*(?<LeadingSign>\+|\-|)\s*(?<Degrees>\d{1,3})\s*[\u00B0]{0,1}\s+(?<Minutes>\d{1,2})\s*[\u2019]{0,1}\s+(?<Seconds>\d{1,2}|\d{1,2}\.\d*)\s*[\u201D]{0,1}\s*(?<Direction>(?i)West|Wst|W|East|Est|E|South|Sth|S|North|Nth|N){0,1}"
            Dim RegExPattern As String = "^\s*(?<LeadingSign>\+|\-|)\s*(?<Degrees>\d{1,3})\s*[\u00B0\u2103\u2109\u00BA\u02DA]{0,1}\s+(?<Minutes>\d{1,2})\s*[\u2019\u0027\u2032]{0,1}\s+(?<Seconds>\d{1,2}\.\d*)\s*[\u201D\u0022\u2033\u02DD\u00A8]{0,1}\s*(?<Direction>(?i)West|Wst|W|East|Est|E|South|Sth|S|North|Nth|N){0,1}"
            Dim myRegEx As New System.Text.RegularExpressions.Regex(RegExPattern)
            Dim myMatch As System.Text.RegularExpressions.Match = myRegEx.Match(StringVal)
            If myMatch.Success Then 'Deg Min Sec Longitude value
                Match = True
                Type = AngleType.Longitude
                Dim LeadingSign As String = myMatch.Groups("LeadingSign").ToString
                Dim Degrees As Integer = Val(myMatch.Groups("Degrees").ToString)
                Dim Minutes As Integer = Val(myMatch.Groups("Minutes").ToString)
                Dim Seconds As Double = Val(myMatch.Groups("Seconds").ToString)
                Dim Direction As String = myMatch.Groups("Direction").ToString
                If Direction.ToLower = "west" Or Direction.ToLower = "wst" Or Direction.ToLower = "w" Then
                    Type = AngleType.Longitude
                    Sign = AngleSign.Negative
                    DecimalDegrees = -Degrees - Minutes / 60 - Seconds / 3600
                ElseIf Direction.ToLower = "east" Or Direction.ToLower = "est" Or Direction.ToLower = "e" Then
                    Type = AngleType.Longitude
                    Sign = AngleSign.Positive
                    DecimalDegrees = Degrees + Minutes / 60 + Seconds / 3600
                ElseIf Direction.ToLower = "south" Or Direction.ToLower = "sth" Or Direction.ToLower = "s" Then
                    Type = AngleType.Latitude
                    Sign = AngleSign.Negative
                    DecimalDegrees = -Degrees - Minutes / 60 - Seconds / 3600
                ElseIf Direction.ToLower = "north" Or Direction.ToLower = "nth" Or Direction.ToLower = "n" Then
                    Type = AngleType.Latitude
                    Sign = AngleSign.Positive
                    DecimalDegrees = Degrees + Minutes / 60 + Seconds / 3600
                Else
                    Type = AngleType.General
                    If LeadingSign = "-" Then
                        Sign = AngleSign.Negative
                        DecimalDegrees = -Degrees - Minutes / 60 - Seconds / 3600
                    Else
                        Sign = AngleSign.Positive
                        DecimalDegrees = Degrees + Minutes / 60 + Seconds / 3600
                    End If

                End If
            Else
                'Attempt to match a Decimal angle value:
                RegExPattern = "^\s*(?<LeadingSign>\+|\-|)\s*(?<DecimalDegrees>\d{1,3}\.\d*|\d{1,3})\s*(?<Direction>(?i)West|Wst|W|East|Est|E|South|Sth|S|North|Nth|N|)"
                'Named capture DecimalDegrees : 1 to 3 digits OR 1 to 3 digits & "." & 0 or more digits
                'Named capture Direction : N or Nth or North or S or Sth or South (case insensitive)
                Dim myRegEx2 As New System.Text.RegularExpressions.Regex(RegExPattern)
                Dim myMatch2 As System.Text.RegularExpressions.Match = myRegEx2.Match(StringVal)
                If myMatch2.Success Then 'Decimal Longitude value
                    Match = True
                    Dim LeadingSign As String = myMatch2.Groups("LeadingSign").ToString
                    Dim DecDegrees As Double = Val(myMatch2.Groups("DecimalDegrees").ToString)
                    Dim Direction As String = myMatch2.Groups("Direction").ToString

                    If Direction.ToLower = "west" Or Direction.ToLower = "wst" Or Direction.ToLower = "w" Then
                        Type = AngleType.Longitude
                        Sign = AngleSign.Negative
                        DecimalDegrees = -DecDegrees

                    ElseIf Direction.ToLower = "east" Or Direction.ToLower = "est" Or Direction.ToLower = "e" Then
                        Type = AngleType.Longitude
                        Sign = AngleSign.Positive
                        DecimalDegrees = DecDegrees

                    ElseIf Direction.ToLower = "south" Or Direction.ToLower = "sth" Or Direction.ToLower = "s" Then
                        Type = AngleType.Latitude
                        Sign = AngleSign.Negative
                        DecimalDegrees = -DecDegrees

                    ElseIf Direction.ToLower = "north" Or Direction.ToLower = "nth" Or Direction.ToLower = "n" Then
                        Type = AngleType.Latitude
                        Sign = AngleSign.Positive
                        DecimalDegrees = DecDegrees

                    Else
                        Type = AngleType.General
                        If LeadingSign = "-" Then
                            Sign = AngleSign.Negative
                            DecimalDegrees = -DecDegrees
                        Else
                            Sign = AngleSign.Positive
                            DecimalDegrees = DecDegrees
                        End If
                    End If
                End If
            End If
            Return Match
        End If
    End Function

#End Region 'Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events - Events that can be triggered by this class." '==========================================================================================================================
    Event ErrorMessage(ByVal Msg As String) 'Send an error message.
    Event Message(ByVal Msg As String) 'Send a normal message.
#End Region 'Events -----------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class 'clsAngle


